const MARKER_PREFIX = "DAYBOOK_EXPORT_ID:";
let operationCancelled = false;

chrome.runtime.onMessage.addListener((message, _sender, sendResponse) => {
  handleMessage(message).then(
    (result) => sendResponse(result),
    (error) => sendResponse({ __daybookError: error instanceof Error ? error.message : String(error) })
  );
  return true;
});

async function handleMessage(message) {
  switch (message.action) {
    case "DISCOVER_HISTORY":
      return { conversations: await discoverConversations(message.date) };
    case "SCRAPE_CONVERSATION":
      return scrapeConversation(message);
    case "EXPORT_DAY":
      operationCancelled = false;
      return exportDay(message);
    case "CANCEL_OPERATION":
      operationCancelled = true;
      return { cancelled: true };
    default:
      throw new Error("Unknown Copilot page request.");
  }
}

async function discoverConversations(selectedDate) {
  const conversations = new Map();
  const sidebar = findHistoryScroller();
  const originalTop = sidebar?.scrollTop || 0;

  for (let pass = 0; pass < 12; pass += 1) {
    collectConversationLinks(conversations, selectedDate);
    if (!sidebar || sidebar.scrollTop + sidebar.clientHeight >= sidebar.scrollHeight - 8) break;
    sidebar.scrollTop += Math.max(sidebar.clientHeight * 0.8, 360);
    await delay(350);
  }

  if (sidebar) sidebar.scrollTop = originalTop;
  return [...conversations.values()].slice(0, 80);
}

function collectConversationLinks(conversations, selectedDate) {
  const anchors = document.querySelectorAll('a[href*="/chat/"], a[href*="/chats/"], a[href*="conversation"]');
  for (const anchor of anchors) {
    const url = new URL(anchor.href, location.href);
    if (url.origin !== location.origin || url.href === location.href) continue;
    const title = cleanText(anchor.getAttribute("aria-label") || anchor.textContent || "Copilot chat");
    if (!title || /new chat|copilot chat$/i.test(title)) continue;
    conversations.set(url.href, {
      url: url.href,
      title,
      dateHint: inferDateHint(anchor, selectedDate)
    });
  }
}

function findHistoryScroller() {
  const candidates = [...document.querySelectorAll('nav, [role="navigation"], [data-testid*="history"], [class*="history"]')];
  return candidates.find((element) => element.scrollHeight > element.clientHeight + 40);
}

function inferDateHint(anchor, selectedDate) {
  let current = anchor.parentElement;
  const chunks = [];
  for (let depth = 0; current && depth < 4; depth += 1, current = current.parentElement) {
    chunks.push(cleanText(current.textContent || ""));
  }
  const text = chunks.join(" ");
  const selected = new Date(`${selectedDate}T12:00:00`);
  const today = new Date();
  if (/\btoday\b/i.test(text) && sameDay(selected, today)) return selectedDate;
  const yesterday = new Date(today); yesterday.setDate(today.getDate() - 1);
  if (/\byesterday\b/i.test(text) && sameDay(selected, yesterday)) return selectedDate;
  const explicit = text.match(/\b(?:Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|Sep(?:tember)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)\s+\d{1,2}(?:,\s+\d{4})?/i);
  if (explicit) {
    const parsed = new Date(explicit[0].includes(",") ? explicit[0] : `${explicit[0]}, ${selected.getFullYear()}`);
    if (!Number.isNaN(parsed.getTime())) return localDate(parsed);
  }
  return undefined;
}

function scrapeConversation({ date, dateHint, titleHint }) {
  const elements = messageElements();
  const pageText = elements.map((element) => cleanText(element.textContent || "")).join("\n");
  const exportId = markerFromText(pageText);
  const conversationId = conversationIdFromUrl(location.href);
  if (exportId) return { generated: true, exportId, conversationId, records: [] };

  const title = conversationTitle(titleHint);
  const records = [];
  elements.forEach((element, index) => {
    const text = cleanMessageText(element.textContent || "");
    if (!text || text.length < 2) return;
    const timestamp = timestampFromElement(element);
    const recordDate = timestamp ? localDate(new Date(timestamp)) : dateHint;
    if (recordDate !== date) return;
    const role = roleFromElement(element, index);
    const links = [...element.querySelectorAll('a[href^="http"]')]
      .map((anchor) => ({ title: cleanText(anchor.textContent || anchor.href), url: anchor.href }))
      .filter((link, linkIndex, allLinks) => link.url !== location.href && allLinks.findIndex((candidate) => candidate.url === link.url) === linkIndex)
      .slice(0, 20);
    records.push({
      id: stableId(`${conversationId}|${role}|${timestamp || recordDate}|${text}`),
      conversationId,
      conversationTitle: title,
      role,
      text,
      date: recordDate,
      timestamp,
      sourceUrl: location.href,
      links,
      generated: false
    });
  });
  return { generated: false, conversationId, records };
}

async function exportDay({ filename, markdown, exportId }) {
  const fileInput = await findUploadInput();
  const file = new File([markdown], filename, { type: "text/markdown" });
  const transfer = new DataTransfer();
  transfer.items.add(file);
  fileInput.files = transfer.files;
  fileInput.dispatchEvent(new Event("change", { bubbles: true }));
  await waitFor(() => document.body.textContent?.includes(filename), 30_000, "Copilot did not accept the Daybook file upload.");

  const prompt = [
    `${MARKER_PREFIX}${exportId}`,
    `Use the attached ${filename} as the only source of truth.`,
    "Create a concise daily brief with these headings: Overview, Important highlights, Decisions, Action items, Blockers, and Open questions.",
    "Include only decisions and action items explicitly supported by the file. Do not invent owners, deadlines, commitments, or facts."
  ].join("\n");
  const composer = await waitForComposer();
  fillComposer(composer, prompt);
  await delay(300);
  submitComposer(composer);

  const markdownResponse = await waitForGeneratedResponse(exportId);
  return {
    markdown: markdownResponse,
    conversationId: conversationIdFromUrl(location.href),
    conversationUrl: location.href
  };
}

async function findUploadInput() {
  let input = preferredFileInput();
  if (input) return input;
  const addButton = findControl(/add content|attach|upload|add a file/i);
  if (!addButton) throw new Error("Copilot's file upload control was not found. Download the Markdown file and attach it manually.");
  addButton.click();
  await delay(350);
  input = preferredFileInput();
  if (!input) {
    const uploadButton = findControl(/upload (?:images and )?files|browse my computer|upload from (?:this )?device/i);
    uploadButton?.click();
  }
  input = await waitFor(preferredFileInput, 8_000, "Copilot's file picker did not open.");
  return input;
}

function preferredFileInput() {
  const inputs = [...document.querySelectorAll('input[type="file"]')];
  return inputs.find((candidate) => /(?:markdown|text|\.md|\.txt|\*)/i.test(candidate.accept || "*")) || inputs[0];
}

function waitForComposer() {
  return waitFor(() => {
    const selectors = [
      'textarea[placeholder*="message" i]',
      'textarea[aria-label*="message" i]',
      '[contenteditable="true"][role="textbox"]',
      '[contenteditable="true"][data-testid*="input"]'
    ];
    return selectors.map((selector) => document.querySelector(selector)).find(Boolean);
  }, 15_000, "Copilot's message box was not found.");
}

function fillComposer(composer, value) {
  composer.focus();
  if (composer instanceof HTMLTextAreaElement || composer instanceof HTMLInputElement) {
    const setter = Object.getOwnPropertyDescriptor(Object.getPrototypeOf(composer), "value")?.set;
    setter?.call(composer, value);
  } else {
    composer.textContent = value;
  }
  composer.dispatchEvent(new InputEvent("input", { bubbles: true, inputType: "insertText", data: value }));
  composer.dispatchEvent(new Event("change", { bubbles: true }));
}

function submitComposer(composer) {
  const send = findControl(/send|submit/i, ['button[type="submit"]', '[data-testid*="send"]', 'button[aria-label*="send" i]']);
  if (send && !send.disabled) send.click();
  else throw new Error("Copilot's Send button was not found. The file and prompt remain ready for manual submission.");
}

async function waitForGeneratedResponse(exportId) {
  let stableText = "";
  let stableCount = 0;
  const deadline = Date.now() + 210_000;
  while (Date.now() < deadline) {
    if (operationCancelled) throw new Error("Copilot operation cancelled.");
    const elements = messageElements();
    const markerIndex = elements.findIndex((element) => (element.textContent || "").includes(`${MARKER_PREFIX}${exportId}`));
    if (markerIndex >= 0) {
      const later = elements.slice(markerIndex + 1).filter((element, index) => roleFromElement(element, markerIndex + 1 + index) === "assistant");
      const text = cleanMessageText(later.at(-1)?.textContent || "");
      if (text.length > 40 && text === stableText) stableCount += 1;
      else { stableText = text; stableCount = 0; }
      if (stableCount >= 3) return stableText;
    }
    await delay(1_500);
  }
  throw new Error("Copilot did not finish a matching response. The generated chat remains open so you can copy it manually.");
}

function messageElements() {
  const selectors = [
    '[data-testid*="chat-turn"]',
    '[data-testid*="message"]',
    '[data-content*="message"]',
    '[role="article"]',
    'main article'
  ];
  for (const selector of selectors) {
    const elements = [...document.querySelectorAll(selector)].filter((element) => {
      const text = cleanText(element.textContent || "");
      return text.length > 1 && !elementsNestedInside(element, selector);
    });
    if (elements.length >= 2) return elements;
  }
  return [];
}

function elementsNestedInside(element, selector) {
  let parent = element.parentElement;
  while (parent) {
    if (parent.matches?.(selector)) return true;
    parent = parent.parentElement;
  }
  return false;
}

function roleFromElement(element, index) {
  const signature = [element.getAttribute("data-testid"), element.getAttribute("data-content"), element.getAttribute("aria-label"), element.className].join(" ").toLowerCase();
  if (/user|human|you said|prompt/.test(signature)) return "user";
  if (/assistant|copilot|bot|response/.test(signature)) return "assistant";
  const text = cleanText(element.textContent || "");
  if (/^(you said|you)\b/i.test(text)) return "user";
  if (/^(copilot said|copilot)\b/i.test(text)) return "assistant";
  return index % 2 === 0 ? "user" : "assistant";
}

function timestampFromElement(element) {
  const time = element.querySelector("time[datetime]") || element.closest("time[datetime]");
  const candidate = time?.getAttribute("datetime") || element.getAttribute("data-timestamp") || element.querySelector("[data-timestamp]")?.getAttribute("data-timestamp");
  if (!candidate) return undefined;
  const parsed = new Date(candidate);
  return Number.isNaN(parsed.getTime()) ? undefined : parsed.toISOString();
}

function conversationTitle(fallback) {
  const heading = [...document.querySelectorAll("h1, h2")].map((element) => cleanText(element.textContent || "")).find((text) => text && !/copilot|chat/i.test(text));
  return heading || fallback || "Copilot chat";
}

function findControl(pattern, selectors = ['button', '[role="button"]']) {
  return selectors.flatMap((selector) => [...document.querySelectorAll(selector)]).find((element) => {
    const label = cleanText(`${element.getAttribute("aria-label") || ""} ${element.getAttribute("title") || ""} ${element.textContent || ""}`);
    return pattern.test(label);
  });
}

function markerFromText(text) {
  return text.match(/DAYBOOK_EXPORT_ID:([0-9a-f-]{20,})/i)?.[1];
}

function conversationIdFromUrl(value) {
  const url = new URL(value);
  const parts = url.pathname.split("/").filter(Boolean);
  return parts.at(-1) || stableId(url.href);
}

function cleanMessageText(value) {
  return cleanText(value)
    .replace(/^(you said|copilot said|you|copilot)\s*/i, "")
    .replace(/\b(copy|thumbs up|thumbs down|more options)\b/gi, "")
    .trim();
}

function cleanText(value) {
  return value.replace(/\s+/g, " ").trim();
}

function stableId(value) {
  let hash = 2166136261;
  for (let index = 0; index < value.length; index += 1) {
    hash ^= value.charCodeAt(index);
    hash = Math.imul(hash, 16777619);
  }
  return (hash >>> 0).toString(16).padStart(8, "0");
}

function localDate(value) {
  const year = value.getFullYear();
  const month = String(value.getMonth() + 1).padStart(2, "0");
  const day = String(value.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function sameDay(a, b) {
  return a.getFullYear() === b.getFullYear() && a.getMonth() === b.getMonth() && a.getDate() === b.getDate();
}

function waitFor(read, timeout, message) {
  return new Promise((resolve, reject) => {
    const started = Date.now();
    const poll = () => {
      const value = read();
      if (value) return resolve(value);
      if (Date.now() - started >= timeout) return reject(new Error(message));
      setTimeout(poll, 250);
    };
    poll();
  });
}

function delay(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}
