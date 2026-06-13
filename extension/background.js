const REQUEST_TYPE = "DAYBOOK_COPILOT_REQUEST";
const EVENT_TYPE = "DAYBOOK_COPILOT_EVENT";
const COPILOT_URL = "https://m365.cloud.microsoft/chat";
const RETENTION_MS = 90 * 24 * 60 * 60 * 1000;
const cancelledRequests = new Set();
const activeRequests = new Set();

chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  if (!message || message.type !== REQUEST_TYPE) return false;
  handleRequest(message, sender).then(
    (data) => sendResponse({ ok: true, data }),
    (error) => sendResponse({ ok: false, error: error instanceof Error ? error.message : String(error) })
  );
  return true;
});

async function handleRequest(message, sender) {
  const appTabId = sender.tab?.id;
  switch (message.action) {
    case "status":
      return { installed: true, version: chrome.runtime.getManifest().version };
    case "getDay":
      await pruneStorage();
      return readDay(message.payload.date);
    case "syncDay":
      return runOperation(message.requestId, () => syncDay(message.payload.date, message.requestId, appTabId));
    case "exportDay":
      return runOperation(message.requestId, () => exportDay(message.payload, message.requestId, appTabId));
    case "cancel":
      for (const requestId of activeRequests) cancelledRequests.add(requestId);
      await cancelCopilotPages();
      return undefined;
    case "clear":
      await chrome.storage.local.clear();
      return undefined;
    default:
      throw new Error("Unknown Daybook extension request.");
  }
}

async function runOperation(requestId, operation) {
  activeRequests.add(requestId);
  try {
    return await operation();
  } finally {
    activeRequests.delete(requestId);
    cancelledRequests.delete(requestId);
  }
}

async function syncDay(date, requestId, appTabId) {
  emit(appTabId, requestId, "opening", "Opening Microsoft 365 Copilot");
  const tab = await ensureCopilotTab();
  assertNotCancelled(requestId);

  emit(appTabId, requestId, "discovering", "Reading Copilot conversation history");
  const discovery = await sendCopilotMessage(tab.id, { action: "DISCOVER_HISTORY", date });
  const conversations = discovery.conversations || [];
  if (!conversations.length) {
    throw new Error("No Copilot conversations were visible. Confirm that Copilot is signed in and its history sidebar is available.");
  }

  const storage = await chrome.storage.local.get(["records", "generatedConversations"]);
  const records = Array.isArray(storage.records) ? storage.records : [];
  const generated = new Set(Array.isArray(storage.generatedConversations) ? storage.generatedConversations : []);
  const collected = [];

  for (let index = 0; index < conversations.length; index += 1) {
    assertNotCancelled(requestId);
    const conversation = conversations[index];
    emit(appTabId, requestId, "syncing", `Reading ${conversation.title || "Copilot chat"}`, index + 1, conversations.length);
    if (tab.url !== conversation.url) {
      await chrome.tabs.update(tab.id, { url: conversation.url });
      await waitForTab(tab.id);
      tab.url = conversation.url;
    }

    const result = await sendCopilotMessage(tab.id, {
      action: "SCRAPE_CONVERSATION",
      date,
      dateHint: conversation.dateHint,
      titleHint: conversation.title
    });

    if (result.generated) {
      generated.add(result.conversationId);
      continue;
    }
    collected.push(...(result.records || []).filter((record) => record.date === date && !record.generated));
  }

  const merged = dedupeRecords([...records, ...collected]).filter((record) => !generated.has(record.conversationId));
  await chrome.storage.local.set({
    records: pruneRecords(merged),
    generatedConversations: [...generated],
    [`lastSynced:${date}`]: new Date().toISOString()
  });
  emit(appTabId, requestId, "complete", "Copilot history synchronized", conversations.length, conversations.length);
  return readDay(date);
}

async function exportDay(payload, requestId, appTabId) {
  assertNotCancelled(requestId);
  const existing = await readDay(payload.date);
  if (existing.brief?.fingerprint === payload.fingerprint) return existing.brief;

  emit(appTabId, requestId, "opening", "Opening a new Microsoft 365 Copilot chat");
  const tab = await ensureCopilotTab(true);
  assertNotCancelled(requestId);
  emit(appTabId, requestId, "uploading", `Uploading ${payload.filename}`);

  const result = await sendCopilotMessage(tab.id, {
    action: "EXPORT_DAY",
    filename: payload.filename,
    markdown: payload.markdown,
    exportId: payload.exportId
  });
  assertNotCancelled(requestId);
  emit(appTabId, requestId, "importing", "Saving the generated brief");

  const brief = {
    id: `brief-${payload.exportId}`,
    date: payload.date,
    exportId: payload.exportId,
    fingerprint: payload.fingerprint,
    markdown: result.markdown,
    conversationId: result.conversationId,
    conversationUrl: result.conversationUrl,
    createdAt: new Date().toISOString()
  };
  const storage = await chrome.storage.local.get(["briefs", "generatedConversations"]);
  const briefs = { ...(storage.briefs || {}), [payload.date]: brief };
  const generated = new Set(Array.isArray(storage.generatedConversations) ? storage.generatedConversations : []);
  if (result.conversationId) generated.add(result.conversationId);
  await chrome.storage.local.set({ briefs, generatedConversations: [...generated] });
  await pruneStorage();
  emit(appTabId, requestId, "complete", "Copilot brief imported");
  return brief;
}

async function readDay(date) {
  const storage = await chrome.storage.local.get(["records", "briefs", `lastSynced:${date}`]);
  return {
    records: (storage.records || []).filter((record) => record.date === date && !record.generated),
    brief: storage.briefs?.[date],
    lastSyncedAt: storage[`lastSynced:${date}`]
  };
}

async function ensureCopilotTab(fresh = false) {
  const tabs = await chrome.tabs.query({ url: ["https://m365.cloud.microsoft/*", "https://m365copilot.com/*"] });
  let tab = tabs[0];
  if (!tab?.id) tab = await chrome.tabs.create({ url: COPILOT_URL, active: true });
  else await chrome.tabs.update(tab.id, { active: true, ...(fresh ? { url: COPILOT_URL } : {}) });
  await waitForTab(tab.id);
  return tab;
}

async function sendCopilotMessage(tabId, message) {
  let lastError;
  for (let attempt = 0; attempt < 6; attempt += 1) {
    try {
      const response = await chrome.tabs.sendMessage(tabId, message);
      if (response?.__daybookError) {
        const error = new Error(response.__daybookError);
        error.daybookNonRetryable = true;
        throw error;
      }
      return response;
    } catch (error) {
      if (error?.daybookNonRetryable) throw error;
      lastError = error;
      await delay(750);
    }
  }
  throw new Error(lastError?.message || "The Copilot page did not become ready.");
}

async function cancelCopilotPages() {
  const tabs = await chrome.tabs.query({ url: ["https://m365.cloud.microsoft/*", "https://m365copilot.com/*"] });
  await Promise.all(tabs.filter((tab) => tab.id).map((tab) => (
    chrome.tabs.sendMessage(tab.id, { action: "CANCEL_OPERATION" }).catch(() => undefined)
  )));
}

function waitForTab(tabId) {
  return new Promise((resolve, reject) => {
    const timeout = setTimeout(() => {
      chrome.tabs.onUpdated.removeListener(listener);
      reject(new Error("Microsoft 365 Copilot took too long to load."));
    }, 45_000);
    const listener = (updatedTabId, changeInfo) => {
      if (updatedTabId !== tabId || changeInfo.status !== "complete") return;
      clearTimeout(timeout);
      chrome.tabs.onUpdated.removeListener(listener);
      setTimeout(resolve, 1_500);
    };
    chrome.tabs.onUpdated.addListener(listener);
    chrome.tabs.get(tabId).then((tab) => {
      if (tab.status === "complete") {
        clearTimeout(timeout);
        chrome.tabs.onUpdated.removeListener(listener);
        setTimeout(resolve, 1_500);
      }
    }).catch(reject);
  });
}

function emit(tabId, requestId, stage, message, current, total) {
  if (!tabId) return;
  chrome.tabs.sendMessage(tabId, {
    type: EVENT_TYPE,
    requestId,
    progress: { stage, message, current, total }
  }).catch(() => undefined);
}

function assertNotCancelled(requestId) {
  if (cancelledRequests.has(requestId)) throw new Error("Copilot operation cancelled.");
}

async function pruneStorage() {
  const storage = await chrome.storage.local.get(["records", "briefs"]);
  const cutoff = Date.now() - RETENTION_MS;
  const records = pruneRecords(storage.records || []);
  const briefs = Object.fromEntries(Object.entries(storage.briefs || {}).filter(([, brief]) => Date.parse(brief.createdAt) >= cutoff));
  await chrome.storage.local.set({ records, briefs });
}

function pruneRecords(records) {
  const cutoffDate = new Date(Date.now() - RETENTION_MS).toISOString().slice(0, 10);
  return records.filter((record) => record.date >= cutoffDate);
}

function dedupeRecords(records) {
  const byId = new Map();
  for (const record of records) byId.set(record.id, record);
  return [...byId.values()];
}

function delay(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}
