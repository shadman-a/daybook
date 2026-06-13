import dayjs from "dayjs";
import { DaybookData, DaybookItem, DaybookPerson } from "../types/daybook";
import { CopilotRecord } from "./types";

export function mergeCopilotRecords(data: DaybookData, records: CopilotRecord[]): DaybookData {
  const items = records
    .filter((record) => !record.generated && record.date === data.date)
    .map(recordToDaybookItem);
  const merged = [...data.items.filter((item) => item.source !== "Copilot"), ...items]
    .sort(compareItems);

  return {
    ...data,
    items: merged,
    copilotCount: items.length
  };
}

export function buildDaybookMarkdown(date: string, items: DaybookItem[], exportId: string): string {
  const sourceItems = items.filter((item) => !item.generated);
  const lines = [
    `<!-- DAYBOOK_EXPORT_ID:${exportId} -->`,
    `# Daybook for ${dayjs(date).format("dddd, MMMM D, YYYY")}`,
    "",
    "This document contains the user's Microsoft 365 activity for one day.",
    "Only report decisions, action items, owners, deadlines, blockers, and open questions explicitly supported by the activity below.",
    "Do not invent missing details.",
    ""
  ];

  for (const item of sourceItems) {
    const time = item.timestamp ? dayjs(item.timestamp).format("h:mm A") : "Time unknown";
    const people = item.people.map(personLabel).filter(Boolean).join(", ");
    lines.push(`## ${time} · ${item.source} · ${clean(item.title)}`);
    if (people) lines.push(`People: ${clean(people)}`);
    if (item.conversationTitle) lines.push(`Conversation: ${clean(item.conversationTitle)}`);
    if (item.preview) lines.push("", clean(item.preview));
    for (const link of copilotLinks(item.raw)) lines.push(`- ${clean(link.title || "Reference")}: ${link.url}`);
    if (item.sourceUrl) lines.push("", `Source: ${item.sourceUrl}`);
    lines.push("");
  }

  return lines.join("\n");
}

function copilotLinks(raw: unknown): Array<{ title: string; url: string }> {
  if (!raw || typeof raw !== "object" || !("links" in raw) || !Array.isArray(raw.links)) return [];
  return raw.links.filter((link): link is { title: string; url: string } => (
    Boolean(link) && typeof link === "object" && typeof link.title === "string" && typeof link.url === "string"
  ));
}

export function fingerprintDaybook(items: DaybookItem[]): string {
  const value = items
    .filter((item) => !item.generated)
    .map((item) => [item.id, item.timestamp, item.title, item.preview].join("|"))
    .sort()
    .join("\n");
  let hash = 2166136261;
  for (let index = 0; index < value.length; index += 1) {
    hash ^= value.charCodeAt(index);
    hash = Math.imul(hash, 16777619);
  }
  return (hash >>> 0).toString(16).padStart(8, "0");
}

export function downloadDaybookMarkdown(filename: string, markdown: string) {
  const url = URL.createObjectURL(new Blob([markdown], { type: "text/markdown;charset=utf-8" }));
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = filename;
  anchor.click();
  URL.revokeObjectURL(url);
}

function recordToDaybookItem(record: CopilotRecord): DaybookItem {
  const author: DaybookPerson = { displayName: record.role === "user" ? "You" : "Microsoft 365 Copilot" };
  return {
    id: `copilot-${record.id}`,
    source: "Copilot",
    timestamp: record.timestamp,
    date: record.date,
    title: record.conversationTitle || (record.role === "user" ? "Copilot prompt" : "Copilot response"),
    preview: record.text,
    people: [author],
    conversationId: record.conversationId,
    conversationTitle: record.conversationTitle,
    sourceUrl: record.sourceUrl,
    generated: false,
    raw: record
  };
}

function compareItems(a: DaybookItem, b: DaybookItem): number {
  if (!a.timestamp && !b.timestamp) return a.title.localeCompare(b.title);
  if (!a.timestamp) return 1;
  if (!b.timestamp) return -1;
  return new Date(a.timestamp).getTime() - new Date(b.timestamp).getTime();
}

function personLabel(person: DaybookPerson): string {
  return person.displayName || person.email || person.id || "";
}

function clean(value: string): string {
  return value.replace(/\s+/g, " ").trim();
}
