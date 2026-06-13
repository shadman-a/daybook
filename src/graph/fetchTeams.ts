import { Client } from "@microsoft/microsoft-graph-client";
import { GraphChat, GraphChatMember, GraphChatMessage, GraphCollection } from "../types/graph";

export type TeamsMessageWithChat = {
  chat: GraphChat;
  message: GraphChatMessage;
  members: GraphChatMember[];
};

export type TeamsFetchResult = {
  messages: TeamsMessageWithChat[];
  activeChatCount: number;
  failedChatCount: number;
  failureReason?: string;
};

const CHAT_BATCH_SIZE = 8;
const memberCache = new Map<string, GraphChatMember[]>();

export async function fetchTeamsMessagesForDay(
  graph: Client,
  startZ: string,
  endZ: string
): Promise<TeamsFetchResult> {
  const activeChats = await fetchActiveChats(graph, startZ);
  const messages: TeamsMessageWithChat[] = [];
  let failedChatCount = 0;
  let failureReason: string | undefined;

  for (let index = 0; index < activeChats.length; index += CHAT_BATCH_SIZE) {
    const batch = activeChats.slice(index, index + CHAT_BATCH_SIZE);
    const settled = await Promise.allSettled(
      batch.map((chat) => fetchChatActivityForDay(graph, chat, startZ, endZ))
    );

    settled.forEach((result) => {
      if (result.status === "fulfilled") messages.push(...result.value);
      else {
        failedChatCount += 1;
        failureReason ??= getGraphFailureReason(result.reason);
      }
    });
  }

  return {
    messages,
    activeChatCount: activeChats.length,
    failedChatCount,
    failureReason
  };
}

async function fetchChatActivityForDay(
  graph: Client,
  chat: GraphChat,
  startZ: string,
  endZ: string
): Promise<TeamsMessageWithChat[]> {
  const [messages, members] = await Promise.all([
    fetchChatMessagesForDay(graph, chat, startZ, endZ),
    fetchChatMembers(graph, chat.id).catch(() => [])
  ]);

  return messages.map((message) => ({ chat, message, members }));
}

async function fetchActiveChats(graph: Client, startZ: string): Promise<GraphChat[]> {
  const activeChats: GraphChat[] = [];
  let response = await graph.api("/me/chats")
    .select("id,topic,chatType,lastUpdatedDateTime,webUrl")
    .expand("lastMessagePreview")
    .orderby("lastMessagePreview/createdDateTime desc")
    .top(50)
    .get() as GraphCollection<GraphChat>;

  while (true) {
    const pageChats = response.value ?? [];
    activeChats.push(...pageChats.filter((chat) => getChatActivityTime(chat) >= startZ));

    // Chats are sorted newest-first. Once a page reaches activity before the
    // selected day, every later page is irrelevant to that day's timeline.
    const crossedStartOfDay = pageChats.some((chat) => {
      const activityTime = getChatActivityTime(chat);
      return activityTime !== "" && activityTime < startZ;
    });
    const nextLink = response["@odata.nextLink"];

    if (crossedStartOfDay || !nextLink) return activeChats;
    response = await graph.api(nextLink).get() as GraphCollection<GraphChat>;
  }
}

async function fetchChatMessagesForDay(
  graph: Client,
  chat: GraphChat,
  startZ: string,
  endZ: string
): Promise<GraphChatMessage[]> {
  const dayMessages: GraphChatMessage[] = [];
  let response = await fetchFirstMessagePage(graph, chat.id, endZ);

  while (true) {
    const pageMessages = response.value ?? [];
    dayMessages.push(...pageMessages
      .filter((message) => {
        const created = message.createdDateTime;
        return created !== undefined && created >= startZ && created < endZ;
      })
    );

    const crossedStartOfDay = pageMessages.some((message) => (
      message.createdDateTime !== undefined && message.createdDateTime < startZ
    ));
    const nextLink = response["@odata.nextLink"];

    if (crossedStartOfDay || !nextLink) return dayMessages;
    response = await graph.api(nextLink).get() as GraphCollection<GraphChatMessage>;
  }
}

async function fetchChatMembers(graph: Client, chatId: string): Promise<GraphChatMember[]> {
  const cached = memberCache.get(chatId);
  if (cached) return cached;

  const encodedChatId = encodeURIComponent(chatId);
  let response: GraphCollection<GraphChatMember>;

  try {
    response = await graph.api(`/me/chats/${encodedChatId}/members`).get() as GraphCollection<GraphChatMember>;
  } catch {
    response = await graph.api(`/chats/${encodedChatId}/members`).get() as GraphCollection<GraphChatMember>;
  }

  const members = [...(response.value ?? [])];
  while (response["@odata.nextLink"]) {
    response = await graph.api(response["@odata.nextLink"]).get() as GraphCollection<GraphChatMember>;
    members.push(...(response.value ?? []));
  }

  memberCache.set(chatId, members);
  return members;
}

async function fetchFirstMessagePage(
  graph: Client,
  chatId: string,
  endZ: string
): Promise<GraphCollection<GraphChatMessage>> {
  const encodedChatId = encodeURIComponent(chatId);

  try {
    return await fetchDateBoundedMessagePage(graph, `/me/chats/${encodedChatId}/messages`, endZ);
  } catch {
    // Some Graph deployments expose the collection only through the canonical
    // chat route even though individual messages support the /me route.
    return fetchDateBoundedMessagePage(graph, `/chats/${encodedChatId}/messages`, endZ);
  }
}

async function fetchDateBoundedMessagePage(
  graph: Client,
  path: string,
  endZ: string
): Promise<GraphCollection<GraphChatMessage>> {
  return graph.api(path)
    .orderby("createdDateTime desc")
    .filter(`createdDateTime lt ${endZ}`)
    .top(50)
    .get() as Promise<GraphCollection<GraphChatMessage>>;
}

function getChatActivityTime(chat: GraphChat): string {
  return chat.lastMessagePreview?.createdDateTime ?? chat.lastUpdatedDateTime ?? "";
}

function getGraphFailureReason(error: unknown): string | undefined {
  if (!error || typeof error !== "object") return undefined;
  const candidate = error as { code?: unknown; message?: unknown; statusCode?: unknown };
  const parts = [
    typeof candidate.code === "string" ? candidate.code : undefined,
    typeof candidate.message === "string" ? candidate.message : undefined,
    typeof candidate.statusCode === "number" ? `HTTP ${candidate.statusCode}` : undefined
  ].filter((part): part is string => Boolean(part));
  return parts.length ? parts.join(": ") : undefined;
}
