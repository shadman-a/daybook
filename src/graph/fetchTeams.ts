import { Client } from "@microsoft/microsoft-graph-client";
import { GraphChat, GraphChatMessage, GraphCollection } from "../types/graph";
import { fetchGraphPages } from "./paginate";

export type TeamsMessageWithChat = { chat: GraphChat; message: GraphChatMessage };

export type TeamsFetchResult = {
  messages: TeamsMessageWithChat[];
  activeChatCount: number;
  failedChatCount: number;
  truncatedChatList: boolean;
  truncatedMessageChatCount: number;
};

const CHAT_PAGE_LIMIT = 10;
const MESSAGE_PAGE_LIMIT = 20;
const CHAT_BATCH_SIZE = 8;

export async function fetchTeamsMessagesForDay(
  graph: Client,
  startZ: string,
  endZ: string
): Promise<TeamsFetchResult> {
  const chatsResult = await fetchGraphPages<GraphChat>(
    graph,
    graph.api("/me/chats")
      .select("id,topic,chatType,lastUpdatedDateTime,webUrl")
      .top(50),
    CHAT_PAGE_LIMIT
  );
  const chats = chatsResult.items;

  // A chat updated after the selected day's start may still contain messages from that day.
  const activeChats = chats.filter((chat) => (
    !chat.lastUpdatedDateTime || chat.lastUpdatedDateTime >= startZ
  ));

  const messages: TeamsMessageWithChat[] = [];
  let failedChatCount = 0;
  let truncatedMessageChatCount = 0;

  for (let index = 0; index < activeChats.length; index += CHAT_BATCH_SIZE) {
    const batch = activeChats.slice(index, index + CHAT_BATCH_SIZE);
    const settled = await Promise.allSettled(batch.map(async (chat) => {
      const chatMessagesResult = await fetchChatMessagesForDay(graph, chat, startZ, endZ);
      return chatMessagesResult;
    }));

    settled.forEach((result) => {
      if (result.status === "fulfilled") {
        messages.push(...result.value.dayMessages);
        if (result.value.truncated) truncatedMessageChatCount += 1;
      }
      else failedChatCount += 1;
    });
  }

  return {
    messages,
    activeChatCount: activeChats.length,
    failedChatCount,
    truncatedChatList: chatsResult.truncated,
    truncatedMessageChatCount
  };
}

async function fetchChatMessagesForDay(
  graph: Client,
  chat: GraphChat,
  startZ: string,
  endZ: string
): Promise<{ dayMessages: TeamsMessageWithChat[]; truncated: boolean }> {
  const dayMessages: TeamsMessageWithChat[] = [];
  let response = await graph.api(`/me/chats/${encodeURIComponent(chat.id)}/messages`)
    .select("id,createdDateTime,lastModifiedDateTime,deletedDateTime,messageType,importance,subject,body,from,webUrl")
    .orderby("createdDateTime desc")
    .filter(`createdDateTime lt ${endZ}`)
    .top(50)
    .get() as GraphCollection<GraphChatMessage>;
  let page = 1;

  while (true) {
    const pageMessages = response.value ?? [];
    dayMessages.push(...pageMessages
      .filter((message) => {
        const created = message.createdDateTime;
        return created !== undefined && created >= startZ && created < endZ;
      })
      .map((message) => ({ chat, message })));

    const crossedStartOfDay = pageMessages.some((message) => (
      message.createdDateTime !== undefined && message.createdDateTime < startZ
    ));
    const nextLink = response["@odata.nextLink"];

    if (crossedStartOfDay || !nextLink) return { dayMessages, truncated: false };
    if (page >= MESSAGE_PAGE_LIMIT) return { dayMessages, truncated: true };

    response = await graph.api(nextLink).get() as GraphCollection<GraphChatMessage>;
    page += 1;
  }
}
