import { Client } from "@microsoft/microsoft-graph-client";
import { GraphCollection, GraphMailMessage } from "../types/graph";

export type MailFetchResult = {
  inbox: GraphMailMessage[];
  sent: GraphMailMessage[];
  failedFolders: string[];
};

export async function fetchMailForDay(graph: Client, startZ: string, endZ: string): Promise<MailFetchResult> {
  const [inboxResponse, sentResponse] = await Promise.allSettled([
    graph.api("/me/messages")
      .filter(`receivedDateTime ge ${startZ} and receivedDateTime lt ${endZ}`)
      .select("id,subject,from,toRecipients,ccRecipients,receivedDateTime,bodyPreview,importance,conversationId,webLink,hasAttachments")
      .top(50)
      .get(),
    graph.api("/me/mailFolders('SentItems')/messages")
      .filter(`sentDateTime ge ${startZ} and sentDateTime lt ${endZ}`)
      .select("id,subject,toRecipients,ccRecipients,sentDateTime,bodyPreview,importance,conversationId,webLink,hasAttachments")
      .top(50)
      .get()
  ]);

  return {
    inbox: inboxResponse.status === "fulfilled"
      ? ((inboxResponse.value as GraphCollection<GraphMailMessage>).value ?? [])
      : [],
    sent: sentResponse.status === "fulfilled"
      ? ((sentResponse.value as GraphCollection<GraphMailMessage>).value ?? [])
      : [],
    failedFolders: [
      ...(inboxResponse.status === "rejected" ? ["Inbox"] : []),
      ...(sentResponse.status === "rejected" ? ["Sent Items"] : [])
    ]
  };
}
