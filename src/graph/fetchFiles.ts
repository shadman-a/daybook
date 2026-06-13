import { Client } from "@microsoft/microsoft-graph-client";
import { GraphCollection, GraphDriveItem, GraphFileActivity, GraphUsedInsight } from "../types/graph";

export type FilesFetchResult = {
  items: GraphFileActivity[];
  failedEndpointCount: number;
};

export async function fetchFilesForDay(graph: Client, startZ: string, endZ: string): Promise<FilesFetchResult> {
  const responses = await Promise.allSettled([
    graph.api("/me/drive/recent").top(50).get(),
    graph.api("/me/insights/used").top(50).get(),
    graph.api("/me/insights/shared").top(50).get()
  ]);

  const files = responses
    .flatMap((result) => result.status === "fulfilled"
      ? ((result.value as GraphCollection<GraphFileActivity>).value ?? [])
      : [])
    .filter((item) => {
      const resource = getResource(item);
      const timestamp = getTimestamp(item, resource);
      return timestamp && timestamp >= startZ && timestamp < endZ;
    });

  const items = [...new Map(files.map((item, index) => {
    const resource = getResource(item);
    return [resource.id ?? `${resource.webUrl ?? resource.name ?? "file"}-${index}`, item];
  })).values()];

  return {
    items,
    failedEndpointCount: responses.filter((result) => result.status === "rejected").length
  };
}

export function getResource(item: GraphFileActivity): GraphDriveItem {
  return "resource" in item && item.resource ? item.resource : item;
}

export function getTimestamp(item: GraphFileActivity, resource = getResource(item)): string | undefined {
  const insight = item as GraphUsedInsight;
  return insight.lastUsed?.lastAccessedDateTime
    ?? resource.lastModifiedDateTime
    ?? resource.createdDateTime;
}
