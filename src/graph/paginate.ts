import { Client, GraphRequest } from "@microsoft/microsoft-graph-client";
import { GraphCollection } from "../types/graph";

export async function fetchGraphPages<T>(
  graph: Client,
  initialRequest: GraphRequest,
  maxPages: number
): Promise<{ items: T[]; truncated: boolean }> {
  const items: T[] = [];
  let response = await initialRequest.get() as GraphCollection<T>;
  let page = 1;

  while (true) {
    items.push(...(response.value ?? []));
    const nextLink = response["@odata.nextLink"];
    if (!nextLink || page >= maxPages) break;
    response = await graph.api(nextLink).get() as GraphCollection<T>;
    page += 1;
  }

  return {
    items,
    truncated: Boolean(response["@odata.nextLink"])
  };
}
