import { Client } from "@microsoft/microsoft-graph-client";
import { GraphCollection, GraphEvent } from "../types/graph";

export async function fetchCalendarForDay(graph: Client, startZ: string, endZ: string): Promise<GraphEvent[]> {
  const response = await graph
    .api("/me/calendarView")
    .query({ startDateTime: startZ, endDateTime: endZ })
    .header("Prefer", 'outlook.timezone="UTC"')
    .select("id,subject,start,end,organizer,attendees,location,bodyPreview,webLink,onlineMeeting,onlineMeetingUrl,isOnlineMeeting")
    .top(100)
    .get() as GraphCollection<GraphEvent>;
  return response.value ?? [];
}
