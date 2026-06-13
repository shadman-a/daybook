import { DaybookData, DaybookItem, DaybookPerson, DaybookWarning } from "../types/daybook";
import {
  GraphEvent,
  GraphFileActivity,
  GraphMailMessage,
  GraphRecipient
} from "../types/graph";
import { TeamsMessageWithChat } from "./fetchTeams";
import { getResource, getTimestamp } from "./fetchFiles";
import { stripHtml } from "./stripHtml";

type NormalizeArgs = {
  date: string;
  teams: TeamsMessageWithChat[];
  events: GraphEvent[];
  inbox: GraphMailMessage[];
  sent: GraphMailMessage[];
  files: GraphFileActivity[];
  warnings: DaybookWarning[];
};

function recipient(value?: GraphRecipient): DaybookPerson {
  return {
    displayName: value?.emailAddress?.name,
    email: value?.emailAddress?.address
  };
}

function hasPerson(person: DaybookPerson): boolean {
  return Boolean(person.id || person.displayName || person.email);
}

function graphDateTime(value?: { dateTime?: string; timeZone?: string }): string {
  const dateTime = value?.dateTime ?? "";
  if (!dateTime) return "";
  if (value?.timeZone === "UTC" && !/(Z|[+-]\d{2}:\d{2})$/.test(dateTime)) return `${dateTime}Z`;
  return dateTime;
}

export function normalizeDaybook(args: NormalizeArgs): DaybookData {
  const teamsItems: DaybookItem[] = args.teams.map(({ chat, message, members }) => {
    const sender = message.from?.user ?? message.from?.application ?? message.from?.device;
    const preview = stripHtml(message.body?.content);
    const memberPeople = members.map((member) => ({
      id: member.userId || member.id,
      displayName: member.displayName,
      email: member.email
    }));
    const people = uniquePeople([
      { id: sender?.id, displayName: sender?.displayName },
      ...memberPeople
    ].filter(hasPerson));
    const participantTitle = memberPeople
      .map((person) => person.displayName || person.email)
      .filter((name): name is string => Boolean(name))
      .slice(0, 3)
      .join(", ");
    const conversationTitle = chat.topic || participantTitle || chat.chatType || "Teams chat";

    return {
      id: `teams-${chat.id}-${message.id}`,
      source: "Teams",
      timestamp: message.createdDateTime ?? "",
      title: chat.topic || message.subject || sender?.displayName || conversationTitle,
      preview: preview || (message.deletedDateTime ? "This message was deleted." : "Teams activity"),
      people,
      conversationId: chat.id,
      conversationTitle,
      sourceUrl: message.webUrl || chat.webUrl,
      importance: message.importance,
      raw: { chat, message }
    };
  });

  const meetingItems: DaybookItem[] = args.events.map((event) => ({
    id: `meeting-${event.id}`,
    source: "Meeting",
    timestamp: graphDateTime(event.start),
    endTimestamp: graphDateTime(event.end),
    title: event.subject || "Meeting",
    preview: event.bodyPreview,
    people: [recipient(event.organizer), ...(event.attendees ?? []).map(recipient)].filter(hasPerson),
    sourceUrl: event.webLink || event.onlineMeeting?.joinUrl || event.onlineMeetingUrl,
    raw: event
  }));

  const emailItems: DaybookItem[] = args.inbox.map((message) => ({
    id: `email-${message.id}`,
    source: "Email",
    timestamp: message.receivedDateTime ?? "",
    title: message.subject || "Email",
    preview: message.bodyPreview,
    people: [recipient(message.from)].filter(hasPerson),
    sourceUrl: message.webLink,
    importance: message.importance,
    raw: message
  }));

  const sentItems: DaybookItem[] = args.sent.map((message) => ({
    id: `sent-${message.id}`,
    source: "Sent Email",
    timestamp: message.sentDateTime ?? "",
    title: message.subject || "Sent email",
    preview: message.bodyPreview,
    people: (message.toRecipients ?? []).map(recipient).filter(hasPerson),
    sourceUrl: message.webLink,
    importance: message.importance,
    raw: message
  }));

  const fileItems: DaybookItem[] = args.files.map((item, index) => {
    const resource = getResource(item);
    return {
      id: `file-${resource.id || index}`,
      source: "File",
      timestamp: getTimestamp(item, resource) ?? "",
      title: resource.name || "File",
      preview: "Opened or updated in Microsoft 365",
      people: [{
        displayName: resource.lastModifiedBy?.user?.displayName
          || resource.createdBy?.user?.displayName
      }].filter(hasPerson),
      sourceUrl: resource.webUrl,
      raw: item
    };
  });

  const items = [...teamsItems, ...meetingItems, ...emailItems, ...sentItems, ...fileItems]
    .filter((item) => item.timestamp)
    .sort((a, b) => new Date(a.timestamp ?? "").getTime() - new Date(b.timestamp ?? "").getTime());

  return {
    date: args.date,
    items,
    teamsCount: teamsItems.length,
    copilotCount: 0,
    meetingCount: meetingItems.length,
    emailCount: emailItems.length + sentItems.length,
    fileCount: fileItems.length,
    warnings: args.warnings
  };
}

function uniquePeople(people: DaybookPerson[]): DaybookPerson[] {
  const seen = new Set<string>();
  return people.filter((person) => {
    const key = person.id || person.email || person.displayName;
    if (!key || seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}
