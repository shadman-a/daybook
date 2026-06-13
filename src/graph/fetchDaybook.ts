import { Client } from "@microsoft/microsoft-graph-client";
import { DaybookWarning, DaybookWarningSource } from "../types/daybook";
import { GraphEvent } from "../types/graph";
import { getDayRange } from "./dateRange";
import { getErrorMessage } from "./errors";
import { fetchCalendarForDay } from "./fetchCalendar";
import { fetchFilesForDay, FilesFetchResult } from "./fetchFiles";
import { fetchMailForDay, MailFetchResult } from "./fetchMail";
import { fetchTeamsMessagesForDay } from "./fetchTeams";
import { normalizeDaybook } from "./normalize";

export async function fetchDaybook(graph: Client, date: string) {
  const { startZ, endZ } = getDayRange(date);

  const [teams, calendarResult, mailResult, filesResult] = await Promise.all([
    fetchTeamsMessagesForDay(graph, startZ, endZ),
    optionalSource("Calendar", fetchCalendarForDay(graph, startZ, endZ), [] as GraphEvent[]),
    optionalSource("Mail", fetchMailForDay(graph, startZ, endZ), { inbox: [], sent: [], failedFolders: [] } as MailFetchResult),
    optionalSource("Files", fetchFilesForDay(graph, startZ, endZ), { items: [], failedEndpointCount: 0 } as FilesFetchResult)
  ]);

  const warnings = [calendarResult.warning, mailResult.warning, filesResult.warning]
    .filter((warning): warning is DaybookWarning => warning !== undefined);

  if (teams.failedChatCount > 0) {
    warnings.unshift({
      source: "Teams",
      message: `${teams.failedChatCount} of ${teams.activeChatCount} active chats could not be read.`
    });
  }

  if (teams.truncatedChatList || teams.truncatedMessageChatCount > 0) {
    warnings.unshift({
      source: "Teams",
      message: "The Teams result was unusually large, so the displayed timeline may be incomplete."
    });
  }

  if (mailResult.value.failedFolders.length > 0) {
    warnings.push({
      source: "Mail",
      message: `${mailResult.value.failedFolders.join(" and ")} could not be read.`
    });
  }

  if (filesResult.value.failedEndpointCount > 0) {
    warnings.push({
      source: "Files",
      message: `${filesResult.value.failedEndpointCount} of 3 recent-file sources could not be read.`
    });
  }

  return normalizeDaybook({
    date,
    teams: teams.messages,
    events: calendarResult.value,
    inbox: mailResult.value.inbox,
    sent: mailResult.value.sent,
    files: filesResult.value.items,
    warnings
  });
}

async function optionalSource<T>(
  source: DaybookWarningSource,
  request: Promise<T>,
  fallback: T
): Promise<{ value: T; warning?: DaybookWarning }> {
  try {
    return { value: await request };
  } catch (error) {
    return {
      value: fallback,
      warning: {
        source,
        message: getErrorMessage(error, `${source} context could not be loaded.`)
      }
    };
  }
}
