import dayjs from "dayjs";
import { CopilotBrief } from "../copilot/types";
import { DaybookData, DaybookItem } from "../types/daybook";

export const DEMO_DATE = "2026-06-12";

export function createDemoData(date = DEMO_DATE): DaybookData {
  const at = (time: string) => dayjs(`${date}T${time}:00`).toISOString();
  const items: DaybookItem[] = [
    {
      id: "demo-teams-1",
      source: "Teams",
      timestamp: at("08:42"),
      title: "Launch readiness",
      preview: "Maya confirmed the analytics migration is complete. The team agreed to keep the Friday rollout and monitor activation through the afternoon.",
      people: [
        { displayName: "Maya Chen", email: "maya@example.com" },
        { displayName: "Jordan Lee", email: "jordan@example.com" },
        { displayName: "You", email: "alex@example.com" }
      ],
      conversationId: "demo-launch",
      conversationTitle: "Launch readiness",
      importance: "high",
      raw: { demo: true }
    },
    {
      id: "demo-meeting-1",
      source: "Meeting",
      timestamp: at("09:30"),
      endTimestamp: at("10:00"),
      title: "Product and design standup",
      preview: "Reviewed onboarding polish, empty states, and the final launch checklist.",
      people: [{ displayName: "Product team" }, { displayName: "Design team" }],
      raw: { demo: true }
    },
    {
      id: "demo-copilot-1",
      source: "Copilot",
      timestamp: at("10:18"),
      date,
      title: "Activation analysis",
      preview: "Compared this week's activation funnel with the previous release and identified setup completion as the strongest leading indicator.",
      people: [{ displayName: "You" }, { displayName: "Microsoft 365 Copilot" }],
      conversationId: "demo-copilot-activation",
      conversationTitle: "Activation analysis",
      raw: { demo: true, links: [{ title: "Activation dashboard", url: "https://example.com/demo-dashboard" }] }
    },
    {
      id: "demo-email-1",
      source: "Email",
      timestamp: at("11:07"),
      title: "Customer preview feedback",
      preview: "The preview group liked the faster navigation and asked for clearer extension setup instructions.",
      people: [{ displayName: "Priya Shah", email: "priya@example.com" }],
      importance: "high",
      raw: { demo: true }
    },
    {
      id: "demo-teams-2",
      source: "Teams",
      timestamp: at("13:24"),
      title: "Growth experiments",
      preview: "Jordan will publish the revised onboarding experiment by 3 PM. Maya will validate the event names before it goes live.",
      people: [
        { displayName: "Jordan Lee", email: "jordan@example.com" },
        { displayName: "Maya Chen", email: "maya@example.com" }
      ],
      conversationId: "demo-growth",
      conversationTitle: "Growth experiments",
      raw: { demo: true }
    },
    {
      id: "demo-file-1",
      source: "File",
      timestamp: at("14:10"),
      title: "Launch checklist.docx",
      preview: "Opened or updated in Microsoft 365",
      people: [{ displayName: "You" }],
      raw: { demo: true }
    },
    {
      id: "demo-copilot-2",
      source: "Copilot",
      timestamp: at("15:36"),
      date,
      title: "Draft launch update",
      preview: "Created a concise stakeholder update covering rollout timing, customer feedback, open risks, and the monitoring plan.",
      people: [{ displayName: "You" }, { displayName: "Microsoft 365 Copilot" }],
      conversationId: "demo-copilot-update",
      conversationTitle: "Draft launch update",
      raw: { demo: true }
    },
    {
      id: "demo-sent-1",
      source: "Sent Email",
      timestamp: at("16:12"),
      title: "Friday rollout is on track",
      preview: "Shared the final rollout plan, owners, and afternoon monitoring window with stakeholders.",
      people: [{ displayName: "Launch stakeholders" }],
      raw: { demo: true }
    }
  ];

  return {
    date,
    items,
    teamsCount: 2,
    copilotCount: 2,
    meetingCount: 1,
    emailCount: 2,
    fileCount: 1,
    warnings: []
  };
}

export function createDemoBrief(date = DEMO_DATE): CopilotBrief {
  return {
    id: "demo-brief",
    date,
    exportId: "demo-export",
    fingerprint: "demo",
    createdAt: dayjs(`${date}T16:30:00`).toISOString(),
    markdown: [
      "## Overview",
      "Launch preparation stayed on track. The analytics migration finished, customer feedback was positive, and the team kept the Friday rollout date.",
      "",
      "## Decisions",
      "- Keep the Friday rollout and monitor activation through the afternoon.",
      "- Use setup completion as the primary early activation signal.",
      "",
      "## Action items",
      "- Jordan: publish the revised onboarding experiment by 3 PM.",
      "- Maya: validate analytics event names before launch.",
      "- You: clarify extension setup instructions before the next preview.",
      "",
      "## Blockers",
      "No active blockers were recorded."
    ].join("\n")
  };
}
