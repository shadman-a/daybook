import { DaybookSource } from "../types/daybook";

const sourceMark: Record<DaybookSource, string> = {
  Teams: "T",
  Meeting: "M",
  Email: "E",
  "Sent Email": "S",
  File: "F",
  Copilot: "C"
};

export function SourceBadge({ source }: { source: DaybookSource }) {
  const slug = source.replace(" ", "-").toLowerCase();
  return <span className={`badge badge-${slug}`}><i aria-hidden="true">{sourceMark[source]}</i>{source}</span>;
}
