import dayjs from "dayjs";
import { DaybookItem } from "../types/daybook";
import { AvatarStack } from "./AvatarStack";
import { SourceBadge } from "./SourceBadge";

export function TimelineItem({
  item,
  selected,
  onSelect
}: {
  item: DaybookItem;
  selected: boolean;
  onSelect: (item: DaybookItem) => void;
}) {
  return (
    <button
      type="button"
      className={`timeline-item source-${slug(item.source)}${selected ? " selected" : ""}`}
      aria-pressed={selected}
      onClick={() => onSelect(item)}
    >
      <div className="time-column">
        <span>{item.timestamp ? dayjs(item.timestamp).format("h:mm") : "—"}</span>
        <small>{item.timestamp ? dayjs(item.timestamp).format("A") : "UNKNOWN"}</small>
      </div>
      <div className="timeline-marker"><i /></div>
      <div className="item-main">
        <div className="item-top">
          <SourceBadge source={item.source} />
          {item.importance === "high" && <span className="importance">Important</span>}
          <span className="open-hint" aria-hidden="true">→</span>
        </div>
        <h3>{item.title}</h3>
        {item.preview && <p>{item.preview}</p>}
        <AvatarStack people={item.people} />
      </div>
    </button>
  );
}

function slug(value: string): string {
  return value.replace(" ", "-").toLowerCase();
}
