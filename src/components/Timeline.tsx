import { DaybookItem } from "../types/daybook";
import { TimelineItem } from "./TimelineItem";

export function Timeline({
  items,
  selectedId,
  onSelect,
  filter
}: {
  items: DaybookItem[];
  selectedId?: string;
  onSelect: (item: DaybookItem) => void;
  filter: string;
}) {
  if (!items.length) {
    return (
      <div className="empty-state">
        <div className="empty-icon" aria-hidden="true">○</div>
        <h3>No {filter === "All" ? "activity" : filter.toLowerCase()} found</h3>
        <p>There is nothing to show for this source on the selected day.</p>
      </div>
    );
  }

  return (
    <div className="timeline">
      {items.map((item) => (
        <TimelineItem
          key={item.id}
          item={item}
          selected={selectedId === item.id}
          onSelect={onSelect}
        />
      ))}
    </div>
  );
}
