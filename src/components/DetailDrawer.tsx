import dayjs from "dayjs";
import { DaybookItem } from "../types/daybook";
import { SourceBadge } from "./SourceBadge";

export function DetailDrawer({
  item,
  onClose,
  copilotUrl
}: {
  item?: DaybookItem;
  onClose: () => void;
  copilotUrl: string;
}) {
  if (!item) {
    return (
      <aside className="drawer empty-drawer">
        <div className="drawer-brand"><div className="brand-mark" aria-hidden="true">D</div><span>Details</span></div>
        <div className="drawer-empty-copy">
          <div className="focus-ring" aria-hidden="true"><span>+</span></div>
          <h2>Select an activity</h2>
          <p>Open any timeline item to see its people, source details, and link back to Microsoft 365.</p>
        </div>
        <a className="copilot-card" href={copilotUrl} target="_blank" rel="noreferrer">
          <span className="copilot-mark" aria-hidden="true">✦</span>
          <span><strong>Microsoft 365 Copilot</strong><small>Open in a new tab</small></span>
          <i aria-hidden="true">↗</i>
        </a>
      </aside>
    );
  }

  return (
    <aside className="drawer">
      <div className="drawer-toolbar">
        <span>Activity details</span>
        <button className="close" type="button" onClick={onClose} aria-label="Close details">×</button>
      </div>
      <div className="drawer-content">
        <SourceBadge source={item.source} />
        <h2>{item.title}</h2>
        <div className="drawer-time">
          <strong>{item.timestamp ? dayjs(item.timestamp).format("dddd, MMMM D") : item.date ? dayjs(item.date).format("dddd, MMMM D") : "Date unknown"}</strong>
          <span>{formatTimeRange(item)}</span>
        </div>
        {item.preview && <p className="drawer-preview">{item.preview}</p>}
        {!!item.people.length && (
          <section className="detail-section">
            <h3>People</h3>
            <div className="person-list">
              {item.people.map((person, index) => {
                const label = person.displayName || person.email || person.id || "Unknown person";
                return (
                  <div className="person" key={`${label}-${index}`}>
                    <span>{label.slice(0, 1).toUpperCase()}</span>
                    <div><strong>{person.displayName || person.email || person.id}</strong>{person.displayName && person.email && <small>{person.email}</small>}</div>
                  </div>
                );
              })}
            </div>
          </section>
        )}
        {item.sourceUrl && <a className="open-source" href={item.sourceUrl} target="_blank" rel="noreferrer">Open in Microsoft 365 <span>↗</span></a>}
        <details className="raw-details"><summary>Technical details</summary><pre>{JSON.stringify(item.raw, null, 2)}</pre></details>
      </div>
    </aside>
  );
}

function formatTimeRange(item: DaybookItem): string {
  if (!item.timestamp) return "Time unknown";
  const start = dayjs(item.timestamp).format("h:mm A");
  return item.endTimestamp ? `${start} – ${dayjs(item.endTimestamp).format("h:mm A")}` : start;
}
