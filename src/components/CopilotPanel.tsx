import dayjs from "dayjs";
import { CopilotBrief, CopilotExtensionStatus, CopilotProgress } from "../copilot/types";

export function CopilotPanel({
  status,
  progress,
  brief,
  busy,
  hasData,
  disclosureAccepted,
  onAcceptDisclosure,
  onSync,
  onExport,
  onCancel,
  onDownload,
  onClear
}: {
  status: CopilotExtensionStatus;
  progress?: CopilotProgress;
  brief?: CopilotBrief;
  busy: boolean;
  hasData: boolean;
  disclosureAccepted: boolean;
  onAcceptDisclosure: () => void;
  onSync: () => void;
  onExport: () => void;
  onCancel: () => void;
  onDownload: () => void;
  onClear: () => void;
}) {
  return (
    <section className="copilot-panel" aria-labelledby="copilot-heading">
      <div className="copilot-panel-heading">
        <div className="copilot-panel-title">
          <span className="copilot-mark" aria-hidden="true">C</span>
          <div>
            <span className="eyebrow">Microsoft 365 Copilot</span>
            <h2 id="copilot-heading">Copilot sync and daily brief</h2>
          </div>
        </div>
        <span className={`extension-status ${status.installed ? "connected" : "missing"}`}>
          {status.installed ? `Extension connected${status.version ? ` · v${status.version}` : ""}` : "Extension not detected"}
        </span>
      </div>

      <p className="copilot-description">
        Sync signed-in Copilot conversations for this date, or send the complete day to a new Copilot chat for a generated brief.
      </p>

      {!disclosureAccepted && (
        <div className="copilot-disclosure">
          <p>Export uploads a generated Markdown file to Microsoft 365 Copilot. Microsoft may store that file in your OneDrive for Business.</p>
          <button type="button" onClick={onAcceptDisclosure}>I understand</button>
        </div>
      )}

      {progress && (
        <div className="copilot-progress" role="status">
          <div><strong>{progress.message}</strong>{progress.total && <span>{progress.current ?? 0} / {progress.total}</span>}</div>
          {progress.total && <progress value={progress.current ?? 0} max={progress.total} />}
        </div>
      )}

      <div className="copilot-actions">
        <button className="secondary-button" type="button" onClick={onSync} disabled={!status.installed || busy}>
          Sync Copilot
        </button>
        <button className="primary-button" type="button" onClick={onExport} disabled={!status.installed || busy || !hasData || !disclosureAccepted}>
          Export to Copilot
        </button>
        {busy && <button className="text-button danger" type="button" onClick={onCancel}>Cancel</button>}
        <button className="text-button" type="button" onClick={onDownload} disabled={!hasData}>Download Markdown</button>
        <button className="text-button" type="button" onClick={onClear} disabled={!status.installed || busy}>Clear Copilot data</button>
      </div>

      {brief && (
        <article className="brief-panel">
          <div className="brief-heading">
            <div><span className="eyebrow">Generated brief</span><h3>Daily summary</h3></div>
            <span>{dayjs(brief.createdAt).format("MMM D, h:mm A")}</span>
          </div>
          <div className="brief-content">{renderMarkdown(brief.markdown)}</div>
          {brief.conversationUrl && <a href={brief.conversationUrl} target="_blank" rel="noreferrer">Open generated Copilot chat ↗</a>}
        </article>
      )}
    </section>
  );
}

function renderMarkdown(markdown: string) {
  return markdown.split(/\n{2,}/).map((block, index) => {
    const value = block.trim();
    if (!value) return null;
    if (value.startsWith("### ")) return <h4 key={index}>{value.slice(4)}</h4>;
    if (value.startsWith("## ")) return <h3 key={index}>{value.slice(3)}</h3>;
    if (value.startsWith("# ")) return <h3 key={index}>{value.slice(2)}</h3>;
    if (/^(?:[-*] .+(?:\n|$))+/.test(value)) {
      return <ul key={index}>{value.split("\n").map((line) => <li key={line}>{line.replace(/^[-*]\s+/, "")}</li>)}</ul>;
    }
    return <p key={index}>{value}</p>;
  });
}
