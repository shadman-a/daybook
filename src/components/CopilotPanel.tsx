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
  onClear,
  onViewBrief
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
  onViewBrief: () => void;
}) {
  return (
    <section className="copilot-panel" aria-labelledby="copilot-heading">
      <div className="copilot-panel-heading">
        <div className="copilot-panel-title">
          <span className="copilot-mark" aria-hidden="true">C</span>
          <div>
            <span className="eyebrow">Microsoft 365 Copilot</span>
            <h2 id="copilot-heading">Copilot tools</h2>
          </div>
        </div>
        <span className={`extension-status ${status.installed ? "connected" : "missing"}`}>
          {status.installed ? `Extension connected${status.version ? ` · v${status.version}` : ""}` : "Extension not detected"}
        </span>
      </div>

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
        {brief && <button className="text-button view-brief-button" type="button" onClick={onViewBrief}>View daily brief →</button>}
      </div>
    </section>
  );
}
