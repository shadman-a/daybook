import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import dayjs from "dayjs";
import { InteractionStatus } from "@azure/msal-browser";
import { useMsal } from "@azure/msal-react";
import { createGraphClient, createGraphClientWithAccessToken } from "./auth/graphClient";
import { getManualTokenIdentity, isExpiredAccessToken, normalizeAccessToken } from "./auth/manualToken";
import { graphScopes, hasMsalClientId } from "./auth/msalConfig";
import { CopilotPanel } from "./components/CopilotPanel";
import { DetailDrawer } from "./components/DetailDrawer";
import { Timeline } from "./components/Timeline";
import {
  cancelCopilotOperation,
  clearCopilotHistory,
  exportDayToCopilot,
  getCopilotDay,
  getCopilotStatus,
  syncCopilotDay
} from "./copilot/bridge";
import {
  buildDaybookMarkdown,
  downloadDaybookMarkdown,
  fingerprintDaybook,
  mergeCopilotRecords
} from "./copilot/daybook";
import { CopilotBrief, CopilotExtensionStatus, CopilotProgress } from "./copilot/types";
import { createDemoBrief, createDemoData, DEMO_DATE } from "./demo/data";
import { getErrorMessage } from "./graph/errors";
import { fetchDaybook } from "./graph/fetchDaybook";
import { DaybookData, DaybookItem } from "./types/daybook";

type TimelineFilter = "All" | "Teams" | "Copilot" | "Context";

const MANUAL_TOKEN_KEY = "daybook.graphExplorerToken";
const COPILOT_DISCLOSURE_KEY = "daybook.copilotUploadDisclosure";
const copilotUrl = "https://m365.cloud.microsoft/chat";

export default function App() {
  const { instance, accounts, inProgress } = useMsal();
  const account = instance.getActiveAccount() ?? accounts[0];
  const [demoMode, setDemoMode] = useState(() => new URLSearchParams(window.location.search).get("demo") === "1");
  const [date, setDate] = useState(() => demoMode ? DEMO_DATE : dayjs().format("YYYY-MM-DD"));
  const [graphData, setGraphData] = useState<DaybookData | null>(null);
  const [data, setData] = useState<DaybookData | null>(null);
  const [selected, setSelected] = useState<DaybookItem>();
  const [filter, setFilter] = useState<TimelineFilter>("All");
  const [search, setSearch] = useState("");
  const [chatFilter, setChatFilter] = useState("");
  const [personFilter, setPersonFilter] = useState("");
  const [loading, setLoading] = useState(false);
  const [signingIn, setSigningIn] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [copilotError, setCopilotError] = useState<string | null>(null);
  const [manualToken, setManualToken] = useState<string | null>(readStoredToken);
  const [extensionStatus, setExtensionStatus] = useState<CopilotExtensionStatus>({ installed: false });
  const [copilotProgress, setCopilotProgress] = useState<CopilotProgress>();
  const [copilotBusy, setCopilotBusy] = useState(false);
  const [brief, setBrief] = useState<CopilotBrief>();
  const [disclosureAccepted, setDisclosureAccepted] = useState(
    () => demoMode || window.localStorage.getItem(COPILOT_DISCLOSURE_KEY) === "accepted"
  );
  const requestId = useRef(0);

  const graph = useMemo(
    () => manualToken
      ? createGraphClientWithAccessToken(manualToken)
      : account
        ? createGraphClient(instance, account)
        : null,
    [instance, account, manualToken]
  );
  const manualIdentity = useMemo(
    () => manualToken ? getManualTokenIdentity(manualToken) : null,
    [manualToken]
  );

  const load = useCallback(async () => {
    if (demoMode) {
      const demoData = createDemoData(date);
      setGraphData(demoData);
      setData(demoData);
      setBrief(createDemoBrief(date));
      setExtensionStatus({ installed: true });
      setError(null);
      setCopilotError(null);
      setSelected((current) => current ? demoData.items.find((item) => item.id === current.id) : undefined);
      return;
    }
    if (!graph) return;
    const currentRequest = ++requestId.current;
    setLoading(true);
    setError(null);
    setCopilotError(null);

    try {
      const [next, status] = await Promise.all([fetchDaybook(graph, date), getCopilotStatus()]);
      if (currentRequest !== requestId.current) return;
      setGraphData(next);
      setExtensionStatus(status);

      let merged = next;
      if (status.installed) {
        try {
          const copilotDay = await getCopilotDay(date);
          merged = mergeCopilotRecords(next, copilotDay.records);
          setBrief(copilotDay.brief);
        } catch (copilotLoadError) {
          setCopilotError(getErrorMessage(copilotLoadError, "Cached Copilot activity could not be loaded."));
        }
      } else {
        setBrief(undefined);
      }

      setData(merged);
      setSelected((current) => current ? merged.items.find((item) => item.id === current.id) : undefined);
    } catch (loadError) {
      if (currentRequest !== requestId.current) return;
      setError(getErrorMessage(loadError, "Daybook could not load Microsoft Graph data."));
    } finally {
      if (currentRequest === requestId.current) setLoading(false);
    }
  }, [date, demoMode, graph]);

  useEffect(() => {
    void load();
  }, [load]);

  async function signIn() {
    setSigningIn(true);
    setError(null);
    try {
      const response = await instance.loginPopup({ scopes: graphScopes, prompt: "select_account" });
      instance.setActiveAccount(response.account);
    } catch (signInError) {
      setError(getErrorMessage(signInError, "Microsoft sign-in did not complete."));
    } finally {
      setSigningIn(false);
    }
  }

  function useAccessToken(value: string) {
    const token = normalizeAccessToken(value);
    setError(null);
    if (!token) return setError("Paste an access token from Graph Explorer.");
    if (isExpiredAccessToken(token)) return setError("That Graph Explorer token has expired. Copy a new access token and try again.");
    window.sessionStorage.setItem(MANUAL_TOKEN_KEY, token);
    setManualToken(token);
  }

  async function signOut() {
    setError(null);
    if (demoMode) {
      window.history.replaceState({}, "", window.location.pathname);
      setDemoMode(false);
      setGraphData(null);
      setData(null);
      setBrief(undefined);
      setSelected(undefined);
      return;
    }
    if (manualToken) {
      requestId.current += 1;
      window.sessionStorage.removeItem(MANUAL_TOKEN_KEY);
      setManualToken(null);
      setGraphData(null);
      setData(null);
      setSelected(undefined);
      setLoading(false);
      return;
    }

    try {
      await instance.logoutPopup({ account });
    } catch (signOutError) {
      setError(getErrorMessage(signOutError, "Microsoft sign-out did not complete."));
    }
  }

  async function syncCopilot() {
    if (demoMode) {
      setCopilotProgress({ stage: "complete", message: "Demo Copilot history is already synchronized" });
      window.setTimeout(() => setCopilotProgress(undefined), 1_200);
      return;
    }
    setCopilotBusy(true);
    setCopilotError(null);
    setCopilotProgress({ stage: "opening", message: "Opening Microsoft 365 Copilot" });
    try {
      const result = await syncCopilotDay(date, setCopilotProgress);
      const base = graphData ?? data;
      if (base) setData(mergeCopilotRecords(base, result.records));
      setBrief(result.brief);
    } catch (syncError) {
      setCopilotError(getErrorMessage(syncError, "Copilot history could not be synchronized."));
    } finally {
      setCopilotBusy(false);
      setCopilotProgress(undefined);
    }
  }

  async function exportToCopilot() {
    if (!data) return;
    if (demoMode) {
      setBrief(createDemoBrief(date));
      setSelected(undefined);
      setCopilotProgress({ stage: "complete", message: "Demo brief generated locally" });
      window.setTimeout(() => setCopilotProgress(undefined), 1_200);
      return;
    }
    const exportId = crypto.randomUUID();
    const filename = `daybook-${date}.md`;
    const markdown = buildDaybookMarkdown(date, data.items, exportId);
    const fingerprint = fingerprintDaybook(data.items);
    setCopilotBusy(true);
    setCopilotError(null);
    setCopilotProgress({ stage: "opening", message: "Opening a new Copilot chat" });

    try {
      const nextBrief = await exportDayToCopilot({ date, exportId, filename, markdown, fingerprint }, setCopilotProgress);
      setBrief(nextBrief);
      setSelected(undefined);
    } catch (exportError) {
      setCopilotError(`${getErrorMessage(exportError, "Copilot export did not complete.")} Download the Markdown file and attach it manually.`);
    } finally {
      setCopilotBusy(false);
      setCopilotProgress(undefined);
    }
  }

  function downloadMarkdown() {
    if (!data) return;
    const exportId = crypto.randomUUID();
    downloadDaybookMarkdown(`daybook-${date}.md`, buildDaybookMarkdown(date, data.items, exportId));
  }

  async function clearCopilot() {
    if (demoMode) {
      setData(createDemoData(date));
      setBrief(undefined);
      return;
    }
    setCopilotError(null);
    try {
      await clearCopilotHistory();
      if (graphData) setData(graphData);
      setBrief(undefined);
    } catch (clearError) {
      setCopilotError(getErrorMessage(clearError, "Local Copilot data could not be cleared."));
    }
  }

  function acceptDisclosure() {
    window.localStorage.setItem(COPILOT_DISCLOSURE_KEY, "accepted");
    setDisclosureAccepted(true);
  }

  function changeFilter(next: TimelineFilter) {
    setFilter(next);
    setSearch("");
    if (next !== "All" && next !== "Teams") {
      setChatFilter("");
      setPersonFilter("");
    }
  }

  function startDemo() {
    window.history.replaceState({}, "", `${window.location.pathname}?demo=1`);
    setDate(DEMO_DATE);
    setDemoMode(true);
    setDisclosureAccepted(true);
  }

  if (!account && !manualToken && !demoMode) {
    return <SignInScreen error={error} busy={signingIn || inProgress !== InteractionStatus.None} msalAvailable={hasMsalClientId} onSignIn={() => void signIn()} onUseToken={useAccessToken} onTryDemo={startDemo} />;
  }

  const displayName = demoMode ? "Alex Morgan" : account?.name || manualIdentity?.name || "Microsoft account";
  const username = demoMode ? "Demo workspace · sample data" : account?.username || manualIdentity?.username || "";
  const allItems = data?.items ?? [];
  const teamItems = allItems.filter((item) => item.source === "Teams");
  const chats = unique(teamItems.map((item) => item.conversationTitle).filter(isString));
  const people = unique(teamItems.flatMap((item) => item.people.map((person) => person.displayName || person.email)).filter(isString));
  const filteredItems = allItems.filter((item) => {
    if (filter === "Teams" && item.source !== "Teams") return false;
    if (filter === "Copilot" && item.source !== "Copilot") return false;
    if (filter === "Context" && (item.source === "Teams" || item.source === "Copilot")) return false;
    if (chatFilter && (item.source !== "Teams" || item.conversationTitle !== chatFilter)) return false;
    if (personFilter && (item.source !== "Teams" || !item.people.some((person) => (person.displayName || person.email) === personFilter))) return false;
    if (!search.trim()) return true;
    const haystack = [item.title, item.preview, item.conversationTitle, ...item.people.map((person) => `${person.displayName ?? ""} ${person.email ?? ""}`)].join(" ").toLowerCase();
    return haystack.includes(search.trim().toLowerCase());
  });
  const contextCount = data ? data.meetingCount + data.emailCount + data.fileCount : 0;

  return (
    <main className="app-shell">
      <section className="main-panel">
        <header className="topbar">
          <div className="brand-lockup"><div className="brand-mark" aria-hidden="true">D</div><div><h1>Daybook {demoMode && <span className="demo-pill">Demo</span>}</h1><p>Microsoft 365 activity, led by Teams and Copilot</p></div></div>
          <div className="account-menu"><div className="account-copy"><strong>{displayName}</strong><span>{username}</span></div><button className="icon-button" type="button" onClick={() => void signOut()} aria-label="Sign out"><span aria-hidden="true">↗</span></button></div>
        </header>

        <section className="day-heading" aria-labelledby="day-title">
          <div><span className="eyebrow">Daily timeline</span><h2 id="day-title">{dayjs(date).format("dddd, MMMM D")}</h2></div>
          <div className="date-controls">
            <div className="date-stepper"><button type="button" onClick={() => setDate((current) => dayjs(current).subtract(1, "day").format("YYYY-MM-DD"))} aria-label="Previous day">←</button><input aria-label="Selected day" type="date" value={date} onChange={(event) => setDate(event.target.value)} /><button type="button" onClick={() => setDate((current) => dayjs(current).add(1, "day").format("YYYY-MM-DD"))} aria-label="Next day">→</button></div>
            <button className="secondary-button" type="button" onClick={() => setDate(dayjs().format("YYYY-MM-DD"))}>Today</button>
            <button className="primary-button" type="button" onClick={() => void load()} disabled={loading}>{loading ? "Syncing…" : "Refresh"}</button>
          </div>
        </section>

        <section className="overview-grid" aria-label="Day summary">
          <div className="teams-card"><div className="source-icon teams-icon" aria-hidden="true">T</div><div><span>Teams messages</span><strong>{data?.teamsCount ?? "—"}</strong><small>Across active chats</small></div><div className="teams-card-glow" aria-hidden="true" /></div>
          <StatCard label="Copilot" value={data?.copilotCount} tone="purple" />
          <StatCard label="Meetings" value={data?.meetingCount} tone="green" />
          <StatCard label="Email" value={data?.emailCount} tone="amber" />
          <StatCard label="Files" value={data?.fileCount} tone="blue" />
        </section>

        {error && <Notice title="Teams timeline unavailable" message={error} tone="error" action={<button type="button" onClick={() => void load()}>Try again</button>} />}
        {copilotError && <Notice title="Copilot is partially unavailable" message={copilotError} tone="warning" />}
        {data?.warnings.map((warning) => <Notice key={`${warning.source}-${warning.message}`} title={`${warning.source} is partially unavailable`} message={warning.message} tone="warning" />)}

        <CopilotPanel
          status={extensionStatus}
          progress={copilotProgress}
          brief={brief}
          busy={copilotBusy}
          hasData={Boolean(data?.items.length)}
          disclosureAccepted={disclosureAccepted}
          onAcceptDisclosure={acceptDisclosure}
          onSync={() => void syncCopilot()}
          onExport={() => void exportToCopilot()}
          onCancel={() => void cancelCopilotOperation()}
          onDownload={downloadMarkdown}
          onClear={() => void clearCopilot()}
          onViewBrief={() => setSelected(undefined)}
        />

        <section className="timeline-section" aria-labelledby="timeline-title">
          <div className="section-heading"><div><span className="eyebrow">Chronological</span><h2 id="timeline-title">Activity</h2></div><div className="filter-tabs" aria-label="Timeline source filter">{(["All", "Teams", "Copilot", "Context"] as TimelineFilter[]).map((option) => { const count = option === "All" ? data?.items.length : option === "Teams" ? data?.teamsCount : option === "Copilot" ? data?.copilotCount : contextCount; return <button type="button" className={filter === option ? "active" : ""} aria-pressed={filter === option} onClick={() => changeFilter(option)} key={option}>{option}<span>{count ?? 0}</span></button>; })}</div></div>
          <div className="activity-tools">
            <label className="search-field"><span>Search {filter.toLowerCase()}</span><input type="search" value={search} onChange={(event) => setSearch(event.target.value)} placeholder={`Search ${filter.toLowerCase()} activity`} /></label>
            <label><span>Teams chat</span><select value={chatFilter} disabled={filter === "Copilot" || filter === "Context"} onChange={(event) => setChatFilter(event.target.value)}><option value="">All chats</option>{chats.map((chat) => <option key={chat}>{chat}</option>)}</select></label>
            <label><span>Person</span><select value={personFilter} disabled={filter === "Copilot" || filter === "Context"} onChange={(event) => setPersonFilter(event.target.value)}><option value="">All people</option>{people.map((person) => <option key={person}>{person}</option>)}</select></label>
          </div>
          {loading && !data ? <TimelineSkeleton /> : error && !data ? null : <Timeline items={filteredItems} selectedId={selected?.id} onSelect={setSelected} filter={filter} />}
        </section>
      </section>
      <DetailDrawer item={selected} brief={brief} onClose={() => setSelected(undefined)} copilotUrl={copilotUrl} />
    </main>
  );
}

function readStoredToken(): string | null {
  const token = window.sessionStorage.getItem(MANUAL_TOKEN_KEY);
  if (!token || isExpiredAccessToken(token)) {
    window.sessionStorage.removeItem(MANUAL_TOKEN_KEY);
    return null;
  }
  return token;
}

function unique(values: string[]): string[] {
  return [...new Set(values)].sort((a, b) => a.localeCompare(b));
}

function isString(value: string | undefined): value is string {
  return Boolean(value);
}

function Notice({ title, message, tone, action }: { title: string; message: string; tone: "error" | "warning"; action?: React.ReactNode }) {
  return <div className={`notice ${tone}-notice`} role={tone === "error" ? "alert" : undefined}><div><strong>{title}</strong><p>{message}</p></div>{action}</div>;
}

function StatCard({ label, value, tone }: { label: string; value?: number; tone: string }) {
  return <div className={`stat-card stat-${tone}`}><span>{label}</span><strong>{value ?? "—"}</strong><small>Context</small></div>;
}

function TimelineSkeleton() {
  return <div className="timeline-skeleton" aria-label="Loading timeline">{[0, 1, 2].map((item) => <div className="skeleton-row" key={item}><span /><div><i /><i /></div></div>)}</div>;
}

function SignInScreen({ error, busy, msalAvailable, onSignIn, onUseToken, onTryDemo }: { error: string | null; busy: boolean; msalAvailable: boolean; onSignIn: () => void; onUseToken: (token: string) => void; onTryDemo: () => void }) {
  const [token, setToken] = useState("");
  return (
    <main className="entry-screen"><section className="signin-layout"><div className="signin-copy"><div className="brand-lockup light"><div className="brand-mark" aria-hidden="true">D</div><strong>Daybook</strong></div><span className="eyebrow">Teams-first work history</span><h1>See the shape of your day.</h1><p>Bring Teams and Copilot conversations into one chronological view, with meetings, email, and files as supporting context.</p><ul><li><span>T</span> Messages from active Teams chats</li><li><span>C</span> Local Copilot history through the extension</li><li><span>24</span> Activity limited to the day you choose</li></ul></div><div className="entry-card signin-card"><span className="eyebrow">Private by design</span><h2>Your workday, in one place</h2><p>Use a temporary Graph Explorer access token. It remains in this browser tab session and is cleared when you sign out or close the session.</p>{error && <div className="compact-error" role="alert">{error}</div>}<form className="token-form" onSubmit={(event) => { event.preventDefault(); onUseToken(token); }}><label htmlFor="graph-token">Graph Explorer access token</label><input id="graph-token" type="password" value={token} onChange={(event) => setToken(event.target.value)} autoComplete="off" spellCheck={false} placeholder="eyJ0eXAiOiJKV1QiLCJ..." /><button className="primary-button token-button" type="submit">Use access token</button></form>{msalAvailable && <><div className="auth-divider"><span>or</span></div><button className="microsoft-button" type="button" onClick={onSignIn} disabled={busy}><span className="microsoft-logo" aria-hidden="true"><i /><i /><i /><i /></span>{busy ? "Opening Microsoft…" : "Sign in with Microsoft"}</button></>}<div className="auth-divider"><span>preview</span></div><button className="demo-button" type="button" onClick={onTryDemo}>Try demo with sample data</button><small>Tokens are sent only to Microsoft Graph and are never shared with the extension.</small></div></section></main>
  );
}
