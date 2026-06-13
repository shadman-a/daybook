import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import dayjs from "dayjs";
import { InteractionStatus } from "@azure/msal-browser";
import { useMsal } from "@azure/msal-react";
import { createGraphClient } from "./auth/graphClient";
import { graphScopes } from "./auth/msalConfig";
import { DetailDrawer } from "./components/DetailDrawer";
import { Timeline } from "./components/Timeline";
import { getErrorMessage } from "./graph/errors";
import { fetchDaybook } from "./graph/fetchDaybook";
import { DaybookData, DaybookItem } from "./types/daybook";

type TimelineFilter = "All" | "Teams" | "Context";

const copilotUrl = "https://m365.cloud.microsoft/chat";

export default function App() {
  const { instance, accounts, inProgress } = useMsal();
  const account = instance.getActiveAccount() ?? accounts[0];
  const [date, setDate] = useState(dayjs().format("YYYY-MM-DD"));
  const [data, setData] = useState<DaybookData | null>(null);
  const [selected, setSelected] = useState<DaybookItem>();
  const [filter, setFilter] = useState<TimelineFilter>("All");
  const [loading, setLoading] = useState(false);
  const [signingIn, setSigningIn] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const requestId = useRef(0);

  const graph = useMemo(
    () => account ? createGraphClient(instance, account) : null,
    [instance, account]
  );

  const load = useCallback(async () => {
    if (!graph) return;
    const currentRequest = ++requestId.current;
    setLoading(true);
    setError(null);

    try {
      const next = await fetchDaybook(graph, date);
      if (currentRequest !== requestId.current) return;
      setData(next);
      setSelected((current) => current
        ? next.items.find((item) => item.id === current.id)
        : undefined);
    } catch (loadError) {
      if (currentRequest !== requestId.current) return;
      setError(getErrorMessage(loadError, "Daybook could not load Microsoft Graph data."));
    } finally {
      if (currentRequest === requestId.current) setLoading(false);
    }
  }, [date, graph]);

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

  async function signOut() {
    setError(null);
    try {
      await instance.logoutPopup({ account });
    } catch (signOutError) {
      setError(getErrorMessage(signOutError, "Microsoft sign-out did not complete."));
    }
  }

  function shiftDate(days: number) {
    setDate((current) => dayjs(current).add(days, "day").format("YYYY-MM-DD"));
  }

  if (!import.meta.env.VITE_MS_CLIENT_ID) {
    return <SetupScreen />;
  }

  if (!account) {
    return (
      <SignInScreen
        error={error}
        busy={signingIn || inProgress !== InteractionStatus.None}
        onSignIn={() => void signIn()}
      />
    );
  }

  const filteredItems = (data?.items ?? []).filter((item) => {
    if (filter === "Teams") return item.source === "Teams";
    if (filter === "Context") return item.source !== "Teams";
    return true;
  });

  const contextCount = data
    ? data.meetingCount + data.emailCount + data.fileCount
    : 0;

  return (
    <main className="app-shell">
      <section className="main-panel">
        <header className="topbar">
          <div className="brand-lockup">
            <div className="brand-mark" aria-hidden="true">D</div>
            <div>
              <h1>Daybook</h1>
              <p>Microsoft 365 activity, led by Teams</p>
            </div>
          </div>
          <div className="account-menu">
            <div className="account-copy">
              <strong>{account.name || "Microsoft account"}</strong>
              <span>{account.username}</span>
            </div>
            <button className="icon-button" type="button" onClick={() => void signOut()} aria-label="Sign out">
              <span aria-hidden="true">↗</span>
            </button>
          </div>
        </header>

        <section className="day-heading" aria-labelledby="day-title">
          <div>
            <span className="eyebrow">Daily timeline</span>
            <h2 id="day-title">{dayjs(date).format("dddd, MMMM D")}</h2>
          </div>
          <div className="date-controls">
            <div className="date-stepper">
              <button type="button" onClick={() => shiftDate(-1)} aria-label="Previous day">←</button>
              <input aria-label="Selected day" type="date" value={date} onChange={(event) => setDate(event.target.value)} />
              <button type="button" onClick={() => shiftDate(1)} aria-label="Next day">→</button>
            </div>
            <button className="secondary-button" type="button" onClick={() => setDate(dayjs().format("YYYY-MM-DD"))}>Today</button>
            <button className="primary-button" type="button" onClick={() => void load()} disabled={loading}>
              {loading ? "Syncing…" : "Refresh"}
            </button>
          </div>
        </section>

        <section className="overview-grid" aria-label="Day summary">
          <div className="teams-card">
            <div className="source-icon teams-icon" aria-hidden="true">T</div>
            <div>
              <span>Teams messages</span>
              <strong>{data?.teamsCount ?? "—"}</strong>
              <small>Across active chats</small>
            </div>
            <div className="teams-card-glow" aria-hidden="true" />
          </div>
          <StatCard label="Meetings" value={data?.meetingCount} tone="green" />
          <StatCard label="Email" value={data?.emailCount} tone="amber" />
          <StatCard label="Files" value={data?.fileCount} tone="blue" />
        </section>

        {error && (
          <div className="notice error-notice" role="alert">
            <div><strong>Teams timeline unavailable</strong><p>{error}</p></div>
            <button type="button" onClick={() => void load()}>Try again</button>
          </div>
        )}

        {data?.warnings.map((warning) => (
          <div className="notice warning-notice" key={`${warning.source}-${warning.message}`}>
            <div><strong>{warning.source} is partially unavailable</strong><p>{warning.message}</p></div>
          </div>
        ))}

        <section className="timeline-section" aria-labelledby="timeline-title">
          <div className="section-heading">
            <div>
              <span className="eyebrow">Chronological</span>
              <h2 id="timeline-title">Activity</h2>
            </div>
            <div className="filter-tabs" aria-label="Timeline source filter">
              {(["All", "Teams", "Context"] as TimelineFilter[]).map((option) => {
                const count = option === "All" ? data?.items.length : option === "Teams" ? data?.teamsCount : contextCount;
                return (
                  <button
                    type="button"
                    className={filter === option ? "active" : ""}
                    aria-pressed={filter === option}
                    onClick={() => setFilter(option)}
                    key={option}
                  >
                    {option}<span>{count ?? 0}</span>
                  </button>
                );
              })}
            </div>
          </div>

          {loading && !data
            ? <TimelineSkeleton />
            : error && !data
              ? null
              : <Timeline items={filteredItems} selectedId={selected?.id} onSelect={setSelected} filter={filter} />}
        </section>
      </section>

      <DetailDrawer item={selected} onClose={() => setSelected(undefined)} copilotUrl={copilotUrl} />
    </main>
  );
}

function StatCard({ label, value, tone }: { label: string; value?: number; tone: string }) {
  return (
    <div className={`stat-card stat-${tone}`}>
      <span>{label}</span>
      <strong>{value ?? "—"}</strong>
      <small>Context</small>
    </div>
  );
}

function TimelineSkeleton() {
  return (
    <div className="timeline-skeleton" aria-label="Loading timeline">
      {[0, 1, 2].map((item) => <div className="skeleton-row" key={item}><span /><div><i /><i /></div></div>)}
    </div>
  );
}

function SetupScreen() {
  return (
    <main className="entry-screen">
      <div className="entry-card setup-card">
        <div className="brand-mark large" aria-hidden="true">D</div>
        <span className="eyebrow">Configuration needed</span>
        <h1>Connect Daybook to Microsoft 365</h1>
        <p>Copy <code>.env.example</code> to <code>.env.local</code>, add your Entra SPA client ID, then restart Vite.</p>
        <pre>VITE_MS_CLIENT_ID=your-client-id-here</pre>
      </div>
    </main>
  );
}

function SignInScreen({ error, busy, onSignIn }: { error: string | null; busy: boolean; onSignIn: () => void }) {
  return (
    <main className="entry-screen">
      <section className="signin-layout">
        <div className="signin-copy">
          <div className="brand-lockup light">
            <div className="brand-mark" aria-hidden="true">D</div>
            <strong>Daybook</strong>
          </div>
          <span className="eyebrow">Teams-first work history</span>
          <h1>See the shape of your day.</h1>
          <p>Bring Teams messages into one chronological view, with meetings, email, and files as supporting context.</p>
          <ul>
            <li><span>T</span> Messages from your active Teams chats</li>
            <li><span>24</span> Activity limited to the day you choose</li>
            <li><span>↗</span> Direct links back to Microsoft 365</li>
          </ul>
        </div>
        <div className="entry-card signin-card">
          <span className="eyebrow">Private by design</span>
          <h2>Your workday, in one place</h2>
          <p>Daybook runs in your browser and requests delegated access to your own Microsoft 365 activity.</p>
          {error && <div className="compact-error" role="alert">{error}</div>}
          <button className="microsoft-button" type="button" onClick={onSignIn} disabled={busy}>
            <span className="microsoft-logo" aria-hidden="true"><i /><i /><i /><i /></span>
            {busy ? "Opening Microsoft…" : "Sign in with Microsoft"}
          </button>
          <small>No backend, database, or application permissions.</small>
        </div>
      </section>
    </main>
  );
}
