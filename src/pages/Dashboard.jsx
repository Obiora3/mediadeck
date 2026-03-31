import Badge from "../components/Badge";
import Empty from "../components/Empty";
import { fmtN } from "../utils/formatters";
import { formatRoleLabel } from "../constants/roles";
import { getMpoWorkflowMeta, isMpoAwaitingUser, MPO_STATUS_LABELS } from "../constants/mpoWorkflow";
import { Card, Stat, Btn } from "../components/ui/primitives";

const isArchived = (item) => Boolean(item?.archivedAt);
const activeOnly = (items = []) => (Array.isArray(items) ? items.filter(item => !isArchived(item)) : []);
const archivedOnly = (items = []) => (Array.isArray(items) ? items.filter(isArchived) : []);

const formatAuditTimestamp = (value) => {
  if (!value) return "—";
  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? String(value) : date.toLocaleString();
};

const Dashboard = ({ user, vendors, clients, campaigns, rates, mpos, notifications, unreadNotifications, setPage, onOpenNotifications }) => {
  const liveVendors = activeOnly(vendors);
  const liveClients = activeOnly(clients);
  const liveCampaigns = activeOnly(campaigns);
  const liveMpos = activeOnly(mpos);
  const totalBudget = liveCampaigns.reduce((s, c) => s + (parseFloat(c.budget) || 0), 0);
  const totalMPOValue = liveMpos.reduce((s, m) => s + (m.netVal || 0), 0);
  const pendingApprovals = liveMpos.filter(m => ["draft", "submitted", "reviewed"].includes(m.status || "draft")).length;
  const unreconciledCount = liveMpos.filter(m => (m.reconciliationStatus || "not_started") !== "completed").length;
  const pendingPaymentCount = liveMpos.filter(m => ["received", "approved", "disputed"].includes(m.invoiceStatus || "pending") && (m.paymentStatus || "unpaid") !== "paid").length;
  const pendingProofCount = liveMpos.filter(m => !["received"].includes(m.proofStatus || "pending")).length;
  const recent = [...liveMpos].sort((a, b) => b.createdAt - a.createdAt).slice(0, 4);
  const myWorkflowQueue = [...liveMpos].filter(m => isMpoAwaitingUser(user, m)).sort((a, b) => (b.updatedAt || b.createdAt || 0) - (a.updatedAt || a.createdAt || 0)).slice(0, 4);
  const readyForDispatch = liveMpos.filter(m => String(m.status || "draft").toLowerCase() === "approved").length;
  const needsRevision = liveMpos.filter(m => String(m.status || "draft").toLowerCase() === "rejected").length;
  const recentNotifications = (notifications || []).slice(0, 5);
  const topVendors = Object.values(liveMpos.reduce((acc, mpo) => {
    const key = mpo.vendorName || "Unknown Vendor";
    acc[key] = acc[key] || { name: key, spend: 0, count: 0 };
    acc[key].spend += parseFloat(mpo.netVal) || 0;
    acc[key].count += 1;
    return acc;
  }, {})).sort((a, b) => b.spend - a.spend).slice(0, 4);
  const budgetUtilization = totalBudget > 0 ? Math.min(100, Math.round((totalMPOValue / totalBudget) * 100)) : 0;

  return (
    <div className="fade">
      <div style={{ marginBottom: 28 }}>
        <h1 style={{ fontFamily: "'Syne',sans-serif", fontWeight: 800, fontSize: 26, letterSpacing: "-.03em" }}>
          Good {new Date().getHours() < 12 ? "morning" : new Date().getHours() < 17 ? "afternoon" : "evening"}, <span style={{ color: "var(--accent)" }}>{user.name?.split(" ")[0]}</span> 👋
        </h1>
        <p style={{ color: "var(--text2)", marginTop: 5 }}>{user.agency} — Media Schedule Platform</p>
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(160px,1fr))", gap: 14, marginBottom: 24 }}>
        <Stat icon="🏢" label="Vendors" value={liveVendors.length} sub={`${archivedOnly(vendors).length} archived`} />
        <Stat icon="👥" label="Clients" value={liveClients.length} sub="Brands & orgs" color="var(--blue)" />
        <Stat icon="📢" label="Campaigns" value={liveCampaigns.length} sub={`${liveCampaigns.filter(c => c.status === "active").length} active`} color="var(--green)" />
        <Stat icon="📄" label="MPOs Issued" value={liveMpos.length} sub={`${pendingApprovals} pending approval`} color="var(--purple)" />
        <Stat icon="💰" label="MPO Value" value={`₦${(totalMPOValue / 1e6).toFixed(1)}M`} sub={`Budget pool ${fmtN(totalBudget)}`} color="var(--teal)" />
        <Stat icon="🔔" label="Unread Alerts" value={unreadNotifications} sub={`${pendingPaymentCount} awaiting payment`} color="var(--orange)" />
        <Stat icon="✅" label="My Queue" value={myWorkflowQueue.length} sub={`${readyForDispatch} approved · ${needsRevision} need changes`} color="var(--blue)" />
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "1.2fr .95fr", gap: 18 }}>
        <div style={{ display: "flex", flexDirection: "column", gap: 18 }}>
          <Card>
            <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 18 }}>
              <h2 style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 15 }}>Operational Watchlist</h2>
              <Btn variant="ghost" size="sm" onClick={() => setPage("mpo")}>Open MPO Workspace →</Btn>
            </div>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(180px,1fr))", gap: 12 }}>
              <div style={{ padding: "14px 16px", borderRadius: 12, background: "var(--bg3)", border: "1px solid var(--border)" }}>
                <div style={{ fontSize: 12, color: "var(--text3)", marginBottom: 6 }}>Pending Approvals</div>
                <div style={{ fontFamily: "'Syne',sans-serif", fontWeight: 800, fontSize: 24 }}>{pendingApprovals}</div>
                <div style={{ fontSize: 12, color: "var(--text2)", marginTop: 6 }}>Draft, submitted, and reviewed MPOs</div>
              </div>
              <div style={{ padding: "14px 16px", borderRadius: 12, background: "var(--bg3)", border: "1px solid var(--border)" }}>
                <div style={{ fontSize: 12, color: "var(--text3)", marginBottom: 6 }}>Needs Reconciliation</div>
                <div style={{ fontFamily: "'Syne',sans-serif", fontWeight: 800, fontSize: 24 }}>{unreconciledCount}</div>
                <div style={{ fontSize: 12, color: "var(--text2)", marginTop: 6 }}>Execution not fully reconciled yet</div>
              </div>
              <div style={{ padding: "14px 16px", borderRadius: 12, background: "var(--bg3)", border: "1px solid var(--border)" }}>
                <div style={{ fontSize: 12, color: "var(--text3)", marginBottom: 6 }}>Awaiting Payment</div>
                <div style={{ fontFamily: "'Syne',sans-serif", fontWeight: 800, fontSize: 24 }}>{pendingPaymentCount}</div>
                <div style={{ fontSize: 12, color: "var(--text2)", marginTop: 6 }}>Invoices received/approved but not paid</div>
              </div>
              <div style={{ padding: "14px 16px", borderRadius: 12, background: "var(--bg3)", border: "1px solid var(--border)" }}>
                <div style={{ fontSize: 12, color: "var(--text3)", marginBottom: 6 }}>Awaiting Proof</div>
                <div style={{ fontFamily: "'Syne',sans-serif", fontWeight: 800, fontSize: 24 }}>{pendingProofCount}</div>
                <div style={{ fontSize: 12, color: "var(--text2)", marginTop: 6 }}>Proof of airing still outstanding</div>
              </div>
            </div>
            <div style={{ marginTop: 16 }}>
              <div style={{ display: "flex", justifyContent: "space-between", fontSize: 12, color: "var(--text2)", marginBottom: 6 }}>
                <span>Budget Utilization</span>
                <strong style={{ color: "var(--text)" }}>{budgetUtilization}%</strong>
              </div>
              <div style={{ height: 10, background: "var(--bg2)", borderRadius: 999, overflow: "hidden", border: "1px solid var(--border)" }}>
                <div style={{ width: `${budgetUtilization}%`, height: "100%", background: "linear-gradient(90deg,var(--accent),var(--purple))" }} />
              </div>
            </div>
          </Card>

          <Card>
            <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 18 }}>
              <h2 style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 15 }}>Recent MPOs</h2>
              <Btn variant="ghost" size="sm" onClick={() => setPage("mpo")}>View all →</Btn>
            </div>
            {recent.length === 0 ? (
              <Empty icon="📄" title="No MPOs yet" sub="Generate your first MPO to see it here" />
            ) : (
              <div style={{ display: "flex", flexDirection: "column", gap: 9 }}>
                {recent.map(m => (
                  <div key={m.id} style={{ display: "flex", alignItems: "center", gap: 12, padding: "11px 14px", background: "var(--bg3)", borderRadius: 10, border: "1px solid var(--border)" }}>
                    <div style={{ width: 38, height: 38, background: "rgba(240,165,0,.12)", borderRadius: 9, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 17, flexShrink: 0 }}>📄</div>
                    <div style={{ flex: 1, minWidth: 0 }}>
                      <div style={{ fontWeight: 600, fontSize: 13, color: "var(--text)", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{m.mpoNo || "MPO"} — {m.vendorName}</div>
                      <div style={{ fontSize: 11, color: "var(--text2)" }}>{m.clientName} · {m.month} {m.year}</div>
                    </div>
                    <div style={{ textAlign: "right", flexShrink: 0 }}>
                      <div style={{ fontWeight: 700, color: "var(--accent)", fontSize: 13 }}>{fmtN(m.netVal)}</div>
                      <Badge color={m.status === "approved" ? "green" : m.status === "sent" ? "blue" : "accent"}>{m.status || "draft"}</Badge>
                    </div>
                  </div>
                ))}
              </div>
            )}
          </Card>
        </div>

        <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
          <Card>
            <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 14 }}>
              <h2 style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 15 }}>My Workflow Queue</h2>
              <Btn variant="ghost" size="sm" onClick={() => setPage("mpo")}>Open MPOs →</Btn>
            </div>
            {myWorkflowQueue.length === 0 ? (
              <Empty icon="✅" title="Nothing waiting on you" sub="Approvals and dispatch work assigned to your role will appear here." />
            ) : (
              <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
                {myWorkflowQueue.map(mpo => {
                  const workflowMeta = getMpoWorkflowMeta(mpo);
                  return (
                    <button key={mpo.id} onClick={() => setPage("mpo")} style={{ textAlign: "left", border: "1px solid var(--border)", background: "var(--bg3)", borderRadius: 10, padding: "10px 12px", cursor: "pointer" }}>
                      <div style={{ display: "flex", justifyContent: "space-between", gap: 8, alignItems: "center" }}>
                        <div style={{ fontSize: 13, fontWeight: 700, color: "var(--text)" }}>{mpo.mpoNo || "MPO"} · {mpo.vendorName || "Vendor"}</div>
                        <Badge color={workflowMeta.color}>Waiting on you</Badge>
                      </div>
                      <div style={{ fontSize: 12, color: "var(--text2)", marginTop: 5 }}>{workflowMeta.hint}</div>
                      <div style={{ fontSize: 11, color: "var(--text3)", marginTop: 7 }}>{MPO_STATUS_LABELS[mpo.status || "draft"] || (mpo.status || "draft")} · {mpo.clientName || "No client"}</div>
                    </button>
                  );
                })}
              </div>
            )}
          </Card>

          <Card>
            <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 14 }}>
              <h2 style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 15 }}>Notifications</h2>
              <Btn variant="ghost" size="sm" onClick={onOpenNotifications}>Open inbox →</Btn>
            </div>
            {recentNotifications.length === 0 ? (
              <Empty icon="🔔" title="No alerts yet" sub="Workflow notifications will appear here." />
            ) : (
              <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
                {recentNotifications.map(notification => (
                  <button key={notification.id} onClick={onOpenNotifications} style={{ textAlign: "left", border: "1px solid var(--border)", background: notification.readAt ? "var(--bg3)" : "rgba(240,165,0,.07)", borderRadius: 10, padding: "10px 12px", cursor: "pointer" }}>
                    <div style={{ display: "flex", justifyContent: "space-between", gap: 8 }}>
                      <div style={{ fontSize: 13, fontWeight: 700, color: "var(--text)" }}>{notification.title}</div>
                      {!notification.readAt ? <Badge color="accent">New</Badge> : null}
                    </div>
                    <div style={{ fontSize: 12, color: "var(--text2)", marginTop: 5 }}>{notification.message || "Open settings notifications to review this alert."}</div>
                    <div style={{ fontSize: 11, color: "var(--text3)", marginTop: 7 }}>{formatAuditTimestamp(notification.createdAt)}</div>
                  </button>
                ))}
              </div>
            )}
          </Card>

          <Card>
            <h2 style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 15, marginBottom: 14 }}>Top Vendor Spend</h2>
            {topVendors.length === 0 ? (
              <Empty icon="💼" title="No spend yet" sub="Vendor spend distribution appears once MPOs are issued." />
            ) : (
              <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
                {topVendors.map((vendor, idx) => (
                  <div key={vendor.name} style={{ display: "flex", justifyContent: "space-between", gap: 10, alignItems: "center", borderBottom: idx === topVendors.length - 1 ? "none" : "1px solid var(--border)", paddingBottom: 10 }}>
                    <div>
                      <div style={{ fontWeight: 700, fontSize: 13, color: "var(--text)" }}>{vendor.name}</div>
                      <div style={{ fontSize: 11, color: "var(--text2)" }}>{vendor.count} MPO{vendor.count !== 1 ? "s" : ""}</div>
                    </div>
                    <div style={{ fontWeight: 700, color: "var(--accent)" }}>{fmtN(vendor.spend)}</div>
                  </div>
                ))}
              </div>
            )}
          </Card>

          <Card style={{ background: "linear-gradient(135deg,rgba(240,165,0,.1),rgba(139,92,246,.07))", border: "1px solid rgba(240,165,0,.18)" }}>
            <div style={{ fontSize: 22, marginBottom: 9 }}>📡</div>
            <div style={{ fontFamily: "'Syne',sans-serif", fontWeight: 700, fontSize: 14, marginBottom: 4, color: "var(--text)" }}>{user.agency}</div>
            <div style={{ fontSize: 11, color: "var(--text2)", marginBottom: 8 }}>{user.email}</div>
            <div style={{ display: "inline-flex", alignItems: "center", gap: 8, padding: "6px 10px", borderRadius: 999, background: "rgba(15,23,42,.06)", color: "var(--text)", fontSize: 11, fontWeight: 700 }}>
              <span>{formatRoleLabel(user.role)}</span>
              {unreadNotifications > 0 ? <span>• {unreadNotifications} unread</span> : <span>• Inbox clear</span>}
            </div>
          </Card>
        </div>
      </div>
    </div>
  );
};

export default Dashboard;
