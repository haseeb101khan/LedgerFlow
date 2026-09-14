const metrics = [
  ["Revenue", "$0"],
  ["Expenses", "$0"],
  ["Net Profit", "$0"],
  ["Inventory Value", "$0"],
];

const phases = [
  "Create organization",
  "Invite team",
  "Import or enter data",
  "Validate mappings",
  "Review dashboard",
  "Ask the analyst",
];

export default function HomePage() {
  return (
    <main className="shell">
      <aside className="sidebar">
        <div className="brand">LedgerFlow</div>
        <nav>
          <a href="#dashboard">Dashboard</a>
          <a href="#imports">Imports</a>
          <a href="#analytics">Analytics</a>
          <a href="#settings">Settings</a>
        </nav>
      </aside>

      <section className="workspace">
        <header className="topbar">
          <div>
            <p>Platform Foundation</p>
            <h1>Company Intelligence Dashboard</h1>
          </div>
          <button type="button">Add Data Source</button>
        </header>

        <section className="metricGrid" id="dashboard">
          {metrics.map(([label, value]) => (
            <article className="card" key={label}>
              <span>{label}</span>
              <strong>{value}</strong>
            </article>
          ))}
        </section>

        <section className="panel" id="imports">
          <div>
            <p>V1 Workflow</p>
            <h2>From messy files to clean business records</h2>
          </div>
          <ol>
            {phases.map((phase) => (
              <li key={phase}>{phase}</li>
            ))}
          </ol>
        </section>
      </section>
    </main>
  );
}
