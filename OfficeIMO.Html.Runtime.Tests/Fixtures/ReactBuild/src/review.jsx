import React from 'react';

export function Review({ title, region, rows, total, onBack }) {
    return <main>
        <h1>Review report</h1>
        <button onClick={onBack}>Back to report</button>
        <h2 id="report-heading">{title} / {region === 'all' ? 'All regions' : region}</h2>
        <table><caption>Regional amounts</caption><thead><tr><th scope="col">Region</th><th scope="col">Amount</th></tr></thead>
            <tbody>{rows.map(row => <tr key={row.region}><th scope="row">{row.region}</th><td>{row.amount}</td></tr>)}</tbody></table>
        <p id="total" role="status" aria-live="polite" className="summary">Total: {total}</p>
    </main>;
}
