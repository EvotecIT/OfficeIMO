import React, { useEffect, useState } from 'react';
import { createRoot } from 'react-dom/client';

function App() {
    const [rows, setRows] = useState(null);
    const [title, setTitle] = useState('Monthly');
    const [region, setRegion] = useState('all');
    const [adjustment, setAdjustment] = useState(0);
    const [route, setRoute] = useState(location.pathname === '/review' ? 'review' : 'dashboard');
    const [reviewComponent, setReviewComponent] = useState(null);

    useEffect(() => {
        fetch('/data.json').then(response => response.json()).then(data => setRows(data));
        const onHistory = () => setRoute(location.pathname === '/review' ? 'review' : 'dashboard');
        window.addEventListener('popstate', onHistory);
        return () => window.removeEventListener('popstate', onHistory);
    }, []);

    const visible = rows === null ? [] : rows.filter(row => region === 'all' || row.region === region);
    const total = visible.reduce((sum, row) => sum + row.amount, adjustment);
    const openReview = () => import('./review.jsx').then(module => {
        setReviewComponent(() => module.Review);
        history.pushState({ route: 'review' }, '', '/review');
        setRoute('review');
    });

    if (route === 'review' && reviewComponent !== null) {
        const Review = reviewComponent;
        return <Review title={title} region={region} rows={visible} total={total}
            onBack={() => { history.back(); setRoute('dashboard'); }} />;
    }

    return <main>
        <h1>Regional report</h1>
        <div className="controls">
            <label htmlFor="report-title">Report title</label>
            <input id="report-title" value={title} onChange={event => setTitle(event.target.value)} />
            <label htmlFor="region">Region</label>
            <select id="region" value={region} onChange={event => setRegion(event.target.value)}>
                <option value="all">All regions</option><option value="North">North</option><option value="South">South</option>
            </select>
            <button onClick={() => setAdjustment(value => value + 3)}>Add adjustment</button>
            <button onClick={openReview}>Prepare review</button>
        </div>
        <h2 id="report-heading">{title} / {region === 'all' ? 'All regions' : region}</h2>
        <table><caption>Regional amounts</caption><thead><tr><th scope="col">Region</th><th scope="col">Amount</th></tr></thead>
            <tbody>{visible.map(row => <tr key={row.region}><th scope="row">{row.region}</th><td>{row.amount}</td></tr>)}</tbody></table>
        <p id="total" role="status" aria-live="polite" className="summary">{rows === null ? 'Loading' : `Total: ${total}`}</p>
    </main>;
}

createRoot(document.getElementById('app')).render(<App />);
