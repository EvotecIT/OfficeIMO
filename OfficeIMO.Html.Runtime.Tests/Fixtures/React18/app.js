(() => {
    const h = React.createElement;
    const { useEffect, useState } = React;

    function Report() {
        const [title, setTitle] = useState(() => localStorage.getItem('report:title') || 'Monthly');
        const [region, setRegion] = useState('all');
        const [rows, setRows] = useState(null);
        const [adjustment, setAdjustment] = useState(0);
        const [review, setReview] = useState(false);

        useEffect(() => {
            fetch('/data.json')
                .then(response => response.json())
                .then(data => { setRows(data); window.reportDataReady = true; });
        }, []);
        useEffect(() => { localStorage.setItem('report:title', title); }, [title]);

        const visible = rows === null ? [] : rows.filter(row => region === 'all' || row.region === region);
        const total = visible.reduce((sum, row) => sum + row.amount, adjustment);
        const items = visible.map(row => h('tr', { key: row.region },
            h('th', { scope: 'row' }, row.region), h('td', null, String(row.amount))));

        return h('main', null,
            h('h1', null, review ? 'Review report' : 'Regional report'),
            review ? null : h('div', { className: 'controls' },
                h('label', { htmlFor: 'report-title' }, 'Report title'),
                h('input', { id: 'report-title', value: title, onChange: event => setTitle(event.target.value) }),
                h('label', { htmlFor: 'region' }, 'Region'),
                h('select', { id: 'region', value: region, onChange: event => setRegion(event.target.value) },
                    h('option', { value: 'all' }, 'All regions'),
                    h('option', { value: 'North' }, 'North'),
                    h('option', { value: 'South' }, 'South')),
                h('button', { id: 'adjust', onClick: () => setAdjustment(value => value + 3) }, 'Add adjustment'),
                h('button', { id: 'review', onClick: () => setReview(true) }, 'Prepare review')),
            h('h2', { id: 'report-heading' }, title + ' / ' + (region === 'all' ? 'All regions' : region)),
            h('table', null,
                h('caption', null, 'Regional amounts'),
                h('thead', null, h('tr', null, h('th', { scope: 'col' }, 'Region'), h('th', { scope: 'col' }, 'Amount'))),
                h('tbody', null, items)),
            h('p', { id: 'total', className: 'summary', role: 'status', 'aria-live': 'polite' },
                rows === null ? 'Loading' : 'Total: ' + total));
    }

    ReactDOM.createRoot(document.getElementById('app')).render(h(Report));
})();
