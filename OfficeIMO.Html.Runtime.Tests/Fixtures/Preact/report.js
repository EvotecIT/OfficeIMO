(() => {
    const { h, render } = window.preact;
    const { useState, useEffect } = window.preactHooks;
    function Report() {
        const [name, setName] = useState(() => localStorage.getItem('report:name') || 'Untitled');
        const [adjustments, setAdjustments] = useState(0);
        const [rows, setRows] = useState(null);
        useEffect(() => {
            const controller = new AbortController();
            fetch('/data.json', { signal: controller.signal })
                .then(response => response.json())
                .then(setRows)
                .catch(error => { if (error.name !== 'AbortError') throw error; });
            return () => controller.abort();
        }, []);
        useEffect(() => { localStorage.setItem('report:name', name); }, [name]);
        return h('section', null,
            h('h1', null, 'Application report'),
            h('label', { htmlFor: 'name' }, 'Report name'),
            h('input', { id: 'name', value: name, onInput: event => setName(event.currentTarget.value) }),
            h('p', { id: 'report-name' }, 'Report: ' + name),
            h('button', { id: 'increment', onClick: () => setAdjustments(value => value + 1) }, 'Add adjustment'),
            h('p', { id: 'adjustments' }, 'Adjustments: ' + adjustments),
            h('ul', null, rows && rows.map(row => h('li', { key: row.name }, row.name + ': ' + row.value))),
            h('p', { id: 'total' }, rows ? 'Total: ' + rows.reduce((sum, row) => sum + row.value, 0) : 'Loading'));
    }
    const root = document.querySelector('#app');
    window.mutationTypes = [];
    new MutationObserver(records => mutationTypes.push(...records.map(record => record.type)))
        .observe(root, { childList: true, subtree: true, attributes: true });
    window.mountReport = () => render(h(Report), root);
    window.unmountReport = () => render(null, root);
})();
