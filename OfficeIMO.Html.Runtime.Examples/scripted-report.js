document.querySelector('#prepare').addEventListener('click', () => {
    fetch('scripted-report.json')
        .then(response => { if (!response.ok) throw new Error('Report data unavailable'); return response.json(); })
        .then(services => setTimeout(() => {
            for (const service of services) {
                const row = document.createElement('tr');
                const name = document.createElement('td'); name.textContent = service.name;
                const count = document.createElement('td'); count.textContent = String(service.count);
                row.appendChild(name); row.appendChild(count); document.querySelector('#rows').appendChild(row);
            }
            document.querySelector('#total').textContent = 'Total: ' + services.reduce((total, service) => total + service.count, 0);
            document.querySelector('#status').textContent = 'Ready — all service results received.';
            window.reportReady = true;
        }, 25));
});
