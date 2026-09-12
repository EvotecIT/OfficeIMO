document.querySelector('#prepare').addEventListener('click', () => {
    Promise.resolve([{ name: 'Document conversion', count: 24 }, { name: 'Content extraction', count: 18 }])
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
