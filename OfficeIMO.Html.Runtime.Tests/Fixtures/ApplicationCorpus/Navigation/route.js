sessionStorage.reportLoads = String(+(sessionStorage.reportLoads || 0) + 1);
if (!history.state) history.replaceState({ phase: 'review' }, '', '/report/review');
function renderRoute() {
  const approved = history.state && history.state.phase === 'approved';
  document.querySelector('#state').textContent = approved ? 'Approved route snapshot' : 'Review pending';
  window.applicationCorpusReady = approved;
}
document.querySelector('#approve').addEventListener('click', function () {
  history.pushState({ phase: 'approved' }, '', '/report/approved');
  renderRoute();
});
document.querySelector('#back').addEventListener('click', () => history.back());
document.querySelector('#forward').addEventListener('click', () => history.forward());
document.querySelector('#reload').addEventListener('click', () => location.reload());
addEventListener('popstate', renderRoute);
addEventListener('beforeunload', () => sessionStorage.beforeUnloadCount = String(+(sessionStorage.beforeUnloadCount || 0) + 1));
addEventListener('pagehide', event => sessionStorage.lastPageHide = String(event.persisted));
addEventListener('unload', () => sessionStorage.lastUnload = 'yes');
renderRoute();
