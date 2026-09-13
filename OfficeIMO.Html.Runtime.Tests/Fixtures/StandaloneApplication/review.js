import { visibleRows, total } from './view.js';

const parameters = new URL(location.href).searchParams;
const title = parameters.get('title');
const region = parameters.get('region');
document.querySelector('#selection').textContent = `${title} / ${region}`;
document.querySelector('#stored').textContent = `Stored: ${localStorage.getItem('application:title')} / ${localStorage.getItem('application:region')}`;
document.querySelector('#lifecycle').textContent = `Lifecycle: pagehide=${sessionStorage.getItem('application:pagehide')}, unload=${sessionStorage.getItem('application:unload')}`;

const dataResponse = await fetch('./data.json');
const rows = await dataResponse.json();
document.querySelector('#review-total').textContent = `Total: ${total(visibleRows(rows, region))}`;
document.querySelector('#back').addEventListener('click', () => history.back());
window.reviewReady = true;
