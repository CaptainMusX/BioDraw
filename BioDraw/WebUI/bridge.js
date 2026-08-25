// BioDraw WebView2 bridge. Messages remain structured objects in both directions.
const BioDraw = {
  send(action, payload) {
    window.chrome.webview.postMessage({ action, payload: payload || {} });
  },
  _handlers: {},
  on(action, fn) { this._handlers[action] = fn; },
  _dispatch(raw) {
    try {
      const data = (raw && typeof raw === 'object') ? raw : JSON.parse(raw);
      const fn = this._handlers[data.action];
      if (fn) fn(data.payload || {});
    } catch (error) {
      console.error('BioDraw bridge message error', error);
    }
  }
};

window.chrome.webview.addEventListener('message', e => BioDraw._dispatch(e.data));

document.addEventListener('DOMContentLoaded', () => {
  const closeBtn = document.querySelector('.tl-close');
  if (closeBtn) closeBtn.addEventListener('click', () => BioDraw.send('cancel'));
  BioDraw.send('ready');
});

document.addEventListener('keydown', event => {
  if (event.key !== 'Escape' || event.defaultPrevented) return;
  if (document.body && document.body.dataset.busy === 'true') return;
  event.preventDefault();
  BioDraw.send('cancel');
});
