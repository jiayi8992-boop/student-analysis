/* Streamlit handshake carries only layout metadata, never student data. */
(() => {
  if (window.parent === window) return;
  const parentOrigin = (() => {
    try { return new URL(document.referrer).origin; } catch { return null; }
  })();
  if (!parentOrigin) return;
  const send = (type, fields = {}) => window.parent.postMessage(
    { isStreamlitMessage: true, type, ...fields }, parentOrigin
  );
  let lastHeight = 0;
  let scheduled = false;
  const resize = () => {
    scheduled = false;
    const height = Math.max(800, Math.ceil(document.querySelector('main').getBoundingClientRect().height));
    if (height !== lastHeight) {
      lastHeight = height;
      send('streamlit:setFrameHeight', { height });
    }
  };
  const schedule = () => {
    if (!scheduled) { scheduled = true; requestAnimationFrame(resize); }
  };
  window.addEventListener('message', event => {
    if (event.source !== window.parent || event.origin !== parentOrigin) return;
    if (event.data?.type === 'streamlit:render') schedule();
  });
  new ResizeObserver(schedule).observe(document.querySelector('main'));
  window.addEventListener('resize', schedule);
  send('streamlit:componentReady', { apiVersion: 1 });
  schedule();
})();
