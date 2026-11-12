// Porta un elemento al centro dello schermo, UNO per tutte le app (Fattura in Excel, Fatture XML
// in bonifici): stesso codice (regola di Davide, 3/11/2026). Uso: mostraInVista(elemento).
// Aspetta che il pop-up degli avvisi (/comune/avviso.js) o una finestra <dialog> aperta si chiudano,
// non si muove se l'elemento è già tutto in vista, e senza animazione se l'utente l'ha chiesto.
(function () {
  function mostraInVista(el) {
    if (!el) return;
    const pop = window.avvisoAperto && window.avvisoAperto();
    const dlg = document.querySelector('dialog[open]');
    if (pop || dlg) { (pop ? document : dlg).addEventListener(pop ? 'avviso-chiuso' : 'close', () => mostraInVista(el), { once: true }); return; }
    const r = el.getBoundingClientRect();
    if (!r.height || (r.top >= 0 && r.bottom <= innerHeight)) return;
    el.scrollIntoView({ block: r.height > innerHeight ? 'start' : 'center', behavior: matchMedia('(prefers-reduced-motion: reduce)').matches ? 'auto' : 'smooth' });
  }
  window.mostraInVista = mostraInVista;
})();
