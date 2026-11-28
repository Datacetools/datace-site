// Il feedback delle app (brief 153): F24, Fattura in Excel, Bonifici, CSV e JSON in Excel. Il modulo
// resta un <form> verso Formspree (campi e controllo del messaggio obbligatorio del browser), ma parte con
// fetch come /raccontaci/: si resta nella pagina e il risultato aperto non si perde. Stessi testi di
// Raccontaci; se l'invio fallisce il testo scritto resta. Nessun evento Umami.
// Uso: <script src="/comune/feedback.js" defer></script> (cerca da sé i moduli verso formspree.io).
(function () {
  const OK = 'Grazie! Leggiamo ogni risposta.';
  const KO = 'Non siamo riusciti a inviare. Riprova tra poco; se il problema continua, scrivi a info@datacetools.com.';

  if (!document.getElementById('fb-stile')) {
    const s = document.createElement('style');
    s.id = 'fb-stile';
    s.textContent =   // gli stessi colori dell'esito di /raccontaci/
      '.fb-esito{border-radius:12px;padding:12px 16px;font-size:14px;}' +
      '.fb-ok{background:rgba(34,197,94,.1);border:1px solid rgba(34,197,94,.25);color:#86efac;}' +
      '.fb-ko{background:rgba(239,68,68,.1);border:1px solid rgba(239,68,68,.25);color:#fca5a5;}';
    document.head.appendChild(s);
  }

  document.querySelectorAll('form[action^="https://formspree.io/"]').forEach(function (form) {
    const bottone = form.querySelector('[type="submit"]');
    const esito = document.createElement('div');
    esito.setAttribute('role', 'status');
    esito.setAttribute('aria-live', 'polite');
    esito.hidden = true;
    form.appendChild(esito);

    function mostra(testo, buono) {
      esito.textContent = testo;
      esito.className = 'fb-esito ' + (buono ? 'fb-ok' : 'fb-ko');
      esito.hidden = false;
    }

    form.addEventListener('submit', async function (e) {
      e.preventDefault();
      bottone.disabled = true;
      try {
        const r = await fetch(form.action, { method: 'POST', body: new FormData(form), headers: { 'Accept': 'application/json' } });
        if (!r.ok) throw new Error('http ' + r.status);
        form.reset();
        mostra(OK, true);
      } catch (_) {
        mostra(KO, false);
      } finally {
        bottone.disabled = false;
      }
    });
  });
})();
