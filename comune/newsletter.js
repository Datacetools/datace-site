// Il modulo della newsletter, UNO per /newsletter/, le guide e Aggiornamenti (brief 105): stesso
// codice e stesso aspetto ovunque. La logica è quella di /raccontaci/ (anti-spam _gotcha, consenso,
// Formspree mreddazr con Accept JSON, esito a schermo, evento Umami solo a invio riuscito); qui
// l'email è obbligatoria e senza la spunta del consenso non parte nulla.
// Uso: <div data-newsletter="origine" [data-tema="chiaro"]></div><script src="/comune/newsletter.js"></script>
// Lo script va subito dopo il segnaposto: il modulo compare mentre la pagina si disegna, senza salti.
(function () {
  const CONSENSO = 'Voglio ricevere la newsletter di Datace';
  const OK = 'Fatto, grazie. Ti scriviamo quando esce qualcosa di nuovo.';
  const KO = 'Non siamo riusciti a inviare. Riprova tra poco; se il problema continua, scrivi a info@datacetools.com.';

  if (!document.getElementById('nl-stile')) {
    const s = document.createElement('style');
    s.id = 'nl-stile';
    s.textContent =
      '.nl{--nl-testo:rgba(255,255,255,.7);--nl-campo:rgba(255,255,255,.05);--nl-bordo:rgba(255,255,255,.12);--nl-scritto:white;--nl-link:#F59E0B;}' +
      '.nl.chiaro{--nl-testo:#475569;--nl-campo:white;--nl-bordo:#cbd5e1;--nl-scritto:#0f2040;--nl-link:#1d7a94;}' +
      '.nl-riga{display:flex;flex-wrap:wrap;gap:10px;}' +
      '.nl-email{flex:1 1 220px;min-width:0;background:var(--nl-campo);border:1px solid var(--nl-bordo);border-radius:12px;padding:10px 14px;font-size:14px;color:var(--nl-scritto);font-family:inherit;}' +
      '.nl-email:focus{outline:2px solid rgba(245,158,11,.5);border-color:transparent;}' +
      '.nl-invia{background:#F59E0B;color:#0f2040;font-weight:700;font-size:14px;padding:10px 24px;border-radius:12px;border:none;cursor:pointer;font-family:inherit;}' +
      '.nl-spunta{display:flex;align-items:flex-start;gap:10px;margin-top:12px;font-size:13px;color:var(--nl-testo);line-height:1.6;cursor:pointer;}' +
      '.nl-spunta input{accent-color:#F59E0B;margin-top:4px;flex-shrink:0;}' +
      '.nl-spunta a{color:var(--nl-link);text-decoration:underline;}' +
      '.nl-esito{border-radius:12px;padding:12px 16px;font-size:14px;margin-top:16px;}' +
      '.nl-ok{background:rgba(34,197,94,.1);border:1px solid rgba(34,197,94,.25);color:#86efac;}' +
      '.nl-ko{background:rgba(239,68,68,.1);border:1px solid rgba(239,68,68,.25);color:#fca5a5;}' +
      '.nl.chiaro .nl-ok{color:#15803d;}.nl.chiaro .nl-ko{color:#b91c1c;}';
    document.head.appendChild(s);
  }

  document.querySelectorAll('[data-newsletter]:not(.nl)').forEach(function (box, i) {
    const origine = box.dataset.newsletter;
    const id = 'nl-email-' + i;
    box.classList.add('nl');
    if (box.dataset.tema === 'chiaro') box.classList.add('chiaro');
    box.innerHTML =
      '<form novalidate style="position:relative;">' +
      '<input type="text" name="_gotcha" tabindex="-1" autocomplete="off" aria-hidden="true" style="position:absolute;left:-9999px;width:1px;height:1px;opacity:0;" />' +
      '<div class="nl-riga">' +
      '<input type="email" id="' + id + '" class="nl-email" placeholder="La tua email" aria-label="La tua email" autocomplete="email" maxlength="200" required />' +
      '<button type="submit" class="nl-invia">Iscriviti</button>' +
      '</div>' +
      '<label class="nl-spunta"><input type="checkbox" class="nl-consenso" />' +
      '<span>' + CONSENSO + ' (<a href="/privacy/">informativa</a>).</span></label>' +
      '<div class="nl-esito" role="status" aria-live="polite" hidden></div>' +
      '</form>';

    const form = box.querySelector('form');
    const email = form.querySelector('.nl-email');
    const consenso = form.querySelector('.nl-consenso');
    const bottone = form.querySelector('.nl-invia');
    const esito = form.querySelector('.nl-esito');

    function mostra(testo, buono) {
      esito.textContent = testo;
      esito.className = 'nl-esito ' + (buono ? 'nl-ok' : 'nl-ko');
      esito.hidden = false;
    }

    form.addEventListener('submit', async function (e) {
      e.preventDefault();
      const mail = email.value.trim();
      if (!mail || !email.checkValidity()) { mostra("L'email non sembra corretta: controllala.", false); email.focus(); return; }
      if (!consenso.checked) { mostra('Per iscriverti serve il consenso: spunta la casella sotto l\'email.', false); consenso.focus(); return; }

      // Il campo nascosto compilato è un robot: non si invia nulla a Formspree
      if (form.elements._gotcha.value) { form.reset(); mostra(OK, true); return; }

      const dati = new FormData();
      dati.append('tool', 'Newsletter');
      dati.append('email', mail);
      dati.append('consenso_newsletter', 'sì');
      dati.append('testo_consenso', CONSENSO + ' (informativa: /privacy/)');   // la notifica in posta è la prova
      dati.append('origine', origine);

      bottone.disabled = true; bottone.style.opacity = '.6';
      try {
        const r = await fetch('https://formspree.io/f/mreddazr', { method: 'POST', body: dati, headers: { 'Accept': 'application/json' } });
        if (!r.ok) throw new Error('http ' + r.status);
        form.reset();
        mostra(OK, true);
        try { if (window.umami) window.umami.track('newsletter-iscritto', { origine: origine }); } catch (_) {}
      } catch (_) {
        mostra(KO, false);
      } finally {
        bottone.disabled = false; bottone.style.opacity = '1';
      }
    });
  });
})();
