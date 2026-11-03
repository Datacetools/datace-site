// Il pop-up degli avvisi, UNO per tutte le app (F24, Fattura, Bonifici): stesso codice,
// stesso aspetto (regola di Davide, 3/11/2026). Modello: il pop-up «Ho capito» di Bonifici.
// Stili scritti qui, non presi dal CSS della pagina: così non cambiano da un'app all'altra.
// Uso: avviso(titolo, testo). Se il pop-up è già aperto, il nuovo avviso va sotto agli altri
// (due problemi dalla stessa scelta di file). Titolo e testo come textContent: mai HTML.
(function () {
  let box, voci, ok;

  function crea() {
    box = document.createElement('div');
    box.id = 'avviso';
    box.setAttribute('role', 'dialog');
    box.setAttribute('aria-modal', 'true');
    box.setAttribute('aria-label', 'Avviso');
    box.style.cssText = 'display:none;position:fixed;inset:0;z-index:60;align-items:center;justify-content:center;' +
      'padding:0 16px;background:rgba(4,8,15,.75);backdrop-filter:blur(6px);-webkit-backdrop-filter:blur(6px);';
    box.addEventListener('click', (e) => { if (e.target === box) chiudi(); });

    const card = document.createElement('div');
    card.style.cssText = 'width:100%;max-width:32rem;padding:28px;border-radius:20px;border:1px solid rgba(245,158,11,.3);' +
      'background:linear-gradient(135deg,rgba(15,32,64,.85),rgba(6,14,31,.95));color:white;';
    voci = document.createElement('div');
    voci.id = 'avviso-voci';
    ok = document.createElement('button');
    ok.id = 'avviso-ok';
    ok.type = 'button';
    ok.textContent = 'Ho capito';
    ok.style.cssText = 'width:100%;padding:12px 24px;border-radius:12px;border:0;font:inherit;font-weight:700;font-size:14px;' +
      'line-height:20px;background:white;color:#0f2040;cursor:pointer;transition:opacity .2s;';
    ok.addEventListener('mouseover', () => { ok.style.opacity = '0.85'; });
    ok.addEventListener('mouseout', () => { ok.style.opacity = '1'; });
    ok.addEventListener('click', chiudi);
    card.append(voci, ok);
    box.append(card);
    document.body.append(box);
    document.addEventListener('keydown', (e) => { if (e.key === 'Escape') chiudi(); });
    box._card = card;
  }

  function avviso(titolo, testo) {
    if (!box) crea();
    const aperto = box.style.display === 'flex';
    if (!aperto) voci.replaceChildren();
    const voce = document.createElement('div');
    voce.style.marginBottom = '16px';
    const h = document.createElement('h2');
    h.textContent = titolo;
    h.style.cssText = 'font-size:17px;font-weight:700;margin:0 0 4px;';
    const p = document.createElement('p');
    p.textContent = testo;
    p.style.cssText = 'font-size:13px;color:rgba(255,255,255,.45);margin:0;';
    voce.append(h, p);
    voci.append(voce);
    if (!aperto) {
      box.style.display = 'flex';
      // l'animazione di Bonifici (bounceIn, 0.4 s)
      box._card.animate([{ transform: 'scale(.8)', opacity: 0 }, { transform: 'scale(1.05)', offset: 0.6 }, { transform: 'scale(1)', opacity: 1 }],
                        { duration: 400, easing: 'ease-out' });
    }
    ok.focus();
  }

  function chiudi() { if (box) box.style.display = 'none'; }

  window.avviso = avviso;
  window.chiudiAvviso = chiudi;
  window.avvisoAperto = () => !!box && box.style.display === 'flex';
})();
