// Controllo delle FUNZIONI che la pagina usa davvero: copia di controllo.js di Fattura PDF (stesso schema,
// elenco dei convertitori; brief 126). JavaScript vecchio stile apposta: deve girare, e mostrare il
// messaggio, anche nei browser che non capiscono il codice moderno di lettore.js.
(function () {
  // Safari (motore WebKit) non è mai stato provato: si blocca per nome, come nelle altre app.
  var BLOCCA_SAFARI = true;
  var ua = navigator.userAgent || "";
  var webkit = /AppleWebKit\//.test(ua) && !/Chrome\/|Chromium\//.test(ua);
  var apple = /Apple/.test(navigator.vendor || "");
  window.__safari = BLOCCA_SAFARI && (webkit || apple);
  function mostra(id) {
    document.addEventListener("DOMContentLoaded", function () {
      var strumento = document.getElementById("strumento");
      if (strumento) strumento.hidden = true;
      var m = document.getElementById(id);
      if (m) m.hidden = false;
    });
  }
  if (window.__safari) {
    window.__browserOk = false;
    window.__funzioniMancanti = ["Safari"];
    mostra("safari");
    return;
  }
  function prova(f) { try { return !!f(); } catch (e) { return false; } }
  var mancano = [];
  var funzioni = {
    "CompressionStream deflate-raw": function () { return new CompressionStream("deflate-raw"); },
    "TextDecoder windows-1252": function () { return new TextDecoder("windows-1252"); },
    "TextDecoder utf-16be": function () { return new TextDecoder("utf-16be"); },
  };
  for (var nome in funzioni) if (!prova(funzioni[nome])) mancano.push(nome);
  window.__browserOk = mancano.length === 0;
  window.__funzioniMancanti = mancano;
  if (!window.__browserOk) mostra("non-supportato");
})();
