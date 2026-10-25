// Controllo delle FUNZIONI che il motore usa davvero (non del nome del browser). Scritto
// in JavaScript vecchio stile apposta: deve girare, e mostrare il messaggio, anche nei
// browser che non capiscono il codice moderno di motore.min.js. Copia di quello di
// Bonifici (bonificixml-web/browser/src/controllo.js): cambiano solo le funzioni.
(function () {
  // Safari (motore WebKit) non è mai stato provato: si blocca per nome, oltre al controllo
  // delle funzioni. Per sbloccarlo un giorno: false.
  var BLOCCA_SAFARI = true;
  var ua = navigator.userAgent || "";
  // Chrome, Edge, Brave e Opera (anche su Mac) hanno "Chrome/" nello user agent e vendor
  // "Google Inc."; Firefox non ha "AppleWebKit". Safari e tutti i browser su iPhone/iPad
  // (che usano WebKit anche se si chiamano Chrome o Firefox) hanno WebKit senza "Chrome/".
  var webkit = /AppleWebKit\//.test(ua) && !/Chrome\/|Chromium\//.test(ua);
  var apple = /Apple/.test(navigator.vendor || "");
  window.__safari = BLOCCA_SAFARI && (webkit || apple);
  if (window.__safari) {
    window.__browserOk = false;
    window.__funzioniMancanti = ["Safari"];
    document.addEventListener("DOMContentLoaded", function () {
      var strumento = document.getElementById("strumento");
      if (strumento) strumento.hidden = true;
      var m = document.getElementById("safari");
      if (m) m.hidden = false;
    });
    return;
  }
  function prova(f) { try { return !!f(); } catch (e) { return false; } }
  var mancano = [];
  var funzioni = {
    "moduli": function () { return "noModule" in document.createElement("script"); },
    "WebAssembly": function () { return typeof WebAssembly === "object" && typeof WebAssembly.instantiate === "function"; },
    // Il Worker «di tipo modulo» non si prova qui: servirebbe crearne uno, e la CSP lo
    // segnerebbe in console. Se manca, il Worker fallisce e la pagina dà l'errore di sempre.
    "Worker": function () { return typeof Worker === "function"; },
    "BigInt64Array": function () { return typeof BigInt64Array === "function"; },
    "TextDecoder": function () { return new TextDecoder("utf-8") && new TextEncoder(); },
    "Response e Blob": function () { return new Response(new Blob(["x"])) && typeof URL.createObjectURL === "function"; },
  };
  for (var nome in funzioni) if (!prova(funzioni[nome])) mancano.push(nome);
  window.__browserOk = mancano.length === 0;
  window.__funzioniMancanti = mancano;
  if (window.__browserOk) return;
  document.addEventListener("DOMContentLoaded", function () {
    var strumento = document.getElementById("strumento");
    if (strumento) strumento.hidden = true;
    var m = document.getElementById("non-supportato");
    if (m) m.hidden = false;
  });
})();
