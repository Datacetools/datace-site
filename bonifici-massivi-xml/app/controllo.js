// Controllo delle FUNZIONI che il motore usa davvero (non del nome del browser). Scritto
// in JavaScript vecchio stile apposta: deve girare, e mostrare il messaggio, anche nei
// browser che non capiscono il codice moderno di bonifici.min.js.
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
    "Worker": function () { return typeof Worker === "function"; },
    "DecompressionStream deflate-raw": function () { return new DecompressionStream("deflate-raw"); },
    "TextDecoder utf-16 e fatal": function () { return new TextDecoder("utf-16le", { fatal: true }) && new TextEncoder(); },
    "BigInt": function () { return typeof BigInt === "function"; },
    "BigInt tra pagina e Worker": function () { return structuredClone(BigInt(1)) === BigInt(1); },
    "crypto.randomUUID": function () { return typeof crypto.randomUUID === "function"; },
    "regex lookbehind e \\p{Nd}": function () { return new RegExp("(?<![_\\\\])a").test("a") && new RegExp("\\p{Nd}", "u").test("١"); },
    "replaceAll": function () { return "a".replaceAll("a", "b") === "b"; },
    "DOMParser XML": function () { return new DOMParser().parseFromString("<a/>", "application/xml").documentElement.nodeName === "a"; },
    "Blob e download": function () { return typeof URL.createObjectURL === "function" && new Blob(["x"]); },
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
