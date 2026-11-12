/* Il codice Umami di un errore dei controlli di Bonifici (bonifici.js, brief 123 e 124):
   Bonifici e Fatture XML in bonifici lo calcolano dal testo del PRIMO errore.
   Solo il tipo: il testo resta nella pagina, a Umami va il codice.
   ponytail: si riconosce dal testo dei messaggi di bonifici.js; se un testo cambia, la pagina
   ricade sul suo codice generico, mai un dato. L'ordine conta. Elenco in docs/09. */
window.ERRORI_BONIFICI = [
  [/formato Excel|non è un Excel valido/, "not_xlsx"],
  [/^Manca il foglio/, "sheet_missing"],
  [/^Beneficiari: manca la colonna/, "column_missing"],
  [/^Distinta: manca il CUC/, "cuc_missing"],
  [/^Distinta: l'IBAN di addebito/, "debit_iban_invalid"],
  [/^Distinta: data di esecuzione/, "date_unreadable"],
  [/^Distinta: la data di esecuzione/, "date_past"],  // solo Fatture XML in bonifici: in Bonifici è un avviso
  [/^Distinta: tipologia/, "type_invalid"],
  [/^Distinta: «[^»]*» deve essere Sì o No/, "yes_no_invalid"],
  [/^Distinta: l'ID distinta/, "batch_id_invalid"],
  [/^Distinta: il nome dell'ordinante/, "payer_name_too_long"],
  [/^Distinta: manca/, "field_missing"],
  [/^Beneficiari: .* di ESEMPIO/, "example_rows"],
  [/^Beneficiari: nessun bonifico|^Beneficiari: il foglio è vuoto/, "no_payments"],
  [/^Beneficiari: (il foglio ha troppe righe|al massimo)/, "too_many_rows"],
  [/^Riga \d+: manca l'intestatario/, "name_missing"],
  [/^Riga \d+: (manca l'IBAN|IBAN «)/, "iban_invalid"],
  [/^Riga \d+: (l'importo|importo)/, "amount_invalid"],
  [/^Riga \d+: Stato «/, "country_invalid"],
  [/^Riga \d+: con l'indirizzo/, "city_missing"],
  [/^Riga \d+: «[^»]*» supera/, "text_too_long"],
];
window.codiceErroreBonifici = function (testo) {
  var c = window.ERRORI_BONIFICI.find(function (x) { return x[0].test(String(testo || "")); });
  return c ? c[1] : null;
};
