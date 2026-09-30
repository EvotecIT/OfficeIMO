namespace OfficeIMO.Invoicing.Pdf;

public sealed partial class InvoicePdfLanguagePack {
    private string AdditionalLabel(InvoicePdfText text) {
        if (EnglishLabels.TryGetValue(text, out string? value)) return value;
        string[] labels = (_builtInAdditionalLabels ? Culture.TwoLetterISOLanguageName : "en") switch {
            "de" => new[] { "Teilrechnung", "Korrigierte Rechnung", "Vorauszahlungsrechnung", "Selbstfakturierte Rechnung", "Einheit", "Positionsnummer", "Beschreibung", "Seite" },
            "pl" => new[] { "Faktura częściowa", "Faktura korygująca", "Faktura zaliczkowa", "Samofakturowanie", "Jednostka", "Numer pozycji", "Opis", "Strona" },
            "fr" => new[] { "Facture partielle", "Facture corrigée", "Facture d'acompte", "Autofacturation", "Unité", "Numéro de ligne", "Description", "Page" },
            "es" => new[] { "Factura parcial", "Factura corregida", "Factura de anticipo", "Autofactura", "Unidad", "Número de línea", "Descripción", "Página" },
            "it" => new[] { "Fattura parziale", "Fattura corretta", "Fattura di acconto", "Autofattura", "Unità", "Numero di riga", "Descrizione", "Pagina" },
            "nl" => new[] { "Deelfactuur", "Gecorrigeerde factuur", "Voorschotfactuur", "Zelf gefactureerde factuur", "Eenheid", "Regelnummer", "Omschrijving", "Pagina" },
            "pt" => new[] { "Fatura parcial", "Fatura corrigida", "Fatura de adiantamento", "Autofaturação", "Unidade", "Número da linha", "Descrição", "Página" },
            "cs" => new[] { "Dílčí faktura", "Opravená faktura", "Zálohová faktura", "Samofakturace", "Jednotka", "Číslo řádku", "Popis", "Strana" },
            "sk" => new[] { "Čiastková faktúra", "Opravená faktúra", "Zálohová faktúra", "Samofakturácia", "Jednotka", "Číslo riadka", "Popis", "Strana" },
            _ => new[] { "Partial invoice", "Corrected invoice", "Prepayment invoice", "Self-billed invoice", "Unit", "Line ID", "Description", "Page" }
        };
        int index = (int)text - (int)InvoicePdfText.PartialInvoice;
        if (index < 0 || index >= labels.Length) throw new ArgumentOutOfRangeException(nameof(text));
        return labels[index];
    }
}
