using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    /// <summary>
    /// Helper methods that normalize documents for better rendering in Word Online and Google Docs.
    /// </summary>
    public partial class WordDocument {
        /// <summary>
        /// Walks all tables in body and headers/footers and normalizes their grid/widths
        /// so online viewers render them consistently.
        /// </summary>
        public void NormalizeTablesForOnline() {
            try {
                var main = _wordprocessingDocument.MainDocumentPart;
                if (main == null) return;

                // Body tables
                foreach (var t in main.Document?.Body?.Descendants<Table>() ?? Enumerable.Empty<Table>()) {
                    try {
                        var wt = new WordTable(this, t, initializeChildren: true);
                        wt.NormalizeForOnline();
                    } catch { }
                }

                // Header tables
                foreach (var hp in main.HeaderParts) {
                    foreach (var t in hp.Header?.Descendants<Table>() ?? Enumerable.Empty<Table>()) {
                        try {
                            var wt = new WordTable(this, t, initializeChildren: true);
                            wt.NormalizeForOnline();
                        } catch { }
                    }
                }

                // Footer tables
                foreach (var fp in main.FooterParts) {
                    foreach (var t in fp.Footer?.Descendants<Table>() ?? Enumerable.Empty<Table>()) {
                        try {
                            var wt = new WordTable(this, t, initializeChildren: true);
                            wt.NormalizeForOnline();
                        } catch { }
                    }
                }
            } catch { }
        }
    }
}
