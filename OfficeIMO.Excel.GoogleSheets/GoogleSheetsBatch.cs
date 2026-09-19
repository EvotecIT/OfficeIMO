using OfficeIMO.GoogleWorkspace;

namespace OfficeIMO.Excel.GoogleSheets {
    /// <summary>
    /// A provider-neutral batch representation of a workbook that can later be translated to Google Sheets API calls.
    /// </summary>
    public sealed class GoogleSheetsBatch {
        private readonly List<GoogleSheetsRequest> _requests = new List<GoogleSheetsRequest>();

        /// <summary>Creates an empty batch; blank titles become <c>Workbook</c>.</summary>
        public GoogleSheetsBatch(
            string title,
            GoogleSheetsTranslationPlan plan,
            TranslationReport report) {
            Title = string.IsNullOrWhiteSpace(title) ? "Workbook" : title;
            Plan = plan ?? throw new ArgumentNullException(nameof(plan));
            Report = report ?? throw new ArgumentNullException(nameof(report));
        }

        /// <summary>Gets the target spreadsheet title.</summary>
        public string Title { get; }
        /// <summary>Gets source counts and pre-export risk classification.</summary>
        public GoogleSheetsTranslationPlan Plan { get; }
        /// <summary>Gets translation notices associated with this batch.</summary>
        public TranslationReport Report { get; }
        /// <summary>Gets provider-neutral requests assembled by the batch compiler.</summary>
        public IReadOnlyList<GoogleSheetsRequest> Requests => _requests;
        internal string? ChartDataSheetName { get; set; }

        internal void Add(GoogleSheetsRequest request) {
            if (request == null) throw new ArgumentNullException(nameof(request));
            _requests.Add(request);
        }
    }
}
