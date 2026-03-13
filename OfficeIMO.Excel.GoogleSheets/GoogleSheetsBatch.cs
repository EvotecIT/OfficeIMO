using OfficeIMO.GoogleWorkspace;

namespace OfficeIMO.Excel.GoogleSheets {
    /// <summary>
    /// A provider-neutral batch representation of a workbook that can later be translated to Google Sheets API calls.
    /// </summary>
    public sealed class GoogleSheetsBatch {
        private readonly List<GoogleSheetsRequest> _requests = new List<GoogleSheetsRequest>();

        public GoogleSheetsBatch(
            string title,
            GoogleSheetsTranslationPlan plan,
            TranslationReport report) {
            Title = string.IsNullOrWhiteSpace(title) ? "Workbook" : title;
            Plan = plan ?? throw new ArgumentNullException(nameof(plan));
            Report = report ?? throw new ArgumentNullException(nameof(report));
        }

        public string Title { get; }
        public GoogleSheetsTranslationPlan Plan { get; }
        public TranslationReport Report { get; }
        public IReadOnlyList<GoogleSheetsRequest> Requests => _requests;

        internal void Add(GoogleSheetsRequest request) {
            if (request == null) throw new ArgumentNullException(nameof(request));
            _requests.Add(request);
        }
    }
}
