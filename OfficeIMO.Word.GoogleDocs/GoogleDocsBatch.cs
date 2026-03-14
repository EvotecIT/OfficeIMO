using OfficeIMO.GoogleWorkspace;
using OfficeIMO.Word;

namespace OfficeIMO.Word.GoogleDocs {
    /// <summary>
    /// A provider-neutral batch representation of a Word document that can later be translated to Google Docs API calls.
    /// </summary>
    public sealed class GoogleDocsBatch {
        private readonly List<GoogleDocsRequest> _requests = new List<GoogleDocsRequest>();
        private readonly List<GoogleDocsSegment> _segments = new List<GoogleDocsSegment>();

        public GoogleDocsBatch(
            string title,
            GoogleDocsTranslationPlan plan,
            TranslationReport report,
            WordDocumentSnapshot snapshot) {
            Title = string.IsNullOrWhiteSpace(title) ? "Document" : title;
            Plan = plan ?? throw new ArgumentNullException(nameof(plan));
            Report = report ?? throw new ArgumentNullException(nameof(report));
            Snapshot = snapshot ?? throw new ArgumentNullException(nameof(snapshot));
        }

        public string Title { get; }
        public GoogleDocsTranslationPlan Plan { get; }
        public TranslationReport Report { get; }
        public WordDocumentSnapshot Snapshot { get; }
        public IReadOnlyList<GoogleDocsRequest> Requests => _requests;
        public IReadOnlyList<GoogleDocsSegment> Segments => _segments;

        internal void Add(GoogleDocsRequest request) {
            if (request == null) throw new ArgumentNullException(nameof(request));
            _requests.Add(request);
        }

        internal void AddSegment(GoogleDocsSegment segment) {
            if (segment == null) throw new ArgumentNullException(nameof(segment));
            _segments.Add(segment);
        }
    }
}
