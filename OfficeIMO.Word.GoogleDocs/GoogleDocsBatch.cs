using OfficeIMO.GoogleWorkspace;
using OfficeIMO.Word;

namespace OfficeIMO.Word.GoogleDocs {
    /// <summary>
    /// A provider-neutral batch representation of a Word document that can later be translated to Google Docs API calls.
    /// </summary>
    public sealed class GoogleDocsBatch {
        private readonly List<GoogleDocsRequest> _requests = new List<GoogleDocsRequest>();
        private readonly List<GoogleDocsSegment> _segments = new List<GoogleDocsSegment>();

        /// <summary>Creates an empty batch; a blank title becomes <c>Document</c>.</summary>
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

        /// <summary>Gets the target Google document title.</summary>
        public string Title { get; }
        /// <summary>Gets counts and fidelity classification for the source.</summary>
        public GoogleDocsTranslationPlan Plan { get; }
        /// <summary>Gets translation notices associated with this batch.</summary>
        public TranslationReport Report { get; }
        /// <summary>Gets the Word snapshot used to compile this batch.</summary>
        public WordDocumentSnapshot Snapshot { get; }
        /// <summary>Gets provider-neutral body requests in source order.</summary>
        public IReadOnlyList<GoogleDocsRequest> Requests => _requests;
        /// <summary>Gets compiled header and footer segments.</summary>
        public IReadOnlyList<GoogleDocsSegment> Segments => _segments;
        internal GoogleDocsWriteControlState? WriteControlState { get; set; }
        internal string? TargetTabId { get; set; }
        internal IReadOnlyList<string> TargetTabIds { get; set; } = Array.Empty<string>();

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
