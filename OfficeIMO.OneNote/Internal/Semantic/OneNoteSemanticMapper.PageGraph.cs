namespace OfficeIMO.OneNote;

internal static partial class OneNoteSemanticMapper {
    private sealed class PageMappingState {
        private readonly OneNoteReaderOptions _options;
        private readonly HashSet<string> _visited = new HashSet<string>(StringComparer.Ordinal);
        private readonly HashSet<string> _active = new HashSet<string>(StringComparer.Ordinal);

        internal PageMappingState(OneNoteReaderOptions options) {
            _options = options;
        }

        internal bool TryEnter(string key) {
            if (_visited.Contains(key)) return false;
            if (_active.Count >= _options.MaxPageRelationshipDepth) {
                throw new OneNoteFormatException(
                    "ONENOTE_PAGE_GRAPH_DEPTH",
                    "The conflict and version-history page relationship depth limit was exceeded.");
            }
            if (!TryVisit(key)) return false;
            _active.Add(key);
            return true;
        }

        internal bool TryVisit(string key) {
            if (_visited.Contains(key)) return false;
            if (_visited.Count >= _options.MaxPageGraphNodes) {
                throw new OneNoteFormatException(
                    "ONENOTE_PAGE_GRAPH_LIMIT",
                    "The distinct page object-space traversal limit was exceeded.");
            }
            _visited.Add(key);
            return true;
        }

        internal void Exit(string key) => _active.Remove(key);
    }
}
