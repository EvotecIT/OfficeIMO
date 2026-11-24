using System.Collections;
using OfficeIMO.Html;

namespace OfficeIMO.Epub;

public sealed partial class EpubPublication {
    // Validates edited content documents against the publication as it will be after the edit, so the
    // cost depends on the edited documents and what they reference, not on the size of the book. Write
    // and Save still validate the whole publication. An edit that removes identifiers or rewrites the
    // navigation document can break links held by other documents, so those edits re-check the whole book.
    private void ValidateContentEdits(IReadOnlyList<(string Path, EpubManifestItem Item, XDocument Document)> edits,
        IReadOnlyDictionary<string, byte[]> payloads, CancellationToken token) {
        var entries = new EntryOverlay(_entries, payloads);
        string navPath = NavigationPath(Root, entries);
        var anchors = new Dictionary<string, HashSet<string>>(StringComparer.Ordinal);
        bool wholeBook = false;
        foreach (var edit in edits) {
            token.ThrowIfCancellationRequested();
            HashSet<string> after = EpubContentIdentifiers.Collect(edit.Document.Root!, edit.Path, false, token);
            anchors[edit.Path] = after;
            if (edit.Path == navPath) { wholeBook = true; continue; }
            HashSet<string> before = EpubContentIdentifiers.Collect(ParseXml(_entries[edit.Path], _maximumEntryBytes).Root!, edit.Path, false, token);
            if (!before.IsSubsetOf(after)) wholeBook = true;
        }
        if (wholeBook) {
            ValidatePublication(_package, entries, new List<OfficeDiagnostic>(), token, changed: true, updateContentProperties: false);
            return;
        }

        ManifestLookup manifest = LiveManifest;
        var stylesheetResults = new Dictionary<string, bool>(StringComparer.Ordinal);
        bool CheckStylesheet(string path) {
            if (!stylesheetResults.TryGetValue(path, out bool remote)) stylesheetResults[path] = remote = ValidateStylesheetClosure(path, entries, manifest, token);
            return remote;
        }
        foreach (var edit in edits) {
            token.ThrowIfCancellationRequested();
            XDocument content = edit.Document;
            IReadOnlyList<EpubManifestItem> declarations = manifest.AtPath(edit.Path);
            if (declarations.Any(item => (HasToken(item.Properties, "scripted") || IsScriptMediaType(item.MediaType) ||
                HasToken(OriginalManifestItem(item.Id)?.Properties, "scripted") || IsScriptMediaType(OriginalManifestItem(item.Id)?.MediaType ?? string.Empty)) &&
                !IsRetainedScriptDeclaration(item, entries)))
                throw new NotSupportedException("Script authoring is outside the EPUB writer contract.");
            // Like Write, an unchanged imported payload under its original media type isn't analyzed again.
            bool payloadChanged = !_originalEntries.TryGetValue(edit.Path, out byte[]? original) || !original.SequenceEqual(payloads[edit.Path]);
            foreach (string mediaType in declarations
                .Where(item => payloadChanged || !HasMediaType(OriginalManifestItem(item.Id)?.MediaType, item.MediaType)).Select(item => item.MediaType)
                .Where(type => !HasMediaType(type, "application/x-dtbncx+xml")).Distinct(StringComparer.OrdinalIgnoreCase)) {
                ValidateContent(content, mediaType);
                string analysisContent = content.ToString(SaveOptions.DisableFormatting);
                var limits = HtmlConversionLimits.UntrustedDefaults;
                limits.MaxInputCharacters = (int)Math.Min(Math.Max(_maximumEntryBytes, analysisContent.Length), int.MaxValue);
                HtmlResourceManifest resources = HtmlResourcePipeline.BuildManifest(
                    HtmlXmlDocumentParser.CreateDocument(content, limits, token, skipProcessingInstructions: true),
                    new HtmlResourcePipelineOptions { Limits = limits });
                if (resources.Resources.Any(resource => !resource.IsAllowed)) throw new NotSupportedException("Authored content contains a URL blocked by the shared HTML policy.");
                ValidateAuthoredResources(content, edit.Path, resources, manifest, token, CheckStylesheet);
            }
            // Fragment targets are checked against the identifiers of only the documents this edit links to.
            foreach (EpubReference reference in ContentResourceReferences(content, edit.Path, token)) {
                string? target = reference.ContainerPath;
                if (reference.Kind != EpubReferenceKind.Container || target == null || reference.Fragment == null ||
                    anchors.ContainsKey(target) || !entries.ContainsKey(target)) continue;
                if (_encryption.Any(encryption => encryption.Path == target && encryption.RequiresDecryption)) continue;
                if (!manifest.AtPath(target).Any(item =>
                    (HasMediaType(item.MediaType, "application/xhtml+xml") || HasMediaType(item.MediaType, "image/svg+xml") ||
                    HasMediaType(item.MediaType, "application/x-dtbncx+xml")))) continue;
                anchors[target] = EpubContentIdentifiers.Collect(ParseXml(entries[target], _maximumEntryBytes).Root!, target, false, token);
            }
            ValidateContentReferences(content, edit.Path, entries, anchors, token);
        }
    }

    // Read-only view of the retained entries with staged payloads layered over them.
    private sealed class EntryOverlay : IReadOnlyDictionary<string, byte[]> {
        private readonly IReadOnlyDictionary<string, byte[]> _entries;
        private readonly IReadOnlyDictionary<string, byte[]> _staged;

        internal EntryOverlay(IReadOnlyDictionary<string, byte[]> entries, IReadOnlyDictionary<string, byte[]> staged) {
            _entries = entries;
            _staged = staged;
        }

        public byte[] this[string key] => TryGetValue(key, out byte[]? value) ? value : throw new KeyNotFoundException(key);
        public IEnumerable<string> Keys => _entries.Keys.Concat(_staged.Keys.Where(key => !_entries.ContainsKey(key)));
        public IEnumerable<byte[]> Values => Keys.Select(key => this[key]);
        public int Count => _entries.Count + _staged.Keys.Count(key => !_entries.ContainsKey(key));
        public bool ContainsKey(string key) => _staged.ContainsKey(key) || _entries.ContainsKey(key);

        public bool TryGetValue(string key, out byte[] value) {
            if (_staged.TryGetValue(key, out byte[]? staged)) { value = staged; return true; }
            if (_entries.TryGetValue(key, out byte[]? retained)) { value = retained; return true; }
            value = Array.Empty<byte>();
            return false;
        }

        public IEnumerator<KeyValuePair<string, byte[]>> GetEnumerator() =>
            Keys.Select(key => new KeyValuePair<string, byte[]>(key, this[key])).GetEnumerator();

        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }
}
