using System.Diagnostics;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public partial class WordDocument {
        private const int MaxSdtId = int.MaxValue - 1;

        private readonly HashSet<int> _usedSdtIds = new();
        private int _nextSdtId = 1;
        private readonly object _sdtIdLock = new();
        private bool _sdtIdsScanned;

        internal int GenerateSdtId() {
            lock (_sdtIdLock) {
                EnsureExistingSdtIdsScanned();
                if (_usedSdtIds.Count >= MaxSdtId) {
                    throw new InvalidOperationException("All available SDT identifiers are exhausted.");
                }

                var candidate = NormalizeCandidate(_nextSdtId);

                while (_usedSdtIds.Contains(candidate)) {
                    candidate = candidate >= MaxSdtId ? 1 : candidate + 1;
                }

                _usedSdtIds.Add(candidate);
                _nextSdtId = candidate >= MaxSdtId ? 1 : candidate + 1;
                return candidate;
            }
        }

        internal void AssignNewSdtIds(OpenXmlElement element) {
            ArgumentNullException.ThrowIfNull(element);

            foreach (var sdtId in element.Descendants<SdtId>()) {
                sdtId.Val = GenerateSdtId();
            }
        }

        internal void AssignNewSdtIds(OfficeXmlNode root) {
            const int W = OfficeXmlNamespaces.WordprocessingML;
            foreach (OfficeXmlNode properties in root.ContentDescendants(W, "sdtPr")) {
                OfficeXmlNode id = properties.Child(W, "id");
                if (id.Exists) id.SetAttribute(W, "val", OfficeXmlViewValues.Format(GenerateSdtId()));
            }
        }

        private void InitializeSdtIdState() {
            lock (_sdtIdLock) {
                _usedSdtIds.Clear();
                _nextSdtId = 1;
                _sdtIdsScanned = false;
            }
        }

        // Reading a document never needs new identifiers, so the existing ones are collected when the first one is generated.
        private void EnsureExistingSdtIdsScanned() {
            if (_sdtIdsScanned) return;
            foreach (var id in EnumerateExistingSdtIds()) {
                if (id <= 0) {
                    continue;
                }

                _usedSdtIds.Add(id);
                if (id >= _nextSdtId) {
                    _nextSdtId = id >= MaxSdtId ? 1 : id + 1;
                }
            }
            // Marked only after every part was read, so a scan that throws is retried on the next id.
            _sdtIdsScanned = true;
        }

        private IEnumerable<int> EnumerateExistingSdtIds() {
            if (_ownedModel != null) {
                const int W = OfficeXmlNamespaces.WordprocessingML;
                foreach (var story in WordFieldInventory.EnumerateStories(this, includeNotes: false)) {
                    // Keep controls in every preserved branch distinct when another reader selects it.
                    foreach (OfficeXmlNode properties in WordFieldInventory.EnumerateFieldContent(story.Root, reserveInactiveCompatibilityIds: true)) {
                        if (properties.Is(W, "sdtPr") &&
                            OfficeXmlViewValues.ParseInt32(properties.Child(W, "id").Attribute(W, "val")) is int id) yield return id;
                    }
                }
                yield break;
            }

            var mainPart = _wordprocessingDocument?.MainDocumentPart;
            if (mainPart == null) {
                yield break;
            }

            foreach (var id in EnumerateSdtIdsSafe(mainPart.Document)) {
                yield return id;
            }

            foreach (var headerPart in mainPart.HeaderParts ?? Enumerable.Empty<HeaderPart>()) {
                foreach (var id in EnumerateSdtIdsSafe(headerPart?.RootElement)) {
                    yield return id;
                }
            }

            foreach (var footerPart in mainPart.FooterParts ?? Enumerable.Empty<FooterPart>()) {
                foreach (var id in EnumerateSdtIdsSafe(footerPart?.RootElement)) {
                    yield return id;
                }
            }
        }

        private static int NormalizeCandidate(int candidate) {
            if (candidate <= 0 || candidate > MaxSdtId) {
                return 1;
            }

            return candidate;
        }

        private static IEnumerable<int> EnumerateSdtIdsSafe(OpenXmlElement? root) {
            if (root == null) {
                yield break;
            }

            for (OpenXmlElement? element = root.FirstChild; element != null;
                 element = WordElementWalk.Next(element, root, descend: !WordElementWalk.IsFormattingOrText(element))) {
                if (!WordElementWalk.IsCommonContent(element) && element is SdtId { Val.HasValue: true } sdtId) {
                    yield return sdtId.Val.Value;
                }
            }
        }
    }
}
