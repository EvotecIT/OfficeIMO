using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public partial class WordDocument {
        private const int MaxSdtId = int.MaxValue - 1;

        private readonly HashSet<int> _usedSdtIds = new();
        private int _nextSdtId = 1;
        private readonly object _sdtIdLock = new();

        internal int GenerateSdtId() {
            lock (_sdtIdLock) {
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
            if (element == null) {
                throw new ArgumentNullException(nameof(element));
            }

            foreach (var sdtId in element.Descendants<SdtId>()) {
                sdtId.Val = GenerateSdtId();
            }
        }

        private void InitializeSdtIdState() {
            lock (_sdtIdLock) {
                _usedSdtIds.Clear();
                _nextSdtId = 1;

                foreach (var id in EnumerateExistingSdtIds()) {
                    if (id <= 0) {
                        continue;
                    }

                    _usedSdtIds.Add(id);
                    if (id >= _nextSdtId) {
                        _nextSdtId = id >= MaxSdtId ? 1 : id + 1;
                    }
                }
            }
        }

        private IEnumerable<int> EnumerateExistingSdtIds() {
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

            IEnumerable<SdtId> query;
            try {
                query = root.Descendants<SdtId>();
            } catch (InvalidOperationException ex) {
                Debug.WriteLine($"Failed to enumerate SDT identifiers: {ex}");
                yield break;
            }

            foreach (var sdtId in query) {
                if (sdtId?.Val?.HasValue == true) {
                    yield return sdtId.Val.Value;
                }
            }
        }
    }
}
