using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public partial class WordDocument {
        private readonly HashSet<int> _usedSdtIds = new();
        private int _nextSdtId = 1;
        private readonly object _sdtIdLock = new();

        internal int GenerateSdtId() {
            lock (_sdtIdLock) {
                var candidate = _nextSdtId;
                if (candidate <= 0) {
                    candidate = 1;
                }

                while (_usedSdtIds.Contains(candidate)) {
                    candidate++;
                }

                _usedSdtIds.Add(candidate);
                _nextSdtId = candidate + 1;
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
                        _nextSdtId = id + 1;
                    }
                }
            }
        }

        private IEnumerable<int> EnumerateExistingSdtIds() {
            if (_wordprocessingDocument?.MainDocumentPart == null) {
                yield break;
            }

            foreach (var sdtId in _wordprocessingDocument.MainDocumentPart.Document.Descendants<SdtId>()) {
                if (sdtId.Val?.HasValue == true) {
                    yield return sdtId.Val.Value;
                }
            }

            foreach (var headerPart in _wordprocessingDocument.MainDocumentPart.HeaderParts) {
                foreach (var sdtId in headerPart.RootElement?.Descendants<SdtId>() ?? Enumerable.Empty<SdtId>()) {
                    if (sdtId.Val?.HasValue == true) {
                        yield return sdtId.Val.Value;
                    }
                }
            }

            foreach (var footerPart in _wordprocessingDocument.MainDocumentPart.FooterParts) {
                foreach (var sdtId in footerPart.RootElement?.Descendants<SdtId>() ?? Enumerable.Empty<SdtId>()) {
                    if (sdtId.Val?.HasValue == true) {
                        yield return sdtId.Val.Value;
                    }
                }
            }
        }
    }
}
