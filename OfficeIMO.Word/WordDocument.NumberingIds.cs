using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public partial class WordDocument {
        private NumberingIdCache? _numberingIdCache;
        private int _numberingIdCacheDepth;

        /// <summary>
        /// Remembers the next free list numbering ids while a converter adds many lists. Without it,
        /// every new list takes the maximum over all numbering definitions, so generation is quadratic
        /// in the number of lists. The remembered ids are used only while the numbering root and its last
        /// child are the ones seen at the previous allocation, so a definition appended by other code
        /// makes the next allocation rescan. Callers own the document for the whole scope and must not
        /// renumber or insert definitions in the middle of the numbering part in the meantime.
        /// </summary>
        internal IDisposable UseNumberingIdCache() {
            _numberingIdCacheDepth++;
            return new NumberingIdCacheScope(this);
        }

        internal void GetNextNumberingIds(Numbering numbering, out int abstractId, out int numberId) {
            NumberingIdCache? cache = _numberingIdCache;
            if (_numberingIdCacheDepth > 0 &&
                cache != null &&
                ReferenceEquals(cache.Numbering, numbering) &&
                ReferenceEquals(cache.LastChild, numbering.LastChild)) {
                abstractId = cache.NextAbstractId;
                numberId = cache.NextNumberId;
                return;
            }

            int? maxAbstractId = null;
            int? maxNumberId = null;
            foreach (OpenXmlElement element in numbering.ChildElements) {
                if (element is AbstractNum abstractNum) {
                    int id = (int)(abstractNum.AbstractNumberId?.Value ?? 0);
                    maxAbstractId = maxAbstractId.HasValue ? Math.Max(maxAbstractId.Value, id) : id;
                } else if (element is NumberingInstance instance) {
                    int id = (int)(instance.NumberID?.Value ?? 0);
                    maxNumberId = maxNumberId.HasValue ? Math.Max(maxNumberId.Value, id) : id;
                }
            }

            abstractId = maxAbstractId.HasValue ? maxAbstractId.Value + 1 : 0;
            numberId = maxNumberId.HasValue ? maxNumberId.Value + 1 : 1;
            if (_numberingIdCacheDepth > 0) {
                _numberingIdCache = new NumberingIdCache(numbering, numbering.LastChild, abstractId, numberId);
            }
        }

        /// <summary>Records definitions appended with ids from <see cref="GetNextNumberingIds"/>.</summary>
        internal void NumberingIdsUsed(Numbering numbering, int abstractId, int numberId) {
            NumberingIdCache? cache = _numberingIdCache;
            if (_numberingIdCacheDepth == 0 || cache == null || !ReferenceEquals(cache.Numbering, numbering)) {
                return;
            }

            _numberingIdCache = new NumberingIdCache(
                numbering,
                numbering.LastChild,
                Math.Max(cache.NextAbstractId, abstractId + 1),
                Math.Max(cache.NextNumberId, numberId + 1));
        }

        private sealed class NumberingIdCache {
            internal NumberingIdCache(Numbering numbering, OpenXmlElement? lastChild, int nextAbstractId, int nextNumberId) {
                Numbering = numbering;
                LastChild = lastChild;
                NextAbstractId = nextAbstractId;
                NextNumberId = nextNumberId;
            }

            internal Numbering Numbering { get; }
            internal OpenXmlElement? LastChild { get; }
            internal int NextAbstractId { get; }
            internal int NextNumberId { get; }
        }

        private sealed class NumberingIdCacheScope : IDisposable {
            private WordDocument? _document;

            internal NumberingIdCacheScope(WordDocument document) => _document = document;

            public void Dispose() {
                WordDocument? document = _document;
                if (document == null) {
                    return;
                }

                _document = null;
                if (--document._numberingIdCacheDepth == 0) {
                    document._numberingIdCache = null;
                }
            }
        }
    }
}
