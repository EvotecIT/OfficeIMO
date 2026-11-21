using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public partial class WordDocument {
        private int _bodySectionSnapshotDepth;
        private BodySectionSnapshot? _bodySectionSnapshot;

        /// <summary>
        /// Shares one split of the body into sections while a document-wide list reads every section.
        /// Each section's list otherwise walks the whole body to find its own part, so joining them costs
        /// sections times body size. Callers only read the document while the scope is open.
        /// </summary>
        internal IDisposable UseBodySectionSnapshot() {
            _bodySectionSnapshotDepth++;
            return new BodySectionSnapshotScope(this);
        }

        /// <summary>Returns the body children of each section when a snapshot scope is open for this body.</summary>
        internal List<List<OpenXmlElement>>? GetBodySectionSnapshot(Body body) {
            if (_bodySectionSnapshotDepth == 0) {
                return null;
            }

            BodySectionSnapshot? snapshot = _bodySectionSnapshot;
            if (snapshot == null || !ReferenceEquals(snapshot.Body, body) || snapshot.SectionCount != Sections.Count) {
                snapshot = new BodySectionSnapshot(body, Sections.Count);
                _bodySectionSnapshot = snapshot;
            }
            return snapshot.ChildrenBySection;
        }

        private sealed class BodySectionSnapshot {
            internal BodySectionSnapshot(Body body, int sectionCount) {
                Body = body;
                SectionCount = sectionCount;
                ChildrenBySection = WordSection.SplitBodyBySection(body, sectionCount);
            }

            internal Body Body { get; }
            internal int SectionCount { get; }
            internal List<List<OpenXmlElement>> ChildrenBySection { get; }
        }

        private sealed class BodySectionSnapshotScope : IDisposable {
            private WordDocument? _document;

            internal BodySectionSnapshotScope(WordDocument document) => _document = document;

            public void Dispose() {
                WordDocument? document = _document;
                if (document == null) {
                    return;
                }

                _document = null;
                if (--document._bodySectionSnapshotDepth == 0) {
                    document._bodySectionSnapshot = null;
                }
            }
        }
    }
}
