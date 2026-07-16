using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word.LegacyDoc.Model;
using System.Text;

namespace OfficeIMO.Word.LegacyDoc.Write {
    internal static partial class LegacyDocWriter {
        private static void AppendSupportedRevisionText(
            StringBuilder text,
            List<LegacyDocWritableRun> runs,
            OpenXmlCompositeElement revisionElement,
            LegacyDocRevisionKind kind,
            LegacyDocWritableFootnotes footnotes,
            LegacyDocWritableEndnotes endnotes,
            LegacyDocWritableFormatting inheritedFormatting,
            LegacyDocWritablePictures? pictures = null,
            OpenXmlPart? ownerPart = null) {
            LegacyDocRevision revision = ReadSupportedRevision(revisionElement, kind);
            LegacyDocWritableFormatting revisionFormatting = inheritedFormatting.WithRevision(revision);
            foreach (OpenXmlElement child in revisionElement.ChildElements) {
                if (child is Run run) {
                    AppendSupportedRunText(
                        text,
                        runs,
                        run,
                        footnotes,
                        endnotes,
                        revisionFormatting,
                        allowHyperlinkRunStyle: false,
                        pictures,
                        ownerPart);
                    continue;
                }

                throw new NotSupportedException($"Native DOC saving supports tracked insertions and deletions only when they contain text runs. Unsupported revision element: {child.LocalName}.");
            }
        }

        private static LegacyDocRevision ReadSupportedRevision(OpenXmlCompositeElement revisionElement, LegacyDocRevisionKind kind) {
            string? author;
            DateTime? date;
            if (revisionElement is InsertedRun insertedRun) {
                author = insertedRun.Author?.Value;
                date = insertedRun.Date?.Value;
            } else if (revisionElement is DeletedRun deletedRun) {
                author = deletedRun.Author?.Value;
                date = deletedRun.Date?.Value;
            } else {
                throw new InvalidOperationException("A DOC revision wrapper must be an insertion or deletion.");
            }

            return new LegacyDocRevision(
                kind,
                string.IsNullOrWhiteSpace(author) ? LegacyDocRevisionAuthorReader.UnknownAuthor : author!,
                date);
        }

        private static void AddRevisionSprms(
            List<byte> grpprl,
            LegacyDocRevision revision,
            IReadOnlyDictionary<string, int>? revisionAuthorIndexes) {
            if (!revision.HasValue) {
                return;
            }

            if (revisionAuthorIndexes == null
                || !revisionAuthorIndexes.TryGetValue(
                    revision.Author ?? LegacyDocRevisionAuthorReader.UnknownAuthor,
                    out int authorIndex)) {
                throw new InvalidOperationException("The DOC revision-author table does not contain a tracked run author.");
            }

            if (authorIndex < short.MinValue || authorIndex > short.MaxValue) {
                throw new NotSupportedException("Native DOC saving supports only revision-author indexes within the Word 97-2003 signed 16-bit range.");
            }

            bool deleted = revision.Kind == LegacyDocRevisionKind.Deleted;
            AddSingleByteSprm(grpprl, deleted ? SprmCFRMarkDel : SprmCFRMarkIns, 1);
            AddUInt16Sprm(grpprl, deleted ? SprmCIbstRMarkDel : SprmCIbstRMark, unchecked((ushort)(short)authorIndex));
            if (revision.Date != null) {
                AddInt32Sprm(
                    grpprl,
                    deleted ? SprmCDttmRMarkDel : SprmCDttmRMark,
                    unchecked((int)CreateDttm(revision.Date.Value)));
            }
        }

        private static uint CreateDttm(DateTime value) {
            if (value.Year < 1900 || value.Year > 2411) {
                throw new NotSupportedException("Native DOC saving supports tracked-revision dates only from 1900 through 2411.");
            }

            return (uint)value.Minute
                | ((uint)value.Hour << 6)
                | ((uint)value.Day << 11)
                | ((uint)value.Month << 16)
                | ((uint)(value.Year - 1900) << 20)
                | ((uint)value.DayOfWeek << 29);
        }

        private static byte[] CreateRevisionAuthorTable(IReadOnlyList<string> authors) {
            if (authors.Count == 0) {
                return Array.Empty<byte>();
            }

            if (authors.Count > ushort.MaxValue) {
                throw new NotSupportedException("Native DOC saving supports only revision-author tables with at most 65,535 names.");
            }

            using var stream = new MemoryStream();
            WriteUInt16(stream, 0xFFFF);
            WriteUInt16(stream, checked((ushort)authors.Count));
            WriteUInt16(stream, 0);
            foreach (string author in authors) {
                if (author.Length > ushort.MaxValue) {
                    throw new NotSupportedException("Native DOC saving cannot write a revision author whose name exceeds 65,535 UTF-16 characters.");
                }

                WriteUInt16(stream, checked((ushort)author.Length));
                byte[] authorBytes = Encoding.Unicode.GetBytes(author);
                stream.Write(authorBytes, 0, authorBytes.Length);
            }

            return stream.ToArray();
        }
    }
}
