using System.Collections.ObjectModel;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint.LegacyPpt.Write {
    internal static partial class LegacyPptWriter {
        internal const ushort RecordProgTags = 0x1388;
        internal const ushort RecordProgBinaryTag = 0x138A;
        internal const ushort RecordBinaryTagDataBlob = 0x138B;
        internal const ushort RecordComment10 = 0x2EE0;
        private const ushort RecordComment10Atom = 0x2EE1;
        private const string Ppt10TagName = "___PPT10";

        internal static bool TryReadAllClassicComments(PowerPointPresentation presentation,
            out IReadOnlyDictionary<string, IReadOnlyList<LegacyPptWriterComment>> commentsBySlide,
            out string? reason) {
            commentsBySlide = new ReadOnlyDictionary<string, IReadOnlyList<LegacyPptWriterComment>>(
                new Dictionary<string, IReadOnlyList<LegacyPptWriterComment>>(StringComparer.Ordinal));
            reason = null;
            PresentationPart? presentationPart = presentation.OpenXmlDocument.PresentationPart;
            if (presentationPart == null) {
                reason = "The Open XML presentation part is missing.";
                return false;
            }

            if (!TryReadClassicAuthors(presentationPart, out Dictionary<uint, LegacyPptWriterAuthor> authors,
                    out reason)) {
                return false;
            }

            var usedAuthorIds = new HashSet<uint>();
            var maximumIndexes = new Dictionary<uint, uint>();
            var result = new Dictionary<string, IReadOnlyList<LegacyPptWriterComment>>(StringComparer.Ordinal);
            foreach (PowerPointSlide slide in presentation.Slides) {
                if (!TryReadClassicComments(slide, authors, usedAuthorIds, maximumIndexes,
                        out IReadOnlyList<LegacyPptWriterComment> comments, out reason)) {
                    return false;
                }
                result.Add(slide.SlidePart.Uri.ToString(), comments);
            }

            foreach (LegacyPptWriterAuthor author in authors.Values) {
                if (!usedAuthorIds.Contains(author.Id)) {
                    reason = $"Classic comment author {author.Id} is unused; the binary format has no standalone author directory.";
                    return false;
                }
                uint maximumIndex = maximumIndexes[author.Id];
                if (author.LastIndex != maximumIndex) {
                    reason = $"Classic comment author {author.Id} has lastIdx={author.LastIndex}, but its greatest comment index is {maximumIndex}.";
                    return false;
                }
            }

            commentsBySlide = new ReadOnlyDictionary<string, IReadOnlyList<LegacyPptWriterComment>>(result);
            return true;
        }

        internal static bool HasModernComments(PowerPointPresentation presentation) {
            PresentationPart? presentationPart = presentation.OpenXmlDocument.PresentationPart;
            if (presentationPart == null) return false;
            if (presentationPart.Parts.Any(pair =>
                    pair.OpenXmlPart is PowerPointAuthorsPart)) {
                return true;
            }
            return presentation.Slides.Any(slide => slide.SlidePart.Parts.Any(pair =>
                pair.OpenXmlPart is PowerPointCommentPart));
        }

        internal static IReadOnlyList<LegacyPptWriterComment> ReadClassicCommentsForSlide(
            PowerPointSlide slide) {
            PresentationPart presentationPart = slide.SlidePart.GetParentParts()
                .OfType<PresentationPart>().Single();
            if (!TryReadClassicAuthors(presentationPart,
                    out Dictionary<uint, LegacyPptWriterAuthor> authors, out string? reason)) {
                throw new NotSupportedException(reason);
            }
            if (!TryReadClassicComments(slide, authors, new HashSet<uint>(),
                    new Dictionary<uint, uint>(), out IReadOnlyList<LegacyPptWriterComment> comments,
                    out reason)) {
                throw new NotSupportedException(reason);
            }
            return comments;
        }

        private static bool TryReadClassicAuthors(PresentationPart presentationPart,
            out Dictionary<uint, LegacyPptWriterAuthor> authors, out string? reason) {
            authors = new Dictionary<uint, LegacyPptWriterAuthor>();
            reason = null;
            CommentAuthorsPart? authorsPart = presentationPart.CommentAuthorsPart;
            if (authorsPart == null) return true;
            P.CommentAuthorList? list = authorsPart.CommentAuthorList;
            if (list == null || list.ChildElements.Any(child => child is not P.CommentAuthor)) {
                reason = "The classic comment author list is missing or contains unsupported extension content.";
                return false;
            }
            var identities = new HashSet<string>(StringComparer.Ordinal);
            foreach (P.CommentAuthor author in list.Elements<P.CommentAuthor>()) {
                if (author.Id?.HasValue != true || author.LastIndex?.HasValue != true
                    || author.ColorIndex?.HasValue != true) {
                    reason = "Every classic comment author must define id, lastIdx, and clrIdx.";
                    return false;
                }
                uint id = author.Id!.Value;
                string name = author.Name?.Value ?? string.Empty;
                string initials = author.Initials?.Value ?? string.Empty;
                if (name.Length > 52 || initials.Length > 52 || name.IndexOf('\0') >= 0
                    || initials.IndexOf('\0') >= 0) {
                    reason = $"Classic comment author {id} exceeds the binary 52-character author or initials limit.";
                    return false;
                }
                if (author.ColorIndex!.Value != id) {
                    reason = $"Classic comment author {id} uses color index {author.ColorIndex.Value}; binary comments can retain only the canonical color index equal to the author id.";
                    return false;
                }
                string identity = name + "\0" + initials;
                if (!identities.Add(identity)) {
                    reason = "Two classic comment author entries have the same name and initials and cannot be distinguished in binary PowerPoint.";
                    return false;
                }
                if (authors.ContainsKey(id)) {
                    reason = $"Classic comment author id {id} is duplicated.";
                    return false;
                }
                authors.Add(id, new LegacyPptWriterAuthor(id, name, initials,
                    author.LastIndex!.Value));
            }
            return true;
        }

        private static bool TryReadClassicComments(PowerPointSlide slide,
            IReadOnlyDictionary<uint, LegacyPptWriterAuthor> authors, ISet<uint> usedAuthorIds,
            IDictionary<uint, uint> maximumIndexes,
            out IReadOnlyList<LegacyPptWriterComment> comments, out string? reason) {
            var result = new List<LegacyPptWriterComment>();
            comments = result;
            reason = null;
            SlideCommentsPart? commentsPart = slide.SlidePart.SlideCommentsPart;
            if (commentsPart == null) return true;
            P.CommentList? list = commentsPart.CommentList;
            if (list == null || list.ChildElements.Any(child => child is not P.Comment)) {
                reason = "A classic slide comment list is missing or contains unsupported extension content.";
                return false;
            }
            foreach (P.Comment comment in list.Elements<P.Comment>()) {
                if (comment.ChildElements.Any(child => child is not P.Position && child is not P.Text)
                    || comment.AuthorId?.HasValue != true || comment.Index?.HasValue != true
                    || comment.Position?.X?.HasValue != true || comment.Position.Y?.HasValue != true) {
                    reason = "Every classic comment must contain only a position and text and define authorId, idx, x, and y.";
                    return false;
                }
                uint authorId = comment.AuthorId!.Value;
                if (!authors.TryGetValue(authorId, out LegacyPptWriterAuthor? author)) {
                    reason = $"Classic comment author id {authorId} is missing from the author list.";
                    return false;
                }
                uint index = comment.Index!.Value;
                if (index > int.MaxValue) {
                    reason = $"Classic comment index {index} exceeds the binary signed 32-bit limit.";
                    return false;
                }
                string text = comment.Text?.Text ?? string.Empty;
                if (text.Length > 32000 || text.IndexOf('\0') >= 0) {
                    reason = $"Classic comment {index} exceeds the binary 32,000-character text limit.";
                    return false;
                }
                long x = comment.Position.X!.Value;
                long y = comment.Position.Y!.Value;
                if (x < int.MinValue || x > int.MaxValue || y < int.MinValue
                    || y > int.MaxValue) {
                    reason = $"Classic comment {index} has a position outside the binary signed 32-bit coordinate range.";
                    return false;
                }
                DateTime? createdAtUtc = NormalizeCommentDate(comment.DateTime?.Value);
                result.Add(new LegacyPptWriterComment(checked((int)index), author.Name,
                    author.Initials, text, createdAtUtc, checked((int)x), checked((int)y)));
                usedAuthorIds.Add(authorId);
                maximumIndexes[authorId] = maximumIndexes.TryGetValue(authorId,
                    out uint currentMaximum) ? Math.Max(currentMaximum, index) : index;
            }
            comments = new ReadOnlyCollection<LegacyPptWriterComment>(result);
            return true;
        }

        private static DateTime? NormalizeCommentDate(DateTime? value) {
            if (!value.HasValue) return null;
            DateTime date = value.Value;
            if (date.Kind == DateTimeKind.Local) return date.ToUniversalTime();
            return DateTime.SpecifyKind(date, DateTimeKind.Utc);
        }

        internal static byte[] BuildCommentProgrammableTagsRecord(
            IReadOnlyList<LegacyPptWriterComment> comments) =>
            BuildContainer(RecordProgTags, instance: 0,
                new[] { BuildPpt10BinaryTagRecord(comments) });

        internal static byte[] BuildPpt10BinaryTagRecord(
            IReadOnlyList<LegacyPptWriterComment> comments) {
            byte[] tagName = BuildRecord(version: 0, instance: 0, RecordCString,
                Encoding.Unicode.GetBytes(Ppt10TagName));
            byte[] data = BuildRecord(version: 0, instance: 0, RecordBinaryTagDataBlob,
                Concat(BuildCommentRecords(comments)));
            return BuildContainer(RecordProgBinaryTag, instance: 0, new[] { tagName, data });
        }

        internal static IReadOnlyList<byte[]> BuildCommentRecords(
            IReadOnlyList<LegacyPptWriterComment> comments) =>
            comments.Select(BuildCommentRecord).ToArray();

        private static byte[] BuildCommentRecord(LegacyPptWriterComment comment) {
            var children = new List<byte[]>(4);
            if (comment.Author.Length > 0) {
                children.Add(BuildRecord(version: 0, instance: 0, RecordCString,
                    Encoding.Unicode.GetBytes(comment.Author)));
            }
            if (comment.Text.Length > 0) {
                children.Add(BuildRecord(version: 0, instance: 1, RecordCString,
                    Encoding.Unicode.GetBytes(comment.Text)));
            }
            if (comment.Initials.Length > 0) {
                children.Add(BuildRecord(version: 0, instance: 2, RecordCString,
                    Encoding.Unicode.GetBytes(comment.Initials)));
            }
            var payload = new byte[28];
            WriteInt32(payload, 0, comment.Index);
            if (comment.CreatedAtUtc.HasValue) {
                DateTime date = comment.CreatedAtUtc.Value;
                WriteUInt16(payload, 4, checked((ushort)date.Year));
                WriteUInt16(payload, 6, checked((ushort)date.Month));
                WriteUInt16(payload, 8, checked((ushort)date.DayOfWeek));
                WriteUInt16(payload, 10, checked((ushort)date.Day));
                WriteUInt16(payload, 12, checked((ushort)date.Hour));
                WriteUInt16(payload, 14, checked((ushort)date.Minute));
                WriteUInt16(payload, 16, checked((ushort)date.Second));
                WriteUInt16(payload, 18, checked((ushort)date.Millisecond));
            }
            WriteInt32(payload, 20, comment.X);
            WriteInt32(payload, 24, comment.Y);
            children.Add(BuildRecord(version: 0, instance: 0, RecordComment10Atom, payload));
            return BuildContainer(RecordComment10, instance: 0, children);
        }

        internal sealed class LegacyPptWriterComment {
            internal LegacyPptWriterComment(int index, string author, string initials, string text,
                DateTime? createdAtUtc, int x, int y) {
                Index = index;
                Author = author ?? string.Empty;
                Initials = initials ?? string.Empty;
                Text = text ?? string.Empty;
                CreatedAtUtc = createdAtUtc;
                X = x;
                Y = y;
            }

            internal int Index { get; }
            internal string Author { get; }
            internal string Initials { get; }
            internal string Text { get; }
            internal DateTime? CreatedAtUtc { get; }
            internal int X { get; }
            internal int Y { get; }
        }

        private sealed class LegacyPptWriterAuthor {
            internal LegacyPptWriterAuthor(uint id, string name, string initials, uint lastIndex) {
                Id = id;
                Name = name;
                Initials = initials;
                LastIndex = lastIndex;
            }

            internal uint Id { get; }
            internal string Name { get; }
            internal string Initials { get; }
            internal uint LastIndex { get; }
        }
    }
}
