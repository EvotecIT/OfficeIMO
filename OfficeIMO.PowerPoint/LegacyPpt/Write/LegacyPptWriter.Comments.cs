using System.Collections.ObjectModel;
using System.Text;
using System.Xml;
using OfficeIMO.OpenXml.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;

namespace OfficeIMO.PowerPoint.LegacyPpt.Write {
    internal static partial class LegacyPptWriter {
        internal const ushort RecordProgTags = 0x1388;
        internal const ushort RecordProgBinaryTag = 0x138A;
        internal const ushort RecordBinaryTagDataBlob = 0x138B;
        internal const ushort RecordComment10 = 0x2EE0;
        private const ushort RecordComment10Atom = 0x2EE1;
        private const string Ppt10TagName = "___PPT10";

        private const string ModernCommentAuthorsRelationshipType =
            "http://schemas.microsoft.com/office/2018/10/relationships/authors";
        private const string ModernCommentsRelationshipType =
            "http://schemas.microsoft.com/office/2018/10/relationships/comments";

        // Callers in LegacyPptWritePreflight.cs and LegacyPptPreservingWriter.ChangePlan.cs not yet passing a view.
        internal static bool TryReadAllClassicComments(PowerPointDocument presentation,
            out IReadOnlyDictionary<string, IReadOnlyList<LegacyPptWriterComment>> commentsBySlide,
            out string? reason) {
            ArgumentNullException.ThrowIfNull(presentation);
            using LegacyPptPackageView view = LegacyPptPackageView.Capture(presentation);
            return TryReadAllClassicComments(view, presentation, out commentsBySlide, out reason);
        }

        internal static bool TryReadAllClassicComments(LegacyPptPackageView view,
            PowerPointDocument presentation,
            out IReadOnlyDictionary<string, IReadOnlyList<LegacyPptWriterComment>> commentsBySlide,
            out string? reason) {
            commentsBySlide = new ReadOnlyDictionary<string, IReadOnlyList<LegacyPptWriterComment>>(
                new Dictionary<string, IReadOnlyList<LegacyPptWriterComment>>(StringComparer.Ordinal));
            reason = null;
            LegacyPptPartView presentationPart = view.Presentation;
            if (presentationPart.IsNone) {
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
                LegacyPptPartView slidePart = SlidePartView(view, slide);
                if (!TryReadClassicComments(slidePart, authors, usedAuthorIds, maximumIndexes,
                        out IReadOnlyList<LegacyPptWriterComment> comments, out reason)) {
                    return false;
                }
                result.Add(slide.PartRef.Name, comments);
            }

            foreach (LegacyPptWriterAuthor author in authors.Values) {
                if (!usedAuthorIds.Contains(author.Id)) {
                    reason = $"Classic comment author {author.Id} is unused; the binary format has no standalone author directory.";
                    return false;
                }
                uint maximumIndex = maximumIndexes[author.Id];
                if (!HasCanonicalClassicAuthorLastIndex(author.Id,
                        author.LastIndex, maximumIndexes)) {
                    reason = $"Classic comment author {author.Id} has lastIdx={author.LastIndex}, but its greatest comment index is {maximumIndex}.";
                    return false;
                }
            }

            commentsBySlide = new ReadOnlyDictionary<string, IReadOnlyList<LegacyPptWriterComment>>(result);
            return true;
        }

        // Callers in LegacyPptWritePreflight.cs and LegacyPptPreservingWriter.ChangePlan.cs not yet passing a view.
        internal static bool HasModernComments(PowerPointDocument presentation) {
            ArgumentNullException.ThrowIfNull(presentation);
            using LegacyPptPackageView view = LegacyPptPackageView.Capture(presentation);
            return HasModernComments(view, presentation);
        }

        internal static bool HasModernComments(LegacyPptPackageView view,
            PowerPointDocument presentation) {
            LegacyPptPartView presentationPart = view.Presentation;
            if (presentationPart.IsNone) return false;
            if (HasInternalRelationship(presentationPart, ModernCommentAuthorsRelationshipType)) {
                return true;
            }
            return presentation.Slides.Any(slide => HasInternalRelationship(
                SlidePartView(view, slide), ModernCommentsRelationshipType));
        }

        private static bool HasInternalRelationship(LegacyPptPartView part, string relationshipType) =>
            part.Relationships.Any(relationship => !relationship.IsExternal
                && !relationship.TargetPart.IsNone && relationship.IsType(relationshipType));

        internal static IReadOnlyList<LegacyPptWriterComment> ReadClassicCommentsForSlide(
            LegacyPptPackageView view, PowerPointSlide slide) {
            if (!TryReadClassicAuthors(view.Presentation,
                    out Dictionary<uint, LegacyPptWriterAuthor> authors, out string? reason)) {
                throw new NotSupportedException(reason);
            }
            if (!TryReadClassicComments(SlidePartView(view, slide), authors, new HashSet<uint>(),
                    new Dictionary<uint, uint>(), out IReadOnlyList<LegacyPptWriterComment> comments,
                    out reason)) {
                throw new NotSupportedException(reason);
            }
            return comments;
        }

        private static bool TryReadClassicAuthors(LegacyPptPartView presentationPart,
            out Dictionary<uint, LegacyPptWriterAuthor> authors, out string? reason) {
            authors = new Dictionary<uint, LegacyPptWriterAuthor>();
            reason = null;
            LegacyPptPartView authorsPart = presentationPart.RelatedPart(
                LegacyPptPackageView.CommentAuthorsRelationshipType);
            if (authorsPart.IsNone) return true;
            OfficeXmlNode list = authorsPart.Root;
            bool foreignChild = false;
            foreach (OfficeXmlNode child in list.Children()) {
                if (!child.Is(Pml, "cmAuthor")) foreignChild = true;
            }
            if (!list.Is(Pml, "cmAuthorLst") || foreignChild) {
                reason = "The classic comment author list is missing or contains unsupported extension content.";
                return false;
            }
            var identities = new HashSet<string>(StringComparer.Ordinal);
            foreach (OfficeXmlNode author in list.Children(Pml, "cmAuthor")) {
                uint? authorId = author.UInt("id");
                uint? lastIndex = author.UInt("lastIdx");
                uint? colorIndex = author.UInt("clrIdx");
                if (!authorId.HasValue || !lastIndex.HasValue || !colorIndex.HasValue) {
                    reason = "Every classic comment author must define id, lastIdx, and clrIdx.";
                    return false;
                }
                uint id = authorId.Value;
                if (!TryGetBinaryCompatibleClassicAuthorIdentity(author,
                        out string name, out string initials,
                        out string identity)) {
                    reason = $"Classic comment author {id} has a blank or binary-incompatible name or initials.";
                    return false;
                }
                if (!HasCanonicalClassicAuthorColorIndex(author)) {
                    reason = $"Classic comment author {id} uses color index {colorIndex.Value}; binary comments can retain only the canonical color index equal to the author id.";
                    return false;
                }
                if (!identities.Add(identity)) {
                    reason = "Two classic comment author entries have the same name and initials and cannot be distinguished in binary PowerPoint.";
                    return false;
                }
                if (authors.ContainsKey(id)) {
                    reason = $"Classic comment author id {id} is duplicated.";
                    return false;
                }
                authors.Add(id, new LegacyPptWriterAuthor(id, name, initials,
                    lastIndex.Value));
            }
            return true;
        }

        /// <summary>Whether a classic comment author (<c>p:cmAuthor</c>) uses the color index equal to its id.</summary>
        internal static bool HasCanonicalClassicAuthorColorIndex(
            OfficeXmlNode author) {
            uint? id = author.UInt("id");
            uint? colorIndex = author.UInt("clrIdx");
            return id.HasValue && colorIndex.HasValue && colorIndex.Value == id.Value;
        }

        /// <summary>
        /// The name, initials and identity of a classic comment author (<c>p:cmAuthor</c>) when binary comments can hold
        /// them.
        /// </summary>
        internal static bool TryGetBinaryCompatibleClassicAuthorIdentity(
            OfficeXmlNode author,
            out string name,
            out string initials,
            out string identity) {
            name = author.Attribute("name") ?? string.Empty;
            string? initialsValue = author.Attribute("initials");
            initials = initialsValue ?? string.Empty;
            identity = string.Empty;
            if (string.IsNullOrWhiteSpace(name)
                || initialsValue == null
                || name.Length > 52 || initials.Length > 52
                || name.IndexOf('\0') >= 0 || initials.IndexOf('\0') >= 0) {
                return false;
            }
            identity = name + "\0" + initials;
            return true;
        }

        internal static bool HasCanonicalClassicAuthorLastIndex(
            uint authorId, uint lastIndex,
            IReadOnlyDictionary<uint, uint> maximumIndexes) =>
            maximumIndexes.TryGetValue(authorId, out uint maximumIndex)
            && lastIndex == maximumIndex;

        private static bool TryReadClassicComments(LegacyPptPartView slidePart,
            IReadOnlyDictionary<uint, LegacyPptWriterAuthor> authors, ISet<uint> usedAuthorIds,
            IDictionary<uint, uint> maximumIndexes,
            out IReadOnlyList<LegacyPptWriterComment> comments, out string? reason) {
            var result = new List<LegacyPptWriterComment>();
            comments = result;
            reason = null;
            LegacyPptPartView commentsPart = slidePart.RelatedPart(
                LegacyPptPackageView.CommentsRelationshipType);
            if (commentsPart.IsNone) return true;
            OfficeXmlNode list = commentsPart.Root;
            bool foreignChild = false;
            foreach (OfficeXmlNode child in list.Children()) {
                if (!child.Is(Pml, "cm")) foreignChild = true;
            }
            if (!list.Is(Pml, "cmLst") || foreignChild) {
                reason = "A classic slide comment list is missing or contains unsupported extension content.";
                return false;
            }
            foreach (OfficeXmlNode comment in list.Children(Pml, "cm")) {
                bool foreignContent = false;
                foreach (OfficeXmlNode child in comment.Children()) {
                    if (!child.Is(Pml, "pos") && !child.Is(Pml, "text")) foreignContent = true;
                }
                OfficeXmlNode position = comment.Child(Pml, "pos");
                uint? commentAuthorId = comment.UInt("authorId");
                uint? commentIndex = comment.UInt("idx");
                long? positionX = position.Long("x");
                long? positionY = position.Long("y");
                if (foreignContent
                    || !commentAuthorId.HasValue || !commentIndex.HasValue
                    || !positionX.HasValue || !positionY.HasValue) {
                    reason = "Every classic comment must contain only a position and text and define authorId, idx, x, and y.";
                    return false;
                }
                uint authorId = commentAuthorId.Value;
                if (!authors.TryGetValue(authorId, out LegacyPptWriterAuthor? author)) {
                    reason = $"Classic comment author id {authorId} is missing from the author list.";
                    return false;
                }
                uint index = commentIndex.Value;
                if (index > int.MaxValue) {
                    reason = $"Classic comment index {index} exceeds the binary signed 32-bit limit.";
                    return false;
                }
                OfficeXmlNode textElement = comment.Child(Pml, "text");
                string text = textElement.IsNone ? string.Empty : textElement.InnerText;
                if (text.Length > 32000 || text.IndexOf('\0') >= 0) {
                    reason = $"Classic comment {index} exceeds the binary 32,000-character text limit.";
                    return false;
                }
                long x = positionX.Value;
                long y = positionY.Value;
                if (x < int.MinValue || x > int.MaxValue || y < int.MinValue
                    || y > int.MaxValue) {
                    reason = $"Classic comment {index} has a position outside the binary signed 32-bit coordinate range.";
                    return false;
                }
                string? dateText = comment.Attribute("dt");
                DateTime? createdAtUtc = NormalizeCommentDate(dateText == null
                    ? null
                    : XmlConvert.ToDateTime(dateText, XmlDateTimeSerializationMode.RoundtripKind));
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
