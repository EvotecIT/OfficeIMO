using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace OfficeIMO.Visio {
    public partial class VisioDocument {
        private readonly struct CommentAuthorKey : IEquatable<CommentAuthorKey> {
            public CommentAuthorKey(string? name, string? initials, string? resolutionId) {
                Name = string.IsNullOrWhiteSpace(name) ? "OfficeIMO" : name!;
                Initials = string.IsNullOrWhiteSpace(initials) ? "OI" : initials!;
                ResolutionId = resolutionId ?? string.Empty;
            }

            public string Name { get; }

            public string Initials { get; }

            public string ResolutionId { get; }

            public bool Equals(CommentAuthorKey other) {
                return string.Equals(Name, other.Name, StringComparison.Ordinal) &&
                       string.Equals(Initials, other.Initials, StringComparison.Ordinal) &&
                       string.Equals(ResolutionId, other.ResolutionId, StringComparison.Ordinal);
            }

            public override bool Equals(object? obj) => obj is CommentAuthorKey other && Equals(other);

            public override int GetHashCode() {
                int hash = StringComparer.Ordinal.GetHashCode(Name);
                hash = (hash * 397) ^ StringComparer.Ordinal.GetHashCode(Initials);
                hash = (hash * 397) ^ StringComparer.Ordinal.GetHashCode(ResolutionId);
                return hash;
            }
        }

        private static void WriteCommentsPart(
            PackagePart commentsPart,
            IReadOnlyList<VisioPage> pages,
            IReadOnlyDictionary<VisioPage, Dictionary<string, VisioMaster>> effectivePageMasters) {
            string serializedComments = BuildCommentsXmlForSave(pages, effectivePageMasters);

            using Stream stream = commentsPart.GetStream(FileMode.Create, FileAccess.Write);
            using StreamWriter writer = new(stream, new UTF8Encoding(false));
            writer.Write(serializedComments);
        }

        private static void ValidateCommentsForSave(
            IReadOnlyList<VisioPage> pages,
            IReadOnlyDictionary<VisioPage, Dictionary<string, VisioMaster>> effectivePageMasters) {
            _ = BuildCommentsXmlForSave(pages, effectivePageMasters);
        }

        private static string BuildCommentsXmlForSave(
            IReadOnlyList<VisioPage> pages,
            IReadOnlyDictionary<VisioPage, Dictionary<string, VisioMaster>> effectivePageMasters) {
            XNamespace ns = VisioNamespace;
            Dictionary<CommentAuthorKey, int> authorIds = new();
            List<(VisioPage Page, VisioComment Comment, int AuthorId)> comments = new();

            foreach (VisioPage page in pages) {
                foreach (VisioComment comment in page.Comments) {
                    ValidateCommentForSave(comment);
                    CommentAuthorKey authorKey = new(comment.AuthorName, comment.AuthorInitials, comment.AuthorResolutionId);
                    if (!authorIds.TryGetValue(authorKey, out int authorId)) {
                        authorId = authorIds.Count + 1;
                        authorIds.Add(authorKey, authorId);
                    }

                    comments.Add((page, comment, authorId));
                    if (comments.Count > MaxLoadedComments) {
                        throw new InvalidDataException($"Visio comments part contains more than {MaxLoadedComments} comments.");
                    }
                }
            }

            XElement authorList = new(ns + "AuthorList");
            foreach (KeyValuePair<CommentAuthorKey, int> author in authorIds.OrderBy(pair => pair.Value)) {
                XElement authorElement = new(ns + "AuthorEntry",
                    new XAttribute("ID", XmlConvert.ToString(author.Value)),
                    new XAttribute("Name", author.Key.Name),
                    new XAttribute("Initials", author.Key.Initials));
                if (!string.IsNullOrWhiteSpace(author.Key.ResolutionId)) {
                    authorElement.Add(new XAttribute("ResolutionID", author.Key.ResolutionId));
                }

                authorList.Add(authorElement);
            }

            XElement commentList = new(ns + "CommentList");
            foreach ((VisioPage page, VisioComment comment, int authorId) in comments) {
                int commentId = comment.Id > 0 ? comment.Id : AssignSaveFallbackCommentId(comment, commentList);
                XElement commentElement = new(ns + "CommentEntry",
                    new XAttribute("IX", XmlConvert.ToString(commentId)),
                    new XAttribute("AuthorID", XmlConvert.ToString(authorId)),
                    new XAttribute("PageID", XmlConvert.ToString(page.Id)),
                    new XAttribute("Date", FormatCommentDate(comment.CreatedAt)),
                    new XAttribute("Done", comment.Done ? "1" : "0"));

                if (comment.EditedAt.HasValue) {
                    commentElement.Add(new XAttribute("EditDate", FormatCommentDate(comment.EditedAt.Value)));
                }

                if (comment.AutoCommentType.HasValue) {
                    commentElement.Add(new XAttribute("AutoCommentType", XmlConvert.ToString(comment.AutoCommentType.Value)));
                }

                if (!string.IsNullOrWhiteSpace(comment.ShapeId)) {
                    Dictionary<string, string> persistedIds = BuildPersistedIdMap(page, effectivePageMasters[page]);
                    commentElement.Add(new XAttribute("ShapeID", GetPersistedId(persistedIds, comment.ShapeId!)));
                }

                commentElement.Value = comment.Text ?? string.Empty;
                commentList.Add(commentElement);
            }

            XDocument commentsXml = new(new XElement(ns + "Comments",
                new XAttribute(XNamespace.Xml + "space", "preserve"),
                authorList,
                commentList));

            string serializedComments = commentsXml.Declaration + Environment.NewLine + commentsXml.ToString(SaveOptions.DisableFormatting);
            ValidateCommentsXmlForSave(serializedComments);
            return serializedComments;
        }

        private static void ValidateCommentForSave(VisioComment comment) {
            string text = comment.Text ?? string.Empty;
            if (text.Length > MaxCommentTextCharacters) {
                throw new InvalidDataException($"Visio comment text exceeds {MaxCommentTextCharacters} characters.");
            }
        }

        private static void ValidateCommentsXmlForSave(string commentsXml) {
            if (commentsXml.Length > MaxCommentsXmlCharacters) {
                throw new InvalidDataException($"Visio comments part exceeds {MaxCommentsXmlCharacters} XML characters.");
            }

            int byteCount = Encoding.UTF8.GetByteCount(commentsXml);
            if (byteCount > MaxCommentsPartBytes) {
                throw new InvalidDataException($"Visio comments part exceeds {MaxCommentsPartBytes} bytes.");
            }
        }

        private static int AssignSaveFallbackCommentId(VisioComment comment, XElement commentList) {
            int commentId = GetSaveFallbackCommentId(commentList);
            comment.Id = commentId;
            return commentId;
        }

        private static int GetSaveFallbackCommentId(XElement commentList) {
            return commentList.Elements()
                .Select(element => (string?)element.Attribute("IX"))
                .Select(value => int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int id) ? id : 0)
                .DefaultIfEmpty(0)
                .Max() + 1;
        }

        private static string FormatCommentDate(DateTimeOffset value) {
            return XmlConvert.ToString(value.UtcDateTime, XmlDateTimeSerializationMode.Utc);
        }
    }
}
