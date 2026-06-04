using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO.Packaging;
using System.Linq;
using System.Xml;
using System.Xml.Linq;

namespace OfficeIMO.Visio {
    public partial class VisioDocument {
        private static void LoadComments(Package package, PackagePart documentPart, VisioDocument document) {
            PackageRelationship? commentsRel = documentPart.GetRelationshipsByType(CommentsRelationshipType).FirstOrDefault();
            if (commentsRel == null) {
                return;
            }

            Uri commentsUri = PackUriHelper.ResolvePartUri(documentPart.Uri, commentsRel.TargetUri);
            if (!package.PartExists(commentsUri)) {
                return;
            }

            PackagePart commentsPart = package.GetPart(commentsUri);
            XDocument commentsXml = XDocument.Load(commentsPart.GetStream());
            XElement? root = commentsXml.Root;
            if (root == null) {
                return;
            }

            Dictionary<int, (string? Name, string? Initials, string? ResolutionId)> authors = new();
            XElement? authorList = root.Elements().FirstOrDefault(element => IsVisioElement(element, "AuthorList"));
            foreach (XElement authorElement in authorList?.Elements().Where(element => IsVisioElement(element, "AuthorEntry")) ?? Enumerable.Empty<XElement>()) {
                if (!TryParseIntAttribute(authorElement, "ID", out int authorId)) {
                    continue;
                }

                authors[authorId] = (
                    authorElement.Attribute("Name")?.Value,
                    authorElement.Attribute("Initials")?.Value,
                    authorElement.Attribute("ResolutionID")?.Value);
            }

            XElement? commentList = root.Elements().FirstOrDefault(element => IsVisioElement(element, "CommentList"));
            foreach (XElement commentElement in commentList?.Elements().Where(element => IsVisioElement(element, "CommentEntry")) ?? Enumerable.Empty<XElement>()) {
                if (!TryParseIntAttribute(commentElement, "PageID", out int pageId)) {
                    continue;
                }

                VisioPage? page = document.Pages.FirstOrDefault(candidate => candidate.Id == pageId);
                if (page == null) {
                    continue;
                }

                TryParseIntAttribute(commentElement, "IX", out int commentId);
                TryParseIntAttribute(commentElement, "AuthorID", out int authorId);
                authors.TryGetValue(authorId, out var author);
                VisioComment comment = new(commentElement.Value) {
                    Id = commentId > 0 ? commentId : GetNextLoadedCommentId(page),
                    AuthorName = author.Name,
                    AuthorInitials = author.Initials,
                    AuthorResolutionId = author.ResolutionId,
                    ShapeId = ResolveCommentShapeId(page, commentElement.Attribute("ShapeID")?.Value),
                    CreatedAt = ParseCommentDate(commentElement.Attribute("Date")?.Value),
                    Done = ParseCommentBool(commentElement.Attribute("Done")?.Value),
                    AutoCommentType = TryParseIntAttribute(commentElement, "AutoCommentType", out int autoCommentType) ? autoCommentType : (int?)null
                };

                string? editDate = commentElement.Attribute("EditDate")?.Value;
                if (!string.IsNullOrWhiteSpace(editDate)) {
                    comment.EditedAt = ParseCommentDate(editDate);
                }

                page.Comments.Add(comment);
            }
        }

        private static bool IsVisioElement(XElement element, string localName) {
            return string.Equals(element.Name.LocalName, localName, StringComparison.OrdinalIgnoreCase);
        }

        private static bool TryParseIntAttribute(XElement element, string attributeName, out int value) {
            return int.TryParse(element.Attribute(attributeName)?.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out value);
        }

        private static int GetNextLoadedCommentId(VisioPage page) {
            int nextId = 1;
            HashSet<int> usedIds = new(page.Comments.Select(comment => comment.Id));
            while (usedIds.Contains(nextId)) {
                nextId++;
            }

            return nextId;
        }

        private static DateTimeOffset ParseCommentDate(string? value) {
            if (string.IsNullOrWhiteSpace(value)) {
                return DateTimeOffset.UtcNow;
            }

            try {
                DateTime parsed = XmlConvert.ToDateTime(value, XmlDateTimeSerializationMode.RoundtripKind);
                return new DateTimeOffset(parsed.ToUniversalTime());
            } catch (FormatException) {
                if (DateTimeOffset.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal, out DateTimeOffset fallback)) {
                    return fallback;
                }

                return DateTimeOffset.UtcNow;
            }
        }

        private static bool ParseCommentBool(string? value) {
            return string.Equals(value, "1", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(value, "true", StringComparison.OrdinalIgnoreCase);
        }

        private static string? ResolveCommentShapeId(VisioPage page, string? persistedId) {
            if (string.IsNullOrWhiteSpace(persistedId)) {
                return null;
            }

            foreach (VisioShape shape in page.AllShapes()) {
                if (string.Equals(shape.PersistedId, persistedId, StringComparison.Ordinal) ||
                    string.Equals(shape.Id, persistedId, StringComparison.Ordinal)) {
                    return shape.Id;
                }
            }

            foreach (VisioConnector connector in page.Connectors) {
                if (string.Equals(connector.PersistedId, persistedId, StringComparison.Ordinal) ||
                    string.Equals(connector.Id, persistedId, StringComparison.Ordinal)) {
                    return connector.Id;
                }
            }

            return persistedId;
        }
    }
}
