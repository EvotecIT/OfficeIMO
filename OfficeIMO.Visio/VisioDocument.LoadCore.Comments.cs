using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Xml;
using System.Xml.Linq;

namespace OfficeIMO.Visio {
    public partial class VisioDocument {
        internal const long MaxCommentsPartBytes = 12_000_000;
        internal const long MaxCommentsXmlCharacters = 10_000_000;
        internal const int MaxLoadedComments = 10_000;
        internal const int MaxCommentTextCharacters = 32_768;

        private static readonly XmlReaderSettings CommentsXmlReaderSettings = new() {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
            MaxCharactersInDocument = MaxCommentsXmlCharacters,
            MaxCharactersFromEntities = 0,
        };

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
            XDocument commentsXml = LoadCommentsXml(commentsPart);
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
            int loadedCommentCount = 0;
            foreach (XElement commentElement in commentList?.Elements().Where(element => IsVisioElement(element, "CommentEntry")) ?? Enumerable.Empty<XElement>()) {
                loadedCommentCount++;
                if (loadedCommentCount > MaxLoadedComments) {
                    throw new InvalidDataException($"Visio comments part contains more than {MaxLoadedComments} comments.");
                }

                string commentText = GetBoundedCommentText(commentElement);
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
                VisioComment comment = new(commentText) {
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

        private static XDocument LoadCommentsXml(PackagePart commentsPart) {
            using Stream commentsStream = commentsPart.GetStream();
            using Stream boundedStream = new BoundedReadStream(commentsStream, MaxCommentsPartBytes);
            using XmlReader reader = XmlReader.Create(boundedStream, CommentsXmlReaderSettings);
            return XDocument.Load(reader);
        }

        private static string GetBoundedCommentText(XElement commentElement) {
            int textLength = 0;
            foreach (XText text in commentElement.DescendantNodes().OfType<XText>()) {
                textLength += text.Value.Length;
                if (textLength > MaxCommentTextCharacters) {
                    throw new InvalidDataException($"Visio comment text exceeds {MaxCommentTextCharacters} characters.");
                }
            }

            return commentElement.Value;
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

        private sealed class BoundedReadStream : Stream {
            private readonly Stream _inner;
            private readonly long _maxBytes;
            private long _bytesRead;

            internal BoundedReadStream(Stream inner, long maxBytes) {
                _inner = inner ?? throw new ArgumentNullException(nameof(inner));
                _maxBytes = maxBytes;
            }

            public override bool CanRead => _inner.CanRead;

            public override bool CanSeek => false;

            public override bool CanWrite => false;

            public override long Length => throw new NotSupportedException();

            public override long Position {
                get => _bytesRead;
                set => throw new NotSupportedException();
            }

            public override void Flush() {
                _inner.Flush();
            }

            public override int Read(byte[] buffer, int offset, int count) {
                int read = _inner.Read(buffer, offset, count);
                _bytesRead += read;
                if (_bytesRead > _maxBytes) {
                    throw new InvalidDataException($"Visio comments part exceeds {_maxBytes} bytes.");
                }

                return read;
            }

            public override long Seek(long offset, SeekOrigin origin) {
                throw new NotSupportedException();
            }

            public override void SetLength(long value) {
                throw new NotSupportedException();
            }

            public override void Write(byte[] buffer, int offset, int count) {
                throw new NotSupportedException();
            }
        }
    }
}
