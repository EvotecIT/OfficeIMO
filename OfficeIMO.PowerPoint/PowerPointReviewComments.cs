using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using P = DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using P188 = DocumentFormat.OpenXml.Office2021.PowerPoint.Comment;

namespace OfficeIMO.PowerPoint {
    /// <summary>Identity used when creating or reassigning a PowerPoint review comment.</summary>
    public sealed class PowerPointCommentAuthor {
        /// <summary>Creates a review author identity.</summary>
        public PowerPointCommentAuthor(string name, string? initials = null,
            string? userId = null, string? providerId = null) {
            Name = string.IsNullOrWhiteSpace(name)
                ? throw new ArgumentException("Author name cannot be empty.", nameof(name))
                : name.Trim();
            Initials = string.IsNullOrWhiteSpace(initials)
                ? CreateInitials(Name)
                : initials!.Trim();
            UserId = string.IsNullOrWhiteSpace(userId) ? null : userId!.Trim();
            ProviderId = string.IsNullOrWhiteSpace(providerId) ? null : providerId!.Trim();
        }

        /// <summary>Display name.</summary>
        public string Name { get; }
        /// <summary>Short initials displayed by classic review surfaces.</summary>
        public string Initials { get; }
        /// <summary>Optional modern author identity.</summary>
        public string? UserId { get; }
        /// <summary>Optional modern identity provider.</summary>
        public string? ProviderId { get; }

        private static string CreateInitials(string name) {
            string[] words = name.Split(new[] { ' ', '\t' },
                StringSplitOptions.RemoveEmptyEntries);
            if (words.Length == 0) return "?";
            return string.Concat(words.Take(2)
                    .Select(word => StringInfo.GetNextTextElement(word)))
                .ToUpperInvariant();
        }
    }

    /// <summary>Status of a modern threaded PowerPoint comment or reply.</summary>
    public enum PowerPointModernCommentStatus {
        /// <summary>The review item is active.</summary>
        Active,
        /// <summary>The review item has been resolved.</summary>
        Resolved,
        /// <summary>The review item has been closed.</summary>
        Closed
    }

    /// <summary>Editable classic PowerPoint comment.</summary>
    public sealed class PowerPointClassicComment {
        private readonly PowerPointPresentation _presentation;
        private readonly PowerPointSlide _slide;
        private readonly P.Comment _comment;

        internal PowerPointClassicComment(PowerPointPresentation presentation,
            PowerPointSlide slide, P.Comment comment) {
            _presentation = presentation;
            _slide = slide;
            _comment = comment;
        }

        /// <summary>Per-author comment index stored by PowerPoint.</summary>
        public uint Index => RequireAttached().Index?.Value ?? 0U;

        /// <summary>Comment author.</summary>
        public PowerPointCommentAuthor Author =>
            _presentation.ResolveClassicCommentAuthor(RequireAttached().AuthorId?.Value);

        /// <summary>Visible review text.</summary>
        public string Text {
            get => RequireAttached().Text?.Text ?? string.Empty;
            set {
                PowerPointPresentation.ValidateClassicCommentText(
                    value, nameof(value));
                RequireAttached().Text = new P.Text(value);
            }
        }

        /// <summary>Creation timestamp when present.</summary>
        public DateTime? Created {
            get => RequireAttached().DateTime?.Value;
            set => RequireAttached().DateTime = value;
        }

        /// <summary>Classic comment X position.</summary>
        public long X {
            get => RequirePosition().X?.Value ?? 0L;
            set {
                PowerPointPresentation.ValidateClassicCommentPosition(
                    value, nameof(value));
                RequirePosition().X = value;
            }
        }

        /// <summary>Classic comment Y position.</summary>
        public long Y {
            get => RequirePosition().Y?.Value ?? 0L;
            set {
                PowerPointPresentation.ValidateClassicCommentPosition(
                    value, nameof(value));
                RequirePosition().Y = value;
            }
        }

        /// <summary>Reassigns the comment to an existing or newly created author.</summary>
        public void SetAuthor(PowerPointCommentAuthor author) {
            if (author == null) throw new ArgumentNullException(nameof(author));
            P.Comment attached = RequireAttached();
            uint? previousAuthorId = attached.AuthorId?.Value;
            P.CommentAuthor target = _presentation.GetOrCreateClassicCommentAuthor(author);
            if (target.Id?.Value == previousAuthorId) return;
            uint nextIndex = _presentation.AllocateClassicCommentIndex(target);
            attached.AuthorId = target.Id?.Value ?? 0U;
            attached.Index = nextIndex;
            _presentation.RemoveClassicCommentAuthorIfUnused(previousAuthorId);
        }

        /// <summary>Removes this comment from its slide.</summary>
        public void Remove() {
            uint? authorId = RequireAttached().AuthorId?.Value;
            SlideCommentsPart? part = _slide.SlidePart.SlideCommentsPart;
            _comment.Remove();
            if (part?.CommentList != null && !part.CommentList.Elements<P.Comment>().Any()) {
                _slide.SlidePart.DeletePart(part);
            }
            _presentation.RemoveClassicCommentAuthorIfUnused(authorId);
        }

        private P.Comment RequireAttached() {
            _presentation.ThrowIfDisposedForCommentApi();
            if (_comment.Parent == null) {
                throw new InvalidOperationException("The classic comment is no longer attached to the presentation.");
            }
            return _comment;
        }

        private P.Position RequirePosition() {
            P.Comment comment = RequireAttached();
            P.Position? position = comment.Position;
            if (position != null) return position;
            position = new P.Position { X = 0L, Y = 0L };
            comment.InsertAt(position, 0);
            return position;
        }

    }

    /// <summary>Editable modern threaded PowerPoint comment.</summary>
    public sealed class PowerPointModernComment {
        private readonly PowerPointPresentation _presentation;
        private readonly PowerPointSlide _slide;
        private readonly PowerPointCommentPart _part;
        private readonly P188.Comment _comment;

        internal PowerPointModernComment(PowerPointPresentation presentation,
            PowerPointSlide slide, PowerPointCommentPart part, P188.Comment comment) {
            _presentation = presentation;
            _slide = slide;
            _part = part;
            _comment = comment;
        }

        /// <summary>Stable modern comment identifier.</summary>
        public string Id => RequireAttached().Id?.Value ?? string.Empty;

        /// <summary>Comment author.</summary>
        public PowerPointCommentAuthor Author =>
            _presentation.ResolveModernCommentAuthor(RequireAttached().AuthorId?.Value);

        /// <summary>Visible review text.</summary>
        public string Text {
            get => PowerPointPresentation.GetModernCommentText(RequireAttached());
            set {
                PowerPointPresentation.ValidateCommentText(value);
                PowerPointPresentation.SetModernCommentText(RequireAttached(), value);
            }
        }

        /// <summary>Review status.</summary>
        public PowerPointModernCommentStatus Status {
            get => PowerPointPresentation.FromModernStatus(RequireAttached().Status?.Value);
            set => RequireAttached().Status = PowerPointPresentation.ToModernStatus(value);
        }

        /// <summary>Creation timestamp.</summary>
        public DateTime? Created {
            get => RequireAttached().Created?.Value;
            set => RequireAttached().Created = value;
        }

        /// <summary>Comment X position.</summary>
        public long X {
            get => RequirePosition().X?.Value ?? 0L;
            set => RequirePosition().X = value;
        }

        /// <summary>Comment Y position.</summary>
        public long Y {
            get => RequirePosition().Y?.Value ?? 0L;
            set => RequirePosition().Y = value;
        }

        /// <summary>Replies in package order.</summary>
        public IReadOnlyList<PowerPointModernCommentReply> Replies {
            get {
                P188.Comment comment = RequireAttached();
                return (IReadOnlyList<PowerPointModernCommentReply>)(comment
                    .GetFirstChild<P188.CommentReplyList>()?
                    .Elements<P188.CommentReply>()
                    .Select(reply => new PowerPointModernCommentReply(
                        _presentation, comment, reply)).ToArray()
                    ?? Array.Empty<PowerPointModernCommentReply>());
            }
        }

        /// <summary>Adds a reply to this comment.</summary>
        public PowerPointModernCommentReply AddReply(PowerPointCommentAuthor author,
            string text, PowerPointModernCommentStatus status = PowerPointModernCommentStatus.Active,
            DateTime? created = null) {
            if (author == null) throw new ArgumentNullException(nameof(author));
            PowerPointPresentation.ValidateCommentText(text);
            P188.Comment comment = RequireAttached();
            P188.Author modernAuthor = _presentation.GetOrCreateModernCommentAuthor(author);
            var reply = new P188.CommentReply(
                PowerPointPresentation.CreateModernCommentTextBody(text)) {
                Id = PowerPointPresentation.CreateModernCommentId(),
                AuthorId = modernAuthor.Id?.Value,
                Status = PowerPointPresentation.ToModernStatus(status),
                Created = created ?? DateTime.UtcNow
            };
            P188.CommentReplyList? replies = comment.GetFirstChild<P188.CommentReplyList>();
            if (replies == null) {
                replies = new P188.CommentReplyList();
                comment.AddChild(replies, true);
            }
            replies.Append(reply);
            return new PowerPointModernCommentReply(_presentation, comment, reply);
        }

        /// <summary>Reassigns the comment to an existing or newly created author.</summary>
        public void SetAuthor(PowerPointCommentAuthor author) {
            if (author == null) throw new ArgumentNullException(nameof(author));
            P188.Comment attached = RequireAttached();
            string? previousAuthorId = attached.AuthorId?.Value;
            string? targetAuthorId = _presentation.GetOrCreateModernCommentAuthor(author).Id?.Value;
            if (string.Equals(previousAuthorId, targetAuthorId,
                    StringComparison.OrdinalIgnoreCase)) return;
            attached.AuthorId = targetAuthorId;
            _presentation.RemoveModernCommentAuthorIfUnused(previousAuthorId);
        }

        /// <summary>Removes the comment and all of its replies.</summary>
        public void Remove() {
            P188.Comment attached = RequireAttached();
            string?[] authorIds = attached.Descendants<P188.CommentReply>()
                .Select(reply => reply.AuthorId?.Value)
                .Concat(new[] { attached.AuthorId?.Value })
                .Where(authorId => !string.IsNullOrWhiteSpace(authorId))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToArray();
            _comment.Remove();
            if (_part.CommentList == null
                || !_part.CommentList.Elements<P188.Comment>().Any()) {
                _slide.SlidePart.DeletePart(_part);
            }
            foreach (string? authorId in authorIds) {
                _presentation.RemoveModernCommentAuthorIfUnused(authorId);
            }
        }

        private P188.Comment RequireAttached() {
            _presentation.ThrowIfDisposedForCommentApi();
            if (_comment.Parent == null) {
                throw new InvalidOperationException("The modern comment is no longer attached to the presentation.");
            }
            return _comment;
        }

        private P188.Point2DType RequirePosition() {
            P188.Comment comment = RequireAttached();
            P188.Point2DType? position = comment.GetFirstChild<P188.Point2DType>();
            if (position != null) return position;
            position = new P188.Point2DType { X = 0L, Y = 0L };
            comment.AddChild(position, true);
            return position;
        }
    }

    /// <summary>Editable reply in a modern PowerPoint comment thread.</summary>
    public sealed class PowerPointModernCommentReply {
        private readonly PowerPointPresentation _presentation;
        private readonly P188.Comment _parent;
        private readonly P188.CommentReply _reply;

        internal PowerPointModernCommentReply(PowerPointPresentation presentation,
            P188.Comment parent, P188.CommentReply reply) {
            _presentation = presentation;
            _parent = parent;
            _reply = reply;
        }

        /// <summary>Stable reply identifier.</summary>
        public string Id => RequireAttached().Id?.Value ?? string.Empty;

        /// <summary>Reply author.</summary>
        public PowerPointCommentAuthor Author =>
            _presentation.ResolveModernCommentAuthor(RequireAttached().AuthorId?.Value);

        /// <summary>Visible reply text.</summary>
        public string Text {
            get => PowerPointPresentation.GetModernCommentText(RequireAttached());
            set {
                PowerPointPresentation.ValidateCommentText(value);
                PowerPointPresentation.SetModernCommentText(RequireAttached(), value);
            }
        }

        /// <summary>Review status.</summary>
        public PowerPointModernCommentStatus Status {
            get => PowerPointPresentation.FromModernStatus(RequireAttached().Status?.Value);
            set => RequireAttached().Status = PowerPointPresentation.ToModernStatus(value);
        }

        /// <summary>Creation timestamp.</summary>
        public DateTime? Created {
            get => RequireAttached().Created?.Value;
            set => RequireAttached().Created = value;
        }

        /// <summary>Reassigns the reply to an existing or newly created author.</summary>
        public void SetAuthor(PowerPointCommentAuthor author) {
            if (author == null) throw new ArgumentNullException(nameof(author));
            P188.CommentReply attached = RequireAttached();
            string? previousAuthorId = attached.AuthorId?.Value;
            string? targetAuthorId = _presentation.GetOrCreateModernCommentAuthor(author).Id?.Value;
            if (string.Equals(previousAuthorId, targetAuthorId,
                    StringComparison.OrdinalIgnoreCase)) return;
            attached.AuthorId = targetAuthorId;
            _presentation.RemoveModernCommentAuthorIfUnused(previousAuthorId);
        }

        /// <summary>Removes this reply from its thread.</summary>
        public void Remove() {
            string? authorId = RequireAttached().AuthorId?.Value;
            P188.CommentReplyList? list = _reply.Parent as P188.CommentReplyList;
            _reply.Remove();
            if (list != null && !list.Elements<P188.CommentReply>().Any()) list.Remove();
            _presentation.RemoveModernCommentAuthorIfUnused(authorId);
        }

        private P188.CommentReply RequireAttached() {
            _presentation.ThrowIfDisposedForCommentApi();
            if (_reply.Parent == null || _parent.Parent == null) {
                throw new InvalidOperationException("The modern comment reply is no longer attached to the presentation.");
            }
            return _reply;
        }
    }

    public sealed partial class PowerPointPresentation {
        /// <summary>Returns editable classic comments for a slide.</summary>
        public IReadOnlyList<PowerPointClassicComment> GetClassicComments(PowerPointSlide slide) {
            EnsureCommentSlide(slide);
            return (IReadOnlyList<PowerPointClassicComment>)(slide.SlidePart.SlideCommentsPart?
                .CommentList?.Elements<P.Comment>()
                .Select(comment => new PowerPointClassicComment(this, slide, comment)).ToArray()
                ?? Array.Empty<PowerPointClassicComment>());
        }

        /// <summary>Adds a classic comment that can round-trip through PPTX and binary PPT.</summary>
        public PowerPointClassicComment AddClassicComment(PowerPointSlide slide,
            PowerPointCommentAuthor author, string text, long x = 0L, long y = 0L,
            DateTime? created = null) {
            EnsureCommentSlide(slide);
            if (author == null) throw new ArgumentNullException(nameof(author));
            ValidateClassicCommentText(text, nameof(text));
            ValidateClassicCommentPosition(x, nameof(x));
            ValidateClassicCommentPosition(y, nameof(y));
            P.CommentAuthor commentAuthor = GetOrCreateClassicCommentAuthor(author);
            var comment = new P.Comment(
                new P.Position { X = x, Y = y }, new P.Text(text)) {
                AuthorId = commentAuthor.Id?.Value ?? 0U,
                Index = AllocateClassicCommentIndex(commentAuthor),
                DateTime = created ?? DateTime.UtcNow
            };
            SlideCommentsPart? commentsPart = slide.SlidePart.SlideCommentsPart;
            if (commentsPart == null) {
                commentsPart = slide.SlidePart.AddNewPart<SlideCommentsPart>();
                commentsPart.CommentList = new P.CommentList();
            }
            if (commentsPart.CommentList == null) commentsPart.CommentList = new P.CommentList();
            commentsPart.CommentList.Append(comment);
            return new PowerPointClassicComment(this, slide, comment);
        }

        /// <summary>Returns editable modern comments for a slide.</summary>
        public IReadOnlyList<PowerPointModernComment> GetModernComments(PowerPointSlide slide) {
            EnsureCommentSlide(slide);
            return slide.SlidePart.Parts.Select(pair => pair.OpenXmlPart)
                .OfType<PowerPointCommentPart>()
                .SelectMany(part => (part.CommentList?.Elements<P188.Comment>()
                    ?? Enumerable.Empty<P188.Comment>())
                    .Select(comment => new PowerPointModernComment(this, slide, part, comment)))
                .ToArray();
        }

        /// <summary>Adds a modern threaded comment to a slide.</summary>
        public PowerPointModernComment AddModernComment(PowerPointSlide slide,
            PowerPointCommentAuthor author, string text,
            PowerPointModernCommentStatus status = PowerPointModernCommentStatus.Active,
            long x = 0L, long y = 0L, DateTime? created = null) {
            EnsureCommentSlide(slide);
            if (author == null) throw new ArgumentNullException(nameof(author));
            ValidateCommentText(text);
            P188.Author modernAuthor = GetOrCreateModernCommentAuthor(author);
            var comment = new P188.Comment(
                new P188.CommentUnknownAnchor(),
                new P188.Point2DType { X = x, Y = y },
                CreateModernCommentTextBody(text)) {
                Id = CreateModernCommentId(),
                AuthorId = modernAuthor.Id?.Value,
                Status = ToModernStatus(status),
                Created = created ?? DateTime.UtcNow
            };
            PowerPointCommentPart? commentsPart = slide.SlidePart.Parts
                .Select(pair => pair.OpenXmlPart).OfType<PowerPointCommentPart>().FirstOrDefault();
            if (commentsPart == null) {
                commentsPart = slide.SlidePart.AddNewPart<PowerPointCommentPart>();
                commentsPart.CommentList = new P188.CommentList();
            }
            if (commentsPart.CommentList == null) commentsPart.CommentList = new P188.CommentList();
            commentsPart.CommentList.Append(comment);
            return new PowerPointModernComment(this, slide, commentsPart, comment);
        }

        internal void ThrowIfDisposedForCommentApi() => ThrowIfDisposed();

        internal P.CommentAuthor GetOrCreateClassicCommentAuthor(PowerPointCommentAuthor author) {
            ValidateClassicCommentAuthor(author);
            CommentAuthorsPart? part = _presentationPart.CommentAuthorsPart;
            if (part == null) {
                part = _presentationPart.AddNewPart<CommentAuthorsPart>();
                part.CommentAuthorList = new P.CommentAuthorList();
            }
            if (part.CommentAuthorList == null) part.CommentAuthorList = new P.CommentAuthorList();
            P.CommentAuthor? existing = part.CommentAuthorList.Elements<P.CommentAuthor>()
                .FirstOrDefault(candidate => string.Equals(candidate.Name?.Value, author.Name,
                    StringComparison.Ordinal) && string.Equals(candidate.Initials?.Value,
                    author.Initials, StringComparison.Ordinal));
            if (existing != null) return existing;
            uint id = AllocateClassicAuthorId(part.CommentAuthorList);
            var created = new P.CommentAuthor {
                Id = id,
                Name = author.Name,
                Initials = author.Initials,
                LastIndex = 0U,
                ColorIndex = id
            };
            part.CommentAuthorList.Append(created);
            return created;
        }

        internal uint AllocateClassicCommentIndex(P.CommentAuthor author) {
            uint? authorId = author.Id?.Value;
            var usedIndexes = new HashSet<uint>(_slides.SelectMany(slide =>
                    slide.SlidePart.SlideCommentsPart?.CommentList?
                        .Elements<P.Comment>() ?? Enumerable.Empty<P.Comment>())
                .Where(comment => comment.AuthorId?.Value == authorId)
                .Where(comment => comment.Index?.Value != null)
                .Select(comment => comment.Index!.Value));
            uint firstCandidate = author.LastIndex?.Value == uint.MaxValue
                ? 0U
                : (author.LastIndex?.Value ?? 0U) + 1U;
            uint next = FindAvailableUInt32Id(usedIndexes, firstCandidate,
                "classic-comment");
            author.LastIndex = usedIndexes.Count == 0
                ? next
                : Math.Max(usedIndexes.Max(), next);
            return next;
        }

        internal PowerPointCommentAuthor ResolveClassicCommentAuthor(uint? authorId) {
            P.CommentAuthor? author = _presentationPart.CommentAuthorsPart?.CommentAuthorList?
                .Elements<P.CommentAuthor>()
                .FirstOrDefault(candidate => candidate.Id?.Value == authorId);
            return author == null
                ? new PowerPointCommentAuthor("Unknown")
                : new PowerPointCommentAuthor(author.Name?.Value ?? "Unknown",
                    author.Initials?.Value);
        }

        internal void RemoveClassicCommentAuthorIfUnused(uint? authorId) {
            if (!authorId.HasValue) return;
            CommentAuthorsPart? part = _presentationPart.CommentAuthorsPart;
            P.CommentAuthor? author = part?.CommentAuthorList?
                .Elements<P.CommentAuthor>()
                .FirstOrDefault(candidate => candidate.Id?.Value == authorId.Value);
            uint[] survivingIndexes = _slides.SelectMany(slide =>
                    slide.SlidePart.SlideCommentsPart?.CommentList?
                        .Elements<P.Comment>() ?? Enumerable.Empty<P.Comment>())
                .Where(comment => comment.AuthorId?.Value == authorId.Value)
                .Select(comment => comment.Index?.Value ?? 0U)
                .ToArray();
            if (survivingIndexes.Length > 0) {
                if (author != null) author.LastIndex = survivingIndexes.Max();
                return;
            }
            author?.Remove();
            if (part?.CommentAuthorList == null
                || part.CommentAuthorList.Elements<P.CommentAuthor>().Any()) return;
            _presentationPart.DeletePart(part);
        }

        internal P188.Author GetOrCreateModernCommentAuthor(PowerPointCommentAuthor author) {
            PowerPointAuthorsPart? firstPart = null;
            foreach (PowerPointAuthorsPart part in _presentationPart.Parts
                         .Select(pair => pair.OpenXmlPart).OfType<PowerPointAuthorsPart>()) {
                if (firstPart == null) firstPart = part;
                P188.Author? existing = part.AuthorList?.Elements<P188.Author>()
                    .FirstOrDefault(candidate => ModernAuthorMatches(candidate, author));
                if (existing != null) return existing;
            }
            PowerPointAuthorsPart target = firstPart ?? _presentationPart.AddNewPart<PowerPointAuthorsPart>();
            if (target.AuthorList == null) target.AuthorList = new P188.AuthorList();
            var created = new P188.Author {
                Id = CreateModernCommentId(),
                Name = author.Name,
                Initials = author.Initials,
                UserId = author.UserId,
                ProviderId = author.ProviderId
            };
            target.AuthorList.Append(created);
            return created;
        }

        internal PowerPointCommentAuthor ResolveModernCommentAuthor(string? authorId) {
            P188.Author? author = _presentationPart.Parts.Select(pair => pair.OpenXmlPart)
                .OfType<PowerPointAuthorsPart>()
                .SelectMany(part => part.AuthorList?.Elements<P188.Author>()
                    ?? Enumerable.Empty<P188.Author>())
                .FirstOrDefault(candidate => string.Equals(candidate.Id?.Value, authorId,
                    StringComparison.OrdinalIgnoreCase));
            return author == null
                ? new PowerPointCommentAuthor("Unknown")
                : new PowerPointCommentAuthor(author.Name?.Value ?? "Unknown",
                    author.Initials?.Value, author.UserId?.Value, author.ProviderId?.Value);
        }

        internal void RemoveModernCommentAuthorIfUnused(string? authorId) {
            if (string.IsNullOrWhiteSpace(authorId)) return;
            bool inUse = _slides.SelectMany(slide => slide.SlidePart.Parts
                    .Select(pair => pair.OpenXmlPart)
                    .OfType<PowerPointCommentPart>())
                .SelectMany(part => part.CommentList?.Elements<P188.Comment>()
                    ?? Enumerable.Empty<P188.Comment>())
                .Any(comment => string.Equals(comment.AuthorId?.Value, authorId,
                        StringComparison.OrdinalIgnoreCase)
                    || comment.Descendants<P188.CommentReply>().Any(reply =>
                        string.Equals(reply.AuthorId?.Value, authorId,
                            StringComparison.OrdinalIgnoreCase)));
            if (inUse) return;

            foreach (PowerPointAuthorsPart part in _presentationPart.Parts
                         .Select(pair => pair.OpenXmlPart)
                         .OfType<PowerPointAuthorsPart>().ToArray()) {
                P188.Author? author = part.AuthorList?.Elements<P188.Author>()
                    .FirstOrDefault(candidate => string.Equals(
                        candidate.Id?.Value, authorId,
                        StringComparison.OrdinalIgnoreCase));
                author?.Remove();
                if (part.AuthorList == null
                    || !part.AuthorList.Elements<P188.Author>().Any()) {
                    _presentationPart.DeletePart(part);
                }
            }
        }

        internal static P188.TextBodyType CreateModernCommentTextBody(string text) {
            var body = new P188.TextBodyType(new A.BodyProperties(),
                new A.ListStyle());
            body.Append(CreateModernCommentParagraphs(text));
            return body;
        }

        internal static string GetModernCommentText(DocumentFormat.OpenXml.OpenXmlElement element) {
            A.Paragraph[] paragraphs = element.GetFirstChild<P188.TextBodyType>()?
                .Elements<A.Paragraph>().ToArray() ?? Array.Empty<A.Paragraph>();
            return string.Join("\n", paragraphs.Select(paragraph => {
                var builder = new System.Text.StringBuilder();
                foreach (DocumentFormat.OpenXml.OpenXmlElement item in
                         paragraph.Descendants()) {
                    if (item is A.Text text) builder.Append(text.Text);
                    else if (item is A.Break) builder.Append('\n');
                }
                return builder.ToString();
            }));
        }

        internal static void SetModernCommentText(DocumentFormat.OpenXml.OpenXmlElement element,
            string text) {
            P188.TextBodyType? body = element.GetFirstChild<P188.TextBodyType>();
            if (body == null) {
                body = CreateModernCommentTextBody(text);
                if (element is P188.Comment comment) comment.AddChild(body, true);
                else if (element is P188.CommentReply reply) reply.AddChild(body, true);
                else element.Append(body);
                return;
            }
            body.RemoveAllChildren<A.Paragraph>();
            body.Append(CreateModernCommentParagraphs(text));
        }

        private static A.Paragraph[] CreateModernCommentParagraphs(string text) =>
            text.Replace("\r\n", "\n").Replace('\r', '\n')
                .Split(new[] { '\n' }, StringSplitOptions.None)
                .Select(line => string.IsNullOrEmpty(line)
                    ? new A.Paragraph()
                    : new A.Paragraph(new A.Run(new A.Text(line))))
                .ToArray();

        internal static P188.CommentStatus ToModernStatus(PowerPointModernCommentStatus status) {
            switch (status) {
                case PowerPointModernCommentStatus.Active: return P188.CommentStatus.Active;
                case PowerPointModernCommentStatus.Resolved: return P188.CommentStatus.Resolved;
                case PowerPointModernCommentStatus.Closed: return P188.CommentStatus.Closed;
                default: throw new ArgumentOutOfRangeException(nameof(status));
            }
        }

        internal static PowerPointModernCommentStatus FromModernStatus(P188.CommentStatus? status) {
            if (status.HasValue && status.Value.Equals(P188.CommentStatus.Resolved)) {
                return PowerPointModernCommentStatus.Resolved;
            }
            if (status.HasValue && status.Value.Equals(P188.CommentStatus.Closed)) {
                return PowerPointModernCommentStatus.Closed;
            }
            return PowerPointModernCommentStatus.Active;
        }

        internal static string CreateModernCommentId() =>
            Guid.NewGuid().ToString("B").ToUpperInvariant();

        internal static void ValidateCommentText(string text) {
            if (string.IsNullOrWhiteSpace(text)) {
                throw new ArgumentException("Comment text cannot be empty.", nameof(text));
            }
            PowerPointXmlValueValidator.ValidateCharacters(text,
                nameof(text), "Comment text");
        }

        internal static void ValidateClassicCommentText(string text,
            string parameterName) {
            if (string.IsNullOrWhiteSpace(text)) {
                throw new ArgumentException("Comment text cannot be empty.",
                    parameterName);
            }
            if (text.Length > 32000) {
                throw new ArgumentException(
                    "Classic comment text cannot exceed 32,000 characters.",
                    parameterName);
            }
            if (text.IndexOf('\0') >= 0) {
                throw new ArgumentException(
                    "Classic comment text cannot contain a NUL character.",
                    parameterName);
            }
            PowerPointXmlValueValidator.ValidateCharacters(text,
                parameterName, "Classic comment text");
        }

        internal static void ValidateClassicCommentPosition(long value,
            string parameterName) {
            if (value < int.MinValue || value > int.MaxValue) {
                throw new ArgumentOutOfRangeException(parameterName,
                    "Classic comment positions must fit in the binary PowerPoint signed 32-bit coordinate range.");
            }
        }

        private static void ValidateClassicCommentAuthor(
            PowerPointCommentAuthor author) {
            if (author.Name.Length > 52 || author.Initials.Length > 52) {
                throw new ArgumentException(
                    "Classic comment author names and initials cannot exceed 52 characters.",
                    nameof(author));
            }
            if (author.Name.IndexOf('\0') >= 0
                || author.Initials.IndexOf('\0') >= 0) {
                throw new ArgumentException(
                    "Classic comment author names and initials cannot contain a NUL character.",
                    nameof(author));
            }
        }

        private void EnsureCommentSlide(PowerPointSlide slide) {
            ThrowIfDisposed();
            if (slide == null) throw new ArgumentNullException(nameof(slide));
            if (!_slides.Contains(slide)) {
                throw new ArgumentException("The slide does not belong to this presentation.", nameof(slide));
            }
        }

        private static uint AllocateClassicAuthorId(P.CommentAuthorList authors) {
            var usedIds = new HashSet<uint>(authors.Elements<P.CommentAuthor>()
                .Where(author => author.Id?.Value != null)
                .Select(author => author.Id!.Value));
            return FindAvailableUInt32Id(usedIds, 0U,
                "classic-comment-author");
        }

        private static bool ModernAuthorMatches(P188.Author candidate,
            PowerPointCommentAuthor author) =>
            string.Equals(candidate.Name?.Value, author.Name, StringComparison.Ordinal)
            && string.Equals(candidate.Initials?.Value, author.Initials, StringComparison.Ordinal)
            && string.Equals(candidate.UserId?.Value, author.UserId, StringComparison.Ordinal)
            && string.Equals(candidate.ProviderId?.Value, author.ProviderId, StringComparison.Ordinal);
    }
}
