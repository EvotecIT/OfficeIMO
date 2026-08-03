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
            string? userId = null, string? providerId = null)
            : this(name, initials, userId, providerId,
                synthesizeMissingInitials: true) {
        }

        private PowerPointCommentAuthor(string name, string? initials,
            string? userId, string? providerId,
            bool synthesizeMissingInitials,
            bool preserveClassicInitials = false) {
            if (name != null) {
                PowerPointXmlValueValidator.ValidateCharacters(name,
                    nameof(name), "Author name");
            }
            if (initials != null) {
                PowerPointXmlValueValidator.ValidateCharacters(initials,
                    nameof(initials), "Author initials");
            }
            if (userId != null) {
                PowerPointXmlValueValidator.ValidateCharacters(userId,
                    nameof(userId), "Author user identifier");
            }
            if (providerId != null) {
                PowerPointXmlValueValidator.ValidateCharacters(providerId,
                    nameof(providerId), "Author provider identifier");
            }
            if (string.IsNullOrWhiteSpace(name)) {
                throw new ArgumentException("Author name cannot be empty.",
                    nameof(name));
            }
            string normalizedName = name!.Trim();
            string? suppliedInitials = string.IsNullOrWhiteSpace(initials)
                ? null : initials!.Trim();
            string normalizedInitials = suppliedInitials
                ?? CreateInitials(normalizedName);
            string? normalizedUserId = string.IsNullOrWhiteSpace(userId)
                ? null : userId!.Trim();
            string? normalizedProviderId = string.IsNullOrWhiteSpace(providerId)
                ? null : providerId!.Trim();
            Name = normalizedName;
            Initials = normalizedInitials;
            ClassicInitials = preserveClassicInitials
                ? initials?.Trim()
                : normalizedInitials;
            ModernInitials = synthesizeMissingInitials
                ? normalizedInitials : suppliedInitials;
            UserId = normalizedUserId;
            ProviderId = normalizedProviderId;
        }

        /// <summary>Display name.</summary>
        public string Name { get; }
        /// <summary>Short initials displayed by classic review surfaces.</summary>
        public string Initials { get; }
        internal string? ClassicInitials { get; }
        internal string? ModernInitials { get; }
        /// <summary>Optional modern author identity.</summary>
        public string? UserId { get; }
        /// <summary>Optional modern identity provider.</summary>
        public string? ProviderId { get; }

        internal static PowerPointCommentAuthor FromImportedModern(
            string name, string? initials, string? userId,
            string? providerId) => new PowerPointCommentAuthor(name,
                initials, userId, providerId,
                synthesizeMissingInitials: false);

        internal static PowerPointCommentAuthor FromImportedClassic(
            string name, string? initials) => new PowerPointCommentAuthor(name,
                initials, userId: null, providerId: null,
                synthesizeMissingInitials: true,
                preserveClassicInitials: true);

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

        /// <summary>Required creation timestamp.</summary>
        public DateTime? Created {
            get => RequireAttached().DateTime?.Value;
            set {
                if (!value.HasValue) {
                    throw new ArgumentNullException(nameof(value),
                        "Classic comments require a creation timestamp.");
                }
                RequireAttached().DateTime = value.Value;
            }
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
            _presentation.EnsureAttachedCommentSlide(_slide);
            if (_comment.Parent == null
                || _slide.SlidePart.SlideCommentsPart?.CommentList?
                    .Elements<P.Comment>().Contains(_comment) != true) {
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
        /// <exception cref="NotSupportedException">The imported comment uses rich text markup that cannot be replaced without losing formatting.</exception>
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
            set {
                if (!value.HasValue) {
                    throw new ArgumentNullException(nameof(value),
                        "Modern comments require a creation timestamp.");
                }
                RequireAttached().Created = value.Value;
            }
        }

        /// <summary>Comment X position within the DrawingML coordinate range.</summary>
        public long X {
            get => RequireAttached().GetFirstChild<P188.Point2DType>()?
                .X?.Value ?? 0L;
            set {
                PowerPointPresentation.ValidateModernCommentPosition(
                    value, nameof(value));
                RequirePosition().X = value;
            }
        }

        /// <summary>Comment Y position within the DrawingML coordinate range.</summary>
        public long Y {
            get => RequireAttached().GetFirstChild<P188.Point2DType>()?
                .Y?.Value ?? 0L;
            set {
                PowerPointPresentation.ValidateModernCommentPosition(
                    value, nameof(value));
                RequirePosition().Y = value;
            }
        }

        /// <summary>Replies in package order.</summary>
        public IReadOnlyList<PowerPointModernCommentReply> Replies {
            get {
                P188.Comment comment = RequireAttached();
                return (IReadOnlyList<PowerPointModernCommentReply>)(comment
                    .GetFirstChild<P188.CommentReplyList>()?
                    .Elements<P188.CommentReply>()
                    .Select(reply => new PowerPointModernCommentReply(
                        _presentation, _slide, _part, comment, reply)).ToArray()
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
            P188.CommentStatus modernStatus =
                PowerPointPresentation.ToModernStatus(status);
            P188.Author modernAuthor = _presentation.GetOrCreateModernCommentAuthor(author);
            var reply = new P188.CommentReply(
                PowerPointPresentation.CreateModernCommentTextBody(text)) {
                Id = PowerPointPresentation.CreateModernCommentId(),
                AuthorId = modernAuthor.Id?.Value,
                Status = modernStatus,
                Created = created ?? DateTime.UtcNow
            };
            P188.CommentReplyList? replies = comment.GetFirstChild<P188.CommentReplyList>();
            if (replies == null) {
                replies = new P188.CommentReplyList();
                comment.AddChild(replies, true);
            }
            replies.Append(reply);
            return new PowerPointModernCommentReply(_presentation, _slide, _part,
                comment, reply);
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
            bool isLastComment = _part.CommentList == null
                || !_part.CommentList.Elements<P188.Comment>()
                    .Any(comment => !ReferenceEquals(comment, attached));
            if (isLastComment && _part.CommentList != null
                && (_part.CommentList.ExtendedAttributes.Any()
                    || _part.CommentList.ChildElements.Any(child =>
                        child is not P188.Comment))) {
                throw new NotSupportedException(
                    "The last modern comment cannot be removed while its comment part contains producer metadata that must be preserved.");
            }
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
            _presentation.EnsureAttachedCommentSlide(_slide);
            if (_comment.Parent == null
                || !_slide.SlidePart.Parts.Any(pair =>
                    ReferenceEquals(pair.OpenXmlPart, _part))
                || _part.CommentList?.Elements<P188.Comment>()
                    .Contains(_comment) != true) {
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
        private readonly PowerPointSlide _slide;
        private readonly PowerPointCommentPart _part;
        private readonly P188.Comment _parent;
        private readonly P188.CommentReply _reply;

        internal PowerPointModernCommentReply(PowerPointPresentation presentation,
            PowerPointSlide slide, PowerPointCommentPart part,
            P188.Comment parent, P188.CommentReply reply) {
            _presentation = presentation;
            _slide = slide;
            _part = part;
            _parent = parent;
            _reply = reply;
        }

        /// <summary>Stable reply identifier.</summary>
        public string Id => RequireAttached().Id?.Value ?? string.Empty;

        /// <summary>Reply author.</summary>
        public PowerPointCommentAuthor Author =>
            _presentation.ResolveModernCommentAuthor(RequireAttached().AuthorId?.Value);

        /// <summary>Visible reply text.</summary>
        /// <exception cref="NotSupportedException">The imported reply uses rich text markup that cannot be replaced without losing formatting.</exception>
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
            set {
                if (!value.HasValue) {
                    throw new ArgumentNullException(nameof(value),
                        "Modern comment replies require a creation timestamp.");
                }
                RequireAttached().Created = value.Value;
            }
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
            _presentation.EnsureAttachedCommentSlide(_slide);
            P188.CommentReplyList? replyList =
                _parent.GetFirstChild<P188.CommentReplyList>();
            if (_reply.Parent == null || _parent.Parent == null
                || !_slide.SlidePart.Parts.Any(pair =>
                    ReferenceEquals(pair.OpenXmlPart, _part))
                || _part.CommentList?.Elements<P188.Comment>()
                    .Contains(_parent) != true
                || replyList?.Elements<P188.CommentReply>()
                    .Contains(_reply) != true) {
                throw new InvalidOperationException("The modern comment reply is no longer attached to the presentation.");
            }
            return _reply;
        }
    }

    public sealed partial class PowerPointPresentation {
        private const int MaxModernCommentTextLength = 32000;
        private const int MaxModernCommentParagraphCount = 1024;
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
            ValidateModernCommentPosition(x, nameof(x));
            ValidateModernCommentPosition(y, nameof(y));
            P188.CommentStatus modernStatus = ToModernStatus(status);
            P188.Author modernAuthor = GetOrCreateModernCommentAuthor(author);
            var comment = new P188.Comment(
                new P188.CommentUnknownAnchor(),
                new P188.Point2DType { X = x, Y = y },
                CreateModernCommentTextBody(text)) {
                Id = CreateModernCommentId(),
                AuthorId = modernAuthor.Id?.Value,
                Status = modernStatus,
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
                    author.ClassicInitials, StringComparison.Ordinal));
            if (existing != null) return existing;
            uint id = AllocateClassicAuthorId(part.CommentAuthorList);
            var created = new P.CommentAuthor {
                Id = id,
                Name = author.Name,
                Initials = author.ClassicInitials,
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
            const uint MaximumClassicCommentIndex = int.MaxValue;
            uint firstCandidate = author.LastIndex?.Value >=
                MaximumClassicCommentIndex
                ? 0U
                : (author.LastIndex?.Value ?? 0U) + 1U;
            uint next = FindAvailableUInt32Id(usedIndexes, firstCandidate,
                "classic-comment", MaximumClassicCommentIndex);
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
                : PowerPointCommentAuthor.FromImportedClassic(
                    author.Name?.Value ?? "Unknown", author.Initials?.Value);
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
            ValidateCommentAuthorIdentity(author);
            PowerPointAuthorsPart[] parts = _presentationPart.Parts
                .Select(pair => pair.OpenXmlPart).OfType<PowerPointAuthorsPart>()
                .ToArray();
            foreach (PowerPointAuthorsPart part in parts) {
                P188.Author? existing = part.AuthorList?.Elements<P188.Author>()
                    .FirstOrDefault(candidate => ModernAuthorExactlyMatches(candidate, author));
                if (existing != null) return existing;
            }
            foreach (PowerPointAuthorsPart part in parts) {
                P188.Author? existing = part.AuthorList?.Elements<P188.Author>()
                    .FirstOrDefault(candidate => ModernAuthorMatchesCreatedDefaults(candidate, author));
                if (existing != null) return existing;
            }
            PowerPointAuthorsPart target = parts.FirstOrDefault()
                ?? _presentationPart.AddNewPart<PowerPointAuthorsPart>();
            if (target.AuthorList == null) target.AuthorList = new P188.AuthorList();
            var created = new P188.Author {
                Id = CreateModernCommentId(),
                Name = author.Name,
                Initials = author.ModernInitials,
                UserId = ResolveModernCommentUserId(author),
                ProviderId = ResolveModernCommentProviderId(author)
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
                : PowerPointCommentAuthor.FromImportedModern(
                    author.Name?.Value ?? "Unknown",
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

        internal void GetCommentAuthorIdsForSlideRemoval(
            PowerPointSlide slide, out uint[] classicAuthorIds,
            out string[] modernAuthorIds) {
            classicAuthorIds = slide.SlidePart.SlideCommentsPart?.CommentList?
                .Elements<P.Comment>()
                .Where(comment => comment.AuthorId?.Value != null)
                .Select(comment => comment.AuthorId!.Value)
                .Distinct()
                .ToArray() ?? Array.Empty<uint>();
            modernAuthorIds = slide.SlidePart.Parts
                .Select(pair => pair.OpenXmlPart)
                .OfType<PowerPointCommentPart>()
                .SelectMany(part => part.CommentList?.Elements<P188.Comment>()
                    ?? Enumerable.Empty<P188.Comment>())
                .SelectMany(comment => new[] { comment.AuthorId?.Value }
                    .Concat(comment.Descendants<P188.CommentReply>()
                        .Select(reply => reply.AuthorId?.Value)))
                .Where(authorId => !string.IsNullOrWhiteSpace(authorId))
                .Select(authorId => authorId!)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToArray();
        }

        internal void ReconcileCommentAuthorsAfterSlideRemoval(
            IEnumerable<uint> classicAuthorIds,
            IEnumerable<string> modernAuthorIds) {
            foreach (uint authorId in classicAuthorIds) {
                RemoveClassicCommentAuthorIfUnused(authorId);
            }
            foreach (string authorId in modernAuthorIds) {
                RemoveModernCommentAuthorIfUnused(authorId);
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
            if (!HasPlainModernCommentTextBody(body)) {
                throw new NotSupportedException(
                    "Rich modern comment text cannot be replaced without discarding its formatting.");
            }
            body.RemoveAllChildren<A.Paragraph>();
            body.Append(CreateModernCommentParagraphs(text));
        }

        internal static bool HasPlainModernCommentTextBody(
            P188.TextBodyType body) {
            if (body == null) return false;
            if (body.Elements<A.BodyProperties>().Count() != 1
                || body.Elements<A.ListStyle>().Count() != 1
                || body.ChildElements.Any(child => child is not A.BodyProperties
                    && child is not A.ListStyle
                    && child is not A.Paragraph)) {
                return false;
            }

            foreach (A.Paragraph paragraph in body.Elements<A.Paragraph>()) {
                if (paragraph.ChildElements.Count == 0) continue;
                if (paragraph.ChildElements.Count != 1
                    || paragraph.FirstChild is not A.Run run
                    || run.ChildElements.Count != 1
                    || run.FirstChild is not A.Text) {
                    return false;
                }
            }
            return true;
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
            if (text.Length > MaxModernCommentTextLength) {
                throw new ArgumentException(
                    $"Modern comment text cannot exceed {MaxModernCommentTextLength:N0} characters.",
                    nameof(text));
            }
            int paragraphCount = 1;
            for (int index = 0; index < text.Length; index++) {
                if (text[index] == '\r') {
                    if (index + 1 < text.Length && text[index + 1] == '\n') index++;
                    paragraphCount++;
                } else if (text[index] == '\n') {
                    paragraphCount++;
                }
                if (paragraphCount > MaxModernCommentParagraphCount) {
                    throw new ArgumentException(
                        $"Modern comment text cannot exceed {MaxModernCommentParagraphCount:N0} paragraphs.",
                        nameof(text));
                }
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

        internal static void ValidateModernCommentPosition(long value,
            string parameterName) => PowerPointDrawingValueValidator
            .ValidateCoordinate(value, parameterName,
                "Modern comment positions");

        private static void ValidateClassicCommentAuthor(
            PowerPointCommentAuthor author) {
            ValidateCommentAuthorIdentity(author);
            if (author.Name.Length > 52
                || (author.ClassicInitials?.Length ?? 0) > 52) {
                throw new ArgumentException(
                    "Classic comment author names and initials cannot exceed 52 characters.",
                    nameof(author));
            }
            if (author.Name.IndexOf('\0') >= 0
                || (author.ClassicInitials?.IndexOf('\0') ?? -1) >= 0) {
                throw new ArgumentException(
                    "Classic comment author names and initials cannot contain a NUL character.",
                    nameof(author));
            }
        }

        private static void ValidateCommentAuthorIdentity(
            PowerPointCommentAuthor author) {
            if (author == null) throw new ArgumentNullException(nameof(author));
            PowerPointXmlValueValidator.ValidateCharacters(author.Name,
                nameof(author), "Author name");
            PowerPointXmlValueValidator.ValidateCharacters(author.Initials,
                nameof(author), "Author initials");
            if (author.UserId != null) {
                PowerPointXmlValueValidator.ValidateCharacters(author.UserId,
                    nameof(author), "Author user identifier");
            }
            if (author.ProviderId != null) {
                PowerPointXmlValueValidator.ValidateCharacters(author.ProviderId,
                    nameof(author), "Author provider identifier");
            }
        }

        private void EnsureCommentSlide(PowerPointSlide slide) {
            ThrowIfDisposed();
            if (slide == null) throw new ArgumentNullException(nameof(slide));
            if (!_slides.Contains(slide)) {
                throw new ArgumentException("The slide does not belong to this presentation.", nameof(slide));
            }
        }

        internal void EnsureAttachedCommentSlide(PowerPointSlide slide) {
            if (!_slides.Contains(slide)) {
                throw new InvalidOperationException(
                    "The comment's slide is no longer attached to the presentation.");
            }
        }

        private static uint AllocateClassicAuthorId(P.CommentAuthorList authors) {
            var usedIds = new HashSet<uint>(authors.Elements<P.CommentAuthor>()
                .Where(author => author.Id?.Value != null)
                .Select(author => author.Id!.Value));
            return FindAvailableUInt32Id(usedIds, 0U,
                "classic-comment-author");
        }

        private static bool ModernAuthorExactlyMatches(P188.Author candidate,
            PowerPointCommentAuthor author) =>
            string.Equals(candidate.Name?.Value, author.Name, StringComparison.Ordinal)
            && string.Equals(candidate.Initials?.Value, author.ModernInitials, StringComparison.Ordinal)
            && string.Equals(candidate.UserId?.Value, author.UserId, StringComparison.Ordinal)
            && string.Equals(candidate.ProviderId?.Value, author.ProviderId, StringComparison.Ordinal);

        private static bool ModernAuthorMatchesCreatedDefaults(P188.Author candidate,
            PowerPointCommentAuthor author) =>
            string.Equals(candidate.Name?.Value, author.Name, StringComparison.Ordinal)
            && string.Equals(candidate.Initials?.Value, author.ModernInitials, StringComparison.Ordinal)
            && string.Equals(candidate.UserId?.Value,
                ResolveModernCommentUserId(author), StringComparison.Ordinal)
            && string.Equals(candidate.ProviderId?.Value,
                ResolveModernCommentProviderId(author), StringComparison.Ordinal);

        private static string ResolveModernCommentUserId(
            PowerPointCommentAuthor author) => author.UserId ?? author.Name;

        private static string ResolveModernCommentProviderId(
            PowerPointCommentAuthor author) => author.ProviderId ?? "OfficeIMO";
    }
}
