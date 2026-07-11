using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using P188 = DocumentFormat.OpenXml.Office2021.PowerPoint.Comment;

namespace OfficeIMO.PowerPoint {
    /// <summary>Comment technology represented by a review item.</summary>
    public enum PowerPointCommentKind {
        /// <summary>Classic PowerPoint comment.</summary>
        Classic,
        /// <summary>Modern threaded PowerPoint comment.</summary>
        Modern,
        /// <summary>Reply to a modern threaded comment.</summary>
        ModernReply
    }

    /// <summary>Read-only typed projection of a classic or modern PowerPoint comment.</summary>
    public sealed class PowerPointReviewComment {
        internal PowerPointReviewComment(PowerPointCommentKind kind, int slideNumber, string id,
            string? parentId, string? authorId, string? authorName, string text, DateTime? created,
            string? status, uint? shapeId, double? x, double? y) {
            Kind = kind;
            SlideNumber = slideNumber;
            Id = id;
            ParentId = parentId;
            AuthorId = authorId;
            AuthorName = authorName;
            Text = text;
            Created = created;
            Status = status;
            ShapeId = shapeId;
            X = x;
            Y = y;
        }

        /// <summary>Comment technology.</summary>
        public PowerPointCommentKind Kind { get; }
        /// <summary>1-based slide number.</summary>
        public int SlideNumber { get; }
        /// <summary>Stable comment identifier within its technology.</summary>
        public string Id { get; }
        /// <summary>Parent comment identifier for replies.</summary>
        public string? ParentId { get; }
        /// <summary>Author identifier.</summary>
        public string? AuthorId { get; }
        /// <summary>Resolved author display name.</summary>
        public string? AuthorName { get; }
        /// <summary>Visible comment text.</summary>
        public string Text { get; }
        /// <summary>Creation timestamp when present.</summary>
        public DateTime? Created { get; }
        /// <summary>Modern comment status when present.</summary>
        public string? Status { get; }
        /// <summary>Anchored shape identifier when present.</summary>
        public uint? ShapeId { get; }
        /// <summary>Classic comment X position when present.</summary>
        public double? X { get; }
        /// <summary>Classic comment Y position when present.</summary>
        public double? Y { get; }
    }

    /// <summary>Read-only review inspection across classic and modern PowerPoint comments.</summary>
    public sealed class PowerPointReviewReport {
        internal PowerPointReviewReport(IList<PowerPointReviewComment> comments) {
            Comments = new ReadOnlyCollection<PowerPointReviewComment>(
                new List<PowerPointReviewComment>(comments));
        }

        /// <summary>Report schema version.</summary>
        public int SchemaVersion => 1;
        /// <summary>Comments in slide and package order.</summary>
        public IReadOnlyList<PowerPointReviewComment> Comments { get; }
        /// <summary>Classic comment count.</summary>
        public int ClassicCount => Comments.Count(comment => comment.Kind == PowerPointCommentKind.Classic);
        /// <summary>Modern comment and reply count.</summary>
        public int ModernCount => Comments.Count(comment => comment.Kind != PowerPointCommentKind.Classic);
        /// <summary>Whether the deck contains review content.</summary>
        public bool HasComments => Comments.Count > 0;

        /// <summary>Serializes review metadata as deterministic JSON.</summary>
        public string ToJson() {
            var json = new StringBuilder();
            json.Append("{\"schemaVersion\":1,\"commentCount\":").Append(Comments.Count)
                .Append(",\"comments\":[");
            for (int index = 0; index < Comments.Count; index++) {
                PowerPointReviewComment comment = Comments[index];
                json.Append('{')
                    .Append("\"kind\":\"").Append(comment.Kind).Append("\",")
                    .Append("\"slideNumber\":").Append(comment.SlideNumber).Append(',')
                    .Append("\"id\":\"").Append(Escape(comment.Id)).Append("\",")
                    .Append("\"authorName\":").Append(JsonString(comment.AuthorName)).Append(',')
                    .Append("\"text\":\"").Append(Escape(comment.Text)).Append("\",")
                    .Append("\"status\":").Append(JsonString(comment.Status)).Append('}');
                if (index < Comments.Count - 1) json.Append(',');
            }
            return json.Append("]}").ToString();
        }

        private static string JsonString(string? value) => value == null ? "null" : "\"" + Escape(value) + "\"";
        private static string Escape(string value) => (value ?? string.Empty).Replace("\\", "\\\\")
            .Replace("\"", "\\\"").Replace("\r", "\\r").Replace("\n", "\\n").Replace("\t", "\\t");
    }

    /// <summary>Detected animation/timing node kind.</summary>
    public enum PowerPointAnimationKind {
        /// <summary>Sequential timing container.</summary>
        Sequence,
        /// <summary>Parallel timing container.</summary>
        Parallel,
        /// <summary>Property animation.</summary>
        Animate,
        /// <summary>Color animation.</summary>
        AnimateColor,
        /// <summary>Visual effect animation.</summary>
        AnimateEffect,
        /// <summary>Motion path animation.</summary>
        AnimateMotion,
        /// <summary>Rotation animation.</summary>
        AnimateRotation,
        /// <summary>Scale animation.</summary>
        AnimateScale,
        /// <summary>Set-value timing action.</summary>
        Set,
        /// <summary>Command timing action.</summary>
        Command,
        /// <summary>Audio or video timing action.</summary>
        Media,
        /// <summary>Unclassified timing node.</summary>
        Other
    }

    /// <summary>Typed read-only projection of one animation/timing node.</summary>
    public sealed class PowerPointAnimationNode {
        internal PowerPointAnimationNode(int slideNumber, PowerPointAnimationKind kind, string elementName,
            string? timingId, uint? shapeId, string? shapeName, string? trigger, string? delay,
            string? duration, string? presetClass, string? presetId, string? presetSubtype) {
            SlideNumber = slideNumber;
            Kind = kind;
            ElementName = elementName;
            TimingId = timingId;
            ShapeId = shapeId;
            ShapeName = shapeName;
            Trigger = trigger;
            Delay = delay;
            Duration = duration;
            PresetClass = presetClass;
            PresetId = presetId;
            PresetSubtype = presetSubtype;
        }

        /// <summary>1-based slide number.</summary>
        public int SlideNumber { get; }
        /// <summary>Detected animation kind.</summary>
        public PowerPointAnimationKind Kind { get; }
        /// <summary>Original OOXML local element name.</summary>
        public string ElementName { get; }
        /// <summary>Common timing node identifier.</summary>
        public string? TimingId { get; }
        /// <summary>Target shape identifier.</summary>
        public uint? ShapeId { get; }
        /// <summary>Resolved target shape name.</summary>
        public string? ShapeName { get; }
        /// <summary>Detected trigger event.</summary>
        public string? Trigger { get; }
        /// <summary>Authored delay value.</summary>
        public string? Delay { get; }
        /// <summary>Authored duration value.</summary>
        public string? Duration { get; }
        /// <summary>PowerPoint preset class.</summary>
        public string? PresetClass { get; }
        /// <summary>PowerPoint preset identifier.</summary>
        public string? PresetId { get; }
        /// <summary>PowerPoint preset subtype.</summary>
        public string? PresetSubtype { get; }
    }

    /// <summary>Read-only animation inspection report.</summary>
    public sealed class PowerPointAnimationReport {
        internal PowerPointAnimationReport(IList<PowerPointAnimationNode> nodes) {
            Nodes = new ReadOnlyCollection<PowerPointAnimationNode>(
                new List<PowerPointAnimationNode>(nodes));
        }

        /// <summary>Report schema version.</summary>
        public int SchemaVersion => 1;
        /// <summary>Timing nodes in slide and tree order.</summary>
        public IReadOnlyList<PowerPointAnimationNode> Nodes { get; }
        /// <summary>Whether animation/timing markup was found.</summary>
        public bool HasAnimations => Nodes.Count > 0;
    }

    public sealed partial class PowerPointPresentation {
        /// <summary>Inspects classic and modern comments without mutating review markup.</summary>
        public PowerPointReviewReport InspectReviewComments() {
            ThrowIfDisposed();
            var comments = new List<PowerPointReviewComment>();
            Dictionary<uint, string> classicAuthors = GetClassicAuthors();
            Dictionary<string, string> modernAuthors = GetModernAuthors();
            for (int index = 0; index < _slides.Count; index++) {
                AddClassicComments(_slides[index], index + 1, classicAuthors, comments);
                AddModernComments(_slides[index], index + 1, modernAuthors, comments);
            }
            return new PowerPointReviewReport(comments);
        }

        /// <summary>Inspects slide timing trees before any animation authoring is attempted.</summary>
        public PowerPointAnimationReport InspectAnimations() {
            ThrowIfDisposed();
            var nodes = new List<PowerPointAnimationNode>();
            for (int index = 0; index < _slides.Count; index++) {
                PowerPointSlide slide = _slides[index];
                Timing? timing = slide.SlidePart.Slide?.Timing;
                if (timing == null) continue;
                foreach (OpenXmlElement element in timing.Descendants().Where(IsAnimationElement)) {
                    OpenXmlElement? common = FindOwnedAnimationDescendant(element, "cTn");
                    OpenXmlElement? target = FindOwnedAnimationDescendant(element, "spTgt");
                    uint? shapeId = ParseUInt(GetAttribute(target, "spid"));
                    string? shapeName = shapeId.HasValue
                        ? slide.EnumerateShapesDeep(slide.Shapes.Concat(slide.GetInheritedShapesForExport()))
                            .FirstOrDefault(shape => shape.Id == shapeId)?.Name
                        : null;
                    OpenXmlElement? condition = FindOwnedAnimationDescendant(element, "cond");
                    nodes.Add(new PowerPointAnimationNode(index + 1, MapAnimationKind(element.LocalName),
                        element.LocalName, GetAttribute(common, "id"), shapeId, shapeName,
                        GetAttribute(condition, "evt"), GetAttribute(condition, "delay"),
                        GetAttribute(common, "dur"), GetAttribute(common, "presetClass"),
                        GetAttribute(common, "presetID"), GetAttribute(common, "presetSubtype")));
                }
            }
            return new PowerPointAnimationReport(nodes);
        }

        private static OpenXmlElement? FindOwnedAnimationDescendant(OpenXmlElement animation, string localName) =>
            animation.Descendants().FirstOrDefault(candidate =>
                candidate.LocalName == localName &&
                ReferenceEquals(candidate.Ancestors().FirstOrDefault(IsAnimationElement), animation));

        private Dictionary<uint, string> GetClassicAuthors() {
            var result = new Dictionary<uint, string>();
            CommentAuthorList? list = _presentationPart.CommentAuthorsPart?.CommentAuthorList;
            if (list == null) return result;
            foreach (CommentAuthor author in list.Elements<CommentAuthor>()) {
                if (author.Id?.Value != null) result[author.Id.Value] = author.Name?.Value ?? string.Empty;
            }
            return result;
        }

        private Dictionary<string, string> GetModernAuthors() {
            var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (PowerPointAuthorsPart part in _presentationPart.Parts.Select(pair => pair.OpenXmlPart)
                         .OfType<PowerPointAuthorsPart>()) {
                P188.AuthorList? list = part.AuthorList;
                if (list == null) continue;
                foreach (P188.Author author in list.Elements<P188.Author>()) {
                    string? id = author.Id?.Value;
                    if (!string.IsNullOrWhiteSpace(id)) result[id!] = author.Name?.Value ?? string.Empty;
                }
            }
            return result;
        }

        private static void AddClassicComments(PowerPointSlide slide, int slideNumber,
            IDictionary<uint, string> authors, IList<PowerPointReviewComment> destination) {
            CommentList? list = slide.SlidePart.SlideCommentsPart?.CommentList;
            if (list == null) return;
            foreach (Comment comment in list.Elements<Comment>()) {
                uint? authorId = comment.AuthorId?.Value;
                destination.Add(new PowerPointReviewComment(PowerPointCommentKind.Classic, slideNumber,
                    (comment.Index?.Value ?? 0U).ToString(CultureInfo.InvariantCulture), null,
                    authorId?.ToString(CultureInfo.InvariantCulture),
                    authorId.HasValue && authors.TryGetValue(authorId.Value, out string? author) ? author : null,
                    comment.Text?.Text ?? comment.InnerText ?? string.Empty, comment.DateTime?.Value,
                    null, null, comment.Position?.X?.Value, comment.Position?.Y?.Value));
            }
        }

        private static void AddModernComments(PowerPointSlide slide, int slideNumber,
            IDictionary<string, string> authors, IList<PowerPointReviewComment> destination) {
            foreach (PowerPointCommentPart part in slide.SlidePart.Parts.Select(pair => pair.OpenXmlPart)
                         .OfType<PowerPointCommentPart>()) {
                P188.CommentList? list = part.CommentList;
                if (list == null) continue;
                foreach (P188.Comment comment in list.Elements<P188.Comment>()) {
                    string id = comment.Id?.Value ?? "modern-" + slideNumber + "-" + destination.Count;
                    string? authorId = comment.AuthorId?.Value;
                    destination.Add(CreateModernComment(PowerPointCommentKind.Modern, slideNumber, id, null,
                        authorId, authors, comment,
                        GetAttribute(comment, "status") ?? comment.Status?.InnerText,
                        comment.Created?.Value));
                    P188.CommentReplyList? replies = comment.GetFirstChild<P188.CommentReplyList>();
                    if (replies == null) continue;
                    foreach (P188.CommentReply reply in replies.Elements<P188.CommentReply>()) {
                        string replyId = reply.Id?.Value ?? "modern-reply-" + slideNumber + "-" + destination.Count;
                        string? replyAuthor = reply.AuthorId?.Value;
                        destination.Add(CreateModernComment(PowerPointCommentKind.ModernReply, slideNumber,
                            replyId, id, replyAuthor, authors, reply,
                            GetAttribute(reply, "status") ?? reply.Status?.InnerText,
                            reply.Created?.Value));
                    }
                }
            }
        }

        private static PowerPointReviewComment CreateModernComment(PowerPointCommentKind kind, int slideNumber,
            string id, string? parentId, string? authorId, IDictionary<string, string> authors,
            OpenXmlElement element, string? status, DateTime? created) {
            string text = string.Concat(element.Descendants<A.Text>()
                .Where(item => item.Ancestors<P188.CommentReply>()
                    .All(reply => ReferenceEquals(reply, element)))
                .Select(item => item.Text));
            uint? shapeId = element.Descendants().SelectMany(item => item.GetAttributes())
                .Where(attribute => attribute.LocalName == "spid" || attribute.LocalName == "shapeId")
                .Select(attribute => ParseUInt(attribute.Value)).FirstOrDefault(value => value.HasValue);
            return new PowerPointReviewComment(kind, slideNumber, id, parentId, authorId,
                authorId != null && authors.TryGetValue(authorId, out string? author) ? author : null,
                text, created, status, shapeId, null, null);
        }

        private static bool IsAnimationElement(OpenXmlElement element) {
            switch (element.LocalName) {
                case "seq": case "par": case "anim": case "animClr": case "animEffect":
                case "animMotion": case "animRot": case "animScale": case "set": case "cmd":
                case "audio": case "video": return true;
                default: return false;
            }
        }

        private static PowerPointAnimationKind MapAnimationKind(string localName) {
            switch (localName) {
                case "seq": return PowerPointAnimationKind.Sequence;
                case "par": return PowerPointAnimationKind.Parallel;
                case "anim": return PowerPointAnimationKind.Animate;
                case "animClr": return PowerPointAnimationKind.AnimateColor;
                case "animEffect": return PowerPointAnimationKind.AnimateEffect;
                case "animMotion": return PowerPointAnimationKind.AnimateMotion;
                case "animRot": return PowerPointAnimationKind.AnimateRotation;
                case "animScale": return PowerPointAnimationKind.AnimateScale;
                case "set": return PowerPointAnimationKind.Set;
                case "cmd": return PowerPointAnimationKind.Command;
                case "audio": case "video": return PowerPointAnimationKind.Media;
                default: return PowerPointAnimationKind.Other;
            }
        }

        private static string? GetAttribute(OpenXmlElement? element, string name) {
            if (element == null) return null;
            OpenXmlAttribute attribute = element.GetAttributes()
                .FirstOrDefault(item => string.Equals(item.LocalName, name, StringComparison.OrdinalIgnoreCase));
            return string.IsNullOrWhiteSpace(attribute.Value) ? null : attribute.Value;
        }

        private static uint? ParseUInt(string? value) => uint.TryParse(value,
            NumberStyles.Integer, CultureInfo.InvariantCulture, out uint parsed) ? parsed : (uint?)null;
    }
}
