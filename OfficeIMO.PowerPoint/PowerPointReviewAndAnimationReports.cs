using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading;
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

    /// <summary>Safety limits for inspecting untrusted presentation timing trees.</summary>
    public sealed class PowerPointAnimationInspectionOptions {
        /// <summary>Maximum XML elements visited across all slide timing trees.</summary>
        public int MaxXmlElements { get; set; } = 100_000;
        /// <summary>Maximum timing-tree nesting depth visited.</summary>
        public int MaxXmlDepth { get; set; } = 128;
        /// <summary>Maximum animation nodes projected into the report.</summary>
        public int MaxAnimationNodes { get; set; } = 10_000;
        /// <summary>Optional cancellation token checked during inspection.</summary>
        public CancellationToken CancellationToken { get; set; }

        internal void Validate() {
            if (MaxXmlElements <= 0) throw new ArgumentOutOfRangeException(nameof(MaxXmlElements));
            if (MaxXmlDepth <= 0) throw new ArgumentOutOfRangeException(nameof(MaxXmlDepth));
            if (MaxAnimationNodes <= 0) throw new ArgumentOutOfRangeException(nameof(MaxAnimationNodes));
        }
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
        public PowerPointAnimationReport InspectAnimations() => InspectAnimations(new PowerPointAnimationInspectionOptions());

        /// <summary>Inspects slide timing trees with explicit traversal and result limits.</summary>
        public PowerPointAnimationReport InspectAnimations(PowerPointAnimationInspectionOptions options) {
            ThrowIfDisposed();
            if (options == null) throw new ArgumentNullException(nameof(options));
            options.Validate();
            var nodes = new List<PowerPointAnimationNode>();
            int visitedElements = 0;
            for (int index = 0; index < _slides.Count; index++) {
                options.CancellationToken.ThrowIfCancellationRequested();
                PowerPointSlide slide = _slides[index];
                Timing? timing = slide.SlidePart.Slide?.Timing;
                if (timing == null) continue;
                var slideItems = new List<AnimationInspectionItem>();
                InspectAnimationTree(timing, options, ref visitedElements, nodes.Count, slideItems);
                uint?[] shapeIds = slideItems
                    .Select(item => ParseUInt(GetAttribute(item.Target, "spid")))
                    .ToArray();
                Dictionary<uint, string?> shapeNames = ResolveAnimationShapeNames(slide, shapeIds);
                for (int itemIndex = 0; itemIndex < slideItems.Count; itemIndex++) {
                    AnimationInspectionItem item = slideItems[itemIndex];
                    uint? shapeId = shapeIds[itemIndex];
                    string? shapeName = shapeId.HasValue && shapeNames.TryGetValue(shapeId.Value, out string? resolvedName)
                        ? resolvedName
                        : null;
                    nodes.Add(new PowerPointAnimationNode(index + 1, MapAnimationKind(item.Element.LocalName),
                        item.Element.LocalName, GetAttribute(item.Common, "id"), shapeId, shapeName,
                        GetAttribute(item.Condition, "evt"), GetAttribute(item.Condition, "delay"),
                        GetAttribute(item.Common, "dur"), GetAttribute(item.Common, "presetClass"),
                        GetAttribute(item.Common, "presetID"), GetAttribute(item.Common, "presetSubtype")));
                }
            }
            return new PowerPointAnimationReport(nodes);
        }

        private static Dictionary<uint, string?> ResolveAnimationShapeNames(
            PowerPointSlide slide,
            IEnumerable<uint?> shapeIds) {
            var unresolvedShapeIds = new HashSet<uint>(shapeIds.Where(shapeId => shapeId.HasValue)
                .Select(shapeId => shapeId!.Value));
            var shapeNames = new Dictionary<uint, string?>();
            if (unresolvedShapeIds.Count == 0) return shapeNames;

            foreach (PowerPointShape shape in slide
                         .EnumerateShapesDeep(slide.Shapes.Concat(slide.GetInheritedShapesForExport()))) {
                if (!shape.Id.HasValue || !unresolvedShapeIds.Remove(shape.Id.Value)) continue;
                shapeNames.Add(shape.Id.Value, shape.Name);
                if (unresolvedShapeIds.Count == 0) break;
            }
            return shapeNames;
        }

        private static void InspectAnimationTree(Timing timing, PowerPointAnimationInspectionOptions options,
            ref int visitedElements, int existingAnimationCount, IList<AnimationInspectionItem> destination) {
            var stack = new Stack<(OpenXmlElement Element, int Depth, AnimationInspectionItem? Owner)>();
            int scheduledElements = visitedElements;
            for (int childIndex = timing.ChildElements.Count - 1; childIndex >= 0; childIndex--) {
                if (scheduledElements >= options.MaxXmlElements) {
                    throw new InvalidDataException("Animation inspection exceeded MaxXmlElements.");
                }
                stack.Push((timing.ChildElements[childIndex], 1, null));
                scheduledElements++;
            }
            while (stack.Count > 0) {
                (OpenXmlElement element, int depth, AnimationInspectionItem? owner) = stack.Pop();
                if ((visitedElements++ & 0xFF) == 0) options.CancellationToken.ThrowIfCancellationRequested();
                if (visitedElements > options.MaxXmlElements) throw new InvalidDataException("Animation inspection exceeded MaxXmlElements.");
                if (depth > options.MaxXmlDepth) throw new InvalidDataException("Animation inspection exceeded MaxXmlDepth.");

                AnimationInspectionItem? childOwner = owner;
                if (IsAnimationElement(element)) {
                    if (existingAnimationCount + destination.Count >= options.MaxAnimationNodes) {
                        throw new InvalidDataException("Animation inspection exceeded MaxAnimationNodes.");
                    }
                    childOwner = new AnimationInspectionItem(element);
                    destination.Add(childOwner);
                } else if (owner != null) {
                    if (element.LocalName == "cTn" && owner.Common == null) owner.Common = element;
                    else if (element.LocalName == "spTgt" && owner.Target == null) owner.Target = element;
                    else if (element.LocalName == "cond" && owner.Condition == null) owner.Condition = element;
                }

                for (int childIndex = element.ChildElements.Count - 1; childIndex >= 0; childIndex--) {
                    if (scheduledElements >= options.MaxXmlElements) {
                        throw new InvalidDataException("Animation inspection exceeded MaxXmlElements.");
                    }
                    if ((scheduledElements & 0xFF) == 0) options.CancellationToken.ThrowIfCancellationRequested();
                    stack.Push((element.ChildElements[childIndex], depth + 1, childOwner));
                    scheduledElements++;
                }
            }
        }

        private sealed class AnimationInspectionItem {
            internal AnimationInspectionItem(OpenXmlElement element) { Element = element; }
            internal OpenXmlElement Element { get; }
            internal OpenXmlElement? Common { get; set; }
            internal OpenXmlElement? Target { get; set; }
            internal OpenXmlElement? Condition { get; set; }
        }

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
