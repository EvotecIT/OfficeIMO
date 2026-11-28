using System.Xml;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;
using OfficeIMO.PowerPoint.OpenXml;

namespace OfficeIMO.PowerPoint {
    /// <summary>Start behavior for an authored timeline action.</summary>
    public enum PowerPointTimelineStart {
        /// <summary>Starts on the next click.</summary>
        OnClick,
        /// <summary>Starts automatically after the previous action.</summary>
        AfterPrevious,
        /// <summary>Starts with the previous action.</summary>
        WithPrevious
    }

    /// <summary>Common settings for incremental timeline authoring.</summary>
    public sealed class PowerPointTimelineOptions {
        /// <summary>Action duration in milliseconds.</summary>
        public uint DurationMilliseconds { get; set; } = 500U;
        /// <summary>Delay in milliseconds.</summary>
        public uint DelayMilliseconds { get; set; }
        /// <summary>Start behavior.</summary>
        public PowerPointTimelineStart Start { get; set; } = PowerPointTimelineStart.OnClick;
    }

    /// <summary>Stable handle for one typed timing action.</summary>
    public sealed class PowerPointTimelineAction {
        internal PowerPointTimelineAction(uint timingId,
            PowerPointAnimationKind kind, uint shapeId) {
            TimingId = timingId;
            Kind = kind;
            ShapeId = shapeId;
        }
        /// <summary>Native common timing-node identifier.</summary>
        public uint TimingId { get; }
        /// <summary>Typed action family.</summary>
        public PowerPointAnimationKind Kind { get; }
        /// <summary>Target shape identifier.</summary>
        public uint ShapeId { get; }
    }

    public partial class PowerPointSlide {
        private const int TimingNamespace = OfficeXmlNamespaces.PresentationML;
        private const int TimingCompatibilityNamespace = OfficeXmlNamespaces.MarkupCompatibility;

        /// <summary>Adds a native motion-path action without replacing the existing timeline.</summary>
        public PowerPointTimelineAction AddMotionAnimation(PowerPointShape shape,
            string path, PowerPointTimelineOptions? options = null) {
            ArgumentException.ThrowIfNullOrWhiteSpace(path);
            uint shapeId = ValidateTimelineTarget(shape);
            (OfficeXmlNode action, uint id) = AppendTimelineAction(shapeId, "animMotion", options);
            action.SetAttribute("origin", "layout");
            action.SetAttribute("path", path);
            action.SetAttribute("pathEditMode", "relative");
            return new PowerPointTimelineAction(id, PowerPointAnimationKind.AnimateMotion, shapeId);
        }

        /// <summary>Adds a native rotation action without replacing the existing timeline.</summary>
        public PowerPointTimelineAction AddRotationAnimation(PowerPointShape shape,
            double degrees, PowerPointTimelineOptions? options = null) {
            if (double.IsNaN(degrees) || double.IsInfinity(degrees)) throw new ArgumentOutOfRangeException(nameof(degrees));
            uint shapeId = ValidateTimelineTarget(shape);
            int angle = checked((int)Math.Round(degrees * 60000D));
            (OfficeXmlNode action, uint id) = AppendTimelineAction(shapeId, "animRot", options);
            action.SetInt32Attribute("by", angle);
            return new PowerPointTimelineAction(id, PowerPointAnimationKind.AnimateRotation, shapeId);
        }

        /// <summary>Adds a native command action targeted at a shape.</summary>
        public PowerPointTimelineAction AddCommandAnimation(PowerPointShape shape,
            string command, PowerPointTimelineOptions? options = null) =>
            AddCommandAnimation(shape, command, PowerPointAnimationCommand.Event, options);

        /// <summary>Adds a native command action of the requested command type.</summary>
        public PowerPointTimelineAction AddCommandAnimation(PowerPointShape shape,
            string command, PowerPointAnimationCommand type,
            PowerPointTimelineOptions? options = null) {
            ArgumentException.ThrowIfNullOrWhiteSpace(command);
            uint shapeId = ValidateTimelineTarget(shape);
            string commandType = type.ToOpenXmlToken();
            (OfficeXmlNode action, uint id) = AppendTimelineAction(shapeId, "cmd", options);
            action.SetAttribute("type", commandType);
            action.SetAttribute("cmd", command);
            return new PowerPointTimelineAction(id, PowerPointAnimationKind.Command, shapeId);
        }

        /// <summary>Changes duration on any typed or imported action with a common timing id.</summary>
        public bool SetAnimationDuration(uint timingId, uint durationMilliseconds) {
            PowerPointPartRef part = PartRef;
            OfficeXmlNode node = FindUniqueCommonTimeNode(part.Root, timingId);
            if (node.IsNone) return false;
            node.SetUInt32Attribute("dur", durationMilliseconds);
            part.MarkEdited();
            return true;
        }

        /// <summary>
        /// Removes exactly one typed action by common timing id. Unrelated siblings and
        /// unmodeled producer sequences are retained.
        /// </summary>
        public bool RemoveAnimation(uint timingId) {
            PowerPointPartRef part = PartRef.ForMovingXmlNodes();
            OfficeXmlNode node = FindUniqueCommonTimeNode(part.Root, timingId);
            OfficeXmlNode behavior = node.Parent;
            OfficeXmlNode action = behavior.Parent;
            if (!behavior.Is(TimingNamespace, "cBhvr") || action.NamespaceId != TimingNamespace ||
                action.LocalName is not ("animMotion" or "animRot" or "cmd")) return false;
            OfficeXmlNode actionList = action.Parent;
            OfficeXmlNode owner = action.AncestorOrSelf(TimingNamespace, "par");
            action.Remove();
            if (actionList.Is(TimingNamespace, "childTnLst") && actionList.FirstChild.IsNone) owner.Remove();
            part.MarkEdited();
            return true;
        }

        private uint ValidateTimelineTarget(PowerPointShape shape) {
            ArgumentNullException.ThrowIfNull(shape);
            if (!ReferenceEquals(shape.OwnerSlide, this)) throw new ArgumentException("The animation target must belong to this slide.", nameof(shape));
            return shape.Id ?? throw new InvalidOperationException("The animation target has no shape identifier.");
        }

        private (OfficeXmlNode Action, uint Id) AppendTimelineAction(uint shapeId, string actionName,
            PowerPointTimelineOptions? options) {
            PowerPointTimelineOptions resolved = options ?? new PowerPointTimelineOptions();
            ValidateTimelineOptions(resolved);

            PowerPointPartRef part = PartRef;
            OfficeXmlNode slide = part.Root;
            OfficeXmlNode timing = slide.Child(TimingNamespace, "timing");
            OfficeXmlNode list = timing.Child(TimingNamespace, "tnLst");
            OfficeXmlNode root = default;
            foreach (OfficeXmlNode parallel in list.Children(TimingNamespace, "par")) {
                OfficeXmlNode node = parallel.Child(TimingNamespace, "cTn");
                if (node.Attribute("nodeType") != "tmRoot") continue;
                root = node;
                break;
            }

            uint actionId = GetNextTimingId(slide);
            uint ownerId = ReserveTimelineOwnerId(actionId, root.IsNone);
            if (timing.IsNone) timing = InsertSlideTiming(slide);
            if (list.IsNone) list = AppendTimingChild(timing, "tnLst");
            if (root.IsNone) {
                root = AppendTimingChild(AppendTimingChild(list, "par"), "cTn");
                root.SetUInt32Attribute("id", actionId + 1);
                root.SetAttribute("dur", "indefinite");
                root.SetAttribute("restart", "never");
                root.SetAttribute("nodeType", "tmRoot");
            }

            OfficeXmlNode children = root.Child(TimingNamespace, "childTnLst");
            if (children.IsNone) children = AppendTimingChild(root, "childTnLst");
            OfficeXmlNode owner = AppendTimingChild(AppendTimingChild(children, "par"), "cTn");
            owner.SetUInt32Attribute("id", ownerId);
            owner.SetAttribute("dur", "indefinite");
            owner.SetAttribute("fill", "hold");
            owner.SetAttribute("nodeType", resolved.Start switch {
                PowerPointTimelineStart.OnClick => "clickEffect",
                PowerPointTimelineStart.WithPrevious => "withEffect",
                _ => "afterEffect"
            });

            OfficeXmlNode condition = AppendTimingChild(AppendTimingChild(owner, "stCondLst"), "cond");
            if (resolved.Start == PowerPointTimelineStart.OnClick) {
                condition.SetAttribute("evt", "onClick");
                AppendTimingChild(AppendTimingChild(condition, "tgtEl"), "sldTgt");
            }
            condition.SetUInt32Attribute("delay", resolved.DelayMilliseconds);
            OfficeXmlNode action = AppendTimingChild(AppendTimingChild(owner, "childTnLst"), actionName);
            OfficeXmlNode behavior = AppendTimingChild(action, "cBhvr");
            OfficeXmlNode actionTiming = AppendTimingChild(behavior, "cTn");
            actionTiming.SetUInt32Attribute("id", actionId);
            actionTiming.SetUInt32Attribute("dur", resolved.DurationMilliseconds);
            actionTiming.SetAttribute("fill", "hold");
            OfficeXmlNode target = AppendTimingChild(AppendTimingChild(behavior, "tgtEl"), "spTgt");
            target.SetUInt32Attribute("spid", shapeId);
            part.MarkEdited();
            return (action, actionId);
        }

        private static OfficeXmlNode AppendTimingChild(OfficeXmlNode parent, string name) =>
            parent.Append(parent.CreateElement(TimingNamespace, name));

        private static OfficeXmlNode InsertSlideTiming(OfficeXmlNode slide) {
            // A trailing unknown element (including a Morph compatibility wrapper) made the SDK prepend timing.
            // A presentation extension list remains the insertion anchor anywhere in the slide.
            OfficeXmlNode before = slide.Child(TimingNamespace, "extLst");
            if (before.IsNone) {
                OfficeXmlNode last = slide.LastChild;
                if (last.NamespaceId != TimingNamespace || last.LocalName is not ("cSld" or "clrMapOvr" or "transition")) {
                    before = slide.FirstChild;
                }
            }
            return slide.InsertBefore(slide.CreateElement(TimingNamespace, "timing"), before);
        }

        private static uint ReserveTimelineOwnerId(uint actionId, bool rootNeeded) {
            ulong ownerId = (ulong)actionId + (rootNeeded ? 2UL : 1UL);
            if (ownerId > uint.MaxValue) throw TimingIdsExhausted();
            return (uint)ownerId;
        }

        private uint GetNextTimingId() => GetNextTimingId(PartRef.Root);

        private static uint GetNextTimingId(OfficeXmlNode slide) {
            uint maxId = MaximumTimingId(slide);
            if (maxId == uint.MaxValue) throw TimingIdsExhausted();
            return maxId + 1;
        }

        private static uint MaximumTimingId(OfficeXmlNode slide, OfficeXmlNode omittedTiming = default) {
            uint maxId = 0;
            foreach (OfficeXmlNode node in TimingNodesNamed(slide, "cTn")) {
                if (omittedTiming.Exists && node.AncestorOrSelf(TimingNamespace, "timing") == omittedTiming) continue;
                if (ReadTimingId(node) is uint id && id > maxId) maxId = id;
            }
            return maxId;
        }

        private static InvalidOperationException TimingIdsExhausted() =>
            new("The imported timing identifier space is exhausted.");

        private static OfficeXmlNode FindUniqueCommonTimeNode(OfficeXmlNode slide, uint timingId) {
            OfficeXmlNode match = default;
            OfficeXmlNode timing = slide.Child(TimingNamespace, "timing");
            foreach (OfficeXmlNode node in TimingNodesNamed(timing, "cTn")) {
                if (ReadTimingId(node) != timingId) continue;
                if (match.Exists) return default;
                match = node;
            }
            return match;
        }

        private static uint? ReadTimingId(OfficeXmlNode node) =>
            node.Attribute("id") is { } value ? XmlConvert.ToUInt32(value) : null;

        internal static OfficeXmlNode FindMediaTimingNode(OfficeXmlNode slide, string shapeId) {
            OfficeXmlNode timing = slide.Child(TimingNamespace, "timing");
            foreach (OfficeXmlNode node in TimingNodesNamed(timing, "cMediaNode")) {
                if (node.Child(TimingNamespace, "tgtEl").Child(TimingNamespace, "spTgt").Attribute("spid") == shapeId) return node;
            }
            return default;
        }

        private static IEnumerable<OfficeXmlNode> TimingNodesNamed(OfficeXmlNode root, string? name) {
            if (root.IsNone) yield break;
            bool includeTargets = name is "sldTgt" or "spTgt";
            var pending = new Stack<(OfficeXmlNode Node, string Context)>();
            PushTimingChildren(root, root.LocalName, pending);
            while (pending.Count > 0) {
                (OfficeXmlNode node, string context) = pending.Pop();
                if (node.Is(TimingCompatibilityNamespace, "AlternateContent")) {
                    // With markup compatibility processing off, the SDK exposes schema-typed branch children
                    // in their parent's context. Unknown wrappers inside a branch still hide their contents.
                    for (OfficeXmlNode branch = node.LastChild; branch.Exists; branch = branch.PreviousSibling) {
                        if (branch.NamespaceId == TimingCompatibilityNamespace && branch.LocalName is "Choice" or "Fallback") {
                            PushTimingChildren(branch, context, pending);
                        }
                    }
                } else if (node.NamespaceId == TimingNamespace &&
                           (IsTimingChild(context, node.LocalName) || includeTargets && IsTimingTargetChild(context, node.LocalName))) {
                    if (name == null || node.LocalName == name) yield return node;
                    PushTimingChildren(node, node.LocalName, pending);
                }
            }
        }

        private static void PushTimingChildren(OfficeXmlNode parent, string context, Stack<(OfficeXmlNode Node, string Context)> pending) {
            for (OfficeXmlNode child = parent.LastChild; child.Exists; child = child.PreviousSibling) {
                pending.Push((child, context));
            }
        }

        private static bool IsTimingChild(string parent, string child) => parent switch {
            "sld" => child == "timing",
            "timing" => child is "tnLst" or "bldLst",
            "bldLst" => child == "bldP",
            "bldP" => child == "tmplLst",
            "tmplLst" => child == "tmpl",
            "tmpl" => child == "tnLst",
            "tnLst" => child == "par",
            "childTnLst" or "subTnLst" => child is "par" or "seq" or "excl" or "anim" or "animClr" or "animEffect"
                or "animMotion" or "animRot" or "animScale" or "cmd" or "set" or "audio" or "video",
            "par" or "seq" or "excl" or "cBhvr" or "cMediaNode" => child == "cTn",
            "anim" or "animClr" or "animEffect" or "animMotion" or "animRot" or "animScale" or "cmd" or "set" => child == "cBhvr",
            "audio" or "video" => child == "cMediaNode",
            "cTn" => child is "childTnLst" or "subTnLst",
            _ => false
        };

        private static bool IsTimingTargetChild(string parent, string child) => parent switch {
            "cBhvr" or "cMediaNode" or "cond" or "endSync" => child == "tgtEl",
            "tgtEl" => child is "spTgt" or "sldTgt",
            "cTn" => child is "stCondLst" or "endCondLst" or "endSync",
            "seq" => child is "prevCondLst" or "nextCondLst",
            "stCondLst" or "endCondLst" or "prevCondLst" or "nextCondLst" => child == "cond",
            _ => false
        };

        private static void ValidateTimelineOptions(PowerPointTimelineOptions options) {
            if (options.DurationMilliseconds == 0) throw new ArgumentOutOfRangeException(nameof(options), "Duration must be positive.");
            if (!Enum.IsDefined(typeof(PowerPointTimelineStart), options.Start)) throw new ArgumentOutOfRangeException(nameof(options));
        }
    }
}
