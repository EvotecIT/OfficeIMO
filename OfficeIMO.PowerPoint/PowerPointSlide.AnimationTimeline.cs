using System;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Presentation;

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
        /// <summary>Adds a native motion-path action without replacing the existing timeline.</summary>
        public PowerPointTimelineAction AddMotionAnimation(PowerPointShape shape,
            string path, PowerPointTimelineOptions? options = null) {
            if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("A motion path is required.", nameof(path));
            uint shapeId = ValidateTimelineTarget(shape);
            uint id = GetNextTimingId();
            CommonBehavior behavior = CreateTimelineBehavior(shapeId, id, options);
            var action = new AnimateMotion(behavior) {
                Origin = AnimateMotionBehaviorOriginValues.Layout,
                Path = path,
                PathEditMode = AnimateMotionPathEditModeValues.Relative
            };
            AppendTimelineAction(action, options);
            return new PowerPointTimelineAction(id,
                PowerPointAnimationKind.AnimateMotion, shapeId);
        }

        /// <summary>Adds a native rotation action without replacing the existing timeline.</summary>
        public PowerPointTimelineAction AddRotationAnimation(PowerPointShape shape,
            double degrees, PowerPointTimelineOptions? options = null) {
            if (double.IsNaN(degrees) || double.IsInfinity(degrees)) throw new ArgumentOutOfRangeException(nameof(degrees));
            uint shapeId = ValidateTimelineTarget(shape);
            uint id = GetNextTimingId();
            var action = new AnimateRotation(CreateTimelineBehavior(shapeId, id, options)) {
                By = checked((int)Math.Round(degrees * 60000D))
            };
            AppendTimelineAction(action, options);
            return new PowerPointTimelineAction(id,
                PowerPointAnimationKind.AnimateRotation, shapeId);
        }

        /// <summary>Adds a native command action targeted at a shape.</summary>
        public PowerPointTimelineAction AddCommandAnimation(PowerPointShape shape,
            string command, PowerPointTimelineOptions? options = null) =>
            AddCommandAnimation(shape, command, CommandValues.Event, options);

        /// <summary>Adds a native command action of the requested command type.</summary>
        public PowerPointTimelineAction AddCommandAnimation(PowerPointShape shape,
            string command, CommandValues type,
            PowerPointTimelineOptions? options = null) {
            if (string.IsNullOrWhiteSpace(command)) throw new ArgumentException("A command is required.", nameof(command));
            uint shapeId = ValidateTimelineTarget(shape);
            uint id = GetNextTimingId();
            var action = new Command(CreateTimelineBehavior(shapeId, id, options)) {
                Type = type,
                CommandName = command
            };
            AppendTimelineAction(action, options);
            return new PowerPointTimelineAction(id,
                PowerPointAnimationKind.Command, shapeId);
        }

        /// <summary>Changes duration on any typed or imported action with a common timing id.</summary>
        public bool SetAnimationDuration(uint timingId, uint durationMilliseconds) {
            CommonTimeNode? node = FindUniqueCommonTimeNode(timingId);
            if (node == null) return false;
            node.Duration = durationMilliseconds.ToString(CultureInfo.InvariantCulture);
            return true;
        }

        /// <summary>
        /// Removes exactly one typed action by common timing id. Unrelated siblings and
        /// unmodeled producer sequences are retained.
        /// </summary>
        public bool RemoveAnimation(uint timingId) {
            CommonTimeNode? node = FindUniqueCommonTimeNode(timingId);
            OpenXmlElement? behavior = node?.Parent;
            OpenXmlElement? action = behavior?.Parent;
            if (action == null || action is not (AnimateMotion or AnimateRotation or Command))
                return false;
            ChildTimeNodeList? actionList = action.Parent as ChildTimeNodeList;
            OpenXmlElement? owner = action.Ancestors<ParallelTimeNode>().FirstOrDefault();
            action.Remove();
            if (owner != null && actionList != null &&
                !actionList.ChildElements.Any()) owner.Remove();
            return true;
        }

        private uint ValidateTimelineTarget(PowerPointShape shape) {
            if (shape == null) throw new ArgumentNullException(nameof(shape));
            if (!ReferenceEquals(shape.OwnerSlide, this)) throw new ArgumentException("The animation target must belong to this slide.", nameof(shape));
            return shape.Id ?? throw new InvalidOperationException("The animation target has no shape identifier.");
        }

        private static CommonBehavior CreateTimelineBehavior(uint shapeId,
            uint timingId, PowerPointTimelineOptions? options) {
            PowerPointTimelineOptions resolved = options ?? new PowerPointTimelineOptions();
            ValidateTimelineOptions(resolved);
            return new CommonBehavior(new CommonTimeNode {
                Id = timingId,
                Duration = resolved.DurationMilliseconds.ToString(CultureInfo.InvariantCulture),
                Fill = TimeNodeFillValues.Hold
            }, new TargetElement(new ShapeTarget {
                ShapeId = shapeId.ToString(CultureInfo.InvariantCulture)
            }));
        }

        private void AppendTimelineAction(OpenXmlElement action,
            PowerPointTimelineOptions? options) {
            PowerPointTimelineOptions resolved = options ?? new PowerPointTimelineOptions();
            ValidateTimelineOptions(resolved);
            uint actionTimingId = action.Descendants<CommonTimeNode>()
                .Select(node => node.Id?.Value ?? 0U).DefaultIfEmpty().Max();
            if (actionTimingId == uint.MaxValue)
                throw new InvalidOperationException("The imported timing identifier space is exhausted.");
            Timing timing = SlideRoot.Timing ??= new Timing();
            TimeNodeList list = timing.GetFirstChild<TimeNodeList>() ?? timing.AppendChild(new TimeNodeList());
            ParallelTimeNode rootParallel = list.Elements<ParallelTimeNode>()
                .FirstOrDefault(node => node.GetFirstChild<CommonTimeNode>()?.NodeType?.Value == TimeNodeValues.TmingRoot)
                ?? list.AppendChild(new ParallelTimeNode(new CommonTimeNode {
                    Id = Math.Max(GetNextTimingId(), actionTimingId + 1U),
                    Duration = "indefinite", Restart = TimeNodeRestartValues.Never,
                    NodeType = TimeNodeValues.TmingRoot
                }));
            CommonTimeNode root = rootParallel.GetFirstChild<CommonTimeNode>()!;
            ChildTimeNodeList children = root.GetFirstChild<ChildTimeNodeList>() ?? root.AppendChild(new ChildTimeNodeList());
            Condition condition = resolved.Start == PowerPointTimelineStart.OnClick
                ? new Condition(new TargetElement(new SlideTarget())) {
                    Event = TriggerEventValues.OnClick,
                    Delay = resolved.DelayMilliseconds.ToString(CultureInfo.InvariantCulture)
                }
                : new Condition { Delay = resolved.DelayMilliseconds.ToString(CultureInfo.InvariantCulture) };
            var actions = new ChildTimeNodeList();
            actions.Append(action);
            uint ownerTimingId = Math.Max(GetNextTimingId(), actionTimingId + 1U);
            var owner = new CommonTimeNode(new StartConditionList(condition), actions) {
                Id = ownerTimingId, Duration = "indefinite", Fill = TimeNodeFillValues.Hold,
                NodeType = resolved.Start == PowerPointTimelineStart.OnClick
                    ? TimeNodeValues.ClickEffect
                    : resolved.Start == PowerPointTimelineStart.WithPrevious
                        ? TimeNodeValues.WithEffect : TimeNodeValues.AfterEffect
            };
            children.Append(new ParallelTimeNode(owner));
        }

        private CommonTimeNode? FindUniqueCommonTimeNode(uint timingId) {
            CommonTimeNode[] matches = SlideRoot.Timing?
                .Descendants<CommonTimeNode>()
                .Where(candidate => candidate.Id?.Value == timingId)
                .Take(2).ToArray() ?? Array.Empty<CommonTimeNode>();
            return matches.Length == 1 ? matches[0] : null;
        }

        private static void ValidateTimelineOptions(PowerPointTimelineOptions options) {
            if (options.DurationMilliseconds == 0) throw new ArgumentOutOfRangeException(nameof(options), "Duration must be positive.");
            if (!Enum.IsDefined(typeof(PowerPointTimelineStart), options.Start)) throw new ArgumentOutOfRangeException(nameof(options));
        }
    }
}
