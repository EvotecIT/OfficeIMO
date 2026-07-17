using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointSlide {
        private const string ClassicAnimationExtensionUri =
            "{5BA743F1-2B69-4BB9-B2E0-4A418B7E7435}";
        private const string ClassicAnimationNamespace =
            "https://schemas.officeimo.net/powerpoint/2026/classic-animations";

        /// <summary>Gets the editable classic animations authored by OfficeIMO or imported from binary PPT.</summary>
        public IReadOnlyList<PowerPointClassicAnimation> ClassicAnimations =>
            ReadClassicAnimations();

        /// <summary>Adds or replaces the classic animation for a shape on this slide.</summary>
        public PowerPointClassicAnimation AddClassicAnimation(PowerPointShape shape,
            PowerPointClassicAnimationEffect effect,
            PowerPointClassicAnimationOptions? options = null) {
            if (shape == null) throw new ArgumentNullException(nameof(shape));
            if (!ReferenceEquals(shape.OwnerSlide, this)) {
                throw new ArgumentException("The animation target must belong to this slide.", nameof(shape));
            }
            uint shapeId = shape.Id
                ?? throw new InvalidOperationException("The animation target has no shape identifier.");
            options ??= new PowerPointClassicAnimationOptions();
            ValidateClassicAnimation(effect, options.Direction, options.BuildType,
                options.DelayMilliseconds, options.AfterEffect, options.TextBuild);

            List<PowerPointClassicAnimation> animations = ReadClassicAnimations().ToList();
            if (SlideRoot.Timing != null && animations.Count == 0) {
                throw new NotSupportedException(
                    "Classic animation authoring cannot replace an existing advanced or media timing tree.");
            }
            animations.RemoveAll(animation => animation.ShapeId == shapeId);
            int nextOrder = animations.Where(animation => animation.Order >= 0)
                .Select(animation => animation.Order)
                .DefaultIfEmpty(-1).Max() + 1;
            var created = new PowerPointClassicAnimation(shapeId, effect, options.Direction,
                options.BuildType, options.Automatic, options.DelayMilliseconds,
                nextOrder, options.Reverse, options.AnimateBackground,
                options.AfterEffect, options.TextBuild, options.RawDimColor,
                stopsSound: options.StopsSound);
            animations.Add(created);
            SetClassicAnimations(animations);
            return created;
        }

        /// <summary>Removes the classic animation attached to a shape.</summary>
        public bool RemoveClassicAnimation(PowerPointShape shape) {
            if (shape == null) throw new ArgumentNullException(nameof(shape));
            if (!ReferenceEquals(shape.OwnerSlide, this)) {
                throw new ArgumentException(
                    "The animation target must belong to this slide.",
                    nameof(shape));
            }
            uint? shapeId = shape.Id;
            if (!shapeId.HasValue) return false;
            List<PowerPointClassicAnimation> animations = ReadClassicAnimations().ToList();
            int removed = animations.RemoveAll(animation => animation.ShapeId == shapeId.Value);
            if (removed == 0) return false;
            SetClassicAnimations(animations);
            return true;
        }

        /// <summary>Removes all OfficeIMO classic animations from the slide.</summary>
        public void ClearClassicAnimations() => SetClassicAnimations(
            Array.Empty<PowerPointClassicAnimation>());

        /// <summary>Sets the embedded WAV or AIFF sound played with a shape's classic animation.</summary>
        public void SetClassicAnimationSound(PowerPointShape shape,
            string audioPath, bool stopExistingSounds = false) {
            if (audioPath == null) throw new ArgumentNullException(nameof(audioPath));
            if (!File.Exists(audioPath)) {
                throw new FileNotFoundException("Audio file not found.", audioPath);
            }
            using FileStream input = new(audioPath, FileMode.Open,
                FileAccess.Read, FileShare.Read);
            SetClassicAnimationSound(shape, input, Path.GetFileName(audioPath),
                GetAudioContentType(audioPath), Path.GetExtension(audioPath),
                stopExistingSounds);
        }

        /// <summary>Sets the embedded WAV or AIFF sound played with a shape's classic animation.</summary>
        public void SetClassicAnimationSound(PowerPointShape shape, Stream audio,
            string name, string contentType = "audio/wav",
            string extension = ".wav", bool stopExistingSounds = false) {
            if (shape == null) throw new ArgumentNullException(nameof(shape));
            if (string.IsNullOrWhiteSpace(name)) {
                throw new ArgumentException("An animation sound name is required.",
                    nameof(name));
            }
            EnsureClassicAnimationTarget(shape);
            string relationshipId = PowerPointEmbeddedSound.Add(_slidePart,
                audio, contentType, extension);
            try {
                UpdateClassicAnimation(shape, animation => CopyClassicAnimation(
                    animation, playsSound: true,
                    stopsSound: stopExistingSounds,
                    soundRelationshipId: relationshipId, soundName: name));
            } catch {
                PowerPointEmbeddedSound.RemoveIfUnused(_slidePart,
                    relationshipId);
                throw;
            }
        }

        /// <summary>Controls whether existing sounds stop when a shape's classic animation starts.</summary>
        public void SetClassicAnimationStopsSound(PowerPointShape shape,
            bool value) => UpdateClassicAnimation(shape, animation =>
                CopyClassicAnimation(animation, stopsSound: value));

        /// <summary>Removes the embedded sound played with a shape's classic animation.</summary>
        public void ClearClassicAnimationSound(PowerPointShape shape) =>
            UpdateClassicAnimation(shape, animation => CopyClassicAnimation(
                animation, playsSound: false, soundRelationshipId: null,
                soundName: null));

        /// <summary>Returns the exact embedded classic-animation sound bytes, when present.</summary>
        public byte[]? GetClassicAnimationSoundBytes(PowerPointShape shape) {
            if (shape == null) throw new ArgumentNullException(nameof(shape));
            uint? shapeId = shape.Id;
            PowerPointClassicAnimation? animation = shapeId.HasValue
                ? ReadClassicAnimations().FirstOrDefault(item =>
                    item.ShapeId == shapeId.Value)
                : null;
            return PowerPointEmbeddedSound.Read(_slidePart,
                animation?.SoundRelationshipId);
        }

        internal void SetClassicAnimations(
            IReadOnlyList<PowerPointClassicAnimation> animations) {
            if (animations == null) throw new ArgumentNullException(nameof(animations));
            IReadOnlyList<PowerPointClassicAnimation> previousAnimations =
                ReadClassicAnimations();
            string[] previousSoundRelationships = previousAnimations
                .Select(animation => animation.SoundRelationshipId)
                .Where(id => !string.IsNullOrEmpty(id))
                .Cast<string>()
                .Distinct(StringComparer.Ordinal)
                .ToArray();
            PowerPointClassicAnimation[] ordered = animations
                .OrderBy(animation => animation.Order)
                .ToArray();
            if (ordered.Length == 0) {
                if (HasOnlyClassicAnimationTiming()) {
                    SlideRoot.Timing = null;
                } else if (previousAnimations.Count > 0) {
                    RemoveClassicTimingNodes(previousAnimations);
                }
                WriteClassicAnimationMetadata(ordered);
                RemoveUnusedClassicAnimationSounds(
                    previousSoundRelationships);
                return;
            }

            bool preserveUnrelatedTiming = SlideRoot.Timing != null
                && previousAnimations.Count > 0
                && !HasOnlyClassicAnimationTiming();
            if (preserveUnrelatedTiming) {
                RemoveClassicTimingNodes(previousAnimations);
            }
            var childNodes = new ChildTimeNodeList();
            uint timingId = preserveUnrelatedTiming
                ? GetNextTimingId()
                : 2U;
            foreach (PowerPointClassicAnimation animation in ordered) {
                uint effectTimingId = timingId++;
                var effectTime = new CommonTimeNode {
                    Id = effectTimingId,
                    Duration = "500",
                    Fill = TimeNodeFillValues.Hold,
                    PresetClass = TimeNodePresetClassValues.Entrance,
                    PresetId = (int)(byte)animation.Effect,
                    PresetSubtype = animation.Direction,
                    AutoReverse = animation.Reverse,
                    AfterEffect = animation.AfterEffect !=
                        PowerPointClassicAnimationAfterEffect.None
                };
                var behavior = new CommonBehavior(effectTime,
                    new TargetElement(new ShapeTarget {
                        ShapeId = animation.ShapeId.ToString(CultureInfo.InvariantCulture)
                    }));
                var effect = new AnimateEffect(behavior) {
                    Transition = AnimateEffectTransitionValues.In,
                    Filter = GetClassicAnimationFilter(animation.Effect,
                        animation.Direction)
                };
                var actions = new ChildTimeNodeList(effect);
                AppendClassicAnimationStopSound(actions, animation,
                    effectTimingId, ref timingId);
                AppendClassicAnimationSound(actions, animation,
                    ref timingId);
                AppendClassicAnimationAfterEffect(actions, animation,
                    effectTimingId, ref timingId);
                var startCondition = animation.Automatic
                    ? new Condition {
                        Delay = animation.DelayMilliseconds.ToString(
                            CultureInfo.InvariantCulture)
                    }
                    : new Condition(new TargetElement(new SlideTarget())) {
                        Event = TriggerEventValues.OnClick,
                        Delay = "0"
                    };
                var owner = new CommonTimeNode(
                    new StartConditionList(startCondition),
                    actions) {
                    Id = timingId++,
                    Duration = "indefinite",
                    Fill = TimeNodeFillValues.Hold,
                    NodeType = animation.Automatic
                        ? TimeNodeValues.AfterEffect
                        : TimeNodeValues.ClickEffect
                };
                childNodes.Append(new ParallelTimeNode(owner));
            }

            var root = new CommonTimeNode(childNodes) {
                Id = 1U,
                Duration = "indefinite",
                Restart = TimeNodeRestartValues.Never,
                NodeType = TimeNodeValues.TmingRoot
            };
            var timing = new Timing(new TimeNodeList(new ParallelTimeNode(root)));
            BuildList? buildList = CreateClassicBuildList(ordered);
            if (buildList != null) timing.Append(buildList);
            if (preserveUnrelatedTiming) {
                AppendClassicTiming(SlideRoot.Timing!, childNodes,
                    buildList);
            } else {
                SlideRoot.Timing?.Remove();
                OpenXmlElement? insertBefore = SlideRoot
                    .GetFirstChild<SlideExtensionList>();
                if (insertBefore == null) SlideRoot.Append(timing);
                else SlideRoot.InsertBefore(timing, insertBefore);
            }
            WriteClassicAnimationMetadata(ordered);
            RemoveUnusedClassicAnimationSounds(previousSoundRelationships);
        }

        private static void AppendClassicAnimationStopSound(
            ChildTimeNodeList actions,
            PowerPointClassicAnimation animation, uint effectTimingId,
            ref uint timingId) {
            if (!animation.StopsSound) return;
            var stopTime = new CommonTimeNode(
                new StartConditionList(new Condition(
                    new TimeNode { Val = effectTimingId }) {
                    Event = TriggerEventValues.Begin,
                    Delay = "0"
                })) {
                Id = timingId++,
                Duration = "1",
                Fill = TimeNodeFillValues.Hold,
                Display = false,
                MasterRelation = TimeNodeMasterRelationValues.SameClick
            };
            var behavior = new CommonBehavior(stopTime,
                new TargetElement(new SlideTarget()));
            actions.Append(new Command(behavior) {
                Type = CommandValues.Event,
                CommandName = "onstopaudio"
            });
        }

        private static void AppendClassicAnimationSound(
            ChildTimeNodeList actions,
            PowerPointClassicAnimation animation, ref uint timingId) {
            if (!animation.PlaysSound
                || string.IsNullOrEmpty(animation.SoundRelationshipId)) {
                return;
            }
            var mediaTime = new CommonTimeNode(
                new StartConditionList(new Condition { Delay = "0" })) {
                Id = timingId++,
                Duration = "media",
                Fill = TimeNodeFillValues.Hold
            };
            var media = new CommonMediaNode(mediaTime,
                new TargetElement(new SoundTarget {
                    Embed = animation.SoundRelationshipId,
                    Name = animation.SoundName ?? "Animation Sound"
                })) {
                Volume = 80000
            };
            actions.Append(new Audio(media));
        }

        private static void AppendClassicAnimationAfterEffect(
            ChildTimeNodeList actions,
            PowerPointClassicAnimation animation, uint effectTimingId,
            ref uint timingId) {
            if (animation.AfterEffect ==
                PowerPointClassicAnimationAfterEffect.None) return;
            CommonTimeNode actionTime = CreateClassicAfterEffectTime(
                animation.AfterEffect, effectTimingId, ref timingId);
            string shapeId = animation.ShapeId.ToString(
                CultureInfo.InvariantCulture);
            var target = new TargetElement(new ShapeTarget {
                ShapeId = shapeId
            });
            if (animation.AfterEffect ==
                PowerPointClassicAnimationAfterEffect.Dim) {
                var behavior = new CommonBehavior(actionTime, target,
                    new AttributeNameList(new AttributeName("style.color")));
                actions.Append(new AnimateColor(behavior,
                    new ToColor(CreateClassicDimColor(
                        animation.RawDimColor))));
                return;
            }
            var visibilityBehavior = new CommonBehavior(actionTime, target,
                new AttributeNameList(new AttributeName(
                    "style.visibility")));
            actions.Append(new SetBehavior(visibilityBehavior,
                new ToVariantValue(new StringVariantValue {
                    Val = "hidden"
                })));
        }

        private static CommonTimeNode CreateClassicAfterEffectTime(
            PowerPointClassicAnimationAfterEffect afterEffect,
            uint effectTimingId, ref uint timingId) {
            Condition condition = afterEffect ==
                    PowerPointClassicAnimationAfterEffect.HideOnNextClick
                ? new Condition(new TargetElement(new SlideTarget())) {
                    Event = TriggerEventValues.OnClick,
                    Delay = "0"
                }
                : new Condition(new TimeNode { Val = effectTimingId }) {
                    Event = TriggerEventValues.End,
                    Delay = "0"
                };
            return new CommonTimeNode(
                new StartConditionList(condition)) {
                Id = timingId++,
                Duration = "1",
                Fill = TimeNodeFillValues.Hold
            };
        }

        private static OpenXmlElement CreateClassicDimColor(uint rawColor) {
            byte index = unchecked((byte)(rawColor >> 24));
            A.SchemeColorValues? scheme = index switch {
                0 => A.SchemeColorValues.Light1,
                1 => A.SchemeColorValues.Dark1,
                2 => A.SchemeColorValues.Accent4,
                3 => A.SchemeColorValues.Dark2,
                4 => A.SchemeColorValues.Light2,
                5 => A.SchemeColorValues.Accent1,
                6 => A.SchemeColorValues.Accent2,
                7 => A.SchemeColorValues.Accent3,
                _ => null
            };
            if (scheme.HasValue) {
                return new A.SchemeColor { Val = scheme.Value };
            }
            return new A.RgbColorModelHex {
                Val = string.Concat(
                    unchecked((byte)rawColor).ToString("X2",
                        CultureInfo.InvariantCulture),
                    unchecked((byte)(rawColor >> 8)).ToString("X2",
                        CultureInfo.InvariantCulture),
                    unchecked((byte)(rawColor >> 16)).ToString("X2",
                        CultureInfo.InvariantCulture))
            };
        }

        private static void AppendClassicTiming(Timing timing,
            ChildTimeNodeList classicNodes, BuildList? classicBuilds) {
            TimeNodeList timeNodeList = timing
                .GetFirstChild<TimeNodeList>()
                ?? timing.AppendChild(new TimeNodeList());
            ParallelTimeNode rootParallel = timeNodeList
                .Elements<ParallelTimeNode>()
                .FirstOrDefault(node => node.GetFirstChild<CommonTimeNode>()?
                    .NodeType?.Value == TimeNodeValues.TmingRoot)
                ?? timeNodeList.AppendChild(new ParallelTimeNode());
            CommonTimeNode root = rootParallel
                .GetFirstChild<CommonTimeNode>()
                ?? rootParallel.AppendChild(new CommonTimeNode {
                    Id = 1U,
                    Duration = "indefinite",
                    Restart = TimeNodeRestartValues.Never,
                    NodeType = TimeNodeValues.TmingRoot
                });
            ChildTimeNodeList targetNodes = root
                .GetFirstChild<ChildTimeNodeList>()
                ?? root.AppendChild(new ChildTimeNodeList());
            foreach (OpenXmlElement node in classicNodes.ChildElements
                         .ToArray()) {
                node.Remove();
                targetNodes.Append(node);
            }
            if (classicBuilds == null) return;
            BuildList targetBuilds = timing.GetFirstChild<BuildList>()
                ?? timing.AppendChild(new BuildList());
            foreach (OpenXmlElement build in classicBuilds.ChildElements
                         .ToArray()) {
                build.Remove();
                targetBuilds.Append(build);
            }
        }

        private void RemoveClassicTimingNodes(
            IReadOnlyList<PowerPointClassicAnimation> animations) {
            Timing? timing = SlideRoot.Timing;
            if (timing == null) return;
            var candidates = timing.Descendants<AnimateEffect>().ToList();
            foreach (PowerPointClassicAnimation animation in animations
                         .OrderBy(item => item.Order)) {
                AnimateEffect? effect = candidates.FirstOrDefault(item =>
                    IsClassicTimingEffect(item, animation));
                if (effect == null) continue;
                OpenXmlElement? owner = effect
                    .Ancestors<ParallelTimeNode>().FirstOrDefault();
                candidates.Remove(effect);
                RemoveClassicTimingCompanions(owner, animation);
                effect.Remove();
                if (owner != null && !HasTimingAction(owner)) {
                    owner.Remove();
                }
            }
            var builds = timing.Descendants<BuildParagraph>().ToList();
            foreach (PowerPointClassicAnimation animation in animations
                         .OrderBy(item => item.Order)) {
                BuildParagraph? build = builds.FirstOrDefault(item =>
                    IsClassicBuild(item, animation));
                if (build == null) continue;
                builds.Remove(build);
                build.Remove();
            }
            foreach (BuildList list in timing.Descendants<BuildList>()
                         .Where(item => !item.HasChildren).ToArray()) {
                list.Remove();
            }
        }

        private static bool HasTimingAction(OpenXmlElement owner) =>
            owner.Descendants().Any(element => element is Animate
                or AnimateColor or AnimateEffect or AnimateMotion
                or AnimateRotation or AnimateScale or Command or Audio
                or Video or SetBehavior);

        private static void RemoveClassicTimingCompanions(
            OpenXmlElement? owner,
            PowerPointClassicAnimation animation) {
            if (owner == null) return;
            string shapeId = animation.ShapeId.ToString(
                CultureInfo.InvariantCulture);
            SetBehavior[] setters = owner.Descendants<SetBehavior>()
                .Where(set => {
                    CommonBehavior? behavior = set.CommonBehavior;
                    ShapeTarget? target = behavior?
                        .GetFirstChild<TargetElement>()?
                        .GetFirstChild<ShapeTarget>();
                    AttributeName[] attributes = behavior?
                        .GetFirstChild<AttributeNameList>()?
                        .Elements<AttributeName>().ToArray()
                        ?? Array.Empty<AttributeName>();
                    StringVariantValue? value = set.ToVariantValue?
                        .GetFirstChild<StringVariantValue>();
                    return string.Equals(target?.ShapeId?.Value, shapeId,
                               StringComparison.Ordinal)
                           && set.ChildElements.Count == 2
                           && behavior != null
                           && behavior.Descendants<ShapeTarget>().Count() == 1
                           && attributes.Length == 1
                           && string.Equals(attributes[0].Text,
                               "style.visibility", StringComparison.Ordinal)
                           && value != null
                           && set.ToVariantValue!.ChildElements.Count == 1
                           && (string.Equals(value.Val?.Value, "visible",
                                   StringComparison.OrdinalIgnoreCase)
                               || string.Equals(value.Val?.Value, "hidden",
                                   StringComparison.OrdinalIgnoreCase));
                }).ToArray();
            foreach (SetBehavior setter in setters) {
                setter.Remove();
            }
            foreach (AnimateColor color in owner.Descendants<AnimateColor>()
                         .Where(item => IsClassicDimAction(item,
                             animation)).ToArray()) {
                color.Remove();
            }
            foreach (Audio audio in owner.Descendants<Audio>()
                         .Where(item => IsClassicSoundAction(item,
                             animation)).ToArray()) {
                audio.Remove();
            }
            foreach (Command command in owner.Descendants<Command>()
                         .Where(item => IsClassicStopSoundAction(item,
                             animation)).ToArray()) {
                command.Remove();
            }
        }

        private static bool IsClassicDimAction(AnimateColor color,
            PowerPointClassicAnimation animation) =>
            animation.AfterEffect ==
                PowerPointClassicAnimationAfterEffect.Dim
            && IsClassicShapeBehavior(color.CommonBehavior,
                animation.ShapeId, "style.color");

        private static bool IsClassicSoundAction(Audio audio,
            PowerPointClassicAnimation animation) {
            SoundTarget? target = audio.CommonMediaNode?
                .GetFirstChild<TargetElement>()?
                .GetFirstChild<SoundTarget>();
            return animation.PlaysSound
                && !string.IsNullOrEmpty(animation.SoundRelationshipId)
                && string.Equals(target?.Embed?.Value,
                    animation.SoundRelationshipId,
                    StringComparison.Ordinal);
        }

        private static bool IsClassicStopSoundAction(Command command,
            PowerPointClassicAnimation animation) {
            CommonBehavior? behavior = command.CommonBehavior;
            CommonTimeNode? owner = command.Ancestors<CommonTimeNode>()
                .FirstOrDefault(node => node.NodeType?.Value ==
                        TimeNodeValues.ClickEffect
                    || node.NodeType?.Value == TimeNodeValues.AfterEffect);
            return animation.StopsSound
                && command.Type?.Value == CommandValues.Event
                && string.Equals(command.CommandName?.Value,
                    "onstopaudio", StringComparison.OrdinalIgnoreCase)
                && command.ChildElements.Count == 1
                && behavior != null
                && behavior.Descendants<SlideTarget>().Count() == 1
                && behavior.Descendants<ShapeTarget>().Count() == 0
                && owner != null
                && owner.Descendants<AnimateEffect>().Count(effect =>
                    uint.TryParse(effect.Descendants<ShapeTarget>()
                            .FirstOrDefault()?.ShapeId?.Value,
                        NumberStyles.Integer, CultureInfo.InvariantCulture,
                        out uint effectShapeId)
                    && effectShapeId == animation.ShapeId) == 1;
        }

        private static bool IsClassicShapeBehavior(CommonBehavior? behavior,
            uint shapeId, string attributeName) {
            ShapeTarget? target = behavior?
                .GetFirstChild<TargetElement>()?
                .GetFirstChild<ShapeTarget>();
            AttributeName[] attributes = behavior?
                .GetFirstChild<AttributeNameList>()?
                .Elements<AttributeName>().ToArray()
                ?? Array.Empty<AttributeName>();
            return string.Equals(target?.ShapeId?.Value,
                       shapeId.ToString(CultureInfo.InvariantCulture),
                       StringComparison.Ordinal)
                   && behavior != null
                   && behavior.Descendants<ShapeTarget>().Count() == 1
                   && attributes.Length == 1
                   && string.Equals(attributes[0].Text, attributeName,
                       StringComparison.Ordinal);
        }

        private static bool IsClassicTimingEffect(AnimateEffect effect,
            PowerPointClassicAnimation animation) {
            CommonTimeNode? effectTime = effect.CommonBehavior?
                .CommonTimeNode;
            string? shapeId = effect.Descendants<ShapeTarget>()
                .FirstOrDefault()?.ShapeId?.Value;
            ParallelTimeNode? owner = effect
                .Ancestors<ParallelTimeNode>().FirstOrDefault();
            CommonTimeNode? ownerTime = owner?
                .GetFirstChild<CommonTimeNode>();
            Condition? start = ownerTime?
                .GetFirstChild<StartConditionList>()?
                .GetFirstChild<Condition>();
            bool expectedAfterEffect = animation.AfterEffect !=
                PowerPointClassicAnimationAfterEffect.None;
            bool targetMatches = string.Equals(shapeId,
                animation.ShapeId.ToString(CultureInfo.InvariantCulture),
                StringComparison.Ordinal);
            bool startMatches = animation.Automatic
                ? start?.Event == null && string.Equals(start?.Delay?.Value,
                    animation.DelayMilliseconds.ToString(
                        CultureInfo.InvariantCulture),
                    StringComparison.Ordinal)
                : start?.Event?.Value == TriggerEventValues.OnClick;
            return targetMatches
                && effect.Transition?.Value ==
                    AnimateEffectTransitionValues.In
                && string.Equals(effect.Filter?.Value,
                    GetClassicAnimationFilter(animation.Effect,
                        animation.Direction), StringComparison.Ordinal)
                && effectTime?.Duration?.Value == "500"
                && effectTime.Fill?.Value == TimeNodeFillValues.Hold
                && effectTime.PresetClass?.Value ==
                    TimeNodePresetClassValues.Entrance
                && effectTime.PresetId?.Value ==
                    (int)(byte)animation.Effect
                && effectTime.PresetSubtype?.Value == animation.Direction
                && effectTime.AutoReverse?.Value == animation.Reverse
                && effectTime.AfterEffect?.Value == expectedAfterEffect
                && ownerTime?.Duration?.Value == "indefinite"
                && ownerTime.Fill?.Value == TimeNodeFillValues.Hold
                && ownerTime.NodeType?.Value == (animation.Automatic
                    ? TimeNodeValues.AfterEffect
                    : TimeNodeValues.ClickEffect)
                && startMatches;
        }

        private static bool IsClassicBuild(BuildParagraph build,
            PowerPointClassicAnimation animation) {
            string shapeId = animation.ShapeId.ToString(
                CultureInfo.InvariantCulture);
            if (!string.Equals(build.ShapeId?.Value, shapeId,
                    StringComparison.Ordinal)
                || build.GroupId?.Value != 0U) return false;
            if (animation.BuildType is >=
                    PowerPointClassicAnimationBuildType.ByLevel1Paragraph
                and <= PowerPointClassicAnimationBuildType.ByLevel5Paragraph) {
                return build.Build?.Value == ParagraphBuildValues.Paragraph
                    && build.BuildLevel?.Value == unchecked((uint)
                        ((byte)animation.BuildType - 1))
                    && build.AnimateBackground?.Value ==
                        animation.AnimateBackground
                    && build.Reverse?.Value == animation.Reverse
                    && string.Equals(build.AutoAdvance?.Value,
                        animation.Automatic
                            ? animation.DelayMilliseconds.ToString(
                                CultureInfo.InvariantCulture)
                            : null,
                        StringComparison.Ordinal);
            }
            return animation.BuildType ==
                    PowerPointClassicAnimationBuildType.AsOneObject
                && build.Build?.Value == ParagraphBuildValues.Whole
                && build.AnimateBackground?.Value ==
                    animation.AnimateBackground;
        }

        private void RemoveUnusedClassicAnimationSounds(
            IEnumerable<string> relationshipIds) {
            foreach (string relationshipId in relationshipIds) {
                PowerPointEmbeddedSound.RemoveIfUnused(_slidePart,
                    relationshipId);
            }
        }

        internal bool HasOnlyClassicAnimationTiming() {
            Timing? timing = SlideRoot.Timing;
            IReadOnlyList<PowerPointClassicAnimation> animations =
                ReadClassicAnimations();
            if (timing == null) return animations.Count == 0;
            if (animations.Count == 0
                || timing.Descendants<AnimateEffect>().Count() != animations.Count) {
                return false;
            }
            Type[] allowedTypes = {
                typeof(TimeNodeList), typeof(ParallelTimeNode),
                typeof(SequenceTimeNode),
                typeof(CommonTimeNode), typeof(ChildTimeNodeList),
                typeof(StartConditionList), typeof(Condition),
                typeof(EndConditionList), typeof(PreviousConditionList),
                typeof(NextConditionList), typeof(TimeNode),
                typeof(TargetElement), typeof(SlideTarget), typeof(ShapeTarget),
                typeof(TextElement), typeof(ParagraphIndexRange),
                typeof(CharRange),
                typeof(AnimateEffect), typeof(CommonBehavior),
                typeof(AnimateColor), typeof(ToColor),
                typeof(A.RgbColorModelHex), typeof(A.SchemeColor),
                typeof(Audio), typeof(CommonMediaNode), typeof(SoundTarget),
                typeof(Command),
                typeof(SetBehavior), typeof(ToVariantValue),
                typeof(StringVariantValue), typeof(AttributeNameList),
                typeof(AttributeName),
                typeof(BuildList), typeof(BuildParagraph)
            };
            if (timing.Descendants().Any(element =>
                    !allowedTypes.Contains(element.GetType()))) return false;
            uint[] targets = timing.Descendants<AnimateEffect>()
                .Select(effect => effect.Descendants<ShapeTarget>().FirstOrDefault()?
                    .ShapeId?.Value)
                .Select(value => uint.TryParse(value, NumberStyles.Integer,
                    CultureInfo.InvariantCulture, out uint parsed)
                    ? parsed : uint.MaxValue)
                .ToArray();
            return targets.SequenceEqual(animations.Select(animation =>
                       animation.ShapeId))
                && HasOnlyClassicStandardActions(timing, animations);
        }

        private void UpdateClassicAnimation(PowerPointShape shape,
            Func<PowerPointClassicAnimation, PowerPointClassicAnimation> update) {
            uint shapeId = EnsureClassicAnimationTarget(shape);
            List<PowerPointClassicAnimation> animations =
                ReadClassicAnimations().ToList();
            int index = animations.FindIndex(animation =>
                animation.ShapeId == shapeId);
            if (index < 0) throw CreateMissingClassicAnimationException();
            animations[index] = update(animations[index]);
            SetClassicAnimations(animations);
        }

        private uint EnsureClassicAnimationTarget(PowerPointShape shape) {
            if (shape == null) throw new ArgumentNullException(nameof(shape));
            if (!ReferenceEquals(shape.OwnerSlide, this)) {
                throw new ArgumentException(
                    "The animation target must belong to this slide.",
                    nameof(shape));
            }
            uint shapeId = shape.Id
                ?? throw new InvalidOperationException(
                    "The animation target has no shape identifier.");
            if (!ReadClassicAnimations().Any(animation =>
                    animation.ShapeId == shapeId)) {
                throw CreateMissingClassicAnimationException();
            }
            return shapeId;
        }

        private static InvalidOperationException
            CreateMissingClassicAnimationException() => new(
                "The target shape has no classic animation.");

        private static PowerPointClassicAnimation CopyClassicAnimation(
            PowerPointClassicAnimation animation, bool? playsSound = null,
            bool? stopsSound = null, string? soundRelationshipId = null,
            string? soundName = null) => new(animation.ShapeId,
                animation.Effect, animation.Direction, animation.BuildType,
                animation.Automatic, animation.DelayMilliseconds,
                animation.Order, animation.Reverse,
                animation.AnimateBackground, animation.AfterEffect,
                animation.TextBuild, animation.RawDimColor,
                playsSound ?? animation.PlaysSound,
                stopsSound ?? animation.StopsSound,
                playsSound == false ? null
                    : soundRelationshipId ?? animation.SoundRelationshipId,
                playsSound == false ? null : soundName ?? animation.SoundName);

        private IReadOnlyList<PowerPointClassicAnimation> ReadClassicAnimations() {
            SlideExtension? extension = SlideRoot.GetFirstChild<SlideExtensionList>()?
                .Elements<SlideExtension>().FirstOrDefault(item => string.Equals(
                    item.Uri?.Value, ClassicAnimationExtensionUri,
                    StringComparison.Ordinal));
            OpenXmlElement? root = extension?.ChildElements.FirstOrDefault(element =>
                element.NamespaceUri == ClassicAnimationNamespace
                && element.LocalName == "classicAnimations");
            if (root == null) return InferClassicAnimations();
            var result = new List<PowerPointClassicAnimation>();
            foreach (OpenXmlElement element in root.ChildElements.Where(item =>
                         item.NamespaceUri == ClassicAnimationNamespace
                         && item.LocalName == "animation")) {
                if (!TryReadUInt(element, "shapeId", out uint shapeId)
                    || !TryReadByte(element, "effect", out byte effect)
                    || !TryReadByte(element, "direction", out byte direction)
                    || !TryReadByte(element, "build", out byte build)
                    || !TryReadInt(element, "delay", out int delay)
                    || !TryReadInt(element, "order", out int order)
                    || !TryReadByte(element, "after", out byte after)
                    || !TryReadByte(element, "textBuild", out byte textBuild)
                    || !TryReadUInt(element, "dimColor", out uint dimColor)
                    || !IsClassicAnimationValue(effect, direction, build,
                        delay, after, textBuild)) continue;
                OpenXmlElement? sound = element.ChildElements.FirstOrDefault(item =>
                    item.NamespaceUri == "http://schemas.openxmlformats.org/presentationml/2006/main"
                    && item.LocalName == "snd");
                result.Add(new PowerPointClassicAnimation(shapeId,
                    (PowerPointClassicAnimationEffect)effect, direction,
                    (PowerPointClassicAnimationBuildType)build,
                    ReadBool(element, "automatic"), delay, order,
                    ReadBool(element, "reverse"),
                    ReadBool(element, "animateBackground"),
                    (PowerPointClassicAnimationAfterEffect)after,
                    (PowerPointClassicTextBuild)textBuild, dimColor,
                    ReadBool(element, "playsSound"),
                    ReadBool(element, "stopsSound"),
                    sound == null ? null : GetAttribute(sound, "embed"),
                    sound == null ? null : GetAttribute(sound, "name")));
            }
            return result.OrderBy(animation => animation.Order).ToArray();
        }

        private void WriteClassicAnimationMetadata(
            IReadOnlyList<PowerPointClassicAnimation> animations) {
            SlideExtensionList? list = SlideRoot
                .GetFirstChild<SlideExtensionList>();
            SlideExtension? existing = list?.Elements<SlideExtension>().FirstOrDefault(item =>
                string.Equals(item.Uri?.Value, ClassicAnimationExtensionUri,
                    StringComparison.Ordinal));
            existing?.Remove();
            if (animations.Count == 0) {
                if (list?.ChildElements.Count == 0) list.Remove();
                return;
            }
            list ??= SlideRoot.AppendChild(new SlideExtensionList());
            var root = new OpenXmlUnknownElement("oimo", "classicAnimations",
                ClassicAnimationNamespace);
            foreach (PowerPointClassicAnimation animation in animations) {
                var element = new OpenXmlUnknownElement("oimo", "animation",
                    ClassicAnimationNamespace);
                SetAttribute(element, "shapeId", animation.ShapeId);
                SetAttribute(element, "effect", (byte)animation.Effect);
                SetAttribute(element, "direction", animation.Direction);
                SetAttribute(element, "build", (byte)animation.BuildType);
                SetAttribute(element, "automatic", animation.Automatic);
                SetAttribute(element, "delay", animation.DelayMilliseconds);
                SetAttribute(element, "order", animation.Order);
                SetAttribute(element, "reverse", animation.Reverse);
                SetAttribute(element, "animateBackground",
                    animation.AnimateBackground);
                SetAttribute(element, "after", (byte)animation.AfterEffect);
                SetAttribute(element, "textBuild", (byte)animation.TextBuild);
                SetAttribute(element, "dimColor", animation.RawDimColor);
                SetAttribute(element, "playsSound", animation.PlaysSound);
                SetAttribute(element, "stopsSound", animation.StopsSound);
                if (animation.PlaysSound
                    && !string.IsNullOrEmpty(animation.SoundRelationshipId)) {
                    element.Append(new Sound {
                        Embed = animation.SoundRelationshipId,
                        Name = animation.SoundName ?? "Animation Sound"
                    });
                }
                root.Append(element);
            }
            var extension = new SlideExtension { Uri = ClassicAnimationExtensionUri };
            extension.Append(root);
            list.Append(extension);
        }

        private static BuildList? CreateClassicBuildList(
            IEnumerable<PowerPointClassicAnimation> animations) {
            var list = new BuildList();
            foreach (PowerPointClassicAnimation animation in animations) {
                if (animation.BuildType is >= PowerPointClassicAnimationBuildType.ByLevel1Paragraph
                    and <= PowerPointClassicAnimationBuildType.ByLevel5Paragraph) {
                    list.Append(new BuildParagraph {
                        ShapeId = animation.ShapeId.ToString(CultureInfo.InvariantCulture),
                        GroupId = 0U,
                        Build = ParagraphBuildValues.Paragraph,
                        BuildLevel = unchecked((uint)((byte)animation.BuildType - 1)),
                        AnimateBackground = animation.AnimateBackground,
                        Reverse = animation.Reverse,
                        AutoAdvance = animation.Automatic
                            ? animation.DelayMilliseconds.ToString(CultureInfo.InvariantCulture)
                            : null
                    });
                } else if (animation.BuildType ==
                           PowerPointClassicAnimationBuildType.AsOneObject) {
                    list.Append(new BuildParagraph {
                        ShapeId = animation.ShapeId.ToString(CultureInfo.InvariantCulture),
                        GroupId = 0U,
                        Build = ParagraphBuildValues.Whole,
                        AnimateBackground = animation.AnimateBackground
                    });
                }
            }
            return list.HasChildren ? list : null;
        }

        private static string GetClassicAnimationFilter(
            PowerPointClassicAnimationEffect effect, byte direction) => effect switch {
                PowerPointClassicAnimationEffect.Cut => direction == 1
                    ? "cut(throughBlack)" : "cut",
                PowerPointClassicAnimationEffect.Random => "random",
                PowerPointClassicAnimationEffect.Blinds => direction == 0
                    ? "blinds(vertical)" : "blinds(horizontal)",
                PowerPointClassicAnimationEffect.Checker => direction == 0
                    ? "checkerboard(across)" : "checkerboard(down)",
                PowerPointClassicAnimationEffect.Cover => "cover(" +
                    GetEightDirection(direction) + ")",
                PowerPointClassicAnimationEffect.Dissolve => "dissolve",
                PowerPointClassicAnimationEffect.Fade => "fade",
                PowerPointClassicAnimationEffect.Pull => "pull(" +
                    GetEightDirection(direction) + ")",
                PowerPointClassicAnimationEffect.RandomBars => direction == 0
                    ? "randomBar(horizontal)" : "randomBar(vertical)",
                PowerPointClassicAnimationEffect.Strips => "strips(" +
                    GetEightDirection(direction) + ")",
                PowerPointClassicAnimationEffect.Wipe => "wipe(" +
                    GetFourDirection(direction) + ")",
                PowerPointClassicAnimationEffect.Zoom => direction == 0
                    ? "box(out)" : "box(in)",
                PowerPointClassicAnimationEffect.Fly => "fly(" + direction
                    .ToString(CultureInfo.InvariantCulture) + ")",
                PowerPointClassicAnimationEffect.Split => "split(" + direction
                    .ToString(CultureInfo.InvariantCulture) + ")",
                PowerPointClassicAnimationEffect.Flash => "flash(" + direction
                    .ToString(CultureInfo.InvariantCulture) + ")",
                PowerPointClassicAnimationEffect.Diamond => "diamond",
                PowerPointClassicAnimationEffect.Plus => "plus",
                PowerPointClassicAnimationEffect.Wedge => "wedge",
                PowerPointClassicAnimationEffect.Wheel => "wheel(" + direction
                    .ToString(CultureInfo.InvariantCulture) + ")",
                PowerPointClassicAnimationEffect.Circle => "circle",
                _ => "cut"
            };

        private static string GetFourDirection(byte direction) => direction switch {
            1 => "u", 2 => "r", 3 => "d", _ => "l"
        };

        private static string GetEightDirection(byte direction) => direction switch {
            1 => "u", 2 => "r", 3 => "d", 4 => "lu", 5 => "ru",
            6 => "ld", 7 => "rd", _ => "l"
        };

        private static void ValidateClassicAnimation(
            PowerPointClassicAnimationEffect effect, byte direction,
            PowerPointClassicAnimationBuildType buildType, int delay,
            PowerPointClassicAnimationAfterEffect afterEffect,
            PowerPointClassicTextBuild textBuild) {
            if (!IsClassicAnimationValue((byte)effect, direction,
                    (byte)buildType, delay, (byte)afterEffect,
                    (byte)textBuild)) {
                throw new ArgumentOutOfRangeException(nameof(direction),
                    "The effect, direction, build, delay, after-effect, or text-build value is not representable by PowerPoint 97-2003.");
            }
        }

        private static bool IsClassicAnimationValue(byte effect, byte direction,
            byte build, int delay, byte after, byte textBuild) => delay >= 0
            && (build <= 0x0A || build == 0xFE)
            && after <= 3 && textBuild <= 2
            && (effect switch {
                0x00 => direction <= 2,
                0x01 => true,
                0x02 or 0x03 or 0x08 or 0x0B => direction <= 1,
                0x04 or 0x07 => direction <= 7,
                0x05 or 0x06 or 0x11 or 0x12 or 0x13 or 0x1B => direction == 0,
                0x09 => direction is >= 4 and <= 7,
                0x0A or 0x0D => direction <= 3,
                0x0C => direction <= 0x1C,
                0x0E => direction <= 2,
                0x1A => direction is 1 or 2 or 3 or 4 or 8,
                _ => false
            });

        private static bool TryReadUInt(OpenXmlElement element, string name,
            out uint value) => uint.TryParse(GetAttribute(element, name),
            NumberStyles.Integer, CultureInfo.InvariantCulture, out value);

        private static bool TryReadByte(OpenXmlElement element, string name,
            out byte value) => byte.TryParse(GetAttribute(element, name),
            NumberStyles.Integer, CultureInfo.InvariantCulture, out value);

        private static bool TryReadInt(OpenXmlElement element, string name,
            out int value) => int.TryParse(GetAttribute(element, name),
            NumberStyles.Integer, CultureInfo.InvariantCulture, out value);

        private static bool ReadBool(OpenXmlElement element, string name) =>
            bool.TryParse(GetAttribute(element, name), out bool value) && value;

        private static string? GetAttribute(OpenXmlElement element, string name) =>
            element.GetAttributes().FirstOrDefault(attribute =>
                attribute.LocalName == name).Value;

        private static void SetAttribute(OpenXmlElement element, string name,
            object value) => element.SetAttribute(new OpenXmlAttribute(string.Empty,
                name, string.Empty, Convert.ToString(value,
                    CultureInfo.InvariantCulture) ?? string.Empty));
    }
}
