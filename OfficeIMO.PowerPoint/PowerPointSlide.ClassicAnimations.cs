using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.PowerPoint.OpenXml;
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
            ArgumentNullException.ThrowIfNull(shape);
            if (!ReferenceEquals(shape.OwnerSlide, this)) {
                throw new ArgumentException("The animation target must belong to this slide.", nameof(shape));
            }
            uint shapeId = shape.Id
                ?? throw new InvalidOperationException("The animation target has no shape identifier.");
            options ??= new PowerPointClassicAnimationOptions();
            ValidateClassicAnimation(effect, options.Direction, options.BuildType,
                options.DelayMilliseconds, options.AfterEffect, options.TextBuild);

            List<PowerPointClassicAnimation> animations = ReadClassicAnimations().ToList();
            if (PartRef.Root.Child(ClassicPresentationNamespace, "timing").Exists && animations.Count == 0) {
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
            ArgumentNullException.ThrowIfNull(shape);
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
            ArgumentNullException.ThrowIfNull(audioPath);
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
            ArgumentNullException.ThrowIfNull(shape);
            if (string.IsNullOrWhiteSpace(name)) {
                throw new ArgumentException("An animation sound name is required.",
                    nameof(name));
            }
            EnsureClassicAnimationTarget(shape);
            string relationshipId = AddSlideSound(PartRef, audio, contentType, extension);
            try {
                UpdateClassicAnimation(shape, animation => CopyClassicAnimation(
                    animation, playsSound: true,
                    stopsSound: stopExistingSounds,
                    soundRelationshipId: relationshipId, soundName: name));
            } catch {
                RemoveUnusedSlideSounds(new[] { relationshipId });
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
            ArgumentNullException.ThrowIfNull(shape);
            uint? shapeId = shape.Id;
            PowerPointClassicAnimation? animation = shapeId.HasValue
                ? ReadClassicAnimations().FirstOrDefault(item =>
                    item.ShapeId == shapeId.Value)
                : null;
            return ReadSlideSound(animation?.SoundRelationshipId);
        }

        internal void SetClassicAnimations(IReadOnlyList<PowerPointClassicAnimation> animations) {
            ArgumentNullException.ThrowIfNull(animations);
            IReadOnlyList<PowerPointClassicAnimation> previous = ReadClassicAnimations();
            string[] previousSounds = previous.Select(animation => animation.SoundRelationshipId)
                .Where(id => !string.IsNullOrEmpty(id)).Cast<string>().Distinct(StringComparer.Ordinal).ToArray();
            PowerPointClassicAnimation[] ordered = animations.OrderBy(animation => animation.Order).ToArray();

            PowerPointPartRef part = PartRef.ForMovingXmlNodes();
            OfficeXmlNode slide = part.Root;
            OfficeXmlNode timing = slide.Child(ClassicPresentationNamespace, "timing");
            bool onlyClassic = part.IsSdk ? HasOnlySdkClassicAnimationTiming(previous) : HasOnlyPackageClassicTiming(timing, previous);
            bool preserveOtherTiming = timing.Exists && previous.Count > 0 && !onlyClassic;
            if (preserveOtherTiming && !part.IsSdk) {
                // Keep complex producer sequences on the existing mixed-timing editing path.
                _ = SlidePart;
                part = part.Current;
                slide = part.Root;
                timing = slide.Child(ClassicPresentationNamespace, "timing");
                onlyClassic = HasOnlySdkClassicAnimationTiming(previous);
                preserveOtherTiming = timing.Exists && !onlyClassic;
            }
            if (preserveOtherTiming) {
                var existing = (Timing)timing.SdkElement!;
                if (ordered.Length > 0) CheckMixedClassicTimingIds(slide, timing, existing, previous, ordered);
                RemoveClassicTimingNodes(existing, previous);
            }

            if (ordered.Length == 0) {
                if (onlyClassic) timing.Remove();
            } else {
                uint firstId = preserveOtherTiming ? GetNextTimingId() : 2U;
                var replacement = CreateClassicTimingMarkup(slide, ordered, firstId);
                if (preserveOtherTiming) {
                    AppendClassicTiming((Timing)timing.SdkElement!, (ChildTimeNodeList)replacement.Children.SdkElement!,
                        replacement.Builds.SdkElement as BuildList);
                } else {
                    timing.Remove();
                    slide.InsertBefore(replacement.Timing, slide.Child(ClassicPresentationNamespace, "extLst"));
                }
            }
            WriteClassicAnimationMetadata(slide, ordered);
            part.MarkEdited();
            RemoveUnusedSlideSounds(previousSounds);
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

        private static void CheckMixedClassicTimingIds(OfficeXmlNode slide, OfficeXmlNode timing, Timing existing,
            IReadOnlyList<PowerPointClassicAnimation> previous, IReadOnlyList<PowerPointClassicAnimation> replacement) {
            ulong count = CountClassicTimingIds(replacement);
            if (MaximumTimingId(slide) + count <= uint.MaxValue) return;

            // The ids at the top of the range can belong to classic actions that the edit will remove.
            Timing staged = CopyClassicTimingForPreflight(existing);
            RemoveClassicTimingNodes(staged, previous);
            uint remaining = Math.Max(MaximumTimingId(slide, timing), MaximumTimingId(OfficeXmlNode.FromSdk(staged)));
            if (remaining + count > uint.MaxValue) throw TimingIdsExhausted();
        }

        private static Timing CopyClassicTimingForPreflight(Timing source) {
            var copy = (Timing)source.CloneNode(false);
            var pending = new Stack<(OpenXmlCompositeElement Source, OpenXmlCompositeElement Target)>();
            pending.Push((source, copy));
            while (pending.Count > 0) {
                (OpenXmlCompositeElement original, OpenXmlCompositeElement target) = pending.Pop();
                foreach (OpenXmlElement child in original.ChildElements) {
                    if (child is OpenXmlCompositeElement composite) {
                        var detached = (OpenXmlCompositeElement)composite.CloneNode(false);
                        target.Append(detached);
                        pending.Push((composite, detached));
                    } else {
                        target.Append(child.CloneNode(true));
                    }
                }
            }
            return copy;
        }

        private static void RemoveClassicTimingNodes(Timing timing, IReadOnlyList<PowerPointClassicAnimation> animations) {
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

        private static bool IsClassicHiddenAction(SetBehavior set, PowerPointClassicAnimation animation) =>
            IsClassicHiddenAction(OfficeXmlNode.FromSdk(set), animation);

        private static bool IsClassicDimAction(AnimateColor color, PowerPointClassicAnimation animation) =>
            IsClassicDimAction(OfficeXmlNode.FromSdk(color), animation);

        private static bool IsClassicSoundAction(Audio audio, PowerPointClassicAnimation animation) =>
            IsClassicSoundAction(OfficeXmlNode.FromSdk(audio), animation);

        private static bool IsClassicStopSoundAction(Command command, PowerPointClassicAnimation animation) =>
            IsClassicStopSoundAction(OfficeXmlNode.FromSdk(command), animation);

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

        internal bool HasOnlyClassicAnimationTiming() {
            IReadOnlyList<PowerPointClassicAnimation> animations = ReadClassicAnimations();
            PowerPointPartRef part = PartRef.ForMovingXmlNodes();
            return part.IsSdk ? HasOnlySdkClassicAnimationTiming(animations)
                : HasOnlyPackageClassicTiming(part.Root.Child(ClassicPresentationNamespace, "timing"), animations);
        }

        private bool HasOnlySdkClassicAnimationTiming(IReadOnlyList<PowerPointClassicAnimation> animations) {
            Timing? timing = SlideRoot.Timing;
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
                // The SDK interop surface can contain a typed action in an unexpected position.
                // Inspect those actions before deciding whether it is safe to replace the entire timing tree.
                && HasOnlyClassicStandardActions(timing.Descendants()
                    .Where(element => element is SetBehavior or AnimateColor or Audio or Command)
                    .Select(OfficeXmlNode.FromSdk).ToArray(), animations);
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
            ArgumentNullException.ThrowIfNull(shape);
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

    }
}
