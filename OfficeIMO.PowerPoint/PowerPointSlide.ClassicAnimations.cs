using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Presentation;

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
                var effectTime = new CommonTimeNode {
                    Id = timingId++,
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
                    new ChildTimeNodeList(effect)) {
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
            foreach (AnimateEffect effect in timing.Descendants<AnimateEffect>()
                         .ToArray()) {
                string? value = effect.Descendants<ShapeTarget>()
                    .FirstOrDefault()?.ShapeId?.Value;
                if (!uint.TryParse(value, NumberStyles.Integer,
                        CultureInfo.InvariantCulture, out uint shapeId)
                    || !animations.Any(animation =>
                        animation.ShapeId == shapeId
                        && effect.CommonBehavior?.CommonTimeNode?.PresetId?.Value
                        == (int)(byte)animation.Effect
                        && effect.CommonBehavior.CommonTimeNode.PresetSubtype?.Value
                        == animation.Direction)) {
                    continue;
                }
                OpenXmlElement? owner = effect
                    .Ancestors<ParallelTimeNode>().FirstOrDefault();
                (owner ?? effect).Remove();
            }
            var shapeIds = new HashSet<string>(animations.Select(animation =>
                animation.ShapeId.ToString(CultureInfo.InvariantCulture)),
                StringComparer.Ordinal);
            foreach (BuildParagraph build in timing.Descendants<BuildParagraph>()
                         .Where(item => shapeIds.Contains(item.ShapeId?.Value
                             ?? string.Empty)).ToArray()) {
                build.Remove();
            }
            foreach (BuildList list in timing.Descendants<BuildList>()
                         .Where(item => !item.HasChildren).ToArray()) {
                list.Remove();
            }
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
                && HasOnlyClassicVisibilitySetBehaviors(timing, animations);
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
