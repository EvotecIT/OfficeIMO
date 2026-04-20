using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Represents a single slide in a presentation.
    /// </summary>
    public partial class PowerPointSlide {
        private readonly List<PowerPointShape> _shapes = new();
        private readonly SlidePart _slidePart;
        private PowerPointNotes? _notes;
        private uint _nextShapeId = 2;
        private const string P14Namespace = "http://schemas.microsoft.com/office/powerpoint/2010/main";
        private const string P159Namespace = "http://schemas.microsoft.com/office/powerpoint/2015/09/main";
        private const string MarkupCompatibilityNamespace = "http://schemas.openxmlformats.org/markup-compatibility/2006";

        internal PowerPointSlide(SlidePart slidePart) {
            _slidePart = slidePart;
            LoadExistingShapes();
        }

        internal SlidePart SlidePart => _slidePart;

        /// <summary>
        ///     Collection of shapes on the slide.
        /// </summary>
        public IReadOnlyList<PowerPointShape> Shapes => _shapes;

        /// <summary>
        ///     Enumerates all textbox shapes on the slide.
        /// </summary>
        public IEnumerable<PowerPointTextBox> TextBoxes => _shapes.OfType<PowerPointTextBox>();

        /// <summary>
        ///     Enumerates all picture shapes on the slide.
        /// </summary>
        public IEnumerable<PowerPointPicture> Pictures => _shapes.OfType<PowerPointPicture>();

        /// <summary>
        ///     Enumerates all table shapes on the slide.
        /// </summary>
        public IEnumerable<PowerPointTable> Tables => _shapes.OfType<PowerPointTable>();

        /// <summary>
        ///     Enumerates all charts on the slide.
        /// </summary>
        public IEnumerable<PowerPointChart> Charts => _shapes.OfType<PowerPointChart>();

        /// <summary>
        ///     Retrieves shapes that are within or intersect the provided bounds.
        /// </summary>
        public IReadOnlyList<PowerPointShape> GetShapesInBounds(PowerPointLayoutBox bounds, bool includePartial = true) {
            if (includePartial) {
                return _shapes
                    .Where(shape =>
                        shape.Right >= bounds.Left &&
                        shape.Left <= bounds.Right &&
                        shape.Bottom >= bounds.Top &&
                        shape.Top <= bounds.Bottom)
                    .ToList();
            }

            return _shapes
                .Where(shape =>
                    shape.Left >= bounds.Left &&
                    shape.Top >= bounds.Top &&
                    shape.Right <= bounds.Right &&
                    shape.Bottom <= bounds.Bottom)
                .ToList();
        }

        /// <summary>
        ///     Retrieves shapes using bounds defined in centimeters.
        /// </summary>
        public IReadOnlyList<PowerPointShape> GetShapesInBoundsCm(double leftCm, double topCm, double widthCm, double heightCm, bool includePartial = true) {
            return GetShapesInBounds(PowerPointLayoutBox.FromCentimeters(leftCm, topCm, widthCm, heightCm), includePartial);
        }

        /// <summary>
        ///     Retrieves shapes using bounds defined in inches.
        /// </summary>
        public IReadOnlyList<PowerPointShape> GetShapesInBoundsInches(double leftInches, double topInches, double widthInches, double heightInches, bool includePartial = true) {
            return GetShapesInBounds(PowerPointLayoutBox.FromInches(leftInches, topInches, widthInches, heightInches), includePartial);
        }

        /// <summary>
        ///     Retrieves shapes using bounds defined in points.
        /// </summary>
        public IReadOnlyList<PowerPointShape> GetShapesInBoundsPoints(double leftPoints, double topPoints, double widthPoints, double heightPoints, bool includePartial = true) {
            return GetShapesInBounds(PowerPointLayoutBox.FromPoints(leftPoints, topPoints, widthPoints, heightPoints), includePartial);
        }

        /// <summary>
        ///     Notes associated with the slide.
        /// </summary>
        public PowerPointNotes Notes => _notes ??= new PowerPointNotes(_slidePart);

        /// <summary>
        ///     Gets or sets the slide background color in hex format (e.g. "FF0000").
        /// </summary>
        public string? BackgroundColor {
            get {
                CommonSlideData? common = SlideRoot.CommonSlideData;
                Background? bg = common?.Background;
                A.SolidFill? solid = bg?.BackgroundProperties?.GetFirstChild<A.SolidFill>();
                return solid?.RgbColorModelHex?.Val;
            }
            set {
                CommonSlideData common = SlideRoot.CommonSlideData ??= new CommonSlideData(new ShapeTree());
                if (value == null) {
                    BackgroundProperties? properties = common.Background?.BackgroundProperties;
                    if (properties == null) {
                        return;
                    }

                    properties.RemoveAllChildren<A.SolidFill>();
                    if (!properties.HasChildren) {
                        common.Background = null;
                    }
                    return;
                }

                Background bg = common.Background ?? new Background();
                BackgroundProperties props = bg.BackgroundProperties ?? new BackgroundProperties();
                RemoveBackgroundFillChildren(props);
                props.Append(new A.SolidFill(new A.RgbColorModelHex { Val = value }));
                bg.BackgroundProperties = props;
                common.Background = bg;
            }
        }

        /// <summary>
        ///     Sets a background image for the slide.
        /// </summary>
        public void SetBackgroundImage(string imagePath) {
            if (imagePath == null) {
                throw new ArgumentNullException(nameof(imagePath));
            }
            if (!File.Exists(imagePath)) {
                throw new FileNotFoundException("Image file not found.", imagePath);
            }

            A.Blip? previousBlip = GetBackgroundBlip();
            string? previousRelationshipId = previousBlip?.Embed?.Value;
            ImagePartType imageType = GetImagePartType(imagePath);
            PartTypeInfo partTypeInfo = imageType.ToPartTypeInfo();
            string imageExtension = PowerPointPartFactory.GetImageExtension(imageType, imagePath);
            string imagePartUri = PowerPointPartFactory.GetIndexedPartUri(
                _slidePart.OpenXmlPackage,
                "ppt/media",
                "image",
                imageExtension,
                allowBaseWithoutIndex: false);

            ImagePart imagePart = PowerPointPartFactory.CreatePart<ImagePart>(
                _slidePart,
                partTypeInfo.ContentType,
                imagePartUri);

            using FileStream stream = new(imagePath, FileMode.Open, FileAccess.Read);
            imagePart.FeedData(stream);
            string relationshipId = _slidePart.GetIdOfPart(imagePart);

            CommonSlideData common = SlideRoot.CommonSlideData ??= new CommonSlideData(new ShapeTree());
            Background background = common.Background ?? new Background();
            BackgroundProperties props = background.BackgroundProperties ?? new BackgroundProperties();

            RemoveBackgroundFillChildren(props);

            props.Append(new A.BlipFill(
                new A.Blip { Embed = relationshipId },
                new A.Stretch(new A.FillRectangle())
            ));

            background.BackgroundProperties = props;
            common.Background = background;

            RemoveUnusedImagePart(previousRelationshipId, previousBlip);
        }

        /// <summary>
        ///     Sets a linear gradient background for the slide using two 6-digit hex colors.
        /// </summary>
        public void SetBackgroundGradient(string startColor, string endColor, double angleDegrees = 135d) {
            if (string.IsNullOrWhiteSpace(startColor)) {
                throw new ArgumentException("Gradient start color cannot be null or empty.", nameof(startColor));
            }

            if (string.IsNullOrWhiteSpace(endColor)) {
                throw new ArgumentException("Gradient end color cannot be null or empty.", nameof(endColor));
            }

            string normalizedStart = NormalizeHexColor(startColor);
            string normalizedEnd = NormalizeHexColor(endColor);

            CommonSlideData common = SlideRoot.CommonSlideData ??= new CommonSlideData(new ShapeTree());
            Background background = common.Background ?? new Background();
            BackgroundProperties props = background.BackgroundProperties ?? new BackgroundProperties();

            RemoveBackgroundFillChildren(props);

            A.GradientFill gradient = new() { RotateWithShape = true };
            A.GradientStopList stops = new();
            stops.Append(
                new A.GradientStop(new A.RgbColorModelHex { Val = normalizedStart }) {
                    Position = 0
                });
            stops.Append(
                new A.GradientStop(new A.RgbColorModelHex { Val = normalizedEnd }) {
                    Position = 100000
                });
            gradient.Append(stops);
            gradient.Append(new A.LinearGradientFill {
                Angle = ToOpenXmlAngle(angleDegrees),
                Scaled = false
            });

            props.Append(gradient);
            background.BackgroundProperties = props;
            common.Background = background;
        }

        /// <summary>
        ///     Clears any background image from the slide.
        /// </summary>
        public void ClearBackgroundImage() {
            CommonSlideData? common = SlideRoot.CommonSlideData;
            if (common?.Background?.BackgroundProperties == null) {
                return;
            }

            A.Blip? previousBlip = GetBackgroundBlip();
            string? previousRelationshipId = previousBlip?.Embed?.Value;
            common.Background.BackgroundProperties.RemoveAllChildren<A.BlipFill>();
            if (!common.Background.BackgroundProperties.HasChildren) {
                common.Background = null;
            }

            RemoveUnusedImagePart(previousRelationshipId, previousBlip);
        }

        /// <summary>
        ///     Transition applied when moving to this slide.
        /// </summary>
        public SlideTransition Transition {
            get {
                Transition? t = GetTransitionElement();
                if (t == null) {
                    return SlideTransition.None;
                }

                if (t.GetFirstChild<FadeTransition>() != null) {
                    return SlideTransition.Fade;
                }

                if (t.GetFirstChild<WipeTransition>() != null) {
                    return SlideTransition.Wipe;
                }

                BlindsTransition? blinds = t.GetFirstChild<BlindsTransition>();
                if (blinds != null) {
                    return blinds.Direction?.Value == DirectionValues.Vertical
                        ? SlideTransition.BlindsVertical
                        : SlideTransition.BlindsHorizontal;
                }

                CombTransition? comb = t.GetFirstChild<CombTransition>();
                if (comb != null) {
                    return comb.Direction?.Value == DirectionValues.Vertical
                        ? SlideTransition.CombVertical
                        : SlideTransition.CombHorizontal;
                }

                PushTransition? push = t.GetFirstChild<PushTransition>();
                if (push != null) {
                    TransitionSlideDirectionValues? direction = push.Direction?.Value;
                    if (direction == TransitionSlideDirectionValues.Up) {
                        return SlideTransition.PushUp;
                    }

                    if (direction == TransitionSlideDirectionValues.Down) {
                        return SlideTransition.PushDown;
                    }

                    if (direction == TransitionSlideDirectionValues.Right) {
                        return SlideTransition.PushRight;
                    }

                    return SlideTransition.PushLeft;
                }

                if (t.GetFirstChild<CutTransition>() != null) {
                    return SlideTransition.Cut;
                }

                if (t.GetFirstChild<P14.FlashTransition>() != null) {
                    return SlideTransition.Flash;
                }

                P14.WarpTransition? warp = t.GetFirstChild<P14.WarpTransition>();
                if (warp != null) {
                    return warp.Direction?.Value == TransitionInOutDirectionValues.Out
                        ? SlideTransition.WarpOut
                        : SlideTransition.WarpIn;
                }

                if (t.GetFirstChild<P14.PrismTransition>() != null) {
                    return SlideTransition.Prism;
                }

                P14.FerrisTransition? ferris = t.GetFirstChild<P14.FerrisTransition>();
                if (ferris != null) {
                    return ferris.Direction?.Value == P14.TransitionLeftRightDirectionTypeValues.Right
                        ? SlideTransition.FerrisRight
                        : SlideTransition.FerrisLeft;
                }

                if (HasMorphTransition(t)) {
                    return SlideTransition.Morph;
                }

                return SlideTransition.None;
            }
            set {
                SlideTransitionSpeed? speed = TransitionSpeed;
                double? durationSeconds = TransitionDurationSeconds;
                bool? advanceOnClick = TransitionAdvanceOnClick;
                double? advanceAfterSeconds = TransitionAdvanceAfterSeconds;

                RemoveTransitionMarkup();
                if (value == SlideTransition.None) {
                    return;
                }

                if (value == SlideTransition.Morph) {
                    SetMorphTransition();
                    return;
                }

                Transition transition = new();
                switch (value) {
                    case SlideTransition.Fade:
                        transition.Append(new FadeTransition());
                        break;
                    case SlideTransition.Wipe:
                        transition.Append(new WipeTransition());
                        break;
                    case SlideTransition.BlindsVertical:
                        transition.Append(new BlindsTransition { Direction = DirectionValues.Vertical });
                        break;
                    case SlideTransition.BlindsHorizontal:
                        transition.Append(new BlindsTransition { Direction = DirectionValues.Horizontal });
                        break;
                    case SlideTransition.CombHorizontal:
                        transition.Append(new CombTransition { Direction = DirectionValues.Horizontal });
                        break;
                    case SlideTransition.CombVertical:
                        transition.Append(new CombTransition { Direction = DirectionValues.Vertical });
                        break;
                    case SlideTransition.PushUp:
                        transition.Append(new PushTransition { Direction = TransitionSlideDirectionValues.Up });
                        break;
                    case SlideTransition.PushDown:
                        transition.Append(new PushTransition { Direction = TransitionSlideDirectionValues.Down });
                        break;
                    case SlideTransition.PushLeft:
                        transition.Append(new PushTransition { Direction = TransitionSlideDirectionValues.Left });
                        break;
                    case SlideTransition.PushRight:
                        transition.Append(new PushTransition { Direction = TransitionSlideDirectionValues.Right });
                        break;
                    case SlideTransition.Cut:
                        transition.Append(new CutTransition());
                        break;
                    case SlideTransition.Flash:
                        transition.AddNamespaceDeclaration("p14", P14Namespace);
                        transition.Append(new P14.FlashTransition());
                        break;
                    case SlideTransition.WarpIn:
                        transition.AddNamespaceDeclaration("p14", P14Namespace);
                        transition.Append(new P14.WarpTransition { Direction = TransitionInOutDirectionValues.In });
                        break;
                    case SlideTransition.WarpOut:
                        transition.AddNamespaceDeclaration("p14", P14Namespace);
                        transition.Append(new P14.WarpTransition { Direction = TransitionInOutDirectionValues.Out });
                        break;
                    case SlideTransition.Prism:
                        transition.AddNamespaceDeclaration("p14", P14Namespace);
                        transition.Append(new P14.PrismTransition { IsContent = true });
                        break;
                    case SlideTransition.FerrisLeft:
                        transition.AddNamespaceDeclaration("p14", P14Namespace);
                        transition.Append(new P14.FerrisTransition { Direction = P14.TransitionLeftRightDirectionTypeValues.Left });
                        break;
                    case SlideTransition.FerrisRight:
                        transition.AddNamespaceDeclaration("p14", P14Namespace);
                        transition.Append(new P14.FerrisTransition { Direction = P14.TransitionLeftRightDirectionTypeValues.Right });
                        break;
                }

                SlideRoot.Transition = transition;
                ApplyTransitionSettings(GetTransitionElement(), speed, durationSeconds, advanceOnClick, advanceAfterSeconds);
            }
        }

        /// <summary>
        ///     Gets or sets the optional transition playback speed.
        /// </summary>
        public SlideTransitionSpeed? TransitionSpeed {
            get {
                Transition? transition = GetTransitionElement();
                if (transition?.Speed?.Value == null) {
                    return null;
                }

                var speedValue = transition.Speed.Value;
                if (speedValue == TransitionSpeedValues.Slow) {
                    return SlideTransitionSpeed.Slow;
                }

                if (speedValue == TransitionSpeedValues.Fast) {
                    return SlideTransitionSpeed.Fast;
                }

                return SlideTransitionSpeed.Medium;
            }
            set {
                Transition? transition = GetTransitionElement();
                if (transition == null) {
                    return;
                }

                transition.Speed = value switch {
                    SlideTransitionSpeed.Slow => TransitionSpeedValues.Slow,
                    SlideTransitionSpeed.Fast => TransitionSpeedValues.Fast,
                    SlideTransitionSpeed.Medium => TransitionSpeedValues.Medium,
                    _ => null
                };
            }
        }

        /// <summary>
        ///     Gets or sets the optional transition duration in seconds.
        /// </summary>
        public double? TransitionDurationSeconds {
            get {
                uint? milliseconds = ParseTransitionMilliseconds(GetTransitionElement()?.Duration?.Value);
                if (milliseconds == null) {
                    return null;
                }

                return milliseconds.Value / 1000.0;
            }
            set {
                Transition? transition = GetTransitionElement();
                if (transition == null) {
                    return;
                }

                if (value.HasValue) {
                    EnsureTransitionNamespace(transition, "p14", P14Namespace);
                }

                transition.Duration = ToMillisecondsString(value);
            }
        }

        /// <summary>
        ///     Gets or sets whether clicking advances the slide.
        /// </summary>
        public bool? TransitionAdvanceOnClick {
            get {
                return GetTransitionElement()?.AdvanceOnClick?.Value;
            }
            set {
                Transition? transition = GetTransitionElement();
                if (transition == null) {
                    return;
                }

                transition.AdvanceOnClick = value;
            }
        }

        /// <summary>
        ///     Gets or sets the optional automatic-advance time in seconds.
        /// </summary>
        public double? TransitionAdvanceAfterSeconds {
            get {
                uint? milliseconds = ParseTransitionMilliseconds(GetTransitionElement()?.AdvanceAfterTime?.Value);
                if (milliseconds == null) {
                    return null;
                }

                return milliseconds.Value / 1000.0;
            }
            set {
                Transition? transition = GetTransitionElement();
                if (transition == null) {
                    return;
                }

                transition.AdvanceAfterTime = ToMillisecondsString(value);
            }
        }

        /// <summary>
        ///     Gets or sets whether the slide is hidden in slide show mode.
        /// </summary>
        public bool Hidden {
            get {
                if (SlideRoot.Show?.Value != null) {
                    return !SlideRoot.Show.Value;
                }

                return IsHiddenShowValue(GetLegacySlideIdShowValue(GetSlideId()));
            }
            set {
                SlideId slideId = GetSlideId();
                slideId.RemoveAttribute("show", string.Empty);

                if (value) {
                    SlideRoot.Show = false;
                } else {
                    SlideRoot.Show = null;
                }
            }
        }

        /// <summary>
        ///     Hides the slide in slide show mode.
        /// </summary>
        public void Hide() => Hidden = true;

        /// <summary>
        ///     Shows the slide in slide show mode.
        /// </summary>
        public void Show() => Hidden = false;

        private static bool HasMorphTransition(Transition transition) {
            return transition.ChildElements.Any(element =>
                (element.LocalName == "morph" && element.NamespaceUri == P159Namespace) ||
                (element.LocalName == "prstTrans" &&
                 element.NamespaceUri == "http://schemas.microsoft.com/office/powerpoint/2012/main" &&
                 string.Equals(element.GetAttribute("prst", string.Empty).Value, "morph", StringComparison.OrdinalIgnoreCase)));
        }

        private static OpenXmlUnknownElement CreateMorphTransition() {
            OpenXmlUnknownElement morph = new OpenXmlUnknownElement("p159", "morph", P159Namespace);
            morph.AddNamespaceDeclaration("p159", P159Namespace);
            morph.SetAttribute(new OpenXmlAttribute("option", string.Empty, "byObject"));
            return morph;
        }

        private static uint? ToMilliseconds(double? seconds) {
            if (!seconds.HasValue) {
                return null;
            }

            if (seconds.Value < 0) {
                return 0;
            }

            return (uint)Math.Round(seconds.Value * 1000.0, MidpointRounding.AwayFromZero);
        }

        private static string? ToMillisecondsString(double? seconds) {
            uint? milliseconds = ToMilliseconds(seconds);
            return milliseconds?.ToString();
        }

        private static uint? ParseTransitionMilliseconds(string? value) {
            if (string.IsNullOrWhiteSpace(value)) {
                return null;
            }

            return uint.TryParse(value, out uint parsed)
                ? parsed
                : null;
        }

        private static void ApplyTransitionSettings(Transition? transition, SlideTransitionSpeed? speed, double? durationSeconds, bool? advanceOnClick, double? advanceAfterSeconds) {
            if (transition == null) {
                return;
            }

            transition.Speed = speed switch {
                SlideTransitionSpeed.Slow => TransitionSpeedValues.Slow,
                SlideTransitionSpeed.Fast => TransitionSpeedValues.Fast,
                SlideTransitionSpeed.Medium => TransitionSpeedValues.Medium,
                _ => null
            };

            if (durationSeconds.HasValue) {
                EnsureTransitionNamespace(transition, "p14", P14Namespace);
            }

            transition.Duration = ToMillisecondsString(durationSeconds);
            transition.AdvanceOnClick = advanceOnClick;
            transition.AdvanceAfterTime = ToMillisecondsString(advanceAfterSeconds);
        }

        private static void EnsureTransitionNamespace(Transition transition, string prefix, string uri) {
            if (!string.Equals(transition.LookupNamespace(prefix), uri, StringComparison.Ordinal)) {
                transition.AddNamespaceDeclaration(prefix, uri);
            }
        }

        private Transition? GetTransitionElement() {
            Transition? transition = SlideRoot.Transition;
            if (transition != null) {
                return transition;
            }

            AlternateContent? alternateContent = GetTransitionAlternateContent();
            if (alternateContent == null) {
                return null;
            }

            foreach (AlternateContentChoice choice in alternateContent.Elements<AlternateContentChoice>()) {
                Transition? choiceTransition = choice.GetFirstChild<Transition>();
                if (choiceTransition != null) {
                    return choiceTransition;
                }
            }

            return alternateContent.GetFirstChild<AlternateContentFallback>()?.GetFirstChild<Transition>();
        }

        private AlternateContent? GetTransitionAlternateContent() {
            return SlideRoot.Elements<AlternateContent>()
                .FirstOrDefault(content =>
                    content.Elements<AlternateContentChoice>().Any(choice => choice.GetFirstChild<Transition>() != null) ||
                    content.GetFirstChild<AlternateContentFallback>()?.GetFirstChild<Transition>() != null);
        }

        private void RemoveTransitionMarkup() {
            SlideRoot.Transition = null;
            GetTransitionAlternateContent()?.Remove();
        }

        private void SetMorphTransition() {
            Slide slide = SlideRoot;
            slide.AddNamespaceDeclaration("mc", MarkupCompatibilityNamespace);
            slide.AddNamespaceDeclaration("p159", P159Namespace);
            slide.MCAttributes = MergeIgnorableNamespace(slide.MCAttributes, "p159");

            Transition morphTransition = new();
            morphTransition.AddNamespaceDeclaration("p159", P159Namespace);
            morphTransition.Append(CreateMorphTransition());

            AlternateContentChoice choice = new() { Requires = "p159" };
            choice.Append(morphTransition);

            AlternateContentFallback fallback = new();
            fallback.Append(new Transition(new FadeTransition()));

            AlternateContent alternateContent = new();
            alternateContent.Append(choice);
            alternateContent.Append(fallback);

            InsertTransitionAlternateContent(alternateContent);
        }

        private void InsertTransitionAlternateContent(AlternateContent alternateContent) {
            Slide slide = SlideRoot;
            OpenXmlElement? insertBefore = slide.GetFirstChild<Timing>();
            insertBefore ??= slide.GetFirstChild<ExtensionListWithModification>();
            if (insertBefore != null) {
                slide.InsertBefore(alternateContent, insertBefore);
                return;
            }

            if (slide.ColorMapOverride != null) {
                slide.InsertAfter(alternateContent, slide.ColorMapOverride);
                return;
            }

            if (slide.CommonSlideData != null) {
                slide.InsertAfter(alternateContent, slide.CommonSlideData);
                return;
            }

            slide.Append(alternateContent);
        }

        private static MarkupCompatibilityAttributes MergeIgnorableNamespace(
            MarkupCompatibilityAttributes? existingAttributes,
            string namespacePrefix) {
            List<string> prefixes = (existingAttributes?.Ignorable?.Value ?? string.Empty)
                .Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries)
                .ToList();

            if (!prefixes.Contains(namespacePrefix, StringComparer.Ordinal)) {
                prefixes.Add(namespacePrefix);
            }

            return new MarkupCompatibilityAttributes {
                Ignorable = string.Join(" ", prefixes)
            };
        }

        private static void RemoveBackgroundFillChildren(BackgroundProperties properties) {
            properties.RemoveAllChildren<A.BlipFill>();
            properties.RemoveAllChildren<A.GradientFill>();
            properties.RemoveAllChildren<A.GroupFill>();
            properties.RemoveAllChildren<A.NoFill>();
            properties.RemoveAllChildren<A.PatternFill>();
            properties.RemoveAllChildren<A.SolidFill>();
        }

        private static string NormalizeHexColor(string value) {
            string normalized = value.Trim();
            if (normalized.StartsWith("#", StringComparison.Ordinal)) {
                normalized = normalized.Substring(1);
            }

            if (normalized.Length != 6 || normalized.Any(c => !Uri.IsHexDigit(c))) {
                throw new ArgumentException("Color must be a 6-digit hex value.", nameof(value));
            }

            return normalized.ToUpperInvariant();
        }

        private static int ToOpenXmlAngle(double degrees) {
            double normalized = degrees % 360d;
            if (normalized < 0) {
                normalized += 360d;
            }

            return (int)Math.Round(normalized * 60000d);
        }

        private A.Blip? GetBackgroundBlip() {
            return SlideRoot.CommonSlideData?.Background?.BackgroundProperties?.GetFirstChild<A.BlipFill>()?.Blip;
        }

        private void RemoveUnusedImagePart(string? relationshipId, A.Blip? currentBlip) {
            string resolvedRelationshipId = relationshipId ?? string.Empty;
            if (string.IsNullOrWhiteSpace(resolvedRelationshipId)) {
                return;
            }
            if (IsImageRelationshipReferenced(resolvedRelationshipId, currentBlip)) {
                return;
            }

            try {
                _slidePart.DeletePart(resolvedRelationshipId);
            } catch (ArgumentOutOfRangeException) {
                // The previous relationship may already be absent on damaged input.
            }
        }

        private bool IsImageRelationshipReferenced(string relationshipId, A.Blip? currentBlip) {
            return SlideRoot
                .Descendants<A.Blip>()
                .Any(blip => !ReferenceEquals(blip, currentBlip) && blip.Embed?.Value == relationshipId);
        }

        private SlideId GetSlideId() {
            PresentationPart presentationPart = _slidePart.GetParentParts()
                .OfType<PresentationPart>()
                .FirstOrDefault()
                ?? throw new InvalidOperationException("Slide is not attached to a presentation.");

            SlideIdList? slideIdList = presentationPart.Presentation?.SlideIdList;
            if (slideIdList == null) {
                throw new InvalidOperationException("Presentation has no slide list.");
            }

            string relId = presentationPart.GetIdOfPart(_slidePart);
            SlideId? slideId = slideIdList.Elements<SlideId>()
                .FirstOrDefault(id => id.RelationshipId?.Value == relId);

            if (slideId == null) {
                throw new InvalidOperationException("Slide not found in presentation.");
            }

            return slideId;
        }

        private static string? GetLegacySlideIdShowValue(SlideId slideId) {
            return slideId.GetAttributes()
                .FirstOrDefault(attribute =>
                    attribute.LocalName == "show" && string.IsNullOrEmpty(attribute.NamespaceUri))
                .Value;
        }

        private static bool IsHiddenShowValue(string? showValue) {
            if (string.IsNullOrEmpty(showValue)) {
                return false;
            }

            return string.Equals(showValue, "0", StringComparison.Ordinal) ||
                   string.Equals(showValue, "false", StringComparison.OrdinalIgnoreCase);
        }

        private void NormalizeHiddenSlideMarkup() {
            SlideId slideId = GetSlideId();
            string? legacyShowValue = GetLegacySlideIdShowValue(slideId);
            if (SlideRoot.Show?.Value == null && IsHiddenShowValue(legacyShowValue)) {
                SlideRoot.Show = false;
            }

            slideId.RemoveAttribute("show", string.Empty);
        }

        /// <summary>
        ///     Gets the index of the layout used by this slide.
        /// </summary>
        public int LayoutIndex {
            get {
                SlideLayoutPart? layoutPart = _slidePart.SlideLayoutPart;
                if (layoutPart == null) {
                    return -1;
                }

                SlideMasterPart master = layoutPart.GetParentParts().OfType<SlideMasterPart>().First();
                SlideLayoutPart[] layouts = master.SlideLayoutParts.ToArray();
                for (int i = 0; i < layouts.Length; i++) {
                    if (layouts[i] == layoutPart) {
                        return i;
                    }
                }

                return -1;
            }
        }

        /// <summary>
        ///     Sets the slide layout using master and layout indexes.
        /// </summary>
        public void SetLayout(int masterIndex, int layoutIndex) {
            PresentationPart presentationPart = _slidePart.GetParentParts().OfType<PresentationPart>().First();

            SlideMasterPart[] masters = presentationPart.SlideMasterParts.ToArray();
            if (masterIndex < 0 || masterIndex >= masters.Length) {
                throw new ArgumentOutOfRangeException(nameof(masterIndex));
            }

            SlideMasterPart masterPart = masters[masterIndex];
            SlideLayoutPart[] layouts = masterPart.SlideLayoutParts.ToArray();
            if (layoutIndex < 0 || layoutIndex >= layouts.Length) {
                throw new ArgumentOutOfRangeException(nameof(layoutIndex));
            }

            SlideLayoutPart layoutPart = layouts[layoutIndex];
            SlideLayoutPart? current = _slidePart.SlideLayoutPart;
            if (current != null) {
                string relId = _slidePart.GetIdOfPart(current);
                _slidePart.DeletePart(relId);
            }

            _slidePart.AddPart(layoutPart);
        }

        /// <summary>
        ///     Sets the slide layout using a layout type.
        /// </summary>
        public void SetLayout(SlideLayoutValues layoutType, int masterIndex = 0) {
            int layoutIndex = GetLayoutIndex(layoutType, masterIndex);
            SetLayout(masterIndex, layoutIndex);
        }

        /// <summary>
        ///     Sets the slide layout using a layout name.
        /// </summary>
        public void SetLayout(string layoutName, int masterIndex = 0, bool ignoreCase = true) {
            int layoutIndex = GetLayoutIndex(layoutName, masterIndex, ignoreCase);
            SetLayout(masterIndex, layoutIndex);
        }

        private int GetLayoutIndex(SlideLayoutValues layoutType, int masterIndex) {
            PresentationPart presentationPart = _slidePart.GetParentParts().OfType<PresentationPart>().First();
            SlideMasterPart[] masters = presentationPart.SlideMasterParts.ToArray();
            if (masterIndex < 0 || masterIndex >= masters.Length) {
                throw new ArgumentOutOfRangeException(nameof(masterIndex));
            }

            SlideLayoutPart[] layouts = masters[masterIndex].SlideLayoutParts.ToArray();
            for (int i = 0; i < layouts.Length; i++) {
                SlideLayoutValues? type = layouts[i].SlideLayout?.Type?.Value;
                if (type == layoutType) {
                    return i;
                }
            }

            throw new InvalidOperationException($"Layout type '{layoutType}' not found for master {masterIndex}.");
        }

        private int GetLayoutIndex(string layoutName, int masterIndex, bool ignoreCase) {
            if (layoutName == null) {
                throw new ArgumentNullException(nameof(layoutName));
            }

            PresentationPart presentationPart = _slidePart.GetParentParts().OfType<PresentationPart>().First();
            SlideMasterPart[] masters = presentationPart.SlideMasterParts.ToArray();
            if (masterIndex < 0 || masterIndex >= masters.Length) {
                throw new ArgumentOutOfRangeException(nameof(masterIndex));
            }

            SlideLayoutPart[] layouts = masters[masterIndex].SlideLayoutParts.ToArray();
            StringComparison comparison = ignoreCase ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;
            for (int i = 0; i < layouts.Length; i++) {
                string name = layouts[i].SlideLayout?.CommonSlideData?.Name?.Value ?? string.Empty;
                if (string.Equals(name, layoutName, comparison)) {
                    return i;
                }
            }

            throw new InvalidOperationException($"Layout '{layoutName}' not found for master {masterIndex}.");
        }

        /// <summary>
        ///     Retrieves a shape by its name.
        /// </summary>
        public PowerPointShape? GetShape(string name) {
            if (name == null) {
                throw new ArgumentNullException(nameof(name));
            }

            return _shapes.FirstOrDefault(s => s.Name == name);
        }

        /// <summary>
        ///     Retrieves a textbox by its name.
        /// </summary>
        public PowerPointTextBox? GetTextBox(string name) {
            if (name == null) {
                throw new ArgumentNullException(nameof(name));
            }

            return TextBoxes.FirstOrDefault(tb => tb.Name == name);
        }

        /// <summary>
        ///     Textboxes that map to placeholders in the slide layout.
        /// </summary>
        public IReadOnlyList<PowerPointTextBox> Placeholders =>
            TextBoxes.Where(tb => tb.IsPlaceholder).ToList();

        /// <summary>
        ///     Retrieves the first placeholder textbox matching the specified type.
        /// </summary>
        public PowerPointTextBox? GetPlaceholder(PlaceholderValues placeholderType, uint? index = null) {
            IEnumerable<PowerPointTextBox> matches = TextBoxes
                .Where(tb => tb.PlaceholderType == placeholderType);

            if (index != null) {
                matches = matches.Where(tb => tb.PlaceholderIndex == index);
            }

            return matches.FirstOrDefault();
        }

        /// <summary>
        ///     Retrieves placeholders defined by the slide layout.
        /// </summary>
        public IReadOnlyList<PowerPointLayoutPlaceholderInfo> GetLayoutPlaceholders() {
            SlideLayoutPart? layoutPart = _slidePart.SlideLayoutPart;
            ShapeTree? shapeTree = layoutPart?.SlideLayout?.CommonSlideData?.ShapeTree;
            if (shapeTree == null) {
                return Array.Empty<PowerPointLayoutPlaceholderInfo>();
            }

            List<PowerPointLayoutPlaceholderInfo> placeholders = new();
            foreach (OpenXmlElement element in shapeTree.ChildElements) {
                PlaceholderShape? placeholder = GetLayoutPlaceholderShape(element);
                if (placeholder == null) {
                    continue;
                }

                string name = GetLayoutElementName(element);
                PowerPointLayoutBox? bounds = GetLayoutElementBounds(element);
                placeholders.Add(new PowerPointLayoutPlaceholderInfo(
                    name,
                    placeholder.Type?.Value,
                    placeholder.Index?.Value,
                    bounds));
            }

            return placeholders;
        }

        /// <summary>
        ///     Retrieves a layout placeholder by type and optional index.
        /// </summary>
        public PowerPointLayoutPlaceholderInfo? GetLayoutPlaceholder(PlaceholderValues placeholderType, uint? index = null) {
            foreach (PowerPointLayoutPlaceholderInfo placeholder in GetLayoutPlaceholders()) {
                if (placeholder.PlaceholderType != placeholderType) {
                    continue;
                }

                if (index != null && placeholder.PlaceholderIndex != index) {
                    continue;
                }

                return placeholder;
            }

            return null;
        }

        /// <summary>
        ///     Retrieves layout placeholder bounds by type and optional index.
        /// </summary>
        public PowerPointLayoutBox? GetLayoutPlaceholderBounds(PlaceholderValues placeholderType, uint? index = null) {
            PowerPointLayoutPlaceholderInfo? placeholder = GetLayoutPlaceholder(placeholderType, index);
            return placeholder?.Bounds;
        }

        /// <summary>
        ///     Retrieves a picture by its name.
        /// </summary>
        public PowerPointPicture? GetPicture(string name) {
            if (name == null) {
                throw new ArgumentNullException(nameof(name));
            }

            return Pictures.FirstOrDefault(p => p.Name == name);
        }

        /// <summary>
        ///     Retrieves a table by its name.
        /// </summary>
        public PowerPointTable? GetTable(string name) {
            if (name == null) {
                throw new ArgumentNullException(nameof(name));
            }

            return Tables.FirstOrDefault(t => t.Name == name);
        }

        /// <summary>
        ///     Replaces text across all textboxes on the slide.
        /// </summary>
        public int ReplaceText(string oldValue, string newValue, bool includeTables = true, bool includeNotes = false) {
            if (oldValue == null) {
                throw new ArgumentNullException(nameof(oldValue));
            }
            if (oldValue.Length == 0) {
                throw new ArgumentException("Old value cannot be empty.", nameof(oldValue));
            }

            string replacement = newValue ?? string.Empty;
            int count = 0;

            foreach (PowerPointTextBox textBox in TextBoxes) {
                count += textBox.ReplaceText(oldValue, replacement);
            }

            if (includeTables) {
                foreach (PowerPointTable table in Tables) {
                    for (int r = 0; r < table.Rows; r++) {
                        for (int c = 0; c < table.Columns; c++) {
                            count += table.GetCell(r, c).ReplaceText(oldValue, replacement);
                        }
                    }
                }
            }

            if (includeNotes && _slidePart.NotesSlidePart != null) {
                string notesText = Notes.Text ?? string.Empty;
                int occurrences = CountOccurrences(notesText, oldValue);
                if (occurrences > 0) {
                    Notes.Text = notesText.Replace(oldValue, replacement);
                    count += occurrences;
                }
            }

            return count;
        }

        /// <summary>
        ///     Retrieves a chart by its name.
        /// </summary>
        public PowerPointChart? GetChart(string name) {
            if (name == null) {
                throw new ArgumentNullException(nameof(name));
            }

            return Charts.FirstOrDefault(c => c.Name == name);
        }

        /// <summary>
        ///     Removes the specified shape from the slide.
        /// </summary>
        public void RemoveShape(PowerPointShape shape) {
            if (shape == null) {
                throw new ArgumentNullException(nameof(shape));
            }

            shape.Element.Remove();
            _shapes.Remove(shape);
        }

        private string GenerateUniqueName(string baseName) {
            int index = 1;
            string name;
            do {
                name = baseName + " " + index++;
            } while (_shapes.Any(s => s.Name == name));

            return name;
        }

        internal void Save() {
            NormalizeHiddenSlideMarkup();
            SlideRoot.Save();
            _notes?.Save();
        }

        private static int CountOccurrences(string value, string oldValue) {
            int count = 0;
            int index = 0;
            while (true) {
                index = value.IndexOf(oldValue, index, StringComparison.Ordinal);
                if (index < 0) {
                    break;
                }
                count++;
                index += oldValue.Length;
            }
            return count;
        }

        private static PlaceholderShape? GetLayoutPlaceholderShape(OpenXmlElement element) {
            return element switch {
                Shape s => s.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape,
                DocumentFormat.OpenXml.Presentation.Picture p => p.NonVisualPictureProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape,
                GraphicFrame g => g.NonVisualGraphicFrameProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape,
                _ => null
            };
        }

        private static string GetLayoutElementName(OpenXmlElement element) {
            return element switch {
                Shape s => s.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value ?? string.Empty,
                DocumentFormat.OpenXml.Presentation.Picture p => p.NonVisualPictureProperties?.NonVisualDrawingProperties?.Name?.Value ?? string.Empty,
                GraphicFrame g => g.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties?.Name?.Value ?? string.Empty,
                _ => string.Empty
            };
        }

        private static PowerPointLayoutBox? GetLayoutElementBounds(OpenXmlElement element) {
            return element switch {
                Shape s => GetLayoutElementBounds(s.ShapeProperties?.Transform2D),
                DocumentFormat.OpenXml.Presentation.Picture p => GetLayoutElementBounds(p.ShapeProperties?.Transform2D),
                GraphicFrame g => GetLayoutElementBounds(g.Transform),
                _ => null
            };
        }

        private static PowerPointLayoutBox? GetLayoutElementBounds(A.Transform2D? transform) {
            long? x = transform?.Offset?.X?.Value;
            long? y = transform?.Offset?.Y?.Value;
            long? cx = transform?.Extents?.Cx?.Value;
            long? cy = transform?.Extents?.Cy?.Value;
            if (x == null || y == null || cx == null || cy == null) {
                return null;
            }

            return new PowerPointLayoutBox(x.Value, y.Value, cx.Value, cy.Value);
        }

        private static PowerPointLayoutBox? GetLayoutElementBounds(Transform? transform) {
            long? x = transform?.Offset?.X?.Value;
            long? y = transform?.Offset?.Y?.Value;
            long? cx = transform?.Extents?.Cx?.Value;
            long? cy = transform?.Extents?.Cy?.Value;
            if (x == null || y == null || cx == null || cy == null) {
                return null;
            }

            return new PowerPointLayoutBox(x.Value, y.Value, cx.Value, cy.Value);
        }

        private void LoadExistingShapes() {
            ShapeTree? tree = SlideRoot.CommonSlideData?.ShapeTree;
            if (tree == null) {
                return;
            }

            uint maxId = 1;
            foreach (OpenXmlElement element in tree.ChildElements) {
                uint? id = element switch {
                    Shape s => s.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value,
                    DocumentFormat.OpenXml.Presentation.Picture p => p.NonVisualPictureProperties?.NonVisualDrawingProperties?.Id?.Value,
                    GraphicFrame g => g.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties?.Id?.Value,
                    _ => null
                };

                if (id.HasValue && id.Value > maxId) {
                    maxId = id.Value;
                }

                PowerPointShape? shape = CreateShapeFromElement(element);
                if (shape != null) {
                    _shapes.Add(shape);
                }
            }

            _nextShapeId = maxId + 1;

            if (_slidePart.NotesSlidePart != null) {
                _notes = new PowerPointNotes(_slidePart);
            }
        }
    }
}
