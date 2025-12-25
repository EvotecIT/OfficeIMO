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
        ///     Notes associated with the slide.
        /// </summary>
        public PowerPointNotes Notes => _notes ??= new PowerPointNotes(_slidePart);

        /// <summary>
        ///     Gets or sets the slide background color in hex format (e.g. "FF0000").
        /// </summary>
        public string? BackgroundColor {
            get {
                CommonSlideData? common = _slidePart.Slide.CommonSlideData;
                Background? bg = common?.Background;
                A.SolidFill? solid = bg?.BackgroundProperties?.GetFirstChild<A.SolidFill>();
                return solid?.RgbColorModelHex?.Val;
            }
            set {
                CommonSlideData common = _slidePart.Slide.CommonSlideData ??= new CommonSlideData(new ShapeTree());
                if (value == null) {
                    common.Background = null;
                    return;
                }

                Background bg = common.Background ?? new Background();
                BackgroundProperties props = bg.BackgroundProperties ?? new BackgroundProperties();
                props.RemoveAllChildren<A.SolidFill>();
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

            CommonSlideData common = _slidePart.Slide.CommonSlideData ??= new CommonSlideData(new ShapeTree());
            Background background = common.Background ?? new Background();
            BackgroundProperties props = background.BackgroundProperties ?? new BackgroundProperties();

            props.RemoveAllChildren<A.SolidFill>();
            props.RemoveAllChildren<A.BlipFill>();

            props.Append(new A.BlipFill(
                new A.Blip { Embed = relationshipId },
                new A.Stretch(new A.FillRectangle())
            ));

            background.BackgroundProperties = props;
            common.Background = background;
        }

        /// <summary>
        ///     Clears any background image from the slide.
        /// </summary>
        public void ClearBackgroundImage() {
            CommonSlideData? common = _slidePart.Slide.CommonSlideData;
            if (common?.Background?.BackgroundProperties == null) {
                return;
            }

            common.Background.BackgroundProperties.RemoveAllChildren<A.BlipFill>();
            if (!common.Background.BackgroundProperties.HasChildren) {
                common.Background = null;
            }
        }

        /// <summary>
        ///     Transition applied when moving to this slide.
        /// </summary>
        public SlideTransition Transition {
            get {
                Transition? t = _slidePart.Slide.Transition;
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
                if (value == SlideTransition.None) {
                    _slidePart.Slide.Transition = null;
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
                    case SlideTransition.Morph:
                        transition.AddNamespaceDeclaration("p159", P159Namespace);
                        transition.Append(CreateMorphTransition());
                        break;
                }

                _slidePart.Slide.Transition = transition;
            }
        }

        /// <summary>
        ///     Gets or sets whether the slide is hidden in slide show mode.
        /// </summary>
        public bool Hidden {
            get {
                SlideId slideId = GetSlideId();
                OpenXmlAttribute? showAttribute = slideId.GetAttributes()
                    .FirstOrDefault(attribute =>
                        attribute.LocalName == "show" && string.IsNullOrEmpty(attribute.NamespaceUri));
                if (showAttribute == null || string.IsNullOrEmpty(showAttribute.Value.Value)) {
                    return false;
                }

                return string.Equals(showAttribute.Value.Value, "0", StringComparison.Ordinal) ||
                       string.Equals(showAttribute.Value.Value, "false", StringComparison.OrdinalIgnoreCase);
            }
            set {
                SlideId slideId = GetSlideId();
                if (value) {
                    slideId.SetAttribute(new OpenXmlAttribute("show", string.Empty, "0"));
                } else {
                    slideId.RemoveAttribute("show", string.Empty);
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
                element.LocalName == "morph" && element.NamespaceUri == P159Namespace);
        }

        private static OpenXmlUnknownElement CreateMorphTransition() {
            OpenXmlUnknownElement morph = new OpenXmlUnknownElement("p159", "morph", P159Namespace);
            morph.SetAttribute(new OpenXmlAttribute("option", string.Empty, "byObject"));
            return morph;
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
            _slidePart.Slide.Save();
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

        private void LoadExistingShapes() {
            ShapeTree? tree = _slidePart.Slide.CommonSlideData?.ShapeTree;
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

                switch (element) {
                    case Shape s:
                        if (s.TextBody != null) {
                            _shapes.Add(new PowerPointTextBox(s));
                        } else {
                            _shapes.Add(new PowerPointAutoShape(s));
                        }
                        break;
                    case DocumentFormat.OpenXml.Presentation.Picture p:
                        _shapes.Add(new PowerPointPicture(p, _slidePart));
                        break;
                    case GraphicFrame g when g.Graphic?.GraphicData?.GetFirstChild<A.Table>() != null:
                        _shapes.Add(new PowerPointTable(g));
                        break;
                    case GraphicFrame g when g.Graphic?.GraphicData?.GetFirstChild<C.ChartReference>() != null:
                        _shapes.Add(new PowerPointChart(g, _slidePart));
                        break;
                }
            }

            _nextShapeId = maxId + 1;

            if (_slidePart.NotesSlidePart != null) {
                _notes = new PowerPointNotes(_slidePart);
            }
        }
    }
}
