using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
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

        internal PowerPointSlide(SlidePart slidePart) {
            _slidePart = slidePart;
            LoadExistingShapes();
        }

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
                }

                _slidePart.Slide.Transition = transition;
            }
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
                    case Shape s when s.TextBody != null:
                        _shapes.Add(new PowerPointTextBox(s));
                        break;
                    case DocumentFormat.OpenXml.Presentation.Picture p:
                        _shapes.Add(new PowerPointPicture(p, _slidePart));
                        break;
                    case GraphicFrame g when g.Graphic?.GraphicData?.GetFirstChild<A.Table>() != null:
                        _shapes.Add(new PowerPointTable(g));
                        break;
                    case GraphicFrame g when g.Graphic?.GraphicData?.GetFirstChild<C.ChartReference>() != null:
                        _shapes.Add(new PowerPointChart(g));
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
