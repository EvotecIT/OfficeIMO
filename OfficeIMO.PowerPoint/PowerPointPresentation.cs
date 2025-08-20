using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml;
using OfficeIMO.PowerPoint.Fluent;
using A = DocumentFormat.OpenXml.Drawing;
using Ap = DocumentFormat.OpenXml.ExtendedProperties;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Represents a PowerPoint presentation providing basic create, open and save operations.
    /// </summary>
    public sealed class PowerPointPresentation : IDisposable {
        private readonly PresentationDocument _document;
        private readonly PresentationPart _presentationPart;
        private readonly List<PowerPointSlide> _slides = new();

        private PowerPointPresentation(PresentationDocument document) {
            _document = document;
            _presentationPart = document.PresentationPart ?? document.AddPresentationPart();
            if (_presentationPart.Presentation == null) {
                _presentationPart.Presentation = new Presentation();
                InitializeDefaultParts();
            }

            if (_presentationPart.Presentation.SlideIdList != null) {
                foreach (SlideId slideId in _presentationPart.Presentation.SlideIdList.Elements<SlideId>()) {
                    string? relId = slideId.RelationshipId;
                    if (!string.IsNullOrEmpty(relId)) {
                        SlidePart slidePart = (SlidePart)_presentationPart.GetPartById(relId!);
                        _slides.Add(new PowerPointSlide(slidePart));
                    }
                }
            }
        }

        /// <summary>
        ///     Collection of slides in the presentation.
        /// </summary>
        public IReadOnlyList<PowerPointSlide> Slides => _slides;

        /// <summary>
        ///     Gets or sets the name of the presentation theme.
        /// </summary>
        public string ThemeName {
            get {
                SlideMasterPart master = _presentationPart.SlideMasterParts.First();
                return master.ThemePart?.Theme?.Name?.Value ?? string.Empty;
            }
            set {
                SlideMasterPart master = _presentationPart.SlideMasterParts.First();
                ThemePart themePart = master.ThemePart ?? master.AddNewPart<ThemePart>();
                if (themePart.Theme == null) {
                    themePart.Theme = new A.Theme { ThemeElements = new A.ThemeElements() };
                }

                themePart.Theme.Name = value;
            }
        }

        /// <inheritdoc />
        public void Dispose() {
            _document.Dispose();
        }

        /// <summary>
        ///     Creates a new PowerPoint presentation at the specified file path.
        /// </summary>
        public static PowerPointPresentation Create(string filePath) {
            PresentationDocument document = PresentationDocument.Create(filePath, PresentationDocumentType.Presentation);
            PowerPointPresentation presentation = new(document);
            presentation._presentationPart.Presentation.Save();
            presentation._document.Save();
            return presentation;
        }

        /// <summary>
        ///     Opens an existing PowerPoint presentation.
        /// </summary>
        public static PowerPointPresentation Open(string filePath) {
            PresentationDocument document = PresentationDocument.Open(filePath, true);
            return new PowerPointPresentation(document);
        }

        /// <summary>
        ///     Adds a new slide using the specified master and layout indexes.
        /// </summary>
        /// <param name="masterIndex">Index of the slide master.</param>
        /// <param name="layoutIndex">Index of the slide layout.</param>
        public PowerPointSlide AddSlide(int masterIndex = 0, int layoutIndex = 0) {
            SlidePart slidePart = _presentationPart.AddNewPart<SlidePart>();
            ShapeTree tree = new(
                new NonVisualGroupShapeProperties(
                    new NonVisualDrawingProperties { Id = 1U, Name = string.Empty },
                    new NonVisualGroupShapeDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties()
                ),
                new GroupShapeProperties(
                    new A.TransformGroup(
                        new A.Offset { X = 0L, Y = 0L },
                        new A.Extents { Cx = 0L, Cy = 0L },
                        new A.ChildOffset { X = 0L, Y = 0L },
                        new A.ChildExtents { Cx = 0L, Cy = 0L }
                    )
                )
            );
            slidePart.Slide = new Slide(new CommonSlideData(tree));

            SlideMasterPart[] masters = _presentationPart.SlideMasterParts.ToArray();
            if (masterIndex < 0 || masterIndex >= masters.Length) {
                throw new ArgumentOutOfRangeException(nameof(masterIndex));
            }

            SlideMasterPart masterPart = masters[masterIndex];

            SlideLayoutPart[] layouts = masterPart.SlideLayoutParts.ToArray();
            if (layoutIndex < 0 || layoutIndex >= layouts.Length) {
                throw new ArgumentOutOfRangeException(nameof(layoutIndex));
            }

            SlideLayoutPart layoutPart = layouts[layoutIndex];
            _ = slidePart.AddPart(layoutPart);

            if (_presentationPart.Presentation.SlideIdList == null) {
                _presentationPart.Presentation.SlideIdList = new SlideIdList();
            }

            uint maxId = 255;
            if (_presentationPart.Presentation.SlideIdList.Elements<SlideId>().Any()) {
                maxId = _presentationPart.Presentation.SlideIdList.Elements<SlideId>().Max(s => s.Id?.Value ?? 0);
            }

            SlideId slideId = new() { Id = maxId + 1, RelationshipId = _presentationPart.GetIdOfPart(slidePart) };
            _presentationPart.Presentation.SlideIdList.Append(slideId);
            _presentationPart.Presentation.Save();

            PowerPointSlide slide = new(slidePart);
            _slides.Add(slide);
            return slide;
        }

        /// <summary>
        ///     Removes the slide at the specified index.
        /// </summary>
        /// <param name="index">Index of the slide to remove.</param>
        public void RemoveSlide(int index) {
            if (index < 0 || index >= _slides.Count) {
                throw new ArgumentOutOfRangeException(nameof(index));
            }

            SlideIdList? slideIdList = _presentationPart.Presentation.SlideIdList;
            if (slideIdList == null) {
                throw new InvalidOperationException("Presentation has no slides.");
            }

            SlideId slideId = slideIdList.Elements<SlideId>().ElementAt(index);
            string? relId = slideId.RelationshipId;

            _slides.RemoveAt(index);
            slideId.Remove();

            if (!string.IsNullOrEmpty(relId)) {
                OpenXmlPart part = _presentationPart.GetPartById(relId);
                _presentationPart.DeletePart(part);
            }

            _presentationPart.Presentation.Save();
        }

        /// <summary>
        ///     Moves a slide from one index to another.
        /// </summary>
        /// <param name="fromIndex">Current index of the slide.</param>
        /// <param name="toIndex">Destination index of the slide.</param>
        public void MoveSlide(int fromIndex, int toIndex) {
            if (fromIndex < 0 || fromIndex >= _slides.Count) {
                throw new ArgumentOutOfRangeException(nameof(fromIndex));
            }

            if (toIndex < 0 || toIndex >= _slides.Count) {
                throw new ArgumentOutOfRangeException(nameof(toIndex));
            }

            if (fromIndex == toIndex) {
                return;
            }

            SlideIdList? slideIdList = _presentationPart.Presentation.SlideIdList;
            if (slideIdList == null) {
                throw new InvalidOperationException("Presentation has no slides.");
            }

            PowerPointSlide slide = _slides[fromIndex];
            _slides.RemoveAt(fromIndex);
            _slides.Insert(toIndex, slide);

            List<SlideId> ids = slideIdList.Elements<SlideId>().ToList();
            SlideId movingId = ids[fromIndex];
            ids.RemoveAt(fromIndex);
            ids.Insert(toIndex, movingId);

            slideIdList.RemoveAllChildren();
            foreach (SlideId id in ids) {
                slideIdList.Append(id);
            }

            _presentationPart.Presentation.Save();
        }

        /// <summary>
        ///     Saves all pending changes to the underlying package.
        /// </summary>
        public void Save() {
            foreach (PowerPointSlide slide in _slides) {
                slide.Save();
            }

            _presentationPart.Presentation.Save();
            _document.Save();
        }

        /// <summary>
        ///     Creates a fluent wrapper for this presentation.
        /// </summary>
        public PowerPointFluentPresentation AsFluent() {
            return new PowerPointFluentPresentation(this);
        }

        private void InitializeDefaultParts() {
            static ShapeTree CreateShapeTree() {
                return new ShapeTree(
                    new NonVisualGroupShapeProperties(
                        new NonVisualDrawingProperties { Id = 1U, Name = string.Empty },
                        new NonVisualGroupShapeDrawingProperties(),
                        new ApplicationNonVisualDrawingProperties()
                    ),
                    new GroupShapeProperties(
                        new A.TransformGroup(
                            new A.Offset { X = 0L, Y = 0L },
                            new A.Extents { Cx = 0L, Cy = 0L },
                            new A.ChildOffset { X = 0L, Y = 0L },
                            new A.ChildExtents { Cx = 0L, Cy = 0L }
                        )
                    )
                );
            }

            SlideMasterPart slideMasterPart = _presentationPart.AddNewPart<SlideMasterPart>();
            SlideMaster slideMaster = new(
                new CommonSlideData(CreateShapeTree()),
                new ColorMap {
                    Background1 = A.ColorSchemeIndexValues.Light1,
                    Text1 = A.ColorSchemeIndexValues.Dark1,
                    Background2 = A.ColorSchemeIndexValues.Light2,
                    Text2 = A.ColorSchemeIndexValues.Dark2,
                    Accent1 = A.ColorSchemeIndexValues.Accent1,
                    Accent2 = A.ColorSchemeIndexValues.Accent2,
                    Accent3 = A.ColorSchemeIndexValues.Accent3,
                    Accent4 = A.ColorSchemeIndexValues.Accent4,
                    Accent5 = A.ColorSchemeIndexValues.Accent5,
                    Accent6 = A.ColorSchemeIndexValues.Accent6,
                    Hyperlink = A.ColorSchemeIndexValues.Hyperlink,
                    FollowedHyperlink = A.ColorSchemeIndexValues.FollowedHyperlink
                },
                new SlideLayoutIdList(),
                new TextStyles(new TitleStyle(), new BodyStyle(), new OtherStyle())
            );
            slideMasterPart.SlideMaster = slideMaster;

            SlideLayoutPart layoutPart0 = slideMasterPart.AddNewPart<SlideLayoutPart>();
            layoutPart0.SlideLayout = new SlideLayout(
                new CommonSlideData(CreateShapeTree()),
                new ColorMapOverride(new A.MasterColorMapping())
            );

            SlideLayoutPart layoutPart1 = slideMasterPart.AddNewPart<SlideLayoutPart>();
            layoutPart1.SlideLayout = new SlideLayout(
                new CommonSlideData(CreateShapeTree()),
                new ColorMapOverride(new A.MasterColorMapping())
            );

            slideMaster.SlideLayoutIdList = new SlideLayoutIdList(
                new SlideLayoutId { Id = 1U, RelationshipId = slideMasterPart.GetIdOfPart(layoutPart0) },
                new SlideLayoutId { Id = 2U, RelationshipId = slideMasterPart.GetIdOfPart(layoutPart1) }
            );

            // theme part is stored under ppt/theme and referenced from both presentation and slide master
            ThemePart themePart = _presentationPart.AddNewPart<ThemePart>();
            themePart.Theme = new A.Theme { Name = "Office Theme", ThemeElements = new A.ThemeElements() };
            slideMasterPart.AddPart(themePart);

            _presentationPart.Presentation.SlideMasterIdList = new SlideMasterIdList(
                new SlideMasterId { Id = 1U, RelationshipId = _presentationPart.GetIdOfPart(slideMasterPart) }
            );

            NotesMasterPart notesMasterPart = _presentationPart.AddNewPart<NotesMasterPart>();
            notesMasterPart.NotesMaster = new NotesMaster(
                new CommonSlideData(CreateShapeTree()),
                new ColorMapOverride(new A.MasterColorMapping()),
                new NotesStyle()
            );

            NotesMasterId notesMasterId = new();
            notesMasterId.SetAttribute(new OpenXmlAttribute("r", "id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships", _presentationPart.GetIdOfPart(notesMasterPart)));
            _presentationPart.Presentation.NotesMasterIdList = new NotesMasterIdList(notesMasterId);

            _presentationPart.Presentation.SlideSize = new SlideSize {
                Cx = 9144000,
                Cy = 6858000,
                Type = SlideSizeValues.Screen4x3
            };

            _presentationPart.Presentation.NotesSize = new NotesSize {
                Cx = 6858000,
                Cy = 9144000
            };

            _presentationPart.Presentation.DefaultTextStyle = new DefaultTextStyle();
            _presentationPart.Presentation.SlideIdList = new SlideIdList();

            // additional parts required by PowerPoint
            _document.PackageProperties.Creator = string.Empty;
            _document.PackageProperties.Created = DateTime.UtcNow;
            _document.PackageProperties.Modified = DateTime.UtcNow;

            ExtendedFilePropertiesPart appPart = _document.AddExtendedFilePropertiesPart();
            appPart.Properties = new Ap.Properties(new Ap.Application { Text = "Microsoft Office PowerPoint" });

            PresentationPropertiesPart presPropsPart = _presentationPart.AddNewPart<PresentationPropertiesPart>();
            presPropsPart.PresentationProperties = new PresentationProperties();

            ViewPropertiesPart viewPropsPart = _presentationPart.AddNewPart<ViewPropertiesPart>();
            viewPropsPart.ViewProperties = new ViewProperties();

            TableStylesPart tableStylesPart = _presentationPart.AddNewPart<TableStylesPart>();
            tableStylesPart.TableStyleList = new A.TableStyleList { Default = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}" };

            _presentationPart.Presentation.Save();
        }
    }
}