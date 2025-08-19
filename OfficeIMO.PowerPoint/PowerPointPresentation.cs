using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.PowerPoint.Fluent;
using A = DocumentFormat.OpenXml.Drawing;

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
                        SlidePart slidePart = (SlidePart)_presentationPart.GetPartById(relId);
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
            PresentationDocument document =
                PresentationDocument.Create(filePath, PresentationDocumentType.Presentation);
            return new PowerPointPresentation(document);
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
            slidePart.Slide = new Slide(new CommonSlideData(new ShapeTree()));

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
            slidePart.AddPart(layoutPart);

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
            SlideMasterPart slideMasterPart = _presentationPart.AddNewPart<SlideMasterPart>();
            slideMasterPart.SlideMaster = new SlideMaster(new CommonSlideData(new ShapeTree()));

            SlideLayoutPart slideLayoutPart1 = slideMasterPart.AddNewPart<SlideLayoutPart>();
            slideLayoutPart1.SlideLayout = new SlideLayout(new CommonSlideData(new ShapeTree()));

            SlideLayoutPart slideLayoutPart2 = slideMasterPart.AddNewPart<SlideLayoutPart>();
            slideLayoutPart2.SlideLayout = new SlideLayout(new CommonSlideData(new ShapeTree()));

            slideMasterPart.SlideMaster.SlideLayoutIdList = new SlideLayoutIdList(
                new SlideLayoutId { Id = 1U, RelationshipId = slideMasterPart.GetIdOfPart(slideLayoutPart1) },
                new SlideLayoutId { Id = 2U, RelationshipId = slideMasterPart.GetIdOfPart(slideLayoutPart2) }
            );

            ThemePart themePart = slideMasterPart.AddNewPart<ThemePart>();
            themePart.Theme = new A.Theme { Name = "Office Theme", ThemeElements = new A.ThemeElements() };

            _presentationPart.Presentation.SlideMasterIdList = new SlideMasterIdList(new SlideMasterId {
                Id = 1U, RelationshipId = _presentationPart.GetIdOfPart(slideMasterPart)
            });

            _presentationPart.Presentation.SlideIdList = new SlideIdList();
            _presentationPart.Presentation.Save();
        }
    }
}