using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Validation;
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
        private bool _initialSlideUntouched = false;
        private List<ValidationErrorInfo>? _validationErrorsCache;
        private bool _validationCacheInvalidated = true;
        private FileFormatVersions _lastValidatedFileFormat = FileFormatVersions.Microsoft365;

        private PowerPointPresentation(PresentationDocument document, bool isNewPresentation) {
            _document = document;
            _presentationPart = document.PresentationPart ?? document.AddPresentationPart();
            if (_presentationPart.Presentation == null) {
                // New presentation - create with required initial structure
                _presentationPart.Presentation = new Presentation();
                InitializeDefaultParts();

                // After initialization, we have one slide created by PowerPointUtils
                // Track it and mark it as untouched
                if (_presentationPart.Presentation.SlideIdList != null) {
                    foreach (SlideId slideId in _presentationPart.Presentation.SlideIdList.Elements<SlideId>()) {
                        string? relId = slideId.RelationshipId;
                        if (!string.IsNullOrEmpty(relId)) {
                            SlidePart slidePart = (SlidePart)_presentationPart.GetPartById(relId!);
                            _slides.Add(new PowerPointSlide(slidePart));
                        }
                    }
                }
                _initialSlideUntouched = isNewPresentation && _slides.Count == 1;
            } else {
                // Loading existing presentation
                LoadExistingSlides();
                _initialSlideUntouched = false; // Existing files don't have untouched initial slide
            }

            PowerPointChartAxisIdGenerator.Initialize(_presentationPart);
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
        /// <param name="filePath">Path where the presentation file will be created.</param>
        public static PowerPointPresentation Create(string filePath) {
            PresentationDocument document = PresentationDocument.Create(filePath, PresentationDocumentType.Presentation);
            PowerPointPresentation presentation = new(document, isNewPresentation: true);
            presentation._presentationPart.Presentation.Save();
            presentation._document.Save();
            return presentation;
        }

        /// <summary>
        ///     Opens an existing PowerPoint presentation.
        /// </summary>
        /// <param name="filePath">Path of the presentation file to open.</param>
        public static PowerPointPresentation Open(string filePath) {
            PresentationDocument document = PresentationDocument.Open(filePath, true);
            return new PowerPointPresentation(document, isNewPresentation: false);
        }

        /// <summary>
        ///     Adds a new slide using the specified master and layout indexes.
        /// </summary>
        /// <param name="masterIndex">Index of the slide master.</param>
        /// <param name="layoutIndex">Index of the slide layout.</param>
        public PowerPointSlide AddSlide(int masterIndex = 0, int layoutIndex = 0) {
            // If we have an untouched initial slide, return it for the user to use
            if (_initialSlideUntouched && _slides.Count == 1) {
                _initialSlideUntouched = false;
                InvalidateValidationCache();
                return _slides[0];
            }

            // Generate a unique relationship ID for the slide
            var existingRelationships = new HashSet<string>(
                _presentationPart.Parts
                    .Select(p => p.RelationshipId)
                    .Union(_presentationPart.ExternalRelationships.Select(r => r.Id))
                    .Union(_presentationPart.HyperlinkRelationships.Select(r => r.Id))
                    .Where(id => !string.IsNullOrEmpty(id))
                    .Select(id => id!)
            );

            // Also check the slide IDs
            if (_presentationPart.Presentation.SlideIdList != null) {
                foreach (SlideId existingSlideId in _presentationPart.Presentation.SlideIdList.Elements<SlideId>()) {
                    if (existingSlideId.RelationshipId is { Value: { Length: > 0 } relId }) {
                        existingRelationships.Add(relId);
                    }
                }
            }

            // Find the next available rId number
            int nextId = 1;
            string slideRelId;
            do {
                slideRelId = "rId" + nextId;
                nextId++;
            } while (existingRelationships.Contains(slideRelId));

            SlidePart slidePart = _presentationPart.AddNewPart<SlidePart>(slideRelId);
            // Create slide exactly like the working example
            slidePart.Slide = new Slide(
                new CommonSlideData(
                    new ShapeTree(
                        new NonVisualGroupShapeProperties(
                            new NonVisualDrawingProperties() { Id = 1U, Name = "" },
                            new NonVisualGroupShapeDrawingProperties(),
                            new ApplicationNonVisualDrawingProperties()),
                        new GroupShapeProperties(new A.TransformGroup()))),
                new ColorMapOverride(new A.MasterColorMapping()));

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

            // Check if this slide part already has a reference to this layout part
            string? existingRelId = null;
            foreach (var partPair in slidePart.Parts) {
                if (partPair.OpenXmlPart == layoutPart) {
                    existingRelId = partPair.RelationshipId;
                    break;
                }
            }

            if (existingRelId == null) {
                // Layout part not yet referenced, add it with a unique relationship ID
                // Check if rId1 is already in use by this slide part
                var slideRelationships = new HashSet<string>(
                    slidePart.Parts.Select(p => p.RelationshipId)
                    .Union(slidePart.ExternalRelationships.Select(r => r.Id))
                    .Union(slidePart.HyperlinkRelationships.Select(r => r.Id))
                    .Where(id => !string.IsNullOrEmpty(id))
                );

                // Find a unique relationship ID for the layout
                string layoutRelId = "rId1";
                if (slideRelationships.Contains(layoutRelId)) {
                    int layoutIdNum = 1;
                    do {
                        layoutRelId = "rId" + layoutIdNum;
                        layoutIdNum++;
                    } while (slideRelationships.Contains(layoutRelId));
                }

                slidePart.AddPart(layoutPart, layoutRelId);
            }
            // If the layout is already referenced, we don't need to add it again

            if (_presentationPart.Presentation.SlideIdList == null) {
                _presentationPart.Presentation.SlideIdList = new SlideIdList();
            }

            uint maxId = 255;
            if (_presentationPart.Presentation.SlideIdList.Elements<SlideId>().Any()) {
                maxId = _presentationPart.Presentation.SlideIdList.Elements<SlideId>().Max(s => s.Id?.Value ?? 255);
            }

            // Ensure ID is at least 256 (PowerPoint convention)
            uint newId = maxId >= 255 ? maxId + 1 : 256;
            SlideId slideId = new() { Id = newId, RelationshipId = slideRelId };
            _presentationPart.Presentation.SlideIdList.Append(slideId);
            _presentationPart.Presentation.Save();

            PowerPointSlide slide = new(slidePart);
            _slides.Add(slide);
            InvalidateValidationCache();
            return slide;
        }

        /// <summary>
        ///     Removes the slide at the specified index.
        /// </summary>
        /// <param name="index">Index of the slide to remove.</param>
        public void RemoveSlide(int index) {
            // If the initial slide is untouched, we pretend there are no slides
            if (_initialSlideUntouched) {
                throw new ArgumentOutOfRangeException(nameof(index), "No slides to remove.");
            }

            if (index < 0 || index >= _slides.Count) {
                throw new ArgumentOutOfRangeException(nameof(index));
            }

            if (_slides.Count == 1) {
                throw new InvalidOperationException("Cannot remove the last slide from the presentation.");
            }

            SlideIdList? slideIdList = _presentationPart.Presentation.SlideIdList;
            if (slideIdList == null) {
                throw new InvalidOperationException("Presentation has no slides.");
            }

            SlideId slideId = slideIdList.Elements<SlideId>().ElementAt(index);
            StringValue? relIdValue = slideId.RelationshipId;

            _slides.RemoveAt(index);
            slideId.Remove();

            if (relIdValue is { Value: { Length: > 0 } relId }) {
                OpenXmlPart part = _presentationPart.GetPartById(relId);
                _presentationPart.DeletePart(part);
            }

            _presentationPart.Presentation.Save();
            InvalidateValidationCache();
        }

        /// <summary>
        ///     Moves a slide from one index to another.
        /// </summary>
        /// <param name="fromIndex">Current index of the slide.</param>
        /// <param name="toIndex">Destination index of the slide.</param>
        public void MoveSlide(int fromIndex, int toIndex) {
            // If the initial slide is untouched, we pretend there are no slides
            if (_initialSlideUntouched) {
                throw new ArgumentOutOfRangeException(nameof(fromIndex), "No slides to move.");
            }

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
            InvalidateValidationCache();
        }

        /// <summary>
        ///     Indicates whether the presentation passes Open XML validation.
        /// </summary>
        public bool DocumentIsValid {
            get {
                return DocumentValidationErrors.Count == 0;
            }
        }

        /// <summary>
        ///     Gets the list of validation errors for the presentation.
        /// </summary>
        public List<ValidationErrorInfo> DocumentValidationErrors {
            get {
                return EnsureValidationCache();
            }
        }

        /// <summary>
        ///     Validates the presentation, refreshes the cached errors, and returns the current validation results.
        /// </summary>
        /// <param name="fileFormatVersions">File format version to validate against.</param>
        public List<ValidationErrorInfo> Validate(FileFormatVersions fileFormatVersions = FileFormatVersions.Microsoft365) {
            _validationCacheInvalidated = false;
            _validationErrorsCache = ValidateDocument(fileFormatVersions);
            _lastValidatedFileFormat = fileFormatVersions;
            return _validationErrorsCache;
        }

        private List<ValidationErrorInfo> EnsureValidationCache(FileFormatVersions fileFormatVersions = FileFormatVersions.Microsoft365) {
            if (_validationCacheInvalidated || _validationErrorsCache is null || _lastValidatedFileFormat != fileFormatVersions) {
                return Validate(fileFormatVersions);
            }

            return _validationErrorsCache;
        }

        internal void InvalidateValidationCache() {
            _validationCacheInvalidated = true;
        }

        /// <summary>
        ///     Validates the presentation using the specified file format version.
        /// </summary>
        /// <param name="fileFormatVersions">File format version to validate against.</param>
        /// <returns>List of validation errors.</returns>
        /// <example>
        /// <code>
        /// using (var presentation = PowerPointPresentation.Create("test.pptx")) {
        ///     var errors = presentation.ValidateDocument();
        ///     if (errors.Count > 0) {
        ///         // Handle validation errors
        ///     }
        /// }
        /// </code>
        /// </example>
        public List<ValidationErrorInfo> ValidateDocument(FileFormatVersions fileFormatVersions = FileFormatVersions.Microsoft365) {
            List<ValidationErrorInfo> listErrors = new List<ValidationErrorInfo>();
            OpenXmlValidator validator = new OpenXmlValidator(fileFormatVersions);
            foreach (ValidationErrorInfo error in validator.Validate(_document)) {
                listErrors.Add(error);
            }

            return listErrors;
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
            _validationCacheInvalidated = true;
        }

        /// <summary>
        ///     Creates a fluent wrapper for this presentation.
        /// </summary>
        public PowerPointFluentPresentation AsFluent() {
            return new PowerPointFluentPresentation(this);
        }

        private void InitializeDefaultParts() {
            // IMPORTANT: PowerPoint requires a very specific initialization pattern to avoid the repair dialog.
            // We must create an initial blank slide with relationship ID "rId2" and then create
            // the slide layout, slide master, and theme in a specific order.
            // DO NOT modify this initialization pattern or PowerPoint will show a repair dialog!
            PowerPointUtils.CreatePresentationParts(_presentationPart);
        }

        private void LoadExistingSlides() {
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
    }
}