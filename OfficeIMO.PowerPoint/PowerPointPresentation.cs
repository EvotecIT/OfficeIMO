using System.IO;
using System.Reflection;
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
        private PresentationDocument? _document;
        private PresentationPart _presentationPart;
        private readonly List<PowerPointSlide> _slides = new();
        private readonly string _filePath;
        private PowerPointSlideSize? _slideSize;
        private bool _initialSlideUntouched = false;
        private bool _disposed = false;
        private static readonly MethodInfo AddNewPartWithContentTypeMethod =
            typeof(OpenXmlPartContainer)
                .GetMethods()
                .Single(m => m.Name == "AddNewPart" &&
                             m.IsGenericMethodDefinition &&
                             m.GetParameters().Length == 2);
        private static readonly MethodInfo AddPartWithIdMethod =
            typeof(OpenXmlPartContainer)
                .GetMethods()
                .Single(m => m.Name == "AddPart" &&
                             m.IsGenericMethodDefinition &&
                             m.GetParameters().Length == 2);

        private PowerPointPresentation(PresentationDocument document, string filePath, bool isNewPresentation) {
            _document = document;
            _filePath = filePath;
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
        public IReadOnlyList<PowerPointSlide> Slides {
            get {
                ThrowIfDisposed();
                return _slides;
            }
        }

        /// <summary>
        ///     Slide size information for the presentation.
        /// </summary>
        public PowerPointSlideSize SlideSize {
            get {
                ThrowIfDisposed();
                return _slideSize ??= new PowerPointSlideSize(_presentationPart);
            }
        }

        /// <summary>
        ///     Gets or sets the name of the presentation theme.
        /// </summary>
        public string ThemeName {
            get {
                ThrowIfDisposed();
                SlideMasterPart master = _presentationPart.SlideMasterParts.First();
                return master.ThemePart?.Theme?.Name?.Value ?? string.Empty;
            }
            set {
                ThrowIfDisposed();
                SlideMasterPart master = _presentationPart.SlideMasterParts.First();
                ThemePart themePart = master.ThemePart ?? master.AddNewPart<ThemePart>();
                if (themePart.Theme == null) {
                    themePart.Theme = new A.Theme { ThemeElements = new A.ThemeElements() };
                }

                themePart.Theme.Name = value;
            }
        }

        /// <summary>
        ///     Gets the list of table styles available in the presentation.    
        /// </summary>
        public IReadOnlyList<PowerPointTableStyleInfo> TableStyles {
            get {
                ThrowIfDisposed();
                TableStylesPart? stylesPart = _presentationPart.TableStylesPart;
                if (stylesPart?.TableStyleList == null) {
                    return Array.Empty<PowerPointTableStyleInfo>();
                }

                List<PowerPointTableStyleInfo> styles = new();
                foreach (A.TableStyle style in stylesPart.TableStyleList.Elements<A.TableStyle>()) {
                    string styleId = style.StyleId?.Value ?? string.Empty;
                    if (string.IsNullOrWhiteSpace(styleId)) {
                        continue;
                    }

                    string name = style.StyleName?.Value ?? string.Empty;
                    styles.Add(new PowerPointTableStyleInfo(styleId, name));
                }

                return styles;
            }
        }

        /// <summary>
        ///     Replaces text across all slides.
        /// </summary>
        public int ReplaceText(string oldValue, string newValue, bool includeTables = true, bool includeNotes = false) {
            ThrowIfDisposed();
            if (oldValue == null) {
                throw new ArgumentNullException(nameof(oldValue));
            }
            if (oldValue.Length == 0) {
                throw new ArgumentException("Old value cannot be empty.", nameof(oldValue));
            }

            int count = 0;
            foreach (PowerPointSlide slide in Slides) {
                count += slide.ReplaceText(oldValue, newValue, includeTables, includeNotes);
            }
            return count;
        }

        /// <inheritdoc />
        public void Dispose() {
            if (_disposed) {
                return;
            }

            try {
                _document?.Dispose();
            } finally {
                _document = null;
                _disposed = true;
            }
        }

        /// <summary>
        ///     Creates a new PowerPoint presentation at the specified file path.
        /// </summary>
        /// <param name="filePath">Path where the presentation file will be created.</param>
        public static PowerPointPresentation Create(string filePath) {
            PresentationDocument document = PresentationDocument.Create(filePath, PresentationDocumentType.Presentation);
            PowerPointPresentation presentation = new(document, filePath, isNewPresentation: true);
            presentation._presentationPart.Presentation.Save();
            presentation._document?.Save();
            return presentation;
        }

        /// <summary>
        ///     Opens an existing PowerPoint presentation.
        /// </summary>
        /// <param name="filePath">Path of the presentation file to open.</param>
        public static PowerPointPresentation Open(string filePath) {
            PresentationDocument document = PresentationDocument.Open(filePath, true);
            return new PowerPointPresentation(document, filePath, isNewPresentation: false);
        }

        /// <summary>
        ///     Adds a new slide using the specified master and layout indexes.
        /// </summary>
        /// <param name="masterIndex">Index of the slide master.</param>
        /// <param name="layoutIndex">Index of the slide layout.</param>
        public PowerPointSlide AddSlide(int masterIndex = 0, int layoutIndex = 0) {
            ThrowIfDisposed();
            // If we have an untouched initial slide, return it for the user to use
            if (_initialSlideUntouched && _slides.Count == 1) {
                _initialSlideUntouched = false;
                if (masterIndex != 0 || layoutIndex != 0) {
                    _slides[0].SetLayout(masterIndex, layoutIndex);
                }
                return _slides[0];
            }

            string slideRelId = GetNextSlideRelationshipId();
            SlidePart slidePart = _presentationPart.AddNewPart<SlidePart>(slideRelId);
            // Create slide exactly like the working example
            slidePart.Slide = new Slide(
                new CommonSlideData(
                    new ShapeTree(
                        new NonVisualGroupShapeProperties(
                            new NonVisualDrawingProperties() { Id = 1U, Name = "" },
                            new NonVisualGroupShapeDrawingProperties(),
                            new ApplicationNonVisualDrawingProperties()),
                        PowerPointUtils.CreateDefaultGroupShapeProperties())),
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

            uint newId = GetNextSlideId();
            SlideId slideId = new() { Id = newId, RelationshipId = slideRelId };
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
        }

        /// <summary>
        ///     Duplicates a slide and inserts it into the presentation.
        /// </summary>
        /// <param name="index">Index of the slide to duplicate.</param>
        /// <param name="insertAt">Index where the duplicate should be inserted. Defaults to index + 1.</param>
        public PowerPointSlide DuplicateSlide(int index, int? insertAt = null) {
            ThrowIfDisposed();
            if (_initialSlideUntouched && _slides.Count == 1) {
                _initialSlideUntouched = false;
            }

            if (index < 0 || index >= _slides.Count) {
                throw new ArgumentOutOfRangeException(nameof(index));
            }

            int targetIndex = insertAt ?? index + 1;
            if (targetIndex < 0 || targetIndex > _slides.Count) {
                throw new ArgumentOutOfRangeException(nameof(insertAt));
            }

            PowerPointSlide sourceSlide = _slides[index];
            SlidePart sourcePart = sourceSlide.SlidePart;

            string slideRelId = GetNextSlideRelationshipId();
            SlidePart slidePart = _presentationPart.AddNewPart<SlidePart>(slideRelId);
            slidePart.Slide = (Slide)sourcePart.Slide.CloneNode(true);

            CloneSlidePartRelationships(sourcePart, slidePart);

            SlideIdList slideIdList = _presentationPart.Presentation.SlideIdList ??= new SlideIdList();
            SlideId slideId = new() { Id = GetNextSlideId(), RelationshipId = slideRelId };
            InsertSlideId(slideIdList, slideId, targetIndex);

            PowerPointSlide duplicate = new(slidePart);
            _slides.Insert(targetIndex, duplicate);
            return duplicate;
        }

        /// <summary>
        ///     Indicates whether the presentation passes Open XML validation.
        /// </summary>
        public bool DocumentIsValid {
            get {
                if (DocumentValidationErrors.Count > 0) {
                    return false;
                }

                return true;
            }
        }

        /// <summary>
        ///     Gets the list of validation errors for the presentation.
        /// </summary>
        public List<ValidationErrorInfo> DocumentValidationErrors {
            get {
                return ValidateDocument();
            }
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
            ThrowIfDisposed();
            List<ValidationErrorInfo> listErrors = new List<ValidationErrorInfo>();
            OpenXmlValidator validator = new OpenXmlValidator(fileFormatVersions);
            foreach (ValidationErrorInfo error in validator.Validate(_document!)) {
                listErrors.Add(error);
            }

            return listErrors;
        }

        /// <summary>
        ///     Saves all pending changes to the underlying package.
        /// </summary>
        public void Save() {
            ThrowIfDisposed();
            foreach (PowerPointSlide slide in _slides) {
                slide.Save();
            }

            _presentationPart.Presentation.Save();
            _document!.Save();
        }

        /// <summary>
        ///     Creates a fluent wrapper for this presentation.
        /// </summary>
        public PowerPointFluentPresentation AsFluent() {
            ThrowIfDisposed();
            return new PowerPointFluentPresentation(this);
        }

        private void InitializeDefaultParts() {
            // IMPORTANT: PowerPoint requires a very specific initialization pattern to avoid the repair dialog.
            // We must create an initial blank slide with relationship ID "rId2" and then create
            // the slide layout, slide master, and theme in a specific order.
            // DO NOT modify this initialization pattern or PowerPoint will show a repair dialog!
            PowerPointUtils.CreatePresentationParts(_document!, _presentationPart);
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

        private string GetNextSlideRelationshipId() {
            var existingRelationships = new HashSet<string>(
                _presentationPart.Parts
                    .Select(p => p.RelationshipId)
                    .Union(_presentationPart.ExternalRelationships.Select(r => r.Id))
                    .Union(_presentationPart.HyperlinkRelationships.Select(r => r.Id))
                    .Where(id => !string.IsNullOrEmpty(id))
                    .Select(id => id!)
            );

            if (_presentationPart.Presentation.SlideIdList != null) {
                foreach (SlideId existingSlideId in _presentationPart.Presentation.SlideIdList.Elements<SlideId>()) {
                    if (existingSlideId.RelationshipId is { Value: { Length: > 0 } relId }) {
                        existingRelationships.Add(relId);
                    }
                }
            }

            int nextId = 1;
            string slideRelId;
            do {
                slideRelId = "rId" + nextId;
                nextId++;
            } while (existingRelationships.Contains(slideRelId));

            return slideRelId;
        }

        private uint GetNextSlideId() {
            uint maxId = 255;
            SlideIdList? slideIdList = _presentationPart.Presentation.SlideIdList;
            if (slideIdList != null && slideIdList.Elements<SlideId>().Any()) {
                maxId = slideIdList.Elements<SlideId>().Max(s => s.Id?.Value ?? 255);
            }

            return maxId >= 255 ? maxId + 1 : 256;
        }

        private static void InsertSlideId(SlideIdList slideIdList, SlideId slideId, int index) {
            List<SlideId> ids = slideIdList.Elements<SlideId>().ToList();
            if (index < 0 || index > ids.Count) {
                throw new ArgumentOutOfRangeException(nameof(index));
            }

            ids.Insert(index, slideId);
            slideIdList.RemoveAllChildren();
            foreach (SlideId id in ids) {
                slideIdList.Append(id);
            }
        }

        private static void CloneSlidePartRelationships(SlidePart source, SlidePart target) {
            foreach (var partPair in source.Parts) {
                ClonePartRecursive(partPair.OpenXmlPart, target, partPair.RelationshipId);
            }

            CloneReferenceRelationships(source, target);
        }

        private static void ClonePartRecursive(OpenXmlPart sourcePart, OpenXmlPartContainer targetContainer, string relationshipId) {
            if (ShouldSharePart(sourcePart)) {
                AddExistingPart(targetContainer, sourcePart, relationshipId);
                return;
            }

            OpenXmlPart newPart = sourcePart is ExtendedPart extendedPart
                ? targetContainer.AddExtendedPart(extendedPart.RelationshipType, extendedPart.ContentType, relationshipId)
                : AddNewPartWithContentType(targetContainer, sourcePart, relationshipId);

            CopyPartData(sourcePart, newPart);
            CloneReferenceRelationships(sourcePart, newPart);

            foreach (var childPair in sourcePart.Parts) {
                ClonePartRecursive(childPair.OpenXmlPart, newPart, childPair.RelationshipId);
            }
        }

        private static OpenXmlPart AddNewPartWithContentType(OpenXmlPartContainer container, OpenXmlPart sourcePart, string relationshipId) {
            MethodInfo method = AddNewPartWithContentTypeMethod.MakeGenericMethod(sourcePart.GetType());
            return (OpenXmlPart)method.Invoke(container, new object[] { sourcePart.ContentType, relationshipId })!;
        }

        private static OpenXmlPart AddExistingPart(OpenXmlPartContainer container, OpenXmlPart sourcePart, string relationshipId) {
            MethodInfo method = AddPartWithIdMethod.MakeGenericMethod(sourcePart.GetType());
            return (OpenXmlPart)method.Invoke(container, new object[] { sourcePart, relationshipId })!;
        }

        private static void CopyPartData(OpenXmlPart sourcePart, OpenXmlPart targetPart) {
            using Stream sourceStream = sourcePart.GetStream(FileMode.Open, FileAccess.Read);
            using Stream targetStream = targetPart.GetStream(FileMode.Create, FileAccess.Write);
            sourceStream.CopyTo(targetStream);
        }

        private static void CloneReferenceRelationships(OpenXmlPartContainer source, OpenXmlPartContainer target) {
            foreach (ExternalRelationship rel in source.ExternalRelationships) {
                target.AddExternalRelationship(rel.RelationshipType, rel.Uri, rel.Id);
            }

            foreach (HyperlinkRelationship rel in source.HyperlinkRelationships) {
                target.AddHyperlinkRelationship(rel.Uri, rel.IsExternal, rel.Id);
            }

            CloneDataPartReferenceRelationships(source, target);
        }

        private static void CloneDataPartReferenceRelationships(OpenXmlPartContainer source, OpenXmlPartContainer target) {
            if (target is not SlidePart slidePart) {
                return;
            }

            foreach (DataPartReferenceRelationship rel in source.DataPartReferenceRelationships) {
                if (rel.DataPart is not MediaDataPart mediaPart) {
                    continue;
                }

                if (rel is AudioReferenceRelationship) {
                    slidePart.AddAudioReferenceRelationship(mediaPart, rel.Id);
                } else if (rel is VideoReferenceRelationship) {
                    slidePart.AddVideoReferenceRelationship(mediaPart, rel.Id);
                } else {
                    slidePart.AddMediaReferenceRelationship(mediaPart, rel.Id);
                }
            }
        }

        private static bool ShouldSharePart(OpenXmlPart part) {
            return part is SlideLayoutPart || part is NotesMasterPart;
        }

        private void ThrowIfDisposed() {
            if (_disposed || _document == null) {
                throw new ObjectDisposedException(nameof(PowerPointPresentation));
            }
        }
    }
}
