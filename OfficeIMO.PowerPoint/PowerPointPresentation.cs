using System.IO;
using System.Reflection;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.PowerPoint.Fluent;
using A = DocumentFormat.OpenXml.Drawing;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;

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
        private const string P14Namespace = "http://schemas.microsoft.com/office/powerpoint/2010/main";
        private const string SectionListUri = "{521415D9-36F7-43E2-AB2F-B90AF26B5E84}";
        private const string DefaultSectionName = "Section 1";
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
        ///     Returns the layouts available for a slide master.
        /// </summary>
        public IReadOnlyList<PowerPointSlideLayoutInfo> GetSlideLayouts(int masterIndex = 0) {
            ThrowIfDisposed();
            SlideMasterPart masterPart = GetSlideMasterPart(masterIndex);
            SlideLayoutPart[] layouts = masterPart.SlideLayoutParts.ToArray();
            var results = new List<PowerPointSlideLayoutInfo>(layouts.Length);

            for (int i = 0; i < layouts.Length; i++) {
                SlideLayoutPart layoutPart = layouts[i];
                SlideLayout? layout = layoutPart.SlideLayout;
                string name = layout?.CommonSlideData?.Name?.Value ?? string.Empty;
                SlideLayoutValues? type = layout?.Type?.Value;
                string? relationshipId = masterPart.GetIdOfPart(layoutPart);
                results.Add(new PowerPointSlideLayoutInfo(masterIndex, i, name, type, relationshipId));
            }

            return results;
        }

        /// <summary>
        ///     Returns the sections defined in the presentation.
        /// </summary>
        public IReadOnlyList<PowerPointSectionInfo> GetSections() {
            ThrowIfDisposed();
            P14.SectionList? sectionList = GetSectionList(create: false);
            if (sectionList == null) {
                return Array.Empty<PowerPointSectionInfo>();
            }

            List<SlideId> slideIds = _presentationPart.Presentation?.SlideIdList?
                .Elements<SlideId>()
                .ToList() ?? new List<SlideId>();
            Dictionary<uint, int> slideIndexMap = BuildSlideIndexMap(slideIds);

            List<PowerPointSectionInfo> sections = new();
            foreach (P14.Section section in sectionList.Elements<P14.Section>()) {
                List<int> indices = new();
                P14.SectionSlideIdList? list = section.SectionSlideIdList;
                if (list != null) {
                    foreach (P14.SectionSlideIdListEntry entry in list.Elements<P14.SectionSlideIdListEntry>()) {
                        uint? slideId = entry.Id?.Value;
                        if (slideId != null && slideIndexMap.TryGetValue(slideId.Value, out int index)) {
                            indices.Add(index);
                        }
                    }
                }

                indices.Sort();
                string name = section.Name?.Value ?? string.Empty;
                string id = section.Id?.Value ?? string.Empty;
                sections.Add(new PowerPointSectionInfo(name, id, indices));
            }

            return sections;
        }

        /// <summary>
        ///     Adds a new section starting at the specified slide index.
        /// </summary>
        public PowerPointSectionInfo AddSection(string name, int startSlideIndex) {
            ThrowIfDisposed();
            if (string.IsNullOrWhiteSpace(name)) {
                throw new ArgumentException("Section name cannot be null or empty.", nameof(name));
            }

            SlideIdList? slideIdList = _presentationPart.Presentation?.SlideIdList;
            if (slideIdList == null) {
                throw new InvalidOperationException("Presentation has no slides.");
            }

            List<SlideId> slideIds = slideIdList.Elements<SlideId>().ToList();
            if (slideIds.Count == 0) {
                throw new InvalidOperationException("Presentation has no slides.");
            }
            if (startSlideIndex < 0 || startSlideIndex >= slideIds.Count) {
                throw new ArgumentOutOfRangeException(nameof(startSlideIndex));
            }

            P14.SectionList sectionList = EnsureSectionList(slideIds);
            uint slideIdValue = slideIds[startSlideIndex].Id?.Value ?? throw new InvalidOperationException("Slide ID is missing.");

            P14.Section? containing = FindSectionBySlideId(sectionList, slideIdValue);
            if (containing == null) {
                P14.Section fallback = sectionList.Elements<P14.Section>().Last();
                EnsureSectionSlideIdList(fallback)
                    .Append(new P14.SectionSlideIdListEntry { Id = slideIdValue });
                return BuildSectionInfo(fallback, slideIds);
            }

            P14.SectionSlideIdList list = EnsureSectionSlideIdList(containing);
            List<P14.SectionSlideIdListEntry> entries = list.Elements<P14.SectionSlideIdListEntry>().ToList();
            int entryIndex = entries.FindIndex(entry => entry.Id?.Value == slideIdValue);
            if (entryIndex <= 0) {
                containing.Name = name;
                return BuildSectionInfo(containing, slideIds);
            }

            List<uint> movedIds = entries
                .Skip(entryIndex)
                .Select(entry => entry.Id?.Value)
                .Where(id => id != null)
                .Select(id => id!.Value)
                .ToList();
            foreach (P14.SectionSlideIdListEntry entry in entries.Skip(entryIndex)) {
                entry.Remove();
            }

            P14.Section newSection = CreateSection(name, movedIds);
            sectionList.InsertAfter(newSection, containing);
            return BuildSectionInfo(newSection, slideIds);
        }

        /// <summary>
        ///     Renames the first section matching the provided name.
        /// </summary>
        public bool RenameSection(string name, string newName, bool ignoreCase = true) {
            ThrowIfDisposed();
            if (name == null) {
                throw new ArgumentNullException(nameof(name));
            }
            if (string.IsNullOrWhiteSpace(newName)) {
                throw new ArgumentException("Section name cannot be null or empty.", nameof(newName));
            }

            P14.SectionList? sectionList = GetSectionList(create: false);
            if (sectionList == null) {
                return false;
            }

            StringComparison comparison = ignoreCase ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;
            foreach (P14.Section section in sectionList.Elements<P14.Section>()) {
                string currentName = section.Name?.Value ?? string.Empty;
                if (string.Equals(currentName, name, comparison)) {
                    section.Name = newName;
                    return true;
                }
            }

            return false;
        }

        /// <summary>
        ///     Finds a layout index by layout type.
        /// </summary>
        public int GetLayoutIndex(SlideLayoutValues layoutType, int masterIndex = 0) {
            ThrowIfDisposed();
            SlideMasterPart masterPart = GetSlideMasterPart(masterIndex);
            SlideLayoutPart[] layouts = masterPart.SlideLayoutParts.ToArray();
            for (int i = 0; i < layouts.Length; i++) {
                SlideLayoutValues? type = layouts[i].SlideLayout?.Type?.Value;
                if (type == layoutType) {
                    return i;
                }
            }

            throw new InvalidOperationException($"Layout type '{layoutType}' not found for master {masterIndex}.");
        }

        /// <summary>
        ///     Finds a layout index by layout name.
        /// </summary>
        public int GetLayoutIndex(string layoutName, int masterIndex = 0, bool ignoreCase = true) {
            ThrowIfDisposed();
            if (layoutName == null) {
                throw new ArgumentNullException(nameof(layoutName));
            }

            SlideMasterPart masterPart = GetSlideMasterPart(masterIndex);
            SlideLayoutPart[] layouts = masterPart.SlideLayoutParts.ToArray();
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
        ///     Adds a slide using a layout type.
        /// </summary>
        public PowerPointSlide AddSlide(SlideLayoutValues layoutType, int masterIndex = 0) {
            int layoutIndex = GetLayoutIndex(layoutType, masterIndex);
            return AddSlide(masterIndex, layoutIndex);
        }

        /// <summary>
        ///     Adds a slide using a layout name.
        /// </summary>
        public PowerPointSlide AddSlide(string layoutName, int masterIndex = 0, bool ignoreCase = true) {
            int layoutIndex = GetLayoutIndex(layoutName, masterIndex, ignoreCase);
            return AddSlide(masterIndex, layoutIndex);
        }

        /// <summary>
        ///     Gets a theme color value in hex format (e.g. "FF0000").
        /// </summary>
        public string? GetThemeColor(PowerPointThemeColor color, int masterIndex = 0) {
            ThrowIfDisposed();
            SlideMasterPart masterPart = GetSlideMasterPart(masterIndex);
            A.ColorScheme? scheme = masterPart.ThemePart?.Theme?.ThemeElements?.ColorScheme;
            if (scheme == null) {
                return null;
            }

            OpenXmlCompositeElement? element = GetColorElement(scheme, color);
            return element?.GetFirstChild<A.RgbColorModelHex>()?.Val;
        }

        /// <summary>
        ///     Sets a theme color value in hex format (e.g. "FF0000").
        /// </summary>
        public void SetThemeColor(PowerPointThemeColor color, string hexValue, int masterIndex = 0) {
            ThrowIfDisposed();
            if (string.IsNullOrWhiteSpace(hexValue)) {
                throw new ArgumentException("Theme color value cannot be null or empty.", nameof(hexValue));
            }

            SlideMasterPart masterPart = GetSlideMasterPart(masterIndex);
            A.ColorScheme scheme = EnsureColorScheme(masterPart);
            OpenXmlCompositeElement element = GetOrCreateColorElement(scheme, color);
            element.RemoveAllChildren<A.RgbColorModelHex>();
            element.RemoveAllChildren<A.SystemColor>();
            element.Append(new A.RgbColorModelHex { Val = hexValue });
        }

        /// <summary>
        ///     Sets multiple theme colors at once.
        /// </summary>
        public void SetThemeColors(IDictionary<PowerPointThemeColor, string> colors, int masterIndex = 0) {
            ThrowIfDisposed();
            if (colors == null) {
                throw new ArgumentNullException(nameof(colors));
            }

            foreach (KeyValuePair<PowerPointThemeColor, string> entry in colors) {
                SetThemeColor(entry.Key, entry.Value, masterIndex);
            }
        }

        /// <summary>
        ///     Gets the major/minor Latin fonts for the theme.
        /// </summary>
        public PowerPointThemeFontInfo GetThemeLatinFonts(int masterIndex = 0) {
            ThrowIfDisposed();
            SlideMasterPart masterPart = GetSlideMasterPart(masterIndex);
            A.FontScheme? scheme = masterPart.ThemePart?.Theme?.ThemeElements?.FontScheme;
            string? majorLatin = scheme?.MajorFont?.LatinFont?.Typeface;
            string? minorLatin = scheme?.MinorFont?.LatinFont?.Typeface;
            return new PowerPointThemeFontInfo(majorLatin, minorLatin);
        }

        /// <summary>
        ///     Sets the major/minor Latin fonts for the theme.
        /// </summary>
        public void SetThemeLatinFonts(string majorLatin, string minorLatin, int masterIndex = 0) {
            ThrowIfDisposed();
            if (string.IsNullOrWhiteSpace(majorLatin)) {
                throw new ArgumentException("Major font cannot be null or empty.", nameof(majorLatin));
            }
            if (string.IsNullOrWhiteSpace(minorLatin)) {
                throw new ArgumentException("Minor font cannot be null or empty.", nameof(minorLatin));
            }

            SlideMasterPart masterPart = GetSlideMasterPart(masterIndex);
            A.FontScheme scheme = EnsureFontScheme(masterPart);
            scheme.MajorFont ??= new A.MajorFont();
            scheme.MinorFont ??= new A.MinorFont();

            scheme.MajorFont.LatinFont ??= new A.LatinFont();
            scheme.MinorFont.LatinFont ??= new A.LatinFont();
            scheme.MajorFont.LatinFont.Typeface = majorLatin;
            scheme.MinorFont.LatinFont.Typeface = minorLatin;
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

            sourceSlide.Save();

            string slideRelId = GetNextSlideRelationshipId();
            SlidePart slidePart = _presentationPart.AddNewPart<SlidePart>(slideRelId);
            slidePart.Slide = (Slide)sourcePart.Slide.CloneNode(true);

            CloneSlidePartRelationships(sourcePart, slidePart, ShouldSharePart, includeDataParts: true);

            SlideIdList slideIdList = _presentationPart.Presentation.SlideIdList ??= new SlideIdList();
            SlideId slideId = new() { Id = GetNextSlideId(), RelationshipId = slideRelId };
            InsertSlideId(slideIdList, slideId, targetIndex);

            PowerPointSlide duplicate = new(slidePart);
            duplicate.Hidden = sourceSlide.Hidden;
            _slides.Insert(targetIndex, duplicate);
            _presentationPart.Presentation.Save();
            return duplicate;
        }

        /// <summary>
        ///     Imports a slide from another presentation and inserts it into the current presentation.
        /// </summary>
        /// <param name="sourcePresentation">Presentation to import from.</param>
        /// <param name="sourceIndex">Index of the slide to import.</param>
        /// <param name="insertAt">Index where the imported slide should be inserted. Defaults to end.</param>
        public PowerPointSlide ImportSlide(PowerPointPresentation sourcePresentation, int sourceIndex, int? insertAt = null) {
            ThrowIfDisposed();
            if (sourcePresentation == null) {
                throw new ArgumentNullException(nameof(sourcePresentation));
            }

            if (ReferenceEquals(sourcePresentation, this)) {
                return DuplicateSlide(sourceIndex, insertAt);
            }

            SlidePart? initialSlidePart = null;
            if (_initialSlideUntouched && _slides.Count == 1) {
                initialSlidePart = _slides[0].SlidePart;
            }

            IReadOnlyList<PowerPointSlide> sourceSlides = sourcePresentation.Slides;
            if (sourceIndex < 0 || sourceIndex >= sourceSlides.Count) {
                throw new ArgumentOutOfRangeException(nameof(sourceIndex));
            }

            int targetIndex = insertAt ?? _slides.Count;
            if (targetIndex < 0 || targetIndex > _slides.Count) {
                throw new ArgumentOutOfRangeException(nameof(insertAt));
            }

            PowerPointSlide sourceSlide = sourceSlides[sourceIndex];
            sourceSlide.Save();

            SlideLayoutPart? sourceLayoutPart = sourceSlide.SlidePart.SlideLayoutPart;
            if (sourceLayoutPart == null) {
                throw new InvalidOperationException("Source slide does not have a layout to import.");
            }

            SlideLayoutPart? targetLayoutPart = FindMatchingLayout(sourceLayoutPart);
            if (targetLayoutPart == null) {
                SlideMasterPart sourceMasterPart = sourceLayoutPart.SlideMasterPart
                    ?? throw new InvalidOperationException("Source slide layout does not have a master.");

                Dictionary<SlideLayoutPart, SlideLayoutPart> layoutMap;
                CloneSlideMasterPart(sourceMasterPart, out layoutMap);

                if (!layoutMap.TryGetValue(sourceLayoutPart, out targetLayoutPart)) {
                    throw new InvalidOperationException("Failed to resolve the imported slide layout.");
                }
            }

            string slideRelId = GetNextSlideRelationshipId();
            SlidePart slidePart = _presentationPart.AddNewPart<SlidePart>(slideRelId);
            slidePart.Slide = (Slide)sourceSlide.SlidePart.Slide.CloneNode(true);

            Dictionary<DataPart, MediaDataPart> mediaPartMap = new();
            CloneSlidePartRelationships(
                sourceSlide.SlidePart,
                slidePart,
                shouldShare: _ => false,
                includeDataParts: true,
                shouldSkip: part => part is SlideLayoutPart || part is NotesSlidePart,
                dataPartMap: mediaPartMap);

            string? layoutRelId = sourceSlide.SlidePart.GetIdOfPart(sourceLayoutPart);
            if (string.IsNullOrWhiteSpace(layoutRelId)) {
                layoutRelId = GetNextRelationshipId(slidePart);
            }

            slidePart.AddPart(targetLayoutPart, layoutRelId);

            SlideIdList slideIdList = _presentationPart.Presentation.SlideIdList ??= new SlideIdList();
            SlideId slideId = new() { Id = GetNextSlideId(), RelationshipId = slideRelId };
            InsertSlideId(slideIdList, slideId, targetIndex);

            PowerPointSlide imported = new(slidePart);
            imported.Hidden = sourceSlide.Hidden;

            if (sourceSlide.SlidePart.NotesSlidePart != null) {
                string notesText = sourceSlide.Notes.Text;
                if (!string.IsNullOrWhiteSpace(notesText)) {
                    imported.Notes.Text = notesText;
                }
            }

            _slides.Insert(targetIndex, imported);
            _presentationPart.Presentation.Save();

            if (initialSlidePart != null) {
                _initialSlideUntouched = false;
                int blankIndex = _slides.FindIndex(slide => ReferenceEquals(slide.SlidePart, initialSlidePart));
                if (blankIndex >= 0) {
                    RemoveSlide(blankIndex);
                }
            }

            return imported;
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

        private SlideMasterPart GetSlideMasterPart(int masterIndex) {
            SlideMasterPart[] masters = _presentationPart.SlideMasterParts.ToArray();
            if (masterIndex < 0 || masterIndex >= masters.Length) {
                throw new ArgumentOutOfRangeException(nameof(masterIndex));
            }
            return masters[masterIndex];
        }

        private static A.Theme EnsureTheme(SlideMasterPart masterPart) {
            ThemePart themePart = masterPart.ThemePart ?? masterPart.AddNewPart<ThemePart>();
            themePart.Theme ??= new A.Theme { ThemeElements = new A.ThemeElements() };
            themePart.Theme.ThemeElements ??= new A.ThemeElements();
            return themePart.Theme;
        }

        private static A.ColorScheme EnsureColorScheme(SlideMasterPart masterPart) {
            A.Theme theme = EnsureTheme(masterPart);
            theme.ThemeElements ??= new A.ThemeElements();
            A.ColorScheme scheme = theme.ThemeElements.ColorScheme ??= new A.ColorScheme { Name = "Office" };
            return scheme;
        }

        private static A.FontScheme EnsureFontScheme(SlideMasterPart masterPart) {
            A.Theme theme = EnsureTheme(masterPart);
            theme.ThemeElements ??= new A.ThemeElements();
            A.FontScheme scheme = theme.ThemeElements.FontScheme ??= new A.FontScheme { Name = "Office" };
            return scheme;
        }

        private static OpenXmlCompositeElement? GetColorElement(A.ColorScheme scheme, PowerPointThemeColor color) {
            return color switch {
                PowerPointThemeColor.Dark1 => scheme.GetFirstChild<A.Dark1Color>(),
                PowerPointThemeColor.Light1 => scheme.GetFirstChild<A.Light1Color>(),
                PowerPointThemeColor.Dark2 => scheme.GetFirstChild<A.Dark2Color>(),
                PowerPointThemeColor.Light2 => scheme.GetFirstChild<A.Light2Color>(),
                PowerPointThemeColor.Accent1 => scheme.GetFirstChild<A.Accent1Color>(),
                PowerPointThemeColor.Accent2 => scheme.GetFirstChild<A.Accent2Color>(),
                PowerPointThemeColor.Accent3 => scheme.GetFirstChild<A.Accent3Color>(),
                PowerPointThemeColor.Accent4 => scheme.GetFirstChild<A.Accent4Color>(),
                PowerPointThemeColor.Accent5 => scheme.GetFirstChild<A.Accent5Color>(),
                PowerPointThemeColor.Accent6 => scheme.GetFirstChild<A.Accent6Color>(),
                PowerPointThemeColor.Hyperlink => scheme.GetFirstChild<A.Hyperlink>(),
                PowerPointThemeColor.FollowedHyperlink => scheme.GetFirstChild<A.FollowedHyperlinkColor>(),
                _ => null
            };
        }

        private static OpenXmlCompositeElement GetOrCreateColorElement(A.ColorScheme scheme, PowerPointThemeColor color) {
            OpenXmlCompositeElement? element = GetColorElement(scheme, color);
            if (element != null) {
                return element;
            }

            element = color switch {
                PowerPointThemeColor.Dark1 => new A.Dark1Color(),
                PowerPointThemeColor.Light1 => new A.Light1Color(),
                PowerPointThemeColor.Dark2 => new A.Dark2Color(),
                PowerPointThemeColor.Light2 => new A.Light2Color(),
                PowerPointThemeColor.Accent1 => new A.Accent1Color(),
                PowerPointThemeColor.Accent2 => new A.Accent2Color(),
                PowerPointThemeColor.Accent3 => new A.Accent3Color(),
                PowerPointThemeColor.Accent4 => new A.Accent4Color(),
                PowerPointThemeColor.Accent5 => new A.Accent5Color(),
                PowerPointThemeColor.Accent6 => new A.Accent6Color(),
                PowerPointThemeColor.Hyperlink => new A.Hyperlink(),
                PowerPointThemeColor.FollowedHyperlink => new A.FollowedHyperlinkColor(),
                _ => new A.Dark1Color()
            };

            scheme.Append(element);
            return element;
        }

        private string GetNextSlideMasterRelationshipId() {
            var existingRelationships = new HashSet<string>(
                _presentationPart.Parts
                    .Select(p => p.RelationshipId)
                    .Union(_presentationPart.ExternalRelationships.Select(r => r.Id))
                    .Union(_presentationPart.HyperlinkRelationships.Select(r => r.Id))
                    .Where(id => !string.IsNullOrEmpty(id))
                    .Select(id => id!)
            );

            if (_presentationPart.Presentation.SlideMasterIdList != null) {
                foreach (SlideMasterId existingMasterId in _presentationPart.Presentation.SlideMasterIdList.Elements<SlideMasterId>()) {
                    if (existingMasterId.RelationshipId is { Value: { Length: > 0 } existingRelId }) {
                        existingRelationships.Add(existingRelId);
                    }
                }
            }

            int nextId = 1;
            string masterRelId;
            do {
                masterRelId = "rId" + nextId;
                nextId++;
            } while (existingRelationships.Contains(masterRelId));

            return masterRelId;
        }

        private uint GetNextSlideMasterId() {
            SlideMasterIdList? slideMasterIdList = _presentationPart.Presentation.SlideMasterIdList;
            if (slideMasterIdList != null && slideMasterIdList.Elements<SlideMasterId>().Any()) {
                uint maxId = slideMasterIdList.Elements<SlideMasterId>().Max(s => s.Id?.Value ?? 0U);
                return maxId >= 2147483648U ? maxId + 1U : 2147483648U;
            }

            return 2147483648U;
        }

        private SlideLayoutPart? FindMatchingLayout(SlideLayoutPart sourceLayoutPart) {
            SlideLayout? sourceLayout = sourceLayoutPart.SlideLayout;
            if (sourceLayout == null) {
                return null;
            }

            string? sourceName = sourceLayout.CommonSlideData?.Name?.Value;
            SlideLayoutValues? sourceType = sourceLayout.Type?.Value;

            foreach (SlideMasterPart masterPart in _presentationPart.SlideMasterParts) {
                foreach (SlideLayoutPart layoutPart in masterPart.SlideLayoutParts) {
                    SlideLayout? candidateLayout = layoutPart.SlideLayout;
                    if (candidateLayout == null) {
                        continue;
                    }

                    SlideLayoutValues? candidateType = candidateLayout.Type?.Value;
                    if (sourceType.HasValue && candidateType != sourceType) {
                        continue;
                    }

                    if (!string.IsNullOrWhiteSpace(sourceName)) {
                        string? candidateName = candidateLayout.CommonSlideData?.Name?.Value;
                        if (!string.Equals(sourceName, candidateName, StringComparison.OrdinalIgnoreCase)) {
                            continue;
                        }
                    }

                    return layoutPart;
                }
            }

            return null;
        }

        private P14.SectionList? GetSectionList(bool create) {
            Presentation presentation = _presentationPart.Presentation ??= new Presentation();
            PresentationExtensionList? extList = presentation.GetFirstChild<PresentationExtensionList>();
            if (extList == null && create) {
                extList = new PresentationExtensionList();
                presentation.Append(extList);
            }

            if (extList == null) {
                return null;
            }

            PresentationExtension? sectionExt = extList.Elements<PresentationExtension>()
                .FirstOrDefault(ext => string.Equals(ext.Uri?.Value, SectionListUri, StringComparison.Ordinal));
            if (sectionExt == null && create) {
                sectionExt = new PresentationExtension { Uri = SectionListUri };
                extList.Append(sectionExt);
            }

            if (sectionExt == null) {
                return null;
            }

            P14.SectionList? sectionList = sectionExt.GetFirstChild<P14.SectionList>();
            if (sectionList == null && create) {
                sectionList = new P14.SectionList();
                sectionList.AddNamespaceDeclaration("p14", P14Namespace);
                sectionExt.Append(sectionList);
            }

            return sectionList;
        }

        private P14.SectionList EnsureSectionList(IReadOnlyList<SlideId> slideIds) {
            P14.SectionList sectionList = GetSectionList(create: true)
                ?? throw new InvalidOperationException("Unable to create a section list.");
            if (!sectionList.Elements<P14.Section>().Any()) {
                List<uint> ids = slideIds
                    .Select(id => id.Id?.Value)
                    .Where(id => id != null)
                    .Select(id => id!.Value)
                    .ToList();
                sectionList.Append(CreateSection(DefaultSectionName, ids));
            }

            EnsureSectionCoverage(sectionList, slideIds);
            return sectionList;
        }

        private static void EnsureSectionCoverage(P14.SectionList sectionList, IReadOnlyList<SlideId> slideIds) {
            Dictionary<uint, int> slideIndexMap = BuildSlideIndexMap(slideIds);
            HashSet<uint> assigned = new();

            foreach (P14.Section section in sectionList.Elements<P14.Section>()) {
                P14.SectionSlideIdList list = EnsureSectionSlideIdList(section);
                List<P14.SectionSlideIdListEntry> entries = list.Elements<P14.SectionSlideIdListEntry>()
                    .Where(entry => entry.Id?.Value != null && slideIndexMap.ContainsKey(entry.Id.Value))
                    .ToList();

                list.RemoveAllChildren();
                foreach (P14.SectionSlideIdListEntry entry in entries
                             .OrderBy(entry => slideIndexMap[entry.Id!.Value])) {
                    list.Append(entry);
                    assigned.Add(entry.Id!.Value);
                }
            }

            P14.Section? lastSection = sectionList.Elements<P14.Section>().LastOrDefault();
            if (lastSection == null) {
                return;
            }

            P14.SectionSlideIdList target = EnsureSectionSlideIdList(lastSection);
            foreach (uint slideId in slideIndexMap.Keys.Where(id => !assigned.Contains(id))) {
                target.Append(new P14.SectionSlideIdListEntry { Id = slideId });
            }
        }

        private static P14.Section CreateSection(string name, IReadOnlyList<uint> slideIds) {
            P14.Section section = new() {
                Id = Guid.NewGuid().ToString("D"),
                Name = name
            };
            P14.SectionSlideIdList list = new();
            foreach (uint slideId in slideIds) {
                list.Append(new P14.SectionSlideIdListEntry { Id = slideId });
            }
            section.Append(list);
            return section;
        }

        private static P14.SectionSlideIdList EnsureSectionSlideIdList(P14.Section section) {
            P14.SectionSlideIdList? list = section.SectionSlideIdList;
            if (list == null) {
                list = new P14.SectionSlideIdList();
                section.Append(list);
            }
            return list;
        }

        private static Dictionary<uint, int> BuildSlideIndexMap(IReadOnlyList<SlideId> slideIds) {
            Dictionary<uint, int> map = new();
            for (int i = 0; i < slideIds.Count; i++) {
                uint? id = slideIds[i].Id?.Value;
                if (id != null) {
                    map[id.Value] = i;
                }
            }
            return map;
        }

        private static P14.Section? FindSectionBySlideId(P14.SectionList sectionList, uint slideId) {
            foreach (P14.Section section in sectionList.Elements<P14.Section>()) {
                P14.SectionSlideIdList? list = section.SectionSlideIdList;
                if (list == null) {
                    continue;
                }

                if (list.Elements<P14.SectionSlideIdListEntry>().Any(entry => entry.Id?.Value == slideId)) {
                    return section;
                }
            }

            return null;
        }

        private PowerPointSectionInfo BuildSectionInfo(P14.Section section, IReadOnlyList<SlideId> slideIds) {
            Dictionary<uint, int> slideIndexMap = BuildSlideIndexMap(slideIds);
            List<int> indices = new();
            P14.SectionSlideIdList? list = section.SectionSlideIdList;
            if (list != null) {
                foreach (P14.SectionSlideIdListEntry entry in list.Elements<P14.SectionSlideIdListEntry>()) {
                    uint? id = entry.Id?.Value;
                    if (id != null && slideIndexMap.TryGetValue(id.Value, out int index)) {
                        indices.Add(index);
                    }
                }
            }

            indices.Sort();
            return new PowerPointSectionInfo(section.Name?.Value ?? string.Empty, section.Id?.Value ?? string.Empty, indices);
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

        private static string GetNextRelationshipId(OpenXmlPartContainer container) {
            var existingRelationships = new HashSet<string>(
                container.Parts.Select(p => p.RelationshipId)
                    .Concat(container.ExternalRelationships.Select(r => r.Id))
                    .Concat(container.HyperlinkRelationships.Select(r => r.Id))
                    .Where(id => !string.IsNullOrEmpty(id)),
                StringComparer.Ordinal);

            int nextId = 1;
            string relId;
            do {
                relId = "rId" + nextId;
                nextId++;
            } while (!existingRelationships.Add(relId));

            return relId;
        }

        private SlideMasterPart CloneSlideMasterPart(
            SlideMasterPart sourceMasterPart,
            out Dictionary<SlideLayoutPart, SlideLayoutPart> layoutMap) {
            layoutMap = new Dictionary<SlideLayoutPart, SlideLayoutPart>();

            if (sourceMasterPart.SlideMaster == null) {
                throw new InvalidOperationException("Source slide master is missing.");
            }

            string masterRelId = GetNextSlideMasterRelationshipId();
            SlideMasterPart targetMasterPart = _presentationPart.AddNewPart<SlideMasterPart>(masterRelId);
            targetMasterPart.SlideMaster = (SlideMaster)sourceMasterPart.SlideMaster.CloneNode(true);

            foreach (var partPair in sourceMasterPart.Parts) {
                OpenXmlPart part = partPair.OpenXmlPart;
                string relId = partPair.RelationshipId;

                if (part is SlideLayoutPart sourceLayoutPart) {
                    SlideLayoutPart clonedLayout = CloneSlideLayoutPart(sourceLayoutPart, targetMasterPart, relId);
                    layoutMap[sourceLayoutPart] = clonedLayout;
                    continue;
                }

                ClonePartRecursive(part, targetMasterPart, relId, _ => false, includeDataParts: false);
            }

            CloneReferenceRelationships(sourceMasterPart, targetMasterPart, includeDataParts: false);

            SlideMasterIdList slideMasterIdList = _presentationPart.Presentation.SlideMasterIdList ??= new SlideMasterIdList();
            slideMasterIdList.Append(new SlideMasterId { Id = GetNextSlideMasterId(), RelationshipId = masterRelId });
            _presentationPart.Presentation.Save();

            return targetMasterPart;
        }

        private static SlideLayoutPart CloneSlideLayoutPart(
            SlideLayoutPart sourceLayoutPart,
            SlideMasterPart targetMasterPart,
            string relationshipId) {
            if (sourceLayoutPart.SlideLayout == null) {
                throw new InvalidOperationException("Source slide layout is missing.");
            }

            SlideLayoutPart targetLayoutPart = targetMasterPart.AddNewPart<SlideLayoutPart>(relationshipId);
            targetLayoutPart.SlideLayout = (SlideLayout)sourceLayoutPart.SlideLayout.CloneNode(true);

            CloneChildParts(
                sourceLayoutPart,
                targetLayoutPart,
                shouldSkip: part => part is SlideMasterPart,
                includeDataParts: false);

            targetLayoutPart.AddPart(targetMasterPart);
            return targetLayoutPart;
        }

        private static void CloneChildParts(
            OpenXmlPart sourcePart,
            OpenXmlPart targetPart,
            Func<OpenXmlPart, bool> shouldSkip,
            bool includeDataParts,
            Dictionary<DataPart, MediaDataPart>? dataPartMap = null) {
            foreach (var childPair in sourcePart.Parts) {
                if (shouldSkip(childPair.OpenXmlPart)) {
                    continue;
                }

                ClonePartRecursive(childPair.OpenXmlPart, targetPart, childPair.RelationshipId, _ => false, includeDataParts, dataPartMap);
            }

            CloneReferenceRelationships(sourcePart, targetPart, includeDataParts, dataPartMap);
        }

        private static void CloneSlidePartRelationships(
            SlidePart source,
            SlidePart target,
            Func<OpenXmlPart, bool> shouldShare,
            bool includeDataParts,
            Func<OpenXmlPart, bool>? shouldSkip = null,
            Dictionary<DataPart, MediaDataPart>? dataPartMap = null) {
            foreach (var partPair in source.Parts) {
                if (shouldSkip != null && shouldSkip(partPair.OpenXmlPart)) {
                    continue;
                }

                ClonePartRecursive(partPair.OpenXmlPart, target, partPair.RelationshipId, shouldShare, includeDataParts, dataPartMap);
            }

            CloneReferenceRelationships(source, target, includeDataParts, dataPartMap);
        }

        private static void ClonePartRecursive(
            OpenXmlPart sourcePart,
            OpenXmlPartContainer targetContainer,
            string relationshipId,
            Func<OpenXmlPart, bool> shouldShare,
            bool includeDataParts,
            Dictionary<DataPart, MediaDataPart>? dataPartMap = null) {
            if (shouldShare(sourcePart)) {
                AddExistingPart(targetContainer, sourcePart, relationshipId);
                return;
            }

            OpenXmlPart newPart = sourcePart is ExtendedPart extendedPart
                ? targetContainer.AddExtendedPart(extendedPart.RelationshipType, extendedPart.ContentType, relationshipId)
                : AddNewPartWithContentType(targetContainer, sourcePart, relationshipId);

            CopyPartData(sourcePart, newPart);
            CloneReferenceRelationships(sourcePart, newPart, includeDataParts, dataPartMap);

            foreach (var childPair in sourcePart.Parts) {
                ClonePartRecursive(childPair.OpenXmlPart, newPart, childPair.RelationshipId, shouldShare, includeDataParts, dataPartMap);
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

        private static void CopyPartData(DataPart sourcePart, DataPart targetPart) {
            using Stream sourceStream = sourcePart.GetStream(FileMode.Open, FileAccess.Read);
            using Stream targetStream = targetPart.GetStream(FileMode.Create, FileAccess.Write);
            sourceStream.CopyTo(targetStream);
        }

        private static void CloneReferenceRelationships(
            OpenXmlPartContainer source,
            OpenXmlPartContainer target,
            bool includeDataParts,
            Dictionary<DataPart, MediaDataPart>? dataPartMap = null) {
            foreach (ExternalRelationship rel in source.ExternalRelationships) {
                target.AddExternalRelationship(rel.RelationshipType, rel.Uri, rel.Id);
            }

            foreach (HyperlinkRelationship rel in source.HyperlinkRelationships) {
                target.AddHyperlinkRelationship(rel.Uri, rel.IsExternal, rel.Id);
            }

            if (includeDataParts) {
                CloneDataPartReferenceRelationships(source, target, dataPartMap);
            }
        }

        private static void CloneDataPartReferenceRelationships(
            OpenXmlPartContainer source,
            OpenXmlPartContainer target,
            Dictionary<DataPart, MediaDataPart>? dataPartMap) {
            OpenXmlPackage? sourcePackage = GetPackage(source);
            OpenXmlPackage? targetPackage = GetPackage(target);
            bool samePackage = sourcePackage != null && targetPackage != null && ReferenceEquals(sourcePackage, targetPackage);

            foreach (DataPartReferenceRelationship rel in source.DataPartReferenceRelationships) {
                if (rel.DataPart is not MediaDataPart mediaPart) {
                    continue;
                }

                MediaDataPart targetMediaPart = mediaPart;
                if (!samePackage) {
                    if (targetPackage == null) {
                        throw new InvalidOperationException("Unable to resolve target package for media import.");
                    }

                    if (dataPartMap != null && dataPartMap.TryGetValue(mediaPart, out MediaDataPart? existing)) {
                        targetMediaPart = existing;
                    } else {
                        targetMediaPart = CreateMediaDataPart(targetPackage, mediaPart.ContentType);
                        CopyPartData(mediaPart, targetMediaPart);
                        dataPartMap?.Add(mediaPart, targetMediaPart);
                    }
                }

                if (rel is AudioReferenceRelationship) {
                    if (TryAddMediaReferenceRelationship(target, "AddAudioReferenceRelationship", targetMediaPart, rel.Id)) {
                        continue;
                    }
                } else if (rel is VideoReferenceRelationship) {
                    if (TryAddMediaReferenceRelationship(target, "AddVideoReferenceRelationship", targetMediaPart, rel.Id)) {
                        continue;
                    }
                } else {
                    if (TryAddMediaReferenceRelationship(target, "AddMediaReferenceRelationship", targetMediaPart, rel.Id)) {
                        continue;
                    }
                }

                if (!samePackage) {
                    throw new InvalidOperationException("Unable to add media reference relationship to the imported slide.");
                }
            }
        }

        private static bool TryAddMediaReferenceRelationship(OpenXmlPartContainer target, string methodName,
            MediaDataPart mediaPart, string relationshipId) {
            MethodInfo? method = target.GetType().GetMethod(methodName,
                new[] { typeof(MediaDataPart), typeof(string) });
            if (method == null) {
                return false;
            }

            method.Invoke(target, new object[] { mediaPart, relationshipId });
            return true;
        }

        private static OpenXmlPackage? GetPackage(OpenXmlPartContainer container) {
            if (container is OpenXmlPackage package) {
                return package;
            }

            if (container is OpenXmlPart part) {
                return part.OpenXmlPackage;
            }

            return null;
        }

        private static MediaDataPart CreateMediaDataPart(OpenXmlPackage targetPackage, string contentType) {
            if (TryInvokeCreateMediaDataPart(targetPackage, new[] { typeof(string) }, new object[] { contentType }, out MediaDataPart? mediaPart) &&
                mediaPart != null) {
                return mediaPart;
            }

            MediaDataPartType? mediaType = TryGetMediaDataPartType(contentType);
            if (mediaType.HasValue &&
                TryInvokeCreateMediaDataPart(targetPackage, new[] { typeof(MediaDataPartType) }, new object[] { mediaType.Value }, out mediaPart) &&
                mediaPart != null) {
                return mediaPart;
            }

            throw new InvalidOperationException($"Unable to create a media data part for content type '{contentType}'.");
        }

        private static bool TryInvokeCreateMediaDataPart(
            OpenXmlPackage targetPackage,
            Type[] parameterTypes,
            object[] args,
            out MediaDataPart? mediaPart) {
            mediaPart = null;
            MethodInfo? method = targetPackage.GetType().GetMethod("CreateMediaDataPart", parameterTypes);
            if (method == null) {
                return false;
            }

            mediaPart = (MediaDataPart?)method.Invoke(targetPackage, args);
            return mediaPart != null;
        }

        private static MediaDataPartType? TryGetMediaDataPartType(string contentType) {
            if (string.IsNullOrWhiteSpace(contentType)) {
                return null;
            }

            return contentType.ToLowerInvariant() switch {
                "audio/aiff" => MediaDataPartType.Aiff,
                "audio/x-aiff" => MediaDataPartType.Aiff,
                "audio/midi" => MediaDataPartType.Midi,
                "audio/x-midi" => MediaDataPartType.Midi,
                "audio/mpeg" => MediaDataPartType.Mp3,
                "audio/mp3" => MediaDataPartType.Mp3,
                "audio/wav" => MediaDataPartType.Wav,
                "audio/x-wav" => MediaDataPartType.Wav,
                "audio/x-ms-wma" => MediaDataPartType.Wma,
                "audio/wma" => MediaDataPartType.Wma,
                "audio/ogg" => MediaDataPartType.OggAudio,
                "application/ogg" => MediaDataPartType.OggAudio,
                "audio/mpegurl" => MediaDataPartType.MpegUrl,
                "application/vnd.ms-asf" => MediaDataPartType.Asx,
                "video/x-msvideo" => MediaDataPartType.Avi,
                "video/avi" => MediaDataPartType.Avi,
                "video/mpeg" => MediaDataPartType.MpegVideo,
                "video/mpg" => MediaDataPartType.Mpg,
                "video/mp4" => MediaDataPartType.MpegVideo,
                "video/quicktime" => MediaDataPartType.Quicktime,
                "video/x-ms-wmv" => MediaDataPartType.Wmv,
                "video/x-ms-wmx" => MediaDataPartType.Wmx,
                "video/x-ms-wvx" => MediaDataPartType.Wvx,
                "video/ogg" => MediaDataPartType.OggVideo,
                "video/vc1" => MediaDataPartType.VC1,
                _ => null
            };
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
