using OfficeIMO.GoogleWorkspace;

namespace OfficeIMO.PowerPoint.GoogleSlides {
    /// <summary>Counts native, rasterized, and unsupported content found while planning a Slides export.</summary>
    public sealed class GoogleSlidesTranslationPlan {
        internal GoogleSlidesTranslationPlan(TranslationReport report) { Report = report; }
        /// <summary>Gets the number of source slides examined.</summary>
        public int SlideCount { get; internal set; }
        /// <summary>Gets the number of text-bearing shapes planned as native Slides elements.</summary>
        public int NativeTextBoxCount { get; internal set; }
        /// <summary>Gets the number of tables planned as native Slides elements.</summary>
        public int NativeTableCount { get; internal set; }
        /// <summary>Gets the number of pictures planned as native Slides elements.</summary>
        public int NativeImageCount { get; internal set; }
        /// <summary>Gets the number of basic shapes planned as native Slides elements.</summary>
        public int NativeShapeCount { get; internal set; }
        /// <summary>Gets the number of complex slides selected for whole-slide rasterization.</summary>
        public int RasterizedSlideCount { get; internal set; }
        /// <summary>Gets the number of source slides with speaker-note text.</summary>
        public int SpeakerNotesCount { get; internal set; }
        /// <summary>Gets the number of unsupported source elements or backgrounds reported.</summary>
        public int UnsupportedElementCount { get; internal set; }
        /// <summary>Gets fidelity notices from planning.</summary>
        public TranslationReport Report { get; }
    }

    /// <summary>Positioned slide content with dimensions in points.</summary>
    public abstract class GoogleSlidesElement {
        /// <summary>Initializes an element's identifier and source geometry in points.</summary>
        protected GoogleSlidesElement(string objectId, double left, double top, double width, double height) {
            ObjectId = objectId; LeftPoints = left; TopPoints = top; WidthPoints = width; HeightPoints = height;
        }
        /// <summary>Gets the identifier assigned for the Slides batch.</summary>
        public string ObjectId { get; }
        /// <summary>Gets the element's left coordinate in points.</summary>
        public double LeftPoints { get; }
        /// <summary>Gets the element's top coordinate in points.</summary>
        public double TopPoints { get; }
        /// <summary>Gets the element's width in points.</summary>
        public double WidthPoints { get; }
        /// <summary>Gets the element's height in points.</summary>
        public double HeightPoints { get; }
        /// <summary>Clockwise rotation of the source PowerPoint element, in degrees.</summary>
        public double RotationDegrees { get; internal set; }
        /// <summary>Whether the source PowerPoint element is reflected across its vertical axis.</summary>
        public bool HorizontalFlip { get; internal set; }
        /// <summary>Whether the source PowerPoint element is reflected across its horizontal axis.</summary>
        public bool VerticalFlip { get; internal set; }
    }

    /// <summary>Basic editable shape appearance supported by Google Slides.</summary>
    public sealed class GoogleSlidesShapeStyle {
        /// <summary>Solid fill color as RGB or RGBA hex.</summary>
        public string? FillColorHex { get; internal set; }
        /// <summary>Fill transparency percentage from 0 (opaque) to 100 (transparent).</summary>
        public int? FillTransparencyPercent { get; internal set; }
        /// <summary>Solid outline color as RGB or RGBA hex.</summary>
        public string? OutlineColorHex { get; internal set; }
        /// <summary>Outline width in points.</summary>
        public double? OutlineWidthPoints { get; internal set; }
    }

    /// <summary>Editable text-bearing shape planned for Google Slides.</summary>
    public sealed class GoogleSlidesTextBox : GoogleSlidesElement {
        /// <summary>Creates a positioned text box with supplied text.</summary>
        public GoogleSlidesTextBox(string id, double left, double top, double width, double height, string text) : base(id, left, top, width, height) { Text = text; }
        /// <summary>Gets the text inserted into the Slides shape.</summary>
        public string Text { get; }
        /// <summary>Gets the Slides shape type; defaults to <c>TEXT_BOX</c>.</summary>
        public string ShapeType { get; internal set; } = "TEXT_BOX";
        /// <summary>Gets the projected first-run bold setting, used when no per-run styles are present.</summary>
        public bool Bold { get; internal set; }
        /// <summary>Gets the projected first-run italic setting, used when no per-run styles are present.</summary>
        public bool Italic { get; internal set; }
        /// <summary>Gets the projected first-run underline setting, used when no per-run styles are present.</summary>
        public bool Underline { get; internal set; }
        /// <summary>Gets the projected first-run strikethrough setting, used when no per-run styles are present.</summary>
        public bool Strikethrough { get; internal set; }
        /// <summary>Gets the projected first-run small-caps setting, used when no per-run styles are present.</summary>
        public bool SmallCaps { get; internal set; }
        /// <summary>Gets the projected first-run Slides baseline offset, when present.</summary>
        public string? BaselineOffset { get; internal set; }
        /// <summary>Gets the projected first-run font size, when present.</summary>
        public int? FontSize { get; internal set; }
        /// <summary>Gets the projected first-run font family, when present.</summary>
        public string? FontFamily { get; internal set; }
        /// <summary>Gets the projected first-run foreground color as a hex value, when present.</summary>
        public string? ForegroundColorHex { get; internal set; }
        /// <summary>Gets the projected first-run hyperlink, when present.</summary>
        public string? Hyperlink { get; internal set; }
        internal List<GoogleSlidesTextStyleRun> TextRuns { get; } = new List<GoogleSlidesTextStyleRun>();
        /// <summary>Editable fill and outline appearance for text-bearing shapes.</summary>
        public GoogleSlidesShapeStyle Style { get; } = new GoogleSlidesShapeStyle();
    }

    internal sealed class GoogleSlidesTextStyleRun {
        internal int StartIndex { get; set; }
        internal int EndIndex { get; set; }
        internal bool Bold { get; set; }
        internal bool Italic { get; set; }
        internal bool Underline { get; set; }
        internal bool Strikethrough { get; set; }
        internal bool SmallCaps { get; set; }
        internal string? BaselineOffset { get; set; }
        internal int? FontSize { get; set; }
        internal string? FontFamily { get; set; }
        internal string? ForegroundColorHex { get; set; }
        internal string? Hyperlink { get; set; }
    }

    internal sealed class GoogleSlidesTableCell {
        internal GoogleSlidesTableCell(string text, IReadOnlyList<GoogleSlidesTextStyleRun> textRuns) {
            Text = text;
            TextRuns = textRuns;
        }
        internal string Text { get; }
        internal IReadOnlyList<GoogleSlidesTextStyleRun> TextRuns { get; }
    }

    /// <summary>Editable table cells projected from a PowerPoint table.</summary>
    public sealed class GoogleSlidesTable : GoogleSlidesElement {
        internal GoogleSlidesTable(string id, double left, double top, double width, double height, IReadOnlyList<IReadOnlyList<GoogleSlidesTableCell>> cells) : base(id, left, top, width, height) {
            StyledCells = cells;
            Cells = cells.Select(row => (IReadOnlyList<string>)row.Select(cell => cell.Text).ToArray()).ToArray();
        }
        /// <summary>Gets the table's row and cell text.</summary>
        public IReadOnlyList<IReadOnlyList<string>> Cells { get; }
        internal IReadOnlyList<IReadOnlyList<GoogleSlidesTableCell>> StyledCells { get; }
    }

    /// <summary>Image content prepared for a Slides batch.</summary>
    public sealed class GoogleSlidesImage : GoogleSlidesElement {
        internal GoogleSlidesImage(string id, double left, double top, double width, double height, byte[] bytes, string contentType, string fileName) : base(id, left, top, width, height) {
            Bytes = bytes; ContentType = contentType; FileName = fileName;
        }
        /// <summary>Gets the image bytes; the batch model does not copy the array.</summary>
        public byte[] Bytes { get; }
        /// <summary>Gets the image media type.</summary>
        public string ContentType { get; }
        /// <summary>Gets the name used for the temporary image content.</summary>
        public string FileName { get; }
    }

    /// <summary>An editable non-text shape planned for Slides.</summary>
    public sealed class GoogleSlidesShape : GoogleSlidesElement {
        internal GoogleSlidesShape(string id, double left, double top, double width, double height, string shapeType) : base(id, left, top, width, height) { ShapeType = shapeType; }
        /// <summary>Gets the Slides shape type.</summary>
        public string ShapeType { get; }
        /// <summary>Editable fill and outline appearance for the shape.</summary>
        public GoogleSlidesShapeStyle Style { get; } = new GoogleSlidesShapeStyle();
    }

    /// <summary>One slide and the elements selected for its Slides batch.</summary>
    public sealed class GoogleSlidesSlide {
        private readonly List<GoogleSlidesElement> _elements = new List<GoogleSlidesElement>();
        internal GoogleSlidesSlide(string objectId, int index) { ObjectId = objectId; Index = index; }
        /// <summary>Gets the identifier assigned for this Slides batch.</summary>
        public string ObjectId { get; }
        /// <summary>Gets the zero-based index in the source presentation.</summary>
        public int Index { get; }
        /// <summary>Gets the projected solid background color as a hex value, when available.</summary>
        public string? BackgroundColorHex { get; internal set; }
        internal GoogleSlidesImage? BackgroundImage { get; set; }
        /// <summary>Gets source speaker-note text, when present.</summary>
        public string? SpeakerNotes { get; internal set; }
        /// <summary>Whether the source slide is hidden and should be skipped during presentation playback.</summary>
        public bool IsSkipped { get; internal set; }
        /// <summary>Gets whether the slide was selected for whole-slide image rendering.</summary>
        public bool IsRasterized { get; internal set; }
        /// <summary>Gets the slide elements prepared for this batch.</summary>
        public IReadOnlyList<GoogleSlidesElement> Elements => _elements;
        internal void Add(GoogleSlidesElement element) => _elements.Add(element);
    }

    /// <summary>Slides page size, selected elements, and fidelity plan for one export.</summary>
    public sealed class GoogleSlidesBatch {
        private readonly List<GoogleSlidesSlide> _slides = new List<GoogleSlidesSlide>();
        internal GoogleSlidesBatch(string title, double width, double height, GoogleSlidesTranslationPlan plan) {
            Title = title; WidthPoints = width; HeightPoints = height; Plan = plan;
        }
        /// <summary>Gets the selected presentation title.</summary>
        public string Title { get; }
        /// <summary>Gets the presentation width in points.</summary>
        public double WidthPoints { get; }
        /// <summary>Gets the presentation height in points.</summary>
        public double HeightPoints { get; }
        /// <summary>Gets counts and fidelity notices produced while constructing the batch.</summary>
        public GoogleSlidesTranslationPlan Plan { get; }
        /// <summary>Gets the slides prepared for this batch in source order.</summary>
        public IReadOnlyList<GoogleSlidesSlide> Slides => _slides;
        internal void Add(GoogleSlidesSlide slide) => _slides.Add(slide);
    }

    /// <summary>Google Drive file reference with Slides revision and translation evidence.</summary>
    public sealed class GooglePresentationReference : GoogleDriveFileReference {
        /// <summary>Gets or sets the Slides presentation identifier.</summary>
        public string? PresentationId { get; set; }
        /// <summary>Gets or sets the Slides revision identifier, when available.</summary>
        public string? RevisionId { get; set; }
        /// <summary>Gets or sets the observed Drive version, when available.</summary>
        public long? DriveVersion { get; set; }
        /// <summary>Gets or sets the observed modification time, when available.</summary>
        public DateTimeOffset? ModifiedTime { get; set; }
        /// <summary>Gets or sets the translation notices associated with the reference.</summary>
        public TranslationReport Report { get; set; } = new TranslationReport();
    }

    /// <summary>An OfficeIMO presentation, its Google source reference, and import notices.</summary>
    public sealed class GoogleSlidesImportResult {
        /// <summary>Creates a result from an imported presentation and its source evidence.</summary>
        public GoogleSlidesImportResult(PowerPointPresentation presentation, GooglePresentationReference source, TranslationReport report) {
            Presentation = presentation; Source = source; Report = report;
        }
        /// <summary>Gets the imported presentation; the caller owns its lifetime.</summary>
        public PowerPointPresentation Presentation { get; }
        /// <summary>Gets the remote presentation reference and observed revision metadata.</summary>
        public GooglePresentationReference Source { get; }
        /// <summary>Gets translation notices produced by the import.</summary>
        public TranslationReport Report { get; }
    }
}
