using OfficeIMO.GoogleWorkspace;

namespace OfficeIMO.PowerPoint.GoogleSlides {
    public sealed class GoogleSlidesTranslationPlan {
        internal GoogleSlidesTranslationPlan(TranslationReport report) { Report = report; }
        public int SlideCount { get; internal set; }
        public int NativeTextBoxCount { get; internal set; }
        public int NativeTableCount { get; internal set; }
        public int NativeImageCount { get; internal set; }
        public int NativeShapeCount { get; internal set; }
        public int RasterizedSlideCount { get; internal set; }
        public int SpeakerNotesCount { get; internal set; }
        public int UnsupportedElementCount { get; internal set; }
        public TranslationReport Report { get; }
    }

    public abstract class GoogleSlidesElement {
        protected GoogleSlidesElement(string objectId, double left, double top, double width, double height) {
            ObjectId = objectId; LeftPoints = left; TopPoints = top; WidthPoints = width; HeightPoints = height;
        }
        public string ObjectId { get; }
        public double LeftPoints { get; }
        public double TopPoints { get; }
        public double WidthPoints { get; }
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

    public sealed class GoogleSlidesTextBox : GoogleSlidesElement {
        public GoogleSlidesTextBox(string id, double left, double top, double width, double height, string text) : base(id, left, top, width, height) { Text = text; }
        public string Text { get; }
        public string ShapeType { get; internal set; } = "TEXT_BOX";
        public bool Bold { get; internal set; }
        public bool Italic { get; internal set; }
        public bool Underline { get; internal set; }
        public int? FontSize { get; internal set; }
        public string? FontFamily { get; internal set; }
        public string? ForegroundColorHex { get; internal set; }
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
        internal int? FontSize { get; set; }
        internal string? FontFamily { get; set; }
        internal string? ForegroundColorHex { get; set; }
        internal string? Hyperlink { get; set; }
    }

    public sealed class GoogleSlidesTable : GoogleSlidesElement {
        internal GoogleSlidesTable(string id, double left, double top, double width, double height, IReadOnlyList<IReadOnlyList<string>> cells) : base(id, left, top, width, height) { Cells = cells; }
        public IReadOnlyList<IReadOnlyList<string>> Cells { get; }
    }

    public sealed class GoogleSlidesImage : GoogleSlidesElement {
        internal GoogleSlidesImage(string id, double left, double top, double width, double height, byte[] bytes, string contentType, string fileName) : base(id, left, top, width, height) {
            Bytes = bytes; ContentType = contentType; FileName = fileName;
        }
        public byte[] Bytes { get; }
        public string ContentType { get; }
        public string FileName { get; }
    }

    public sealed class GoogleSlidesShape : GoogleSlidesElement {
        internal GoogleSlidesShape(string id, double left, double top, double width, double height, string shapeType) : base(id, left, top, width, height) { ShapeType = shapeType; }
        public string ShapeType { get; }
        /// <summary>Editable fill and outline appearance for the shape.</summary>
        public GoogleSlidesShapeStyle Style { get; } = new GoogleSlidesShapeStyle();
    }

    public sealed class GoogleSlidesSlide {
        private readonly List<GoogleSlidesElement> _elements = new List<GoogleSlidesElement>();
        internal GoogleSlidesSlide(string objectId, int index) { ObjectId = objectId; Index = index; }
        public string ObjectId { get; }
        public int Index { get; }
        public string? BackgroundColorHex { get; internal set; }
        internal GoogleSlidesImage? BackgroundImage { get; set; }
        public string? SpeakerNotes { get; internal set; }
        /// <summary>Whether the source slide is hidden and should be skipped during presentation playback.</summary>
        public bool IsSkipped { get; internal set; }
        public bool IsRasterized { get; internal set; }
        public IReadOnlyList<GoogleSlidesElement> Elements => _elements;
        internal void Add(GoogleSlidesElement element) => _elements.Add(element);
    }

    public sealed class GoogleSlidesBatch {
        private readonly List<GoogleSlidesSlide> _slides = new List<GoogleSlidesSlide>();
        internal GoogleSlidesBatch(string title, double width, double height, GoogleSlidesTranslationPlan plan) {
            Title = title; WidthPoints = width; HeightPoints = height; Plan = plan;
        }
        public string Title { get; }
        public double WidthPoints { get; }
        public double HeightPoints { get; }
        public GoogleSlidesTranslationPlan Plan { get; }
        public IReadOnlyList<GoogleSlidesSlide> Slides => _slides;
        internal void Add(GoogleSlidesSlide slide) => _slides.Add(slide);
    }

    public sealed class GooglePresentationReference : GoogleDriveFileReference {
        public string? PresentationId { get; set; }
        public string? RevisionId { get; set; }
        public long? DriveVersion { get; set; }
        public DateTimeOffset? ModifiedTime { get; set; }
        public TranslationReport Report { get; set; } = new TranslationReport();
    }

    public sealed class GoogleSlidesImportResult {
        public GoogleSlidesImportResult(PowerPointPresentation presentation, GooglePresentationReference source, TranslationReport report) {
            Presentation = presentation; Source = source; Report = report;
        }
        public PowerPointPresentation Presentation { get; }
        public GooglePresentationReference Source { get; }
        public TranslationReport Report { get; }
    }
}
