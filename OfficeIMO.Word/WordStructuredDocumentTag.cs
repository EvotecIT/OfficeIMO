using DocumentFormat.OpenXml.Wordprocessing;
using ImageSharpColor = SixLabors.ImageSharp.Color;

namespace OfficeIMO.Word {
    /// <summary>
    /// Represents a structured document tag (content control) within a Word document.
    /// </summary>
    public class WordStructuredDocumentTag : WordElement {
        private WordDocument _document;
        private Paragraph? _paragraph;
        private SdtRun? _stdRun;
        private SdtBlock? _sdtBlock;

        /// <summary>
        /// Gets the alias associated with this content control.
        /// </summary>
        public string? Alias {
            get {
                if (_stdRun != null) {
                    var sdtAlias = _stdRun.SdtProperties?.OfType<SdtAlias>().FirstOrDefault();
                    return sdtAlias?.Val;
                }

                if (_sdtBlock != null) {
                    var sdtAlias = _sdtBlock.SdtProperties?.OfType<SdtAlias>().FirstOrDefault();
                    return sdtAlias?.Val;
                }

                return null;
            }
        }

        /// <summary>
        /// Gets or sets the tag value for this content control.
        /// </summary>
        public string? Tag {
            get {
                if (_stdRun != null) {
                    var tag = _stdRun.SdtProperties?.OfType<Tag>().FirstOrDefault();
                    return tag?.Val;
                }

                if (_sdtBlock != null) {
                    var tag = _sdtBlock.SdtProperties?.OfType<Tag>().FirstOrDefault();
                    return tag?.Val;
                }

                return null;
            }
            set {
                if (_stdRun != null) {
                    var properties = _stdRun.SdtProperties ??= new SdtProperties();
                    var tag = properties.OfType<Tag>().FirstOrDefault();
                    if (tag == null) {
                        tag = new Tag();
                        properties.Append(tag);
                    }
                    tag.Val = value;
                } else if (_sdtBlock != null) {
                    var properties = _sdtBlock.SdtProperties ??= new SdtProperties();
                    var tag = properties.OfType<Tag>().FirstOrDefault();
                    if (tag == null) {
                        tag = new Tag();
                        properties.Append(tag);
                    }
                    tag.Val = value;
                }
            }
        }

        /// <summary>
        /// Gets or sets the inner text of the content control.
        /// </summary>
        public string? Text {
            get {
                if (_stdRun != null) {
                    var run = _stdRun.SdtContentRun?.ChildElements.OfType<Run>().FirstOrDefault();
                    var text = run?.OfType<Text>().FirstOrDefault();
                    return text?.Text;
                }

                if (_sdtBlock != null) {
                    var paragraph = _sdtBlock.SdtContentBlock?.ChildElements.OfType<Paragraph>().FirstOrDefault();
                    var run = paragraph?.ChildElements.OfType<Run>().FirstOrDefault();
                    var text = run?.OfType<Text>().FirstOrDefault();
                    return text?.Text;
                }

                return null;
            }
            set {
                if (_stdRun != null) {
                    var run = _stdRun.SdtContentRun?.ChildElements.OfType<Run>().FirstOrDefault();
                    if (run == null) {
                        run = new Run();
                        _stdRun.SdtContentRun ??= new SdtContentRun();
                        _stdRun.SdtContentRun.Append(run);
                    }

                    var text = run.OfType<Text>().FirstOrDefault();
                    if (text == null) {
                        text = new Text { Space = SpaceProcessingModeValues.Preserve };
                        run.Append(text);
                    }

                    text.Text = value ?? string.Empty;
                } else if (_sdtBlock != null) {
                    _sdtBlock.SdtContentBlock ??= new SdtContentBlock();
                    var paragraph = _sdtBlock.SdtContentBlock.ChildElements.OfType<Paragraph>().FirstOrDefault();
                    if (paragraph == null) {
                        paragraph = new Paragraph();
                        _sdtBlock.SdtContentBlock.Append(paragraph);
                    }

                    var run = paragraph.ChildElements.OfType<Run>().FirstOrDefault();
                    if (run == null) {
                        run = new Run();
                        paragraph.Append(run);
                    }

                    var text = run.OfType<Text>().FirstOrDefault();
                    if (text == null) {
                        text = new Text { Space = SpaceProcessingModeValues.Preserve };
                        run.Append(text);
                    }

                    text.Text = value ?? string.Empty;
                }
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether the content control text is bold.
        /// </summary>
        public bool Bold {
            get {
                var runProperties = GetRunProperties();
                return runProperties != null && runProperties.Bold != null;
            }
            set {
                var runProperties = VerifyRunProperties();
                if (value) {
                    runProperties.Bold = new Bold();
                    runProperties.BoldComplexScript = new BoldComplexScript();
                } else {
                    runProperties.BoldComplexScript?.Remove();
                    runProperties.Bold?.Remove();
                }
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether the content control text is italic.
        /// </summary>
        public bool Italic {
            get {
                var runProperties = GetRunProperties();
                return runProperties != null && runProperties.Italic != null;
            }
            set {
                var runProperties = VerifyRunProperties();
                if (value) {
                    runProperties.Italic = new Italic();
                    runProperties.ItalicComplexScript = new ItalicComplexScript();
                } else {
                    runProperties.ItalicComplexScript?.Remove();
                    runProperties.Italic?.Remove();
                }
            }
        }

        /// <summary>
        /// Gets or sets the underline style for the content control text.
        /// </summary>
        public UnderlineValues? Underline {
            get {
                var runProperties = GetRunProperties();
                return runProperties?.Underline?.Val?.Value;
            }
            set {
                var runProperties = VerifyRunProperties();
                if (value != null) {
                    runProperties.Underline ??= new Underline();
                    runProperties.Underline.Val = value;
                } else {
                    runProperties.Underline?.Remove();
                }
            }
        }

        /// <summary>
        /// Gets or sets the font size in points for the content control text.
        /// </summary>
        public int? FontSize {
            get {
                var runProperties = GetRunProperties();
                if (runProperties?.FontSize?.Val != null &&
                    int.TryParse(runProperties.FontSize.Val, out var halfPoints)) {
                    return halfPoints / 2;
                }
                return null;
            }
            set {
                var runProperties = VerifyRunProperties();
                if (value != null) {
                    runProperties.FontSize = new FontSize { Val = (value * 2).ToString() };
                } else {
                    runProperties.FontSize?.Remove();
                }
            }
        }

        /// <summary>
        /// Gets or sets the font family for the content control text.
        /// </summary>
        public string? FontFamily {
            get {
                var runProperties = GetRunProperties();
                return runProperties?.RunFonts?.Ascii;
            }
            set {
                var runProperties = VerifyRunProperties();
                runProperties.RunFonts ??= new RunFonts();

                if (string.IsNullOrEmpty(value)) {
                    runProperties.RunFonts.Ascii = null;
                    runProperties.RunFonts.HighAnsi = null;
                    runProperties.RunFonts.ComplexScript = null;
                    runProperties.RunFonts.EastAsia = null;
                } else {
                    runProperties.RunFonts.Ascii = value;
                    runProperties.RunFonts.HighAnsi = value;
                    runProperties.RunFonts.ComplexScript = value;
                    runProperties.RunFonts.EastAsia = value;
                }
            }
        }

        /// <summary>
        /// Gets or sets the text color using <see cref="Color"/>.
        /// </summary>
        public ImageSharpColor? Color {
            get {
                if (string.IsNullOrEmpty(ColorHex)) {
                    return null;
                }
                return Helpers.ParseColor(ColorHex);
            }
            set {
                ColorHex = value?.ToHexColor() ?? string.Empty;
            }
        }

        /// <summary>
        /// Gets or sets the text color as a hexadecimal string.
        /// </summary>
        public string ColorHex {
            get {
                var runProperties = GetRunProperties();
                return runProperties?.Color?.Val?.Value ?? string.Empty;
            }
            set {
                var runProperties = VerifyRunProperties();
                if (!string.IsNullOrEmpty(value)) {
                    runProperties.Color = new DocumentFormat.OpenXml.Wordprocessing.Color {
                        Val = value.Replace("#", "").ToLowerInvariant()
                    };
                } else {
                    runProperties.Color?.Remove();
                }
            }
        }

        /// <summary>
        /// Gets or sets the highlight color applied to the content control text.
        /// </summary>
        public HighlightColorValues? Highlight {
            get {
                var runProperties = GetRunProperties();
                return runProperties?.Highlight?.Val?.Value;
            }
            set {
                var runProperties = VerifyRunProperties();
                if (value.HasValue) {
                    runProperties.Highlight = new Highlight { Val = value.Value };
                } else {
                    runProperties.Highlight?.Remove();
                }
            }
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="WordStructuredDocumentTag"/> class.
        /// </summary>
        /// <param name="document">Parent document.</param>
        /// <param name="paragraph">Paragraph that contains the content control.</param>
        /// <param name="stdRun">Underlying structured document run.</param>
        public WordStructuredDocumentTag(WordDocument document, Paragraph paragraph, SdtRun stdRun) {
            this._document = document;
            this._paragraph = paragraph;
            this._stdRun = stdRun;
        }

        /// <summary>
        /// Initializes a new instance from a structured document block.
        /// </summary>
        /// <param name="document">Parent document.</param>
        /// <param name="sdtBlock">Structured document block.</param>
        public WordStructuredDocumentTag(WordDocument document, SdtBlock sdtBlock) {
            this._document = document;
            this._sdtBlock = sdtBlock;
        }

        /// <summary>
        /// Removes the structured document tag from the document.
        /// </summary>
        public void Remove() {
            if (this._stdRun != null) {
                this._stdRun.Remove();
            } else if (this._sdtBlock != null) {
                this._sdtBlock.Remove();
            }
        }

        /// <summary>
        /// Sets the content control text to bold and returns the instance for chaining.
        /// </summary>
        public WordStructuredDocumentTag SetBold(bool isBold = true) {
            Bold = isBold;
            return this;
        }

        /// <summary>
        /// Sets the content control text to italic and returns the instance for chaining.
        /// </summary>
        public WordStructuredDocumentTag SetItalic(bool isItalic = true) {
            Italic = isItalic;
            return this;
        }

        /// <summary>
        /// Sets the underline style and returns the instance for chaining.
        /// </summary>
        public WordStructuredDocumentTag SetUnderline(UnderlineValues? underline) {
            Underline = underline;
            return this;
        }

        /// <summary>
        /// Sets the font size in points and returns the instance for chaining.
        /// </summary>
        public WordStructuredDocumentTag SetFontSize(int fontSize) {
            FontSize = fontSize;
            return this;
        }

        /// <summary>
        /// Sets the font family and returns the instance for chaining.
        /// </summary>
        public WordStructuredDocumentTag SetFontFamily(string fontFamily) {
            FontFamily = fontFamily;
            return this;
        }

        /// <summary>
        /// Sets the text color using a hexadecimal value and returns the instance for chaining.
        /// </summary>
        public WordStructuredDocumentTag SetColorHex(string color) {
            ColorHex = color;
            return this;
        }

        /// <summary>
        /// Sets the text color and returns the instance for chaining.
        /// </summary>
        public WordStructuredDocumentTag SetColor(ImageSharpColor? color) {
            Color = color;
            return this;
        }

        /// <summary>
        /// Sets the highlight color and returns the instance for chaining.
        /// </summary>
        public WordStructuredDocumentTag SetHighlight(HighlightColorValues? highlight) {
            Highlight = highlight;
            return this;
        }

        private RunProperties? GetRunProperties() {
            var run = GetRun();
            return run?.GetFirstChild<RunProperties>();
        }

        private RunProperties VerifyRunProperties() {
            var run = VerifyRun();
            var runProperties = run.GetFirstChild<RunProperties>();
            return runProperties ?? run.PrependChild(new RunProperties());
        }

        private Run? GetRun() {
            if (_stdRun != null) {
                return _stdRun.SdtContentRun?.ChildElements.OfType<Run>().FirstOrDefault();
            }

            if (_sdtBlock != null) {
                var paragraph = _sdtBlock.SdtContentBlock?.ChildElements.OfType<Paragraph>().FirstOrDefault();
                return paragraph?.ChildElements.OfType<Run>().FirstOrDefault();
            }

            return null;
        }

        private Run VerifyRun() {
            if (_stdRun != null) {
                _stdRun.SdtContentRun ??= new SdtContentRun();
                var run = _stdRun.SdtContentRun.ChildElements.OfType<Run>().FirstOrDefault();
                if (run == null) {
                    run = new Run();
                    _stdRun.SdtContentRun.Append(run);
                }
                return run;
            }

            if (_sdtBlock != null) {
                _sdtBlock.SdtContentBlock ??= new SdtContentBlock();
                var paragraph = _sdtBlock.SdtContentBlock.ChildElements.OfType<Paragraph>().FirstOrDefault();
                if (paragraph == null) {
                    paragraph = new Paragraph();
                    _sdtBlock.SdtContentBlock.Append(paragraph);
                }

                var run = paragraph.ChildElements.OfType<Run>().FirstOrDefault();
                if (run == null) {
                    run = new Run();
                    paragraph.Append(run);
                }

                return run;
            }

            throw new InvalidOperationException("Structured document tag is not initialized.");
        }
    }
}
