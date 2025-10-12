using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    /// <summary>
    /// Defines template styles that can be used when generating a table of contents.
    /// </summary>
    public enum TableOfContentStyle {
        /// <summary>
        /// Built-in layout with a heading followed by entries.
        /// </summary>
        Template1,

        /// <summary>
        /// Alternative layout with the same structure but different identifiers.
        /// </summary>
        Template2
    }

    /// <summary>
    /// Represents a table of contents within a Word document.
    /// </summary>
    public class WordTableOfContent : WordElement {
        private readonly WordDocument _document;
        private readonly SdtBlock _sdtBlock;

        /// <summary>
        /// Exposes the underlying structured document tag for this table of contents.
        /// </summary>
        internal SdtBlock SdtBlock => _sdtBlock;

        /// <summary>
        /// Gets the template style used to create this table of contents.
        /// </summary>
        public TableOfContentStyle Style { get; }

        /// <summary>
        /// Gets or sets the heading text displayed for the table of contents.
        /// </summary>
        public string Text {
            get {
                var contentBlock = _sdtBlock?.SdtContentBlock;
                if (contentBlock != null) {
                    foreach (var paragraph in contentBlock.ChildElements.OfType<Paragraph>()) {
                        var text = paragraph
                            .OfType<Run>()
                            .FirstOrDefault()?
                            .OfType<Text>()
                            .FirstOrDefault()?.Text;
                        if (text != null) {
                            return text;
                        }
                    }
                }
                return string.Empty;
            }
            set {
                var contentBlock = _sdtBlock?.SdtContentBlock;
                if (contentBlock != null) {
                    foreach (var paragraph in contentBlock.ChildElements.OfType<Paragraph>()) {
                        var text = paragraph
                            .OfType<Run>()
                            .FirstOrDefault()?
                            .OfType<Text>()
                            .FirstOrDefault();
                        if (text != null) {
                            text.Text = value;
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Gets or sets the text shown when the document has no entries for the table of contents.
        /// </summary>
        public string TextNoContent {
            get {
                var contentBlock = _sdtBlock?.SdtContentBlock;
                if (contentBlock != null) {
                    foreach (var paragraph in contentBlock.ChildElements.OfType<Paragraph>()) {
                        var text = paragraph
                            .OfType<SimpleField>()
                            .FirstOrDefault()?
                            .OfType<Run>()
                            .FirstOrDefault()?
                            .OfType<Text>()
                            .FirstOrDefault()?.Text;
                        if (text != null) {
                            return text;
                        }
                    }
                }
                return string.Empty;
            }
            set {
                var contentBlock = _sdtBlock?.SdtContentBlock;
                if (contentBlock != null) {
                    foreach (var paragraph in contentBlock.ChildElements.OfType<Paragraph>()) {
                        var text = paragraph
                            .OfType<SimpleField>()
                            .FirstOrDefault()?
                            .OfType<Run>()
                            .FirstOrDefault()?
                            .OfType<Text>()
                            .FirstOrDefault();
                        if (text != null) {
                            text.Text = value;
                        }
                    }
                }
            }
        }



        /// <summary>
        /// Initializes a new instance of the <see cref="WordTableOfContent"/> class and appends it to the document body.
        /// </summary>
        /// <param name="wordDocument">Parent document where the table of contents will be created.</param>
        /// <param name="tableOfContentStyle">Template style used to generate the table of contents.</param>
        public WordTableOfContent(WordDocument wordDocument, TableOfContentStyle tableOfContentStyle)
            : this(wordDocument, GetStyle(tableOfContentStyle), tableOfContentStyle, appendToBody: true, queueUpdateOnOpen: true) {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="WordTableOfContent"/> class using an existing structured document tag.
        /// </summary>
        /// <param name="wordDocument">Parent document that owns the table of contents.</param>
        /// <param name="sdtBlock">Structured document tag representing the table of contents.</param>
        public WordTableOfContent(WordDocument wordDocument, SdtBlock sdtBlock)
            : this(wordDocument, sdtBlock, TableOfContentStyle.Template1, appendToBody: false, queueUpdateOnOpen: true) {
        }

        internal WordTableOfContent(WordDocument wordDocument, SdtBlock sdtBlock, bool queueUpdateOnOpen)
            : this(wordDocument, sdtBlock, TableOfContentStyle.Template1, appendToBody: false, queueUpdateOnOpen: queueUpdateOnOpen) {
        }

        private WordTableOfContent(WordDocument wordDocument, SdtBlock sdtBlock, TableOfContentStyle style, bool appendToBody, bool queueUpdateOnOpen) {
            this._document = wordDocument ?? throw new ArgumentNullException(nameof(wordDocument));
            this._sdtBlock = sdtBlock ?? throw new ArgumentNullException(nameof(sdtBlock));
            this.Style = style;

            if (appendToBody) {
                var body = this._document._wordprocessingDocument?.MainDocumentPart?.Document?.Body;
                body?.Append(_sdtBlock);
            }

            if (queueUpdateOnOpen) {
                QueueUpdateOnOpen();
            }
        }

        /// <summary>
        /// Flags the document to update this table of contents when the file is opened.
        /// </summary>
        public void Update() {
            QueueUpdateOnOpen(force: true);
        }

        internal void QueueUpdateOnOpen(bool force = false) {
            if (force) {
                _ = MarkFieldsAsDirty();
            } else if (!MarkFieldsAsDirty()) {
                return;
            }

            this._document.Settings.UpdateFieldsOnOpen = true;
            this._document.NotifyTableOfContentUpdateQueued();
        }

        private bool MarkFieldsAsDirty() {
            if (_sdtBlock == null) {
                return false;
            }

            var marked = false;

            foreach (var simpleField in _sdtBlock.Descendants<SimpleField>()) {
                if (simpleField.Dirty?.Value != true) {
                    simpleField.Dirty = true;
                    marked = true;
                }
            }

            foreach (var fieldChar in _sdtBlock.Descendants<FieldChar>()) {
                if (fieldChar.Dirty?.Value != true) {
                    fieldChar.Dirty = true;
                    marked = true;
                }
            }

            return marked;
        }

        /// <summary>
        /// Deletes this table of contents from the parent document.
        /// </summary>
        public void Remove() {
            _document.RemoveTableOfContent();
        }

        /// <summary>
        /// Removes this table of contents and creates a new one in the same location.
        /// </summary>
        /// <returns>The newly created <see cref="WordTableOfContent"/> instance.</returns>
        public WordTableOfContent Regenerate() {
            return _document.RegenerateTableOfContent();
        }

        /// <summary>
        /// Returns a predefined structured document tag matching the chosen style.
        /// </summary>
        /// <param name="style">Template identifier to retrieve.</param>
        private static SdtBlock GetStyle(TableOfContentStyle style) {
            switch (style) {
                case TableOfContentStyle.Template1: return Template1;
                case TableOfContentStyle.Template2: return Template2;
            }
            throw new ArgumentOutOfRangeException(nameof(style));
        }
        /// <summary>
        /// Structured document tag implementing the default table-of-contents layout.
        /// </summary>
        private static SdtBlock Template1 {
            get {
                SdtBlock sdtBlock1 = new SdtBlock();

                SdtProperties sdtProperties1 = new SdtProperties();
                SdtId sdtId1 = new SdtId() { Val = -619995952 };

                SdtContentDocPartObject sdtContentDocPartObject1 = new SdtContentDocPartObject();
                DocPartGallery docPartGallery1 = new DocPartGallery() { Val = "Table of Contents" };
                DocPartUnique docPartUnique1 = new DocPartUnique();

                sdtContentDocPartObject1.Append(docPartGallery1);
                sdtContentDocPartObject1.Append(docPartUnique1);

                sdtProperties1.Append(sdtId1);
                sdtProperties1.Append(sdtContentDocPartObject1);

                SdtEndCharProperties sdtEndCharProperties1 = new SdtEndCharProperties();

                RunProperties runProperties1 = new RunProperties();
                RunFonts runFonts1 = new RunFonts() { AsciiTheme = ThemeFontValues.MinorHighAnsi, HighAnsiTheme = ThemeFontValues.MinorHighAnsi, EastAsiaTheme = ThemeFontValues.MinorHighAnsi, ComplexScriptTheme = ThemeFontValues.MinorBidi };
                Bold bold1 = new Bold();
                BoldComplexScript boldComplexScript1 = new BoldComplexScript();
                NoProof noProof1 = new NoProof();
                Color color1 = new Color() { Val = "auto" };
                FontSize fontSize1 = new FontSize() { Val = "22" };
                FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "22" };

                runProperties1.Append(runFonts1);
                runProperties1.Append(bold1);
                runProperties1.Append(boldComplexScript1);
                runProperties1.Append(noProof1);
                runProperties1.Append(color1);
                runProperties1.Append(fontSize1);
                runProperties1.Append(fontSizeComplexScript1);

                sdtEndCharProperties1.Append(runProperties1);

                SdtContentBlock sdtContentBlock1 = new SdtContentBlock();

                Paragraph paragraph1 = new Paragraph() { RsidParagraphAddition = "00327375", RsidRunAdditionDefault = "00327375", ParagraphId = "2054FA1D", TextId = "07CA9725" };

                ParagraphProperties paragraphProperties1 = new ParagraphProperties();
                ParagraphStyleId paragraphStyleId1 = new ParagraphStyleId() { Val = "TOCHeading" };

                paragraphProperties1.Append(paragraphStyleId1);

                Run run1 = new Run();
                Text text1 = new Text();
                text1.Text = "Table of Contents";

                run1.Append(text1);

                paragraph1.Append(paragraphProperties1);
                paragraph1.Append(run1);

                Paragraph paragraph2 = new Paragraph() { RsidParagraphAddition = "00327375", RsidRunAdditionDefault = "00327375", ParagraphId = "770BF6F5", TextId = "2FB682F4" };

                SimpleField simpleField1 = new SimpleField() { Instruction = " TOC \\o \"1-3\" \\h \\z \\u " };

                Run run2 = new Run();

                RunProperties runProperties2 = new RunProperties();
                Bold bold2 = new Bold();
                BoldComplexScript boldComplexScript2 = new BoldComplexScript();
                NoProof noProof2 = new NoProof();

                runProperties2.Append(bold2);
                runProperties2.Append(boldComplexScript2);
                runProperties2.Append(noProof2);
                Text text2 = new Text();
                text2.Text = "No table of contents entries found.";

                run2.Append(runProperties2);
                run2.Append(text2);

                simpleField1.Append(run2);

                paragraph2.Append(simpleField1);

                sdtContentBlock1.Append(paragraph1);
                sdtContentBlock1.Append(paragraph2);

                sdtBlock1.Append(sdtProperties1);
                sdtBlock1.Append(sdtEndCharProperties1);
                sdtBlock1.Append(sdtContentBlock1);
                return sdtBlock1;

            }
        }
        /// <summary>
        /// Alternative layout used for table-of-contents generation.
        /// </summary>
        private static SdtBlock Template2 {
            get {

                SdtBlock sdtBlock1 = new SdtBlock();

                SdtProperties sdtProperties1 = new SdtProperties();
                SdtId sdtId1 = new SdtId() { Val = -909075344 };

                SdtContentDocPartObject sdtContentDocPartObject1 = new SdtContentDocPartObject();
                DocPartGallery docPartGallery1 = new DocPartGallery() { Val = "Table of Contents" };
                DocPartUnique docPartUnique1 = new DocPartUnique();

                sdtContentDocPartObject1.Append(docPartGallery1);
                sdtContentDocPartObject1.Append(docPartUnique1);

                sdtProperties1.Append(sdtId1);
                sdtProperties1.Append(sdtContentDocPartObject1);

                SdtEndCharProperties sdtEndCharProperties1 = new SdtEndCharProperties();

                RunProperties runProperties1 = new RunProperties();
                RunFonts runFonts1 = new RunFonts() { AsciiTheme = ThemeFontValues.MinorHighAnsi, HighAnsiTheme = ThemeFontValues.MinorHighAnsi, EastAsiaTheme = ThemeFontValues.MinorHighAnsi, ComplexScriptTheme = ThemeFontValues.MinorBidi };
                Bold bold1 = new Bold();
                BoldComplexScript boldComplexScript1 = new BoldComplexScript();
                NoProof noProof1 = new NoProof();
                Color color1 = new Color() { Val = "auto" };
                FontSize fontSize1 = new FontSize() { Val = "22" };
                FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "22" };

                runProperties1.Append(runFonts1);
                runProperties1.Append(bold1);
                runProperties1.Append(boldComplexScript1);
                runProperties1.Append(noProof1);
                runProperties1.Append(color1);
                runProperties1.Append(fontSize1);
                runProperties1.Append(fontSizeComplexScript1);

                sdtEndCharProperties1.Append(runProperties1);

                SdtContentBlock sdtContentBlock1 = new SdtContentBlock();

                Paragraph paragraph1 = new Paragraph() { RsidParagraphAddition = "00CD6D01", RsidRunAdditionDefault = "00CD6D01", ParagraphId = "5645B277", TextId = "289B3262" };

                ParagraphProperties paragraphProperties1 = new ParagraphProperties();
                ParagraphStyleId paragraphStyleId1 = new ParagraphStyleId() { Val = "TOCHeading" };

                paragraphProperties1.Append(paragraphStyleId1);

                Run run1 = new Run();
                Text text1 = new Text();
                text1.Text = "Table of Contents";

                run1.Append(text1);

                paragraph1.Append(paragraphProperties1);
                paragraph1.Append(run1);

                Paragraph paragraph2 = new Paragraph() { RsidParagraphAddition = "00CD6D01", RsidRunAdditionDefault = "00CD6D01", ParagraphId = "4ADC5792", TextId = "715D6FE9" };

                SimpleField simpleField1 = new SimpleField() { Instruction = " TOC \\o \"1-3\" \\h \\z \\u " };

                Run run2 = new Run();

                RunProperties runProperties2 = new RunProperties();
                Bold bold2 = new Bold();
                BoldComplexScript boldComplexScript2 = new BoldComplexScript();
                NoProof noProof2 = new NoProof();

                runProperties2.Append(bold2);
                runProperties2.Append(boldComplexScript2);
                runProperties2.Append(noProof2);
                Text text2 = new Text();
                text2.Text = "No table of contents entries found.";

                run2.Append(runProperties2);
                run2.Append(text2);

                simpleField1.Append(run2);

                paragraph2.Append(simpleField1);

                sdtContentBlock1.Append(paragraph1);
                sdtContentBlock1.Append(paragraph2);

                sdtBlock1.Append(sdtProperties1);
                sdtBlock1.Append(sdtEndCharProperties1);
                sdtBlock1.Append(sdtContentBlock1);
                return sdtBlock1;
            }
        }
    }
}
