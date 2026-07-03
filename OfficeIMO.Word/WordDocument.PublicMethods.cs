using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Net.Http;

namespace OfficeIMO.Word {
    /// <summary>
    /// Provides public methods for manipulating a Word document.
    /// </summary>
    public partial class WordDocument {
        /// <summary>
        /// Appends a paragraph to the document body.
        /// </summary>
        /// <param name="wordParagraph">Optional paragraph to append. When <c>null</c> a new paragraph is created.</param>
        /// <returns>The added <see cref="WordParagraph"/> instance.</returns>
        public WordParagraph AddParagraph(WordParagraph? wordParagraph = null) {
            if (wordParagraph is null) {
                // we create paragraph (and within that add it to document)
                wordParagraph = new WordParagraph(this, newParagraph: true, newRun: false);
            }

            var body = _wordprocessingDocument?.MainDocumentPart?.Document?.Body
                ?? throw new InvalidOperationException("Document body is missing.");

            WordParagraph.EnsureParagraphCanBeInserted(this, body, wordParagraph,
                "append a paragraph to the document body");
            AppendBlockToBody(wordParagraph._paragraph);
            wordParagraph.RefreshParent();
            return wordParagraph;
        }

        /// <summary>
        /// Adds a new paragraph containing the specified text.
        /// </summary>
        /// <param name="text">Text for the paragraph.</param>
        /// <returns>The created <see cref="WordParagraph"/>.</returns>
        public WordParagraph AddParagraph(string text) {
            //return AddParagraph().SetText(text);
            return AddParagraph().AddText(text);
        }

        /// <summary>
        /// Inserts a page break into the document.
        /// </summary>
        /// <returns>The created <see cref="WordParagraph"/> representing the page break.</returns>
        public WordParagraph AddPageBreak() {
            WordParagraph newWordParagraph = new WordParagraph {
                _run = new Run(new Break() { Type = BreakValues.Page }),
                _document = this
            };
            newWordParagraph._paragraph = new Paragraph(newWordParagraph._run);

            AppendBlockToBody(newWordParagraph._paragraph);
            newWordParagraph.RefreshParent();
            return newWordParagraph;
        }

        /// <summary>
        /// Adds default headers and footers to the document.
        /// </summary>
        public void AddHeadersAndFooters() {
            WordHeadersAndFooters.AddHeadersAndFooters(this);
        }

        /// <summary>
        /// Inserts a break into the document.
        /// </summary>
        /// <param name="breakType">Type of break to insert.</param>
        /// <returns>The created <see cref="WordParagraph"/> containing the break.</returns>
        public WordParagraph AddBreak(BreakValues? breakType = null) {
            breakType ??= BreakValues.Page;
            WordParagraph newWordParagraph = new WordParagraph {
                _run = new Run(new Break() { Type = breakType }),
                _document = this
            };
            newWordParagraph._paragraph = new Paragraph(newWordParagraph._run);

            var currentSection = this.Sections.LastOrDefault();
            if (currentSection != null) {
                currentSection.AppendParagraphToSection(newWordParagraph);
            } else {
                AppendBlockToBody(newWordParagraph._paragraph);
                newWordParagraph.RefreshParent();
            }
            return newWordParagraph;
        }

        /// <summary>
        /// Determines whether a paragraph style with the specified identifier exists in the document.
        /// </summary>
        /// <param name="styleId">The style identifier to look for.</param>
        /// <returns><c>true</c> if the style exists; otherwise, <c>false</c>.</returns>
        public bool StyleExists(string styleId) {
            if (string.IsNullOrWhiteSpace(styleId)) {
                return false;
            }
            var styles = _wordprocessingDocument?.MainDocumentPart?.StyleDefinitionsPart?.Styles;
            return styles != null && styles.OfType<DocumentFormat.OpenXml.Wordprocessing.Style>().Any(s => string.Equals(s.StyleId, styleId, StringComparison.OrdinalIgnoreCase));
        }

        /// <summary>
        /// Adds a hyperlink pointing to an external URI.
        /// </summary>
        /// <param name="text">Display text for the hyperlink.</param>
        /// <param name="uri">Target URI.</param>
        /// <param name="addStyle">Whether to apply hyperlink style.</param>
        /// <param name="tooltip">Tooltip for the hyperlink.</param>
        /// <param name="history">Whether to track link history.</param>
        /// <returns>The created <see cref="WordParagraph"/>.</returns>
        public WordParagraph AddHyperLink(string text, Uri uri, bool addStyle = false, string tooltip = "", bool history = true) {
            if (string.IsNullOrWhiteSpace(text)) {
                throw new ArgumentException("Text cannot be null or empty.", nameof(text));
            }

            if (uri == null) {
                throw new ArgumentException("Uri cannot be null.", nameof(uri));
            }

            return this.AddParagraph().AddHyperLink(text, uri, addStyle, tooltip, history);
        }

        /// <summary>
        /// Adds an internal hyperlink pointing to a bookmark.
        /// </summary>
        /// <param name="text">Display text for the hyperlink.</param>
        /// <param name="anchor">Bookmark anchor.</param>
        /// <param name="addStyle">Whether to apply hyperlink style.</param>
        /// <param name="tooltip">Tooltip for the hyperlink.</param>
        /// <param name="history">Whether to track link history.</param>
        /// <returns>The created <see cref="WordParagraph"/>.</returns>
        public WordParagraph AddHyperLink(string text, string anchor, bool addStyle = false, string tooltip = "", bool history = true) {
            if (string.IsNullOrWhiteSpace(text)) {
                throw new ArgumentException("Text cannot be null or empty.", nameof(text));
            }

            if (string.IsNullOrWhiteSpace(anchor)) {
                throw new ArgumentException("Anchor cannot be null or empty.", nameof(anchor));
            }

            return this.AddParagraph().AddHyperLink(text, anchor, addStyle, tooltip, history);
        }

        /// <summary>
        /// Downloads an image from the specified URL and inserts it into the document.
        /// </summary>
        /// <param name="url">URL of the image to download.</param>
        /// <param name="width">Optional width for the image.</param>
        /// <param name="height">Optional height for the image.</param>
        /// <returns>The created <see cref="WordImage"/>.</returns>
        public WordImage AddImageFromUrl(string url, double? width = null, double? height = null) {
            if (string.IsNullOrWhiteSpace(url)) {
                throw new ArgumentException("URL cannot be null or empty.", nameof(url));
            }

            using HttpClient client = new HttpClient();
            var data = client.GetByteArrayAsync(url).GetAwaiter().GetResult();
            using var ms = new MemoryStream(data);

            string fileName = "image";
            try {
                var uri = new Uri(url);
                fileName = Path.GetFileName(uri.LocalPath);
                if (string.IsNullOrEmpty(fileName)) {
                    fileName = "image";
                }
            } catch (UriFormatException) {
                // ignore and use default filename
            }

            var paragraph = AddParagraph();
            paragraph.AddImage(ms, fileName, width, height);
            return paragraph.Image ?? throw new InvalidOperationException("Image was not added to the paragraph.");
        }

        /// <summary>
        /// Inserts a VML image into the document.
        /// </summary>
        public WordImage AddImageVml(string filePathImage, double? width = null, double? height = null) {
            var paragraph = AddParagraph();
            paragraph.AddImageVml(filePathImage, width, height);
            return paragraph.Image ?? throw new InvalidOperationException("Image was not added to the paragraph.");
        }

        /// <summary>
        /// Adds the chart to the document. The type of chart is determined by the type of data passed in.
        /// You can use multiple:
        /// .AddBar() to add a bar chart
        /// .AddLine() to add a line chart
        /// .AddPie() to add a pie chart
        /// .AddArea() to add an area chart
        /// .AddScatter() to add a scatter chart
        /// .AddRadar() to add a radar chart
        /// .AddBar3D() to add a 3-D bar chart.
        /// .AddPie3D() to add a 3-D pie chart.
        /// .AddLine3D() to add a 3-D line chart.
        /// You can't mix and match the types of charts, except bar and line which can coexist in a combo chart.
        /// </summary>
        /// <param name="title">The title.</param>
        /// <param name="roundedCorners">if set to <c>true</c> [rounded corners].</param>
        /// <param name="width">The width.</param>
        /// <param name="height">The height.</param>
        /// <returns>WordChart</returns>
        public WordChart AddChart(string title = "", bool roundedCorners = false, int width = 600, int height = 600) {
            var paragraph = this.AddParagraph();
            var chartInstance = new WordChart(this, paragraph, title, roundedCorners, width, height);
            return chartInstance;
        }

        /// <summary>
        /// Creates a chart ready for combining bar and line series.
        /// Use <see cref="WordChart.AddChartAxisX"/> to supply category labels
        /// and then call <see cref="WordChart.AddBar(string,int,OfficeIMO.Drawing.OfficeColor)"/> or
        /// <see cref="WordChart.AddLine"/> to add data. The call to <c>AddChartAxisX</c> must be performed
        /// before adding any series so both chart types share the same axes.
        /// </summary>
        public WordChart AddComboChart(string title = "", bool roundedCorners = false, int width = 600, int height = 600) {
            return AddChart(title, roundedCorners, width, height);
        }

        /// <summary>
        /// Creates a list using one of the built-in numbering styles.
        /// For manually configured lists prefer <see cref="AddCustomList"/>.
        /// </summary>
        public WordList AddList(WordListStyle style) {
            WordList wordList = new WordList(this);
            wordList.AddList(style);
            return wordList;
        }

        /// <summary>
        /// Adds a bulleted list using the default bulleted style.
        /// </summary>
        /// <returns>The created <see cref="WordList"/>.</returns>
        public WordList AddListBulleted() {
            return AddList(WordListStyle.Bulleted);
        }

        /// <summary>
        /// Adds a numbered list using the default numbering style.
        /// </summary>
        /// <returns>The created <see cref="WordList"/>.</returns>
        public WordList AddListNumbered() {
            return AddList(WordListStyle.Numbered);
        }

        /// <summary>
        /// Adds a custom bullet list with formatting options.
        /// </summary>
        /// <param name="symbol">Bullet symbol.</param>
        /// <param name="fontName">Font name for the symbol.</param>
        /// <param name="colorHex">Hex color of the symbol.</param>
        /// <param name="fontSize">Font size in points.</param>
        /// <returns>The created <see cref="WordList"/>.</returns>
        public WordList AddCustomBulletList(char symbol, string fontName, string colorHex, int? fontSize = null) {
            return WordList.AddCustomBulletList(this, symbol, fontName, colorHex, fontSize);
        }

        /// <summary>
        /// Adds a custom bullet list with formatting options.
        /// </summary>
        /// <param name="symbol">Bullet symbol.</param>
        /// <param name="fontName">Font name for the symbol.</param>
        /// <param name="color">Bullet color.</param>
        /// <param name="colorHex">Hex color fallback.</param>
        /// <param name="fontSize">Font size in points.</param>
        /// <returns>The created <see cref="WordList"/>.</returns>
        public WordList AddCustomBulletList(WordBulletSymbol symbol, string fontName, OfficeIMO.Drawing.OfficeColor? color = null, string? colorHex = null, int? fontSize = null) {
            return WordList.AddCustomBulletList(this, symbol, fontName, color, colorHex, fontSize);
        }

        /// <summary>
        /// Adds a custom bullet list using the specified formatting.
        /// </summary>
        /// <param name="kind">Bullet level kind.</param>
        /// <param name="fontName">Font name for the bullet.</param>
        /// <param name="color">Bullet color.</param>
        /// <param name="colorHex">Hex color fallback.</param>
        /// <param name="fontSize">Font size in points.</param>
        /// <returns>The created <see cref="WordList"/>.</returns>
        /// <example>
        /// <code><![CDATA[
        /// using WordDocument document = WordDocument.Create(filePath);
        ///
        /// WordList list = document.AddCustomBulletList(
        ///     WordListLevelKind.BulletSquareSymbol,
        ///     "Courier New",
        ///     OfficeIMO.Drawing.OfficeColor.Red,
        ///     fontSize: 16);
        ///
        /// list.AddItem("Custom bullet item");
        /// document.Save();
        /// ]]></code>
        /// </example>
        public WordList AddCustomBulletList(WordListLevelKind kind, string fontName, OfficeIMO.Drawing.OfficeColor? color = null, string? colorHex = null, int? fontSize = null) {
            return WordList.AddCustomBulletList(this, kind, fontName, color, colorHex, fontSize);
        }

        /// <summary>
        /// Creates a bullet list where the bullet symbol is provided as an image.
        /// </summary>
        /// <param name="imageStream">Stream containing the image data.</param>
        /// <param name="fileName">File name used to determine the image type.</param>
        /// <returns>The created <see cref="WordList"/>.</returns>
        public WordList AddPictureBulletList(Stream imageStream, string fileName) {
            return WordList.AddPictureBulletList(this, imageStream, fileName);
        }

        /// <summary>
        /// Creates a bullet list where the bullet symbol is loaded from an image file.
        /// </summary>
        /// <param name="imagePath">Path to the image file.</param>
        /// <returns>The created <see cref="WordList"/>.</returns>
        public WordList AddPictureBulletList(string imagePath) {
            return WordList.AddPictureBulletList(this, imagePath);
        }

        /// <summary>
        /// Creates a custom list with no predefined levels for manual configuration.
        /// </summary>
        /// <returns>The created <see cref="WordList"/>.</returns>
        /// <example>
        /// <code><![CDATA[
        /// using WordDocument document = WordDocument.Create(filePath);
        ///
        /// WordList list = document.AddCustomList()
        ///     .AddListLevel(1, WordListLevelKind.BulletSquareSymbol, "Courier New", colorHex: "#FF0000", fontSize: 14)
        ///     .AddListLevel(5, WordListLevelKind.BulletBlackCircle, "Arial", colorHex: "#00FF00", fontSize: 10);
        ///
        /// list.AddItem("First level item");
        /// list.AddItem("Fifth level item", 4);
        /// document.Save();
        /// ]]></code>
        /// </example>
        public WordList AddCustomList() {
            return WordList.AddCustomList(this);
        }

        /// <summary>
        /// Creates a list configured for a table of contents.
        /// </summary>
        /// <param name="style">Numbering style to apply.</param>
        /// <returns>The created <see cref="WordList"/>.</returns>
        public WordList AddTableOfContentList(WordListStyle style) {
            WordList wordList = new WordList(this, true);
            wordList.AddList(style);
            return wordList;
        }

        /// <summary>
        /// Creates a numbering definition that can be customized and reused.
        /// </summary>
        /// <returns>The created <see cref="WordListNumbering"/>.</returns>
        public WordListNumbering CreateNumberingDefinition() {
            return WordListNumbering.CreateNumberingDefinition(this);
        }

        /// <summary>
        /// Retrieves a numbering definition by its identifier.
        /// </summary>
        /// <param name="abstractNumberId">Identifier of the numbering definition.</param>
        /// <returns>The <see cref="WordListNumbering"/> if found; otherwise, <c>null</c>.</returns>
        public WordListNumbering? GetNumberingDefinition(int abstractNumberId) {
            return WordListNumbering.GetNumberingDefinition(this, abstractNumberId);
        }

        /// <summary>
        /// Adds a table to the end of the document body.
        /// </summary>
        /// <param name="rows">Number of rows to create.</param>
        /// <param name="columns">Number of columns to create.</param>
        /// <param name="tableStyle">Optional table style to apply.</param>
        /// <returns>The inserted <see cref="WordTable"/> instance.</returns>
        public WordTable AddTable(int rows, int columns, WordTableStyle tableStyle = WordTableStyle.TableGrid) {
            WordTable wordTable = new WordTable(this, rows, columns, tableStyle);
            return wordTable;
        }

        /// <summary>
        /// Creates a table without inserting it into the document.
        /// </summary>
        /// <param name="rows">Number of rows to create.</param>
        /// <param name="columns">Number of columns to create.</param>
        /// <param name="tableStyle">Optional table style to apply.</param>
        /// <returns>The newly created <see cref="WordTable"/>.</returns>
        public WordTable CreateTable(int rows, int columns, WordTableStyle tableStyle = WordTableStyle.TableGrid) {
            return WordTable.Create(this, rows, columns, tableStyle);
        }

        /// <summary>
        /// Inserts an existing table after the provided paragraph.
        /// </summary>
        /// <param name="anchor">Paragraph after which the table will be inserted.</param>
        /// <param name="table">Table instance to insert.</param>
        /// <returns>The inserted <see cref="WordTable"/>.</returns>
        public WordTable InsertTableAfter(WordParagraph anchor, WordTable table) {
            if (anchor is null) throw new ArgumentNullException(nameof(anchor));
            if (table is null) throw new ArgumentNullException(nameof(table));

            anchor._paragraph.InsertAfterSelf(table._table);
            return table;
        }

        /// <summary>
        /// Inserts a paragraph at the specified index within the body.
        /// </summary>
        /// <param name="index">Zero-based position at which to insert the paragraph.</param>
        /// <param name="paragraph">Optional paragraph to insert. When <c>null</c> a new paragraph is created.</param>
        /// <returns>The inserted <see cref="WordParagraph"/>.</returns>
        public WordParagraph InsertParagraphAt(int index, WordParagraph? paragraph = null) {
            if (paragraph is null) {
                paragraph = new WordParagraph(this, true, false);
            }

            var body = _document.Body ?? throw new InvalidOperationException("Document body is missing.");
            var paragraphs = body.Elements<Paragraph>().ToList();
            if (index < 0 || index > paragraphs.Count) {
                throw new ArgumentOutOfRangeException(nameof(index));
            }

            if (index == paragraphs.Count) {
                var sectPr = body.Elements<SectionProperties>().FirstOrDefault();
                if (sectPr != null) {
                    body.InsertBefore(paragraph._paragraph, sectPr);
                } else {
                    body.Append(paragraph._paragraph);
                }
            } else {
                body.InsertBefore(paragraph._paragraph, paragraphs[index]);
            }
            return paragraph;
        }

        /// <summary>
        /// Updates page and total page number fields.
        /// When a table of contents is present the document is flagged to refresh
        /// fields on open so Word can update the TOC.
        /// </summary>
        public void UpdateFields() {
            WordFieldUpdater.UpdatePageCounters(this);
        }

        /// <summary>
        /// Updates page and total page number fields, using the supplied options for supported deterministic field refresh.
        /// </summary>
        /// <param name="options">Options controlling deterministic field refresh behavior.</param>
        public void UpdateFields(WordFieldUpdateOptions options) {
            UpdateFieldsAndGetReport(options);
        }

        /// <summary>
        /// Adds a table of contents to the current document.
        /// </summary>
        /// <param name="tableOfContentStyle">Optional style to use when creating the table of contents.</param>
        /// <param name="minLevel">Minimum heading level to include (1..9).</param>
        /// <param name="maxLevel">Maximum heading level to include (1..9).</param>
        /// <returns>The created <see cref="WordTableOfContent"/> instance.</returns>
        public WordTableOfContent AddTableOfContent(
            TableOfContentStyle tableOfContentStyle = TableOfContentStyle.Template1,
            int minLevel = 1,
            int maxLevel = 3) {
            WordTableOfContent wordTableContent = new WordTableOfContent(this, tableOfContentStyle, minLevel, maxLevel);
            var body = _document.Body ?? throw new InvalidOperationException("Document body is missing.");
            _tableOfContentIndex = body.ChildElements.ToList().IndexOf(wordTableContent.SdtBlock);
            _tableOfContentStyle = tableOfContentStyle;
            return wordTableContent;
        }

        /// <summary>
        /// Removes the current table of contents from the document if one exists.
        /// </summary>
        public void RemoveTableOfContent() {
            var toc = TableOfContent;
            if (toc is not null) {
                toc.SdtBlock.Remove();
                _tableOfContentIndex = null;
            }
        }

        /// <summary>
        /// Removes the existing table of contents and creates a new one at the same location.
        /// </summary>
        /// <returns>The newly created <see cref="WordTableOfContent"/>.</returns>
        public WordTableOfContent RegenerateTableOfContent() {
            var toc = TableOfContent;
            var style = _tableOfContentStyle ?? TableOfContentStyle.Template1;
            var body = _document.Body ?? throw new InvalidOperationException("Document body is missing.");
            int index = _tableOfContentIndex ?? (toc != null ? body.ChildElements.ToList().IndexOf(toc.SdtBlock) : -1);
            RemoveTableOfContent();
            var newToc = new WordTableOfContent(this, style);
            if (index >= 0 && index < body.ChildElements.Count - 1) {
                var block = newToc.SdtBlock;
                block.Remove();
                body.InsertAt(block, index);
                _tableOfContentIndex = index;
            } else {
                _tableOfContentIndex = body.ChildElements.ToList().IndexOf(newToc.SdtBlock);
            }
            return newToc;
        }

        /// <summary>
        /// Adds a built-in cover page to the document.
        /// </summary>
        /// <param name="coverPageTemplate">Cover page template to use.</param>
        /// <returns>The created <see cref="WordCoverPage"/>.</returns>
        public WordCoverPage AddCoverPage(CoverPageTemplate coverPageTemplate) {
            WordCoverPage wordCoverPage = new WordCoverPage(this, coverPageTemplate);
            return wordCoverPage;
        }

        /// <summary>
        /// Inserts a text box into the document.
        /// </summary>
        /// <param name="text">Initial text for the text box.</param>
        /// <param name="wrapTextImage">Text wrapping option.</param>
        /// <returns>The created <see cref="WordTextBox"/>.</returns>
        public WordTextBox AddTextBox(string text, WrapTextImage wrapTextImage = WrapTextImage.Square) {
            WordTextBox wordTextBox = new WordTextBox(this, text, wrapTextImage);
            return wordTextBox;
        }

        /// <summary>
        /// Inserts a VML text box into the document.
        /// </summary>
        public WordTextBox AddTextBoxVml(string text) {
            var paragraph = this.AddParagraph();
            return paragraph.AddTextBoxVml(text);
        }

        /// <summary>
        /// Adds a basic shape to the document in a new paragraph.
        /// </summary>
        /// <param name="shapeType">Type of shape to create.</param>
        /// <param name="widthPt">Width in points or line end X.</param>
        /// <param name="heightPt">Height in points or line end Y.</param>
        /// <param name="fillColor">Fill color in hex format.</param>
        /// <param name="strokeColor">Stroke color in hex format.</param>
        /// <param name="strokeWeightPt">Stroke weight in points.</param>
        /// <param name="arcSize">Corner roundness fraction for rounded rectangles.</param>
        /// <returns>The created <see cref="WordShape"/>.</returns>
        public WordShape AddShape(ShapeType shapeType, double widthPt, double heightPt,
            string fillColor = "#FFFFFF", string strokeColor = "#000000", double strokeWeightPt = 1, double arcSize = 0.25) {
            var paragraph = AddParagraph();
            return paragraph.AddShape(shapeType, widthPt, heightPt, fillColor, strokeColor, strokeWeightPt, arcSize);
        }

        /// <summary>
        /// Adds a basic shape to the document using <see cref="OfficeIMO.Drawing.OfficeColor"/> values.
        /// </summary>
        public WordShape AddShape(ShapeType shapeType, double widthPt, double heightPt,
            OfficeIMO.Drawing.OfficeColor fillColor, OfficeIMO.Drawing.OfficeColor strokeColor, double strokeWeightPt = 1, double arcSize = 0.25) {
            return AddShape(shapeType, widthPt, heightPt, fillColor.ToHexColor(), strokeColor.ToHexColor(), strokeWeightPt, arcSize);
        }

        /// <summary>
        /// Adds a DrawingML shape to the document in a new paragraph.
        /// </summary>
        /// <param name="shapeType">Type of shape to create.</param>
        /// <param name="widthPt">Width in points.</param>
        /// <param name="heightPt">Height in points.</param>
        public WordShape AddShapeDrawing(ShapeType shapeType, double widthPt, double heightPt) {
            var paragraph = AddParagraph();
            return paragraph.AddShapeDrawing(shapeType, widthPt, heightPt);
        }

        /// <summary>
        /// Adds a DrawingML shape anchored at an absolute position on the page in a new paragraph.
        /// </summary>
        public WordShape AddShapeDrawing(ShapeType shapeType, double widthPt, double heightPt, double leftPt, double topPt) {
            var paragraph = AddParagraph();
            return paragraph.AddShapeDrawing(shapeType, widthPt, heightPt, leftPt, topPt);
        }

        /// <summary>
        /// Inserts a SmartArt diagram into the document.
        /// </summary>
        /// <param name="type">Layout type of the SmartArt.</param>
        /// <returns>The created <see cref="WordSmartArt"/> instance.</returns>
        public WordSmartArt AddSmartArt(SmartArtType type) {
            var paragraph = AddParagraph();
            var smartArt = new WordSmartArt(this, paragraph, type);
            return smartArt;
        }


        /// <summary>
        /// Inserts a horizontal line into the document.
        /// </summary>
        /// <param name="lineType">Border style of the line.</param>
        /// <param name="color">Line color.</param>
        /// <param name="size">Line width in eighths of a point.</param>
        /// <param name="space">Space above and below the line.</param>
        /// <returns>The paragraph containing the line.</returns>
        public WordParagraph AddHorizontalLine(BorderValues? lineType = null, OfficeIMO.Drawing.OfficeColor? color = null, uint size = 12, uint space = 1) {
            lineType ??= BorderValues.Single;
            return this.AddParagraph().AddHorizontalLine(lineType.Value, color, size, space);
        }

        /// <summary>
        /// Adds a new section to the document.
        /// </summary>
        /// <param name="sectionMark">Section break type.</param>
        /// <returns>The created <see cref="WordSection"/>.</returns>
        public WordSection AddSection(SectionMarkValues? sectionMark = null) {
            Paragraph paragraph = new Paragraph();

            ParagraphProperties paragraphProperties = new ParagraphProperties();

            SectionProperties sectionProperties = WordHeadersAndFooters.CreateSectionProperties();

            if (sectionMark != null) {
                SectionType sectionType = new SectionType() { Val = sectionMark };
                sectionProperties.Append(sectionType);
            }

            paragraphProperties.Append(sectionProperties);
            paragraph.Append(paragraphProperties);


            AppendBlockToBody(paragraph);


            WordSection wordSection = new WordSection(this, paragraph);

            return wordSection;
        }

        /// <summary>
        /// Removes the section at the specified index.
        /// </summary>
        /// <param name="index">Zero based index of the section to remove.</param>
        public void RemoveSection(int index) {
            if (index < 0 || index >= this.Sections.Count) {
                throw new ArgumentOutOfRangeException(nameof(index));
            }

            this.Sections[index].RemoveSection();
        }

        /// <summary>
        /// Clones the section at the specified index and inserts the clone after it.
        /// </summary>
        /// <param name="index">Zero based index of the section to clone.</param>
        /// <returns>The cloned <see cref="WordSection"/>.</returns>
        public WordSection CloneSection(int index) {
            if (index < 0 || index >= this.Sections.Count) {
                throw new ArgumentOutOfRangeException(nameof(index));
            }

            return this.Sections[index].CloneSection();
        }

        /// <summary>
        /// Inserts a bookmark in a new paragraph.
        /// </summary>
        /// <param name="bookmarkName">Name of the bookmark.</param>
        /// <returns>The created <see cref="WordParagraph"/>.</returns>
        public WordParagraph AddBookmark(string bookmarkName) {
            return this.AddParagraph().AddBookmark(bookmarkName);
        }

        /// <summary>
        /// Inserts a citation field referencing the specified source tag.
        /// </summary>
        /// <param name="sourceTag">Tag of the bibliographic source.</param>
        /// <returns>The created <see cref="WordParagraph"/>.</returns>
        public WordParagraph AddCitation(string sourceTag) {
            var field = new CitationField { SourceTag = sourceTag };
            return this.AddParagraph().AddField(field);
        }

        /// <summary>
        /// Adds a field to the document in a new paragraph.
        /// </summary>
        /// <param name="wordFieldType">Type of field to insert.</param>
        /// <param name="wordFieldFormat">Optional field format.</param>
        /// <param name="customFormat">Custom format string for date or time fields.</param>
        /// <param name="advanced">Whether to use advanced formatting.</param>
        /// <param name="parameters">Additional switch parameters.</param>
        /// <returns>The created <see cref="WordParagraph"/>.</returns>
        public WordParagraph AddField(WordFieldType wordFieldType, WordFieldFormat? wordFieldFormat = null, string? customFormat = null, bool advanced = false, List<string>? parameters = null) {
            return this.AddParagraph().AddField(wordFieldType, wordFieldFormat, customFormat!, advanced, parameters!);
        }

        /// <summary>
        /// Adds a field represented by a <see cref="WordFieldCode"/> to the document in a new paragraph.
        /// </summary>
        /// <param name="fieldCode">Field code instance describing instructions and switches.</param>
        /// <param name="wordFieldFormat">Optional field format.</param>
        /// <param name="customFormat">Custom format string for date or time fields.</param>
        /// <param name="advanced">Whether to use advanced formatting.</param>
        /// <returns>The created <see cref="WordParagraph"/>.</returns>
        public WordParagraph AddField(WordFieldCode fieldCode, WordFieldFormat? wordFieldFormat = null, string? customFormat = null, bool advanced = false) {
            return this.AddParagraph().AddField(fieldCode, wordFieldFormat, customFormat!, advanced);
        }

        /// <summary>
        /// Adds a field built using <see cref="WordFieldBuilder"/>.
        /// </summary>
        /// <param name="builder">Field builder instance.</param>
        /// <param name="advanced">Whether to use advanced formatting.</param>
        /// <returns>The created <see cref="WordParagraph"/>.</returns>
        public WordParagraph AddField(WordFieldBuilder builder, bool advanced = false) {
            return this.AddParagraph().AddField(builder, advanced);
        }

        /// <summary>
        /// Inserts an equation specified in OMML format.
        /// </summary>
        /// <param name="omml">OMML markup for the equation.</param>
        /// <returns>The created <see cref="WordParagraph"/>.</returns>
        public WordParagraph AddEquation(string omml) {
            return this.AddParagraph().AddEquation(omml);
        }

        /// <summary>
        /// Embeds an object with a preview image.
        /// </summary>
        /// <param name="filePath">Path to the object file.</param>
        /// <param name="imageFilePath">Preview image path.</param>
        /// <param name="width">Optional width in points.</param>
        /// <param name="height">Optional height in points.</param>
        /// <returns>The paragraph containing the embedded object.</returns>
        public WordParagraph AddEmbeddedObject(string filePath, string imageFilePath, double? width = null, double? height = null) {
            return this.AddParagraph().AddEmbeddedObject(filePath, imageFilePath, width, height);
        }

        /// <summary>
        /// Embeds an object with custom options.
        /// </summary>
        /// <param name="filePath">Path to the object file.</param>
        /// <param name="options">Embedding options.</param>
        /// <returns>The paragraph containing the embedded object.</returns>
        public WordParagraph AddEmbeddedObject(string filePath, WordEmbeddedObjectOptions options) {
            return this.AddParagraph().AddEmbeddedObject(filePath, options);
        }

        /// <summary>
        /// Removes all watermarks from the document including headers.
        /// </summary>
        public void RemoveWatermark() {
            foreach (var section in this.Sections) {
                section.RemoveWatermark();
            }
        }

    }
}
