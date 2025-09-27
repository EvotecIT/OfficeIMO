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
            body.AppendChild(wordParagraph._paragraph);
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

            var body = _document.Body ?? throw new InvalidOperationException("Document body is missing.");
            body.Append(newWordParagraph._paragraph);
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
                var body = _document.Body ?? throw new InvalidOperationException("Document body is missing.");
                body.Append(newWordParagraph._paragraph);
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
        /// and then call <see cref="WordChart.AddBar(string,int,SixLabors.ImageSharp.Color)"/> or
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
        public WordList AddCustomBulletList(WordBulletSymbol symbol, string fontName, SixLabors.ImageSharp.Color? color = null, string? colorHex = null, int? fontSize = null) {
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
        public WordList AddCustomBulletList(WordListLevelKind kind, string fontName, SixLabors.ImageSharp.Color? color = null, string? colorHex = null, int? fontSize = null) {
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
            int page = 1;
            foreach (var paragraph in Paragraphs) {
                var field = paragraph.Field;
                if (field != null && field.FieldType == WordFieldType.Page) {
                    field.Text = page.ToString();
                }

                if (paragraph.IsPageBreak) {
                    page++;
                }
            }

            foreach (var field in Fields.Where(f => f.FieldType == WordFieldType.NumPages)) {
                field.Text = page.ToString();
            }

            TableOfContent?.Update();
        }

        /// <summary>
        /// Adds a table of contents to the current document.
        /// </summary>
        /// <param name="tableOfContentStyle">Optional style to use when creating the table of contents.</param>
        /// <returns>The created <see cref="WordTableOfContent"/> instance.</returns>
        public WordTableOfContent AddTableOfContent(TableOfContentStyle tableOfContentStyle = TableOfContentStyle.Template1) {
            WordTableOfContent wordTableContent = new WordTableOfContent(this, tableOfContentStyle);
            var body = _document.Body ?? throw new InvalidOperationException("Document body is missing.");
            _tableOfContentIndex = body.ChildElements.Count - 1;
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
                _tableOfContentIndex = body.ChildElements.Count - 1;
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
        /// Adds a basic shape to the document using <see cref="SixLabors.ImageSharp.Color"/> values.
        /// </summary>
        public WordShape AddShape(ShapeType shapeType, double widthPt, double heightPt,
            SixLabors.ImageSharp.Color fillColor, SixLabors.ImageSharp.Color strokeColor, double strokeWeightPt = 1, double arcSize = 0.25) {
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
        public WordParagraph AddHorizontalLine(BorderValues? lineType = null, SixLabors.ImageSharp.Color? color = null, uint size = 12, uint space = 1) {
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


            var body = _document.Body ?? throw new InvalidOperationException("Document body is missing.");
            body.Append(paragraph);


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
        /// Adds a new paragraph with a content control (structured document tag).
        /// </summary>
        /// <param name="text">Initial text of the control.</param>
        /// <param name="alias">Optional alias for the control.</param>
        /// <param name="tag">Optional tag for the control.</param>
        /// <returns>The created <see cref="WordStructuredDocumentTag"/>.</returns>
        public WordStructuredDocumentTag AddStructuredDocumentTag(string text, string? alias = null, string? tag = null) {
            return this.AddParagraph().AddStructuredDocumentTag(text, alias!, tag!);
        }

        /// <summary>
        /// Adds a new paragraph with a repeating section content control.
        /// </summary>
        /// <param name="sectionTitle">Optional title of the repeating section.</param>
        /// <param name="alias">Optional alias for the control.</param>
        /// <param name="tag">Optional tag for the control.</param>
        /// <returns>The created <see cref="WordRepeatingSection"/>.</returns>
        public WordRepeatingSection AddRepeatingSection(string? sectionTitle = null, string? alias = null, string? tag = null) {
            return this.AddParagraph().AddRepeatingSection(sectionTitle!, alias!, tag!);
        }

        /// <summary>
        /// Embeds another document as an alternative format part.
        /// </summary>
        /// <param name="fileName">Path to the document.</param>
        /// <param name="type">Optional format part type.</param>
        /// <returns>The created <see cref="WordEmbeddedDocument"/>.</returns>
        public WordEmbeddedDocument AddEmbeddedDocument(string fileName, WordAlternativeFormatImportPartType? type = null) {
            return new WordEmbeddedDocument(this, fileName, type, false);
        }

        /// <summary>
        /// Embeds HTML content as an alternative format part.
        /// </summary>
        /// <param name="htmlContent">HTML content to embed.</param>
        /// <param name="type">Format part type.</param>
        /// <returns>The created <see cref="WordEmbeddedDocument"/>.</returns>
        public WordEmbeddedDocument AddEmbeddedFragment(string htmlContent, WordAlternativeFormatImportPartType type) {
            return new WordEmbeddedDocument(this, htmlContent, type, true);
        }

        /// <summary>
        /// Retrieves a structured document tag by its tag value.
        /// </summary>
        /// <param name="tag">Tag value of the control.</param>
        /// <returns>The matching <see cref="WordStructuredDocumentTag"/> or <c>null</c>.</returns>
        public WordStructuredDocumentTag? GetStructuredDocumentTagByTag(string tag) {
            return this.StructuredDocumentTags.FirstOrDefault(sdt => sdt.Tag == tag);
        }

        /// <summary>
        /// Retrieves a structured document tag by its alias.
        /// </summary>
        /// <param name="alias">Alias of the control.</param>
        /// <returns>The matching <see cref="WordStructuredDocumentTag"/> or <c>null</c>.</returns>
        public WordStructuredDocumentTag? GetStructuredDocumentTagByAlias(string alias) {
            return this.StructuredDocumentTags.FirstOrDefault(sdt => sdt.Alias == alias);
        }

        /// <summary>
        /// Retrieves a checkbox control by its tag value.
        /// </summary>
        /// <param name="tag">Tag value of the checkbox.</param>
        /// <returns>The matching <see cref="WordCheckBox"/> or <c>null</c>.</returns>
        public WordCheckBox? GetCheckBoxByTag(string tag) {
            return this.CheckBoxes.FirstOrDefault(cb => cb.Tag == tag);
        }

        /// <summary>
        /// Retrieves a checkbox control by its alias.
        /// </summary>
        /// <param name="alias">Alias of the checkbox.</param>
        /// <returns>The matching <see cref="WordCheckBox"/> or <c>null</c>.</returns>
        public WordCheckBox? GetCheckBoxByAlias(string alias) {
            return this.CheckBoxes.FirstOrDefault(cb => cb.Alias == alias);
        }

        /// <summary>
        /// Retrieves a date picker control by its tag value.
        /// </summary>
        /// <param name="tag">Tag value of the date picker.</param>
        /// <returns>The matching <see cref="WordDatePicker"/> or <c>null</c>.</returns>
        public WordDatePicker? GetDatePickerByTag(string tag) {
            return this.DatePickers.FirstOrDefault(dp => dp.Tag == tag);
        }

        /// <summary>
        /// Retrieves a date picker control by its alias.
        /// </summary>
        /// <param name="alias">Alias of the date picker.</param>
        /// <returns>The matching <see cref="WordDatePicker"/> or <c>null</c>.</returns>
        public WordDatePicker? GetDatePickerByAlias(string alias) {
            return this.DatePickers.FirstOrDefault(dp => dp.Alias == alias);
        }

        /// <summary>
        /// Retrieves a dropdown list control by its tag value.
        /// </summary>
        /// <param name="tag">Tag value of the dropdown list.</param>
        /// <returns>The matching <see cref="WordDropDownList"/> or <c>null</c>.</returns>
        public WordDropDownList? GetDropDownListByTag(string tag) {
            return this.DropDownLists.FirstOrDefault(dl => dl.Tag == tag);
        }

        /// <summary>
        /// Retrieves a dropdown list control by its alias.
        /// </summary>
        /// <param name="alias">Alias of the dropdown list.</param>
        /// <returns>The matching <see cref="WordDropDownList"/> or <c>null</c>.</returns>
        public WordDropDownList? GetDropDownListByAlias(string alias) {
            return this.DropDownLists.FirstOrDefault(dl => dl.Alias == alias);
        }

        /// <summary>
        /// Retrieves a combo box control by its tag value.
        /// </summary>
        public WordComboBox? GetComboBoxByTag(string tag) {
            return this.ComboBoxes.FirstOrDefault(cb => cb.Tag == tag);
        }

        /// <summary>
        /// Retrieves a combo box control by its alias.
        /// </summary>
        public WordComboBox? GetComboBoxByAlias(string alias) {
            return this.ComboBoxes.FirstOrDefault(cb => cb.Alias == alias);
        }

        /// <summary>
        /// Retrieves a picture control by its tag value.
        /// </summary>
        public WordPictureControl? GetPictureControlByTag(string tag) {
            return this.PictureControls.FirstOrDefault(pc => pc.Tag == tag);
        }

        /// <summary>
        /// Retrieves a picture control by its alias.
        /// </summary>
        public WordPictureControl? GetPictureControlByAlias(string alias) {
            return this.PictureControls.FirstOrDefault(pc => pc.Alias == alias);
        }

        /// <summary>
        /// Retrieves a repeating section control by its tag value.
        /// </summary>
        public WordRepeatingSection? GetRepeatingSectionByTag(string tag) {
            return this.RepeatingSections.FirstOrDefault(rs => rs.Tag == tag);
        }

        /// <summary>
        /// Retrieves a repeating section control by its alias.
        /// </summary>
        public WordRepeatingSection? GetRepeatingSectionByAlias(string alias) {
            return this.RepeatingSections.FirstOrDefault(rs => rs.Alias == alias);
        }
        /// <summary>
        /// Removes an embedded document from the document.
        /// </summary>
        /// <param name="embeddedDocument">Embedded document to remove.</param>
        public void RemoveEmbeddedDocument(WordEmbeddedDocument embeddedDocument) {
            if (embeddedDocument == null) {
                throw new ArgumentNullException(nameof(embeddedDocument));
            }

            embeddedDocument.Remove();
        }

        /// <summary>
        /// Removes all watermarks from the document including headers.
        /// </summary>
        public void RemoveWatermark() {
            foreach (var section in this.Sections) {
                section.RemoveWatermark();
            }
        }


        private int CombineRuns(WordHeaderFooter wordHeaderFooter) {
            int count = 0;
            if (wordHeaderFooter != null) {
                var defaultHeader = this.Header?.Default;
                if (defaultHeader != null) {
                    foreach (var p in defaultHeader.Paragraphs) count += CombineIdenticalRuns(p._paragraph);
                    foreach (var table in defaultHeader.Tables) {
                        table.Paragraphs.ForEach(p => count += CombineIdenticalRuns(p._paragraph));
                    }
                }
            }

            return count;
        }


        /// <summary>
        /// This method will combine identical runs in a paragraph.
        /// This is useful when you have a paragraph with multiple runs of the same style, that Microsoft Word creates.
        /// This feature is *EXPERIMENTAL* and may not work in all cases.
        /// It may impact on how your document looks like, please do extensive testing before using this feature.
        /// </summary>
        /// <returns></returns>
        public int CleanupDocument(DocumentCleanupOptions options = DocumentCleanupOptions.All) {
            int count = 0;

            if (_wordprocessingDocument?.MainDocumentPart?.Document?.Body != null) {
                foreach (var paragraph in _wordprocessingDocument.MainDocumentPart.Document.Body.Descendants<Paragraph>().ToList()) {
                    count += CleanupParagraph(paragraph, options);
                }
            }

            foreach (var header in _wordprocessingDocument?.MainDocumentPart?.HeaderParts ?? Enumerable.Empty<HeaderPart>()) {
                foreach (var paragraph in header.Header.Descendants<Paragraph>().ToList()) {
                    count += CleanupParagraph(paragraph, options);
                }
            }

            foreach (var footer in _wordprocessingDocument?.MainDocumentPart?.FooterParts ?? Enumerable.Empty<FooterPart>()) {
                foreach (var paragraph in footer.Footer.Descendants<Paragraph>().ToList()) {
                    count += CleanupParagraph(paragraph, options);
                }
            }

            return count;
        }

        /// <summary>
        /// Searches the document for paragraphs containing the specified text.
        /// </summary>
        /// <param name="text">Text to search for.</param>
        /// <param name="stringComparison">Comparison rules for the search.</param>
        /// <returns>A list of found <see cref="WordParagraph"/> instances.</returns>
        public List<WordParagraph> Find(string text, StringComparison stringComparison = StringComparison.OrdinalIgnoreCase) {
            int count = 0;
            List<WordParagraph> list = FindAndReplaceInternal(text, "", ref count, false, stringComparison);
            return list;
        }

        /// <summary>
        /// FindAdnReplace from the whole doc
        /// </summary>
        /// <param name="textToFind"></param>
        /// <param name="textToReplace"></param>
        /// <param name="stringComparison"></param>
        /// <returns></returns>
        public int FindAndReplace(string textToFind, string textToReplace, StringComparison stringComparison = StringComparison.OrdinalIgnoreCase) {
            int countFind = 0;
            FindAndReplaceInternal(textToFind, textToReplace, ref countFind, true, stringComparison);
            return countFind;
        }

        /// <summary>
        /// FindAdnReplace from the range parparagraphs
        /// </summary>
        /// <param name="paragraphs"></param>
        /// <param name="textToFind"></param>
        /// <param name="textToReplace"></param>
        /// <param name="stringComparison"></param>
        /// <returns></returns>
        public static int FindAndReplace(List<WordParagraph> paragraphs, string textToFind, string textToReplace, StringComparison stringComparison = StringComparison.OrdinalIgnoreCase) {
            int countFind = 0;
            FindAndReplaceNested(paragraphs, textToFind, textToReplace, ref countFind, true, stringComparison);
            return countFind;
        }


        private static List<WordParagraph> FindAndReplaceNested(List<WordParagraph> paragraphs, string textToFind, string textToReplace, ref int count, bool replace, StringComparison stringComparison = StringComparison.OrdinalIgnoreCase) {
            List<WordParagraph> foundParagraphs = ReplaceText(paragraphs, textToFind, textToReplace, ref count, replace, stringComparison);
            return foundParagraphs;
        }


        /// <summary>
        /// Replace text inside each paragraph
        /// </summary>
        /// <param name="paragraphs"></param>
        /// <param name="oldText"></param>
        /// <param name="newText"></param>
        /// <param name="count"></param>
        /// <param name="replace"></param>
        /// <param name="stringComparison"></param>
        /// <returns></returns>
        /// <exception cref="ArgumentNullException"></exception>
        private static List<WordParagraph> ReplaceText(List<WordParagraph> paragraphs, string oldText, string newText, ref int count, bool replace, StringComparison stringComparison = StringComparison.OrdinalIgnoreCase) {
            if (string.IsNullOrEmpty(oldText)) {
                throw new ArgumentNullException("oldText should not be null");
            }
            List<WordParagraph> foundParagraphs = new List<WordParagraph>();
            var removeParas = new List<int>();
            var foundList = SearchText(paragraphs, oldText, new WordPositionInParagraph() { Paragraph = 0 }, stringComparison);

            if (foundList?.Count > 0) {
                count += foundList.Count;
                foreach (var ts in foundList) {
                    if (!IsSegmentValid(paragraphs, ts))
                        continue;
                    if (ts.BeginIndex == ts.EndIndex) {
                        var p = paragraphs[ts.BeginIndex];
                        if (p is not null) {
                            if (replace) {
                                int replaceCount = 0;
                                p.Text = p.Text.FindAndReplace(oldText, newText, stringComparison, ref replaceCount);
                            }
                            if (!foundParagraphs.Any(fp => ReferenceEquals(fp._paragraph, p._paragraph))) {
                                foundParagraphs.Add(p);
                            }
                        }
                    } else {
                        if (replace) {
                            var beginPara = paragraphs[ts.BeginIndex];
                            var endPara = paragraphs[ts.EndIndex];
                            if (beginPara is not null && endPara is not null) {
                                beginPara.Text = beginPara.Text.Replace(beginPara.Text.Substring(ts.BeginChar), newText);
                                endPara.Text = endPara.Text.Replace(endPara.Text.Substring(0, ts.EndChar + 1), "");
                                if (!foundParagraphs.Any(fp => ReferenceEquals(fp._paragraph, beginPara._paragraph))) {
                                    foundParagraphs.Add(beginPara);
                                }
                            }
                            for (int i = ts.EndIndex - 1; i > ts.BeginIndex; i--) {
                                removeParas.Add(i);
                            }
                        }

                    }
                }
            }

            if (replace) {
                if (removeParas.Count > 0) {
                    removeParas = removeParas.Distinct().OrderByDescending(i => i).ToList();// Need remove by descending
                    foreach (var index in removeParas) {
                        paragraphs[index].Remove();//Remove blank paragraph
                    }
                }
            }
            return foundParagraphs;
        }

        private static List<WordTextSegment> SearchText(List<WordParagraph> paragraphs, String searched, WordPositionInParagraph startPos, StringComparison stringComparison = StringComparison.OrdinalIgnoreCase) {

            var segList = new List<WordTextSegment>();
            int startRun = startPos.Paragraph,
            startText = startPos.Text,
            startChar = startPos.Char;
            int beginRunPos = 0, beginCharPos = 0, candCharPos = 0;
            bool newList = false;
            for (int runPos = startRun; runPos < paragraphs.Count; runPos++) {
                int textPos = 0, charPos = 0;
                var p = paragraphs[runPos];

                if (!string.IsNullOrEmpty(p.Text)) {
                    if (textPos >= startText) {
                        string candidate = p.Text;
                        if (runPos == startRun)
                            charPos = startChar;
                        else
                            charPos = 0;
                        for (; charPos < candidate.Length; charPos++) {
                            if (string.Compare(candidate[charPos].ToString(), searched[0].ToString(), stringComparison) == 0 && (candCharPos == 0)) {
                                beginCharPos = charPos;
                                beginRunPos = runPos;
                                newList = true;
                            }
                            if (string.Compare(candidate[charPos].ToString(), searched[candCharPos].ToString(), stringComparison) == 0) {
                                if (candCharPos + 1 < searched.Length) {
                                    candCharPos++;
                                } else if (newList) {
                                    WordTextSegment segement = new WordTextSegment();
                                    segement.BeginIndex = (beginRunPos);
                                    segement.BeginChar = (beginCharPos);
                                    segement.EndIndex = (runPos);
                                    segement.EndChar = (charPos);
                                    segList.Add(segement);
                                    //Reset
                                    startChar = charPos;
                                    startText = textPos;
                                    startRun = runPos;
                                    newList = false;
                                    candCharPos = 0;
                                }
                            } else
                                candCharPos = 0;
                        }

                    }
                    textPos++;
                }


            }
            return segList;
        }

        private List<WordParagraph> FindAndReplaceInternal(string textToFind, string textToReplace, ref int count, bool replace, StringComparison stringComparison = StringComparison.OrdinalIgnoreCase) {
            WordFind wordFind = new WordFind();
            List<WordParagraph> list = new List<WordParagraph>();
            list.AddRange(FindAndReplaceNested(this.Paragraphs, textToFind, textToReplace, ref count, replace, stringComparison));

            foreach (var table in this.Tables) {
                list.AddRange(FindAndReplaceNested(table.Paragraphs, textToFind, textToReplace, ref count, replace, stringComparison));
            }

            if (this.Header?.Default != null) {
                list.AddRange(FindAndReplaceNested(this.Header.Default.Paragraphs, textToFind, textToReplace, ref count, replace, stringComparison));
                foreach (var table in this.Header.Default.Tables) {
                    list.AddRange(FindAndReplaceNested(table.Paragraphs, textToFind, textToReplace, ref count, replace, stringComparison));
                }
            }

            if (this.Header?.Even != null) {
                list.AddRange(FindAndReplaceNested(this.Header.Even.Paragraphs, textToFind, textToReplace, ref count, replace, stringComparison));
                foreach (var table in this.Header.Even.Tables) {
                    list.AddRange(FindAndReplaceNested(table.Paragraphs, textToFind, textToReplace, ref count, replace, stringComparison));
                }
            }

            if (this.Header?.First != null) {
                list.AddRange(FindAndReplaceNested(this.Header.First.Paragraphs, textToFind, textToReplace, ref count, replace, stringComparison));
                foreach (var table in this.Header.First.Tables) {
                    list.AddRange(FindAndReplaceNested(table.Paragraphs, textToFind, textToReplace, ref count, replace, stringComparison));
                }
            }

            if (this.Footer?.Default != null) {
                list.AddRange(FindAndReplaceNested(this.Footer.Default.Paragraphs, textToFind, textToReplace, ref count, replace, stringComparison));
                foreach (var table in this.Footer.Default.Tables) {
                    list.AddRange(FindAndReplaceNested(table.Paragraphs, textToFind, textToReplace, ref count, replace, stringComparison));
                }
            }

            if (this.Footer?.Even != null) {
                list.AddRange(FindAndReplaceNested(this.Footer.Even.Paragraphs, textToFind, textToReplace, ref count, replace, stringComparison));
                foreach (var table in this.Footer.Even.Tables) {
                    list.AddRange(FindAndReplaceNested(table.Paragraphs, textToFind, textToReplace, ref count, replace, stringComparison));
                }
            }

            if (this.Footer?.First != null) {
                list.AddRange(FindAndReplaceNested(this.Footer.First.Paragraphs, textToFind, textToReplace, ref count, replace, stringComparison));
                foreach (var table in this.Footer.First.Tables) {
                    list.AddRange(FindAndReplaceNested(table.Paragraphs, textToFind, textToReplace, ref count, replace, stringComparison));
                }
            }

            return list;
        }

        private static bool IsSegmentValid(List<WordParagraph> paragraphs, WordTextSegment ts) {
            if (paragraphs is null || ts is null) {
                return false;
            }

            if (ts.BeginIndex < 0 || ts.EndIndex < ts.BeginIndex || ts.EndIndex >= paragraphs.Count) {
                return false;
            }

            var beginPara = paragraphs[ts.BeginIndex];
            var endPara = paragraphs[ts.EndIndex];

            if (beginPara is null || endPara is null) {
                return false;
            }

            if (ts.BeginChar < 0 || ts.BeginChar >= beginPara.Text.Length) {
                return false;
            }

            if (ts.EndChar < 0 || ts.EndChar >= endPara.Text.Length) {
                return false;
            }

            return true;
        }
    }
}
