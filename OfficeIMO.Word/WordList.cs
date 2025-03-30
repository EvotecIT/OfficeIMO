using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word;

public partial class WordList : WordElement {
    private readonly WordprocessingDocument _wordprocessingDocument;
    private readonly WordDocument _document;
    // private readonly WordSection _section;
    private int _abstractId;
    internal int _numberId;

    /// <summary>
    /// This provides a way to set items to be treated with heading style
    /// </summary>
    private readonly bool _isToc;

    private WordParagraph _wordParagraph;
    private readonly WordHeaderFooter _headerFooter;

    /// <summary>
    /// Indicates whether the list is treated as a Table of Contents (TOC).
    /// </summary>
    public bool IsToc {
        get {
            return ListItems
                .Select(paragraph => paragraph.Style.ToString())
                .Any(style => style.StartsWith("Heading", StringComparison.Ordinal));
        }
    }

    /// <summary>
    /// Gets all the list items associated with this WordList.
    /// </summary>
    //private string NsidId {
    //    get {
    //        if (AbstractNum == null) {
    //            return null;
    //        }

    //        return AbstractNum.Nsid.Val;

    //    }
    //    set {
    //        if (AbstractNum != null) {
    //            AbstractNum.Nsid.Val = value;
    //        }
    //    }
    //}

    //private string GenerateNsidId() {
    //    // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.nsid?view=openxml-2.8.1
    //    // Specifies a number value specified as a four digit hexadecimal number),
    //    // whose contents of this decimal number are interpreted based on the context of the parent XML element.
    //    // for example FFFFFF89 or D9842532
    //    return Guid.NewGuid().ToString().ToUpper().Substring(0, 8);

    //}

    //private AbstractNum AbstractNum {
    //    get {
    //        var numbering = _document._wordprocessingDocument.MainDocumentPart!.NumberingDefinitionsPart!.Numbering;
    //        var abstractNumList = numbering.ChildElements.OfType<AbstractNum>();
    //        foreach (var abstractNum in abstractNumList) {
    //            if (abstractNum.AbstractNumberId == _abstractId) {
    //                return abstractNum;
    //            }
    //        }

    //        return null;
    //    }
    //}
    public List<WordParagraph> ListItems {
        get {
            List<WordParagraph> list = new List<WordParagraph>();
            foreach (var paragraph in _document.Paragraphs) {
                if (paragraph.IsListItem == true && paragraph._listNumberId == _numberId) {
                    list.Add(paragraph);
                }
            }

            foreach (var table in _document.Tables) {
                foreach (var paragraph in table.Paragraphs) {
                    if (paragraph.IsListItem == true && paragraph._listNumberId == _numberId) {
                        list.Add(paragraph);
                    }
                }
            }

            if (_document.Header.Default != null) {
                foreach (var paragraph in _document.Header.Default.Paragraphs) {
                    if (paragraph.IsListItem == true && paragraph._listNumberId == _numberId) {
                        list.Add(paragraph);
                    }
                }
                foreach (var table in _document.Header.Default.Tables) {
                    foreach (var paragraph in table.Paragraphs) {
                        if (paragraph.IsListItem == true && paragraph._listNumberId == _numberId) {
                            list.Add(paragraph);
                        }
                    }
                }
            }

            if (_document.Header.Even != null) {
                foreach (var paragraph in _document.Header.Even.Paragraphs) {
                    if (paragraph.IsListItem == true && paragraph._listNumberId == _numberId) {
                        list.Add(paragraph);
                    }
                }
                foreach (var table in _document.Header.Even.Tables) {
                    foreach (var paragraph in table.Paragraphs) {
                        if (paragraph.IsListItem == true && paragraph._listNumberId == _numberId) {
                            list.Add(paragraph);
                        }
                    }
                }
            }

            if (_document.Header.First != null) {
                foreach (var paragraph in _document.Header.First.Paragraphs) {
                    if (paragraph.IsListItem == true && paragraph._listNumberId == _numberId) {
                        list.Add(paragraph);
                    }
                }
                foreach (var table in _document.Header.First.Tables) {
                    foreach (var paragraph in table.Paragraphs) {
                        if (paragraph.IsListItem == true && paragraph._listNumberId == _numberId) {
                            list.Add(paragraph);
                        }
                    }
                }
            }


            if (_document.Footer.Default != null) {
                foreach (var paragraph in _document.Footer.Default.Paragraphs) {
                    if (paragraph.IsListItem == true && paragraph._listNumberId == _numberId) {
                        list.Add(paragraph);
                    }
                }
                foreach (var table in _document.Footer.Default.Tables) {
                    foreach (var paragraph in table.Paragraphs) {
                        if (paragraph.IsListItem == true && paragraph._listNumberId == _numberId) {
                            list.Add(paragraph);
                        }
                    }
                }
            }

            if (_document.Footer.Even != null) {
                foreach (var paragraph in _document.Footer.Even.Paragraphs) {
                    if (paragraph.IsListItem == true && paragraph._listNumberId == _numberId) {
                        list.Add(paragraph);
                    }
                }
                foreach (var table in _document.Footer.Even.Tables) {
                    foreach (var paragraph in table.Paragraphs) {
                        if (paragraph.IsListItem == true && paragraph._listNumberId == _numberId) {
                            list.Add(paragraph);
                        }
                    }
                }
            }

            if (_document.Footer.First != null) {
                foreach (var paragraph in _document.Footer.First.Paragraphs) {
                    if (paragraph.IsListItem == true && paragraph._listNumberId == _numberId) {
                        list.Add(paragraph);
                    }
                }
                foreach (var table in _document.Footer.First.Tables) {
                    foreach (var paragraph in table.Paragraphs) {
                        if (paragraph.IsListItem == true && paragraph._listNumberId == _numberId) {
                            list.Add(paragraph);
                        }
                    }
                }
            }
            return list;


            //if (_wordParagraph != null) {
            //    var list = new List<Paragraph>();
            //    var parent = _wordParagraph._paragraph.Parent;
            //    var elementsAfter = parent.ChildElements.OfType<Paragraph>();
            //    foreach (var element in elementsAfter) {
            //        if (element.ParagraphProperties != null && element.ParagraphProperties.NumberingProperties != null) {
            //            if (element.ParagraphProperties.NumberingProperties.NumberingId.Val == _numberId) {
            //                list.Add(element);
            //            }
            //        }
            //    }
            //    var listWord = WordSection.ConvertParagraphsToWordParagraphs(_document, list);
            //    return listWord;
            //} else {
            //    return new List<WordParagraph>();
            //}
            //elementsAfter.Where(paragraph => paragraph.IsListItem && paragraph._listNumberId == _numberId).ToList();
            //return _document.Paragraphs
            //    .Where(paragraph => paragraph.IsListItem && paragraph._listNumberId == _numberId)
            //    .ToList();
        }
    }

    /// <summary>
    /// Restarts numbering of a list after a break. Requires a list to be set to RestartNumbering overall.
    /// </summary>
    public bool RestartNumberingAfterBreak {
        get {
            var numbering = _document._wordprocessingDocument.MainDocumentPart!.NumberingDefinitionsPart!.Numbering;
            var listAbstracts = numbering.ChildElements.OfType<AbstractNum>();
            foreach (var abstractInstance in listAbstracts) {
                if (abstractInstance.AbstractNumberId == _abstractId) {
                    var currentValue = abstractInstance.GetAttribute("restartNumberingAfterBreak", "http://schemas.microsoft.com/office/word/2012/wordml");
                    return currentValue.Value != "0";
                }
            }
            return false;
        }
        set {
            var numbering = _document._wordprocessingDocument.MainDocumentPart!.NumberingDefinitionsPart!.Numbering;
            var listAbstracts = numbering.ChildElements.OfType<AbstractNum>();
            foreach (var abstractInstance in listAbstracts) {
                if (abstractInstance.AbstractNumberId == _abstractId) {
                    var setValue = value ? "1" : "0";
                    abstractInstance.SetAttribute(new OpenXmlAttribute("w15", "restartNumberingAfterBreak", "http://schemas.microsoft.com/office/word/2012/wordml", setValue));
                }
            }
        }
    }

    /// <summary>
    /// Exposes the numbering properties of the list, allowing for customization.
    /// </summary>
    public WordListNumbering Numbering {
        get {
            var abstractNum = _document._wordprocessingDocument.MainDocumentPart!.NumberingDefinitionsPart!.Numbering
                .ChildElements.OfType<AbstractNum>()
                .FirstOrDefault(a => a.AbstractNumberId == _abstractId);
            return new WordListNumbering(abstractNum);
        }
    }

    /// <summary>
    /// Gets or sets a value indicating whether the numbering symbols are bold.
    /// </summary>
    public bool Bold {
        get { return GetNumberingProperty<bool>(props => props.Elements<Bold>().Any(), false); }
        set {
            SetNumberingProperty(props => {
                props.RemoveAllChildren<Bold>();
                props.RemoveAllChildren<BoldComplexScript>();
                if (value) {
                    props.Append(new Bold());
                    props.Append(new BoldComplexScript());
                }
            }, value);
        }
    }

    /// <summary>
    /// Gets or sets the font size of the numbering symbols in points.
    /// </summary>
    public int? FontSize {
        get {
            return GetNumberingProperty<int?>(props => {
                var fontSize = props.Elements<FontSize>().FirstOrDefault();
                // Convert from half-points to points
                return fontSize?.Val != null ? int.Parse(fontSize.Val) / 2 : null;
            });
        }
        set {
            SetNumberingProperty(props => {
                props.RemoveAllChildren<FontSize>();
                props.RemoveAllChildren<FontSizeComplexScript>();
                if (value.HasValue) {
                    // Convert from points to half-points
                    var halfPoints = (value.Value * 2).ToString();
                    props.Append(new FontSize { Val = halfPoints });
                    props.Append(new FontSizeComplexScript { Val = halfPoints });
                }
            }, value.HasValue);
        }
    }

    /// <summary>
    /// Gets or sets the color of the numbering symbols.
    /// </summary>
    public SixLabors.ImageSharp.Color? Color {
        get {
            if (ColorHex == "") {
                return null;
            }
            return SixLabors.ImageSharp.Color.Parse("#" + ColorHex);
        }
        set {
            if (value != null) {
                this.ColorHex = value.Value.ToHexColor();
            } else {
                this.ColorHex = "";
            }
        }
    }

    /// <summary>
    /// Gets or sets the hexadecimal color value of the numbering symbols.
    /// </summary>
    public string ColorHex {
        get {
            return GetNumberingProperty<string>(props => {
                var color = props.Elements<DocumentFormat.OpenXml.Wordprocessing.Color>().FirstOrDefault();
                return color?.Val ?? "";
            });
        }
        set {
            SetNumberingProperty(props => {
                props.RemoveAllChildren<DocumentFormat.OpenXml.Wordprocessing.Color>();
                if (!string.IsNullOrEmpty(value)) {
                    props.Append(new DocumentFormat.OpenXml.Wordprocessing.Color {
                        Val = value.Replace("#", "")
                    });
                }
            }, !string.IsNullOrEmpty(value));
        }
    }

    /// <summary>
    /// Gets or sets a value indicating whether the numbering symbols are italicized.
    /// </summary>
    public bool Italic {
        get => GetNumberingProperty(props => props.Elements<Italic>().Any(), false);
        set => SetNumberingProperty(props => {
            props.RemoveAllChildren<Italic>();
            props.RemoveAllChildren<ItalicComplexScript>();
            if (value) {
                props.Append(new Italic());
                props.Append(new ItalicComplexScript());
            }
        }, value);
    }

    /// <summary>
    /// Gets or sets the underline style of the numbering symbols.
    /// </summary>
    public UnderlineValues? Underline {
        get => GetNumberingProperty<UnderlineValues?>(props =>
            props.Elements<Underline>().FirstOrDefault()?.Val);
        set => SetNumberingProperty(props => {
            props.RemoveAllChildren<Underline>();
            if (value.HasValue) {
                props.Append(new Underline { Val = value.Value });
            }
        }, value.HasValue);
    }

    /// <summary>
    /// Gets or sets a value indicating whether the numbering symbols have a strikethrough.
    /// </summary>
    public bool Strike {
        get => GetNumberingProperty(props => props.Elements<Strike>().Any(), false);
        set => SetNumberingProperty(props => {
            props.RemoveAllChildren<Strike>();
            if (value) {
                props.Append(new Strike());
            }
        }, value);
    }

    /// <summary>
    /// Gets or sets a value indicating whether the numbering symbols have a double strikethrough.
    /// </summary>
    public bool DoubleStrike {
        get => GetNumberingProperty(props => props.Elements<DoubleStrike>().Any(), false);
        set => SetNumberingProperty(props => {
            props.RemoveAllChildren<DoubleStrike>();
            if (value) {
                props.Append(new DoubleStrike());
            }
        }, value);
    }

    /// <summary>
    /// Gets or sets the font name of the numbering symbols.
    /// </summary>
    public string FontName {
        get => GetNumberingProperty<string>(props =>
            props.Elements<RunFonts>().FirstOrDefault()?.Ascii);
        set => SetNumberingProperty(props => {
            props.RemoveAllChildren<RunFonts>();
            if (!string.IsNullOrEmpty(value)) {
                props.Append(new RunFonts { Ascii = value, HighAnsi = value });
            }
        }, !string.IsNullOrEmpty(value));
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="WordList"/> class.
    /// </summary>
    /// <param name="wordDocument">The Word document.</param>
    /// <param name="isToc">Indicates if the list should be treated as a TOC.</param>
    public WordList(WordDocument wordDocument, bool isToc = false) {
        _document = wordDocument;
        _wordprocessingDocument = wordDocument._wordprocessingDocument;
        //_section = section;
        _isToc = isToc;
        // section.Lists.Add(this);
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="WordList"/> class with a starting paragraph.
    /// </summary>
    /// <param name="wordDocument">The Word document.</param>
    /// <param name="paragraph">The starting paragraph.</param>
    /// <param name="isToc">Indicates if the list should be treated as a TOC.</param>
    public WordList(WordDocument wordDocument, WordParagraph paragraph, bool isToc = false) {
        _document = wordDocument;
        _wordprocessingDocument = wordDocument._wordprocessingDocument;
        //_section = section;
        _isToc = isToc;
        _wordParagraph = paragraph;
        // section.Lists.Add(this);
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="WordList"/> class with a specific number ID.
    /// </summary>
    /// <param name="wordDocument">The Word document.</param>
    /// <param name="numberId">The numbering ID.</param>
    public WordList(WordDocument wordDocument, int numberId) {
        _document = wordDocument;
        _wordprocessingDocument = wordDocument._wordprocessingDocument;
        //  _section = section;
        _numberId = numberId;
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="WordList"/> class for headers and footers.
    /// </summary>
    /// <param name="wordDocument">The Word document.</param>
    /// <param name="headerFooter">The header or footer.</param>
    public WordList(WordDocument wordDocument, WordHeaderFooter headerFooter) {
        _document = wordDocument;
        _wordprocessingDocument = wordDocument._wordprocessingDocument;
        _headerFooter = headerFooter;
    }

    /// <summary>
    /// Adds an item to the list using an existing paragraph.
    /// </summary>
    /// <param name="wordParagraph">The paragraph to add.</param>
    /// <param name="level">The list level.</param>
    /// <returns>The added <see cref="WordParagraph"/>.</returns>
    public WordParagraph AddItem(WordParagraph wordParagraph, int level = 0) {
        return AddItem(null, level, wordParagraph);
    }

    /// <summary>
    /// Adds an item to the list with specified text.
    /// </summary>
    /// <param name="text">The text of the list item.</param>
    /// <param name="level">The list level.</param>
    /// <param name="wordParagraph">An optional existing paragraph.</param>
    /// <returns>The added <see cref="WordParagraph"/>.</returns>
    public WordParagraph AddItem(string text, int level = 0, WordParagraph wordParagraph = null) {
        var paragraph = new Paragraph();
        var run = new Run();
        run.Append(new RunProperties());
        run.Append(new Text { Space = SpaceProcessingModeValues.Preserve });

        var paragraphProperties = new ParagraphProperties();
        paragraphProperties.Append(new ParagraphStyleId { Val = "ListParagraph" });
        paragraphProperties.Append(
            new NumberingProperties(
                new NumberingLevelReference { Val = level },
                new NumberingId { Val = _numberId }
            ));
        paragraph.Append(paragraphProperties);
        paragraph.Append(run);

        // Determine proper placement for the paragraph
        if (wordParagraph != null) {
            // If a specific paragraph reference is provided, insert after it
            wordParagraph._paragraph.InsertAfterSelf(paragraph);
        } else if (this.ListItems.Count == 0 && _wordParagraph != null) {
            // First item in a paragraph-referenced list - insert after reference paragraph
            _wordParagraph._paragraph.InsertAfterSelf(paragraph);
        } else if (_isToc || IsToc) {
            // TOC list items should be placed at the end of the document
            _wordprocessingDocument.MainDocumentPart!.Document.Body!.AppendChild(paragraph);
        } else if (_headerFooter != null) {
            // Header/footer list items
            if (_headerFooter._header != null) {
                _headerFooter._header.Append(paragraph);
            } else if (_headerFooter._footer != null) {
                _headerFooter._footer.Append(paragraph);
            }
        } else if (_wordParagraph != null && _wordParagraph._paragraph.Parent is TableCell) {
            // Handle table cell lists
            var parent = _wordParagraph._paragraph.Parent;
            if (this.ListItems.Count > 0) {
                var lastItem = this.ListItems.Last();
                lastItem._paragraph.InsertAfterSelf(paragraph);
            } else {
                parent.Append(paragraph);
            }
        } else {
            // For standard lists without specific placement, add at the end
            _wordprocessingDocument.MainDocumentPart!.Document.Body!.AppendChild(paragraph);
        }

        var newParagraph = new WordParagraph(_document, paragraph, run);
        if (text != null) {
            newParagraph.Text = text;
        }

        // Handle TOC styling
        if (_isToc || IsToc) {
            newParagraph.Style = WordParagraphStyle.GetStyle(level);
        }

        return newParagraph;
    }

    /// <summary>
    /// Removes the list and its items from the document.
    /// </summary>
    public void Remove() {
        // Get the Numbering part from the document
        var numbering = _document._wordprocessingDocument.MainDocumentPart.NumberingDefinitionsPart.Numbering;

        // Find and remove the AbstractNum associated with this list
        var abstractNum = numbering.Elements<AbstractNum>().FirstOrDefault(a => a.AbstractNumberId.Value == _abstractId);
        if (abstractNum != null) {
            numbering.RemoveChild(abstractNum);
        }

        // Find and remove the NumberingInstance associated with this list
        var numberingInstance = numbering.Elements<NumberingInstance>().FirstOrDefault(n => n.NumberID.Value == _numberId);
        if (numberingInstance != null) {
            numbering.RemoveChild(numberingInstance);
        }

        // Remove the list items from the document
        foreach (var listItem in ListItems) {
            listItem.Remove();
        }
    }

    /// <summary>
    /// Merges another list into this list.
    /// </summary>
    /// <param name="documentList">The list to merge.</param>
    public void Merge(WordList documentList) {
        // Reattach all items from the other list to this list
        foreach (var item in documentList.ListItems) {
            var numberingProperties = item._paragraphProperties.NumberingProperties;
            // Change the NumId to the NumId of this list
            if (numberingProperties != null && numberingProperties.NumberingId != null) {
                numberingProperties.NumberingId.Val = this._numberId;
            }
        }
        // Remove the other list
        documentList.Remove();
    }
}
