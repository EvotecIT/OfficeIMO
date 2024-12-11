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
    /// This provides a way to set items to be treated with heading style during load
    /// </summary>
    public bool IsToc {
        get {
            return ListItems
                .Select(paragraph => paragraph.Style.ToString())
                .Any(style => style.StartsWith("Heading", StringComparison.Ordinal));
        }
    }

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
    /// Exposes the numbering properties of the list allowing for customizations of lists
    /// </summary>
    /// <value>
    /// The numbering.
    /// </value>
    public WordListNumbering Numbering {
        get {
            var abstractNum = _document._wordprocessingDocument.MainDocumentPart!.NumberingDefinitionsPart!.Numbering
                .ChildElements.OfType<AbstractNum>()
                .FirstOrDefault(a => a.AbstractNumberId == _abstractId);
            return new WordListNumbering(abstractNum);
        }
    }

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

    public bool Strike {
        get => GetNumberingProperty(props => props.Elements<Strike>().Any(), false);
        set => SetNumberingProperty(props => {
            props.RemoveAllChildren<Strike>();
            if (value) {
                props.Append(new Strike());
            }
        }, value);
    }

    public bool DoubleStrike {
        get => GetNumberingProperty(props => props.Elements<DoubleStrike>().Any(), false);
        set => SetNumberingProperty(props => {
            props.RemoveAllChildren<DoubleStrike>();
            if (value) {
                props.Append(new DoubleStrike());
            }
        }, value);
    }

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

    public WordList(WordDocument wordDocument, bool isToc = false) {
        _document = wordDocument;
        _wordprocessingDocument = wordDocument._wordprocessingDocument;
        //_section = section;
        _isToc = isToc;
        // section.Lists.Add(this);
    }

    public WordList(WordDocument wordDocument, WordParagraph paragraph, bool isToc = false) {
        _document = wordDocument;
        _wordprocessingDocument = wordDocument._wordprocessingDocument;
        //_section = section;
        _isToc = isToc;
        _wordParagraph = paragraph;
        // section.Lists.Add(this);
    }


    public WordList(WordDocument wordDocument, int numberId) {
        _document = wordDocument;
        _wordprocessingDocument = wordDocument._wordprocessingDocument;
        //  _section = section;
        _numberId = numberId;
    }

    public WordList(WordDocument wordDocument, WordHeaderFooter headerFooter) {
        _document = wordDocument;
        _wordprocessingDocument = wordDocument._wordprocessingDocument;
        _headerFooter = headerFooter;
    }

    public WordParagraph AddItem(WordParagraph wordParagraph, int level = 0) {
        return AddItem(null, level, wordParagraph);
    }

    public WordParagraph AddItem(string text, int level = 0, WordParagraph wordParagraph = null) {
        if (wordParagraph != null) {
            wordParagraph._paragraphProperties.Append(new ParagraphStyleId { Val = "ListParagraph" });
            wordParagraph._paragraphProperties.Append(
                new NumberingProperties(
                    new NumberingLevelReference { Val = level },
                    new NumberingId { Val = _numberId }
                ));
            if (text != null) {
                wordParagraph.Text = text;
            }
        } else {
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

            if (_wordParagraph != null) {

                if (this.ListItems.Count > 0) {
                    var lastItem = this.ListItems.Last();
                    var allElements = lastItem._paragraph.Parent.ChildElements.OfType<Paragraph>();
                    if (allElements.Count() > 0) {
                        var lastParagraph = allElements.Last();
                        lastParagraph.Parent.Append(paragraph);
                    }
                } else {
                    var allElements = _wordParagraph._paragraph.Parent.ChildElements.OfType<Paragraph>();
                    var lastElement = allElements.Last();
                    lastElement.Parent.Append(paragraph);
                }

                // _wordParagraph._paragraph.Append(paragraph);
            } else {
                if (this.ListItems.Count > 0) {
                    var lastItem = this.ListItems.Last();
                    var allElementsAfter = lastItem._paragraph.ElementsAfter();
                    if (allElementsAfter.Count() > 0) {
                        var lastParagraph = allElementsAfter.Last();
                        lastParagraph.InsertAfterSelf(paragraph);
                    } else {
                        lastItem._paragraph.InsertAfterSelf(paragraph);
                    }
                } else {
                    if (_headerFooter != null && _headerFooter._header != null) {
                        _headerFooter._header.Append(paragraph);
                    } else if (_headerFooter != null && _headerFooter._footer != null) {
                        _headerFooter._footer.Append(paragraph);
                    } else {
                        _wordprocessingDocument.MainDocumentPart!.Document.Body!.AppendChild(paragraph);
                    }
                }
            }
            wordParagraph = new WordParagraph(_document, paragraph, run) {
                Text = text
            };
        }

        // this simplifies TOC for user usage
        if (_isToc || IsToc) {
            wordParagraph.Style = WordParagraphStyle.GetStyle(level);
        }

        if (_wordParagraph == null) {
            _wordParagraph = wordParagraph;
        }

        return wordParagraph;
    }
    
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
