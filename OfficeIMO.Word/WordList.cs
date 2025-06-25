using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using System.Linq;
using System.Collections.Generic;

namespace OfficeIMO.Word;

/// <summary>
/// Represents a collection of paragraphs formatted as a list and
/// exposes methods to manipulate the list's numbering and style.
/// </summary>
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

    private readonly List<WordParagraph> _listItems = new();

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
    public List<WordParagraph> ListItems {
        get {
            return _listItems;
        }
    }

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
            EnsureW15Namespace(numbering);
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
            return Helpers.ParseColor(ColorHex);
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
                        Val = value.Replace("#", "").ToLowerInvariant()
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
    /// Gets the list style or <see cref="WordListStyle.Custom"/> when the list does not match a built-in style.
    /// </summary>
    public WordListStyle Style => WordListStyles.MatchStyle(GetAbstractNum());

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
        foreach (var p in _document.EnumerateAllParagraphs().Where(p => p.IsListItem && p._listNumberId == _numberId)) {
            p._list = this;
            _listItems.Add(p);
        }
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

}
