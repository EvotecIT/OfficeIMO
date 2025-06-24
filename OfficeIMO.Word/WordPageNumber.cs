using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using Wp = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using Wpg = DocumentFormat.OpenXml.Office2010.Word.DrawingGroup;
using Wps = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;
using A14 = DocumentFormat.OpenXml.Office2010.Drawing;
using Wp14 = DocumentFormat.OpenXml.Office2010.Word.Drawing;
using V = DocumentFormat.OpenXml.Vml;
using Ovml = DocumentFormat.OpenXml.Vml.Office;
using Wvml = DocumentFormat.OpenXml.Vml.Wordprocessing;

namespace OfficeIMO.Word;

public partial class WordPageNumber {
    private WordDocument _document;
    private SdtBlock _sdtBlock;
    private WordHeader _wordHeader;
    private WordFooter _wordFooter;
    private WordParagraph _wordParagraph;
    private readonly List<WordParagraph> _listParagraphs;

    public JustificationValues? ParagraphAlignment {
        get {
            return this._wordParagraph.ParagraphAlignment;
        }
        set {
            this._wordParagraph.ParagraphAlignment = value;
        }
    }

    /// <summary>
    /// Gets the primary paragraph containing the page number field.
    /// </summary>
    public WordParagraph Paragraph {
        get { return _wordParagraph; }
    }

    /// <summary>
    /// Gets all paragraphs that make up the page number content.
    /// </summary>
    public IReadOnlyList<WordParagraph> Paragraphs {
        get { return _listParagraphs; }
    }

    public WordField Field {
        get { return _wordParagraph.Field; }
    }

    public int? Number {
        get {
            return int.TryParse(Field.Text, out int result) ? result : null;
        }
    }
}
