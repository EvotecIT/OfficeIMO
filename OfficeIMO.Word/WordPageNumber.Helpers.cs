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
    /// <summary>
    /// Appends text to the last paragraph of the page number.
    /// </summary>
    /// <param name="text">Text to append.</param>
    /// <returns>The paragraph that received the text.</returns>
    public WordParagraph AppendText(string text) {
        if (string.IsNullOrEmpty(text)) {
            throw new ArgumentNullException(nameof(text));
        }
        var paragraph = _listParagraphs.Last();
        return paragraph.AddText(text);
    }
    public WordPageNumber(WordDocument wordDocument, WordHeader wordHeader, WordPageNumberStyle wordPageNumberStyle) {
        this._document = wordDocument;
        this._wordHeader = wordHeader;
        this._sdtBlock = GetStyle(wordPageNumberStyle);

        if (_sdtBlock != null) {
            _sdtBlock.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            _sdtBlock.AddNamespaceDeclaration("wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup");
            _sdtBlock.AddNamespaceDeclaration("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");
            try {
                _sdtBlock.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            } catch (InvalidOperationException) {
                // prefix already defined
            }

            _listParagraphs = WordSection.ConvertParagraphsToWordParagraphs(_document, _sdtBlock.SdtContentBlock.ChildElements.OfType<Paragraph>());
            this._wordParagraph = _listParagraphs[0];
        }
        wordHeader._header.Append(_sdtBlock);
    }
    public WordPageNumber(WordDocument wordDocument, WordFooter wordFooter, WordPageNumberStyle wordPageNumberStyle) {
        this._document = wordDocument;
        this._wordFooter = wordFooter;
        this._sdtBlock = GetStyle(wordPageNumberStyle);

        if (_sdtBlock != null) {
            _sdtBlock.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            _sdtBlock.AddNamespaceDeclaration("wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup");
            _sdtBlock.AddNamespaceDeclaration("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");
            try {
                _sdtBlock.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            } catch (InvalidOperationException) {
                // prefix already defined
            }
            //var paragraphs = _sdtBlock.SdtContentBlock.ChildElements.OfType<Paragraph>();
            //foreach (var paragraph in paragraphs) {
            //    this._wordParagraph = new WordParagraph(_document, paragraph);
            //}
            _listParagraphs = WordSection.ConvertParagraphsToWordParagraphs(_document, _sdtBlock.SdtContentBlock.ChildElements.OfType<Paragraph>());
            this._wordParagraph = _listParagraphs[0];
        }
        wordFooter._footer.Append(_sdtBlock);
    }

    private static SdtBlock GetStyle(WordPageNumberStyle style) {
        switch (style) {
            case WordPageNumberStyle.PlainNumber: return PlainNumber1;
            case WordPageNumberStyle.AccentBar: return AccentBar1;
            case WordPageNumberStyle.PageNumberXofY: return PageNumberXofY1;
            case WordPageNumberStyle.Brackets1: return Brackets1;
            case WordPageNumberStyle.Brackets2: return Brackets2;
            case WordPageNumberStyle.Dots: return Dots1;
            case WordPageNumberStyle.LargeItalics: return LargeItalics1;
            case WordPageNumberStyle.Roman: return Roman1;
            case WordPageNumberStyle.Tildes: return Tildes1;
            case WordPageNumberStyle.TwoBars: return FooterTwoBars1;
            case WordPageNumberStyle.TopLine: return TopLine1;
            case WordPageNumberStyle.Tab: return Tab1;
            case WordPageNumberStyle.ThickLine: return ThickLine1;
            //case WordPageNumberStyle.ThinLine: return ThinLine1;
            case WordPageNumberStyle.RoundedRectangle: return RoundedRectangle1;
            case WordPageNumberStyle.Circle: return Circle1;
            case WordPageNumberStyle.VeryLarge: return VeryLarge1;
            case WordPageNumberStyle.VerticalOutline1: return VerticalOutline1;
            case WordPageNumberStyle.VerticalOutline2: return VerticalOutline2;
        }
        throw new ArgumentOutOfRangeException(nameof(style));
    }
}
