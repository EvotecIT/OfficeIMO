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

namespace OfficeIMO {
    public enum WordPageNumberStyle {
        PlainNumber,
        AccentBar,
        PageNumberXofY,
        Brackets1,
        /// <summary>
        /// Due to way this page number style is built the location is always header (center) regardless of placement in document
        /// </summary>
        Brackets2,
        Dots,
        LargeItalics,
        Roman,
        Tildes,
        /// <summary>
        /// Due to way this page number style is built the location is always footer regardless of placement in document
        /// </summary>
        TwoBars,
        TopLine,
        Tab,
        ThickLine,
        //ThinLine,
        RoundedRectangle,
        /// <summary>
        /// Due to way this page number style is built the location is always header (center) regardless of placement in document
        /// </summary>
        Circle,
        /// <summary>
        /// Due to way this page number style is built the location is always header (right) regardless of placement in document
        /// </summary>
        VeryLarge,
        /// <summary>
        /// Due to way this page number style is built the location is always header (left) regardless of placement in document
        /// </summary>
        VerticalOutline1,
        /// <summary>
        /// Due to way this page number style is built the location is always header (right) regardless of placement in document
        /// </summary>
        VerticalOutline2
    }
    public class WordPageNumber {
        private WordDocument _document;
        private SdtBlock _sdtBlock;
        private WordHeader _wordHeader;
        private WordFooter _wordFooter;
        private WordParagraph _wordParagraph;

        public JustificationValues ParagraphAlignment {
            get {
                return this._wordParagraph.ParagraphAlignment;
            }
            set {
                this._wordParagraph.ParagraphAlignment = value;
            }
        }
        public WordPageNumber(WordDocument wordDocument, WordHeader wordHeader, WordPageNumberStyle wordPageNumberStyle) {
            this._document = wordDocument;
            this._wordHeader = wordHeader;
            this._sdtBlock = GetStyle(wordPageNumberStyle);

            if (_sdtBlock != null) {
                var paragraphs = _sdtBlock.SdtContentBlock.ChildElements.OfType<Paragraph>();
                foreach (var paragraph in paragraphs) {
                   this._wordParagraph = new WordParagraph(_document, paragraph);
                }
            }
            wordHeader._header.Append(_sdtBlock);
        }
        public WordPageNumber(WordDocument wordDocument, WordFooter wordFooter, WordPageNumberStyle wordPageNumberStyle) {
            this._document = wordDocument;
            this._wordFooter = wordFooter;
            this._sdtBlock = GetStyle(wordPageNumberStyle);

            if (_sdtBlock != null) {
                var paragraphs = _sdtBlock.SdtContentBlock.ChildElements.OfType<Paragraph>();
                foreach (var paragraph in paragraphs) {
                    this._wordParagraph = new WordParagraph(_document, paragraph);
                }
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
        private static SdtBlock PlainNumber1 {
            get {
                SdtBlock sdtBlock1 = new SdtBlock();

                SdtProperties sdtProperties1 = new SdtProperties();
                SdtId sdtId1 = new SdtId() { Val = -94168831 };

                SdtContentDocPartObject sdtContentDocPartObject1 = new SdtContentDocPartObject();
                DocPartGallery docPartGallery1 = new DocPartGallery() { Val = "Page Numbers (Top of Page)" };
                DocPartUnique docPartUnique1 = new DocPartUnique();

                sdtContentDocPartObject1.Append(docPartGallery1);
                sdtContentDocPartObject1.Append(docPartUnique1);

                sdtProperties1.Append(sdtId1);
                sdtProperties1.Append(sdtContentDocPartObject1);

                SdtEndCharProperties sdtEndCharProperties1 = new SdtEndCharProperties();

                RunProperties runProperties1 = new RunProperties();
                NoProof noProof1 = new NoProof();

                runProperties1.Append(noProof1);

                sdtEndCharProperties1.Append(runProperties1);

                SdtContentBlock sdtContentBlock1 = new SdtContentBlock();

                Paragraph paragraph1 = new Paragraph() { RsidParagraphAddition = "000003E1", RsidRunAdditionDefault = "000003E1", ParagraphId = "2D5F018E", TextId = "0961F169" };

                ParagraphProperties paragraphProperties1 = new ParagraphProperties();
                ParagraphStyleId paragraphStyleId1 = new ParagraphStyleId() { Val = "Header" };

                paragraphProperties1.Append(paragraphStyleId1);

                Run run1 = new Run();
                FieldChar fieldChar1 = new FieldChar() { FieldCharType = FieldCharValues.Begin };

                run1.Append(fieldChar1);

                Run run2 = new Run();
                FieldCode fieldCode1 = new FieldCode() { Space = SpaceProcessingModeValues.Preserve };
                fieldCode1.Text = " PAGE   \\* MERGEFORMAT ";

                run2.Append(fieldCode1);

                Run run3 = new Run();
                FieldChar fieldChar2 = new FieldChar() { FieldCharType = FieldCharValues.Separate };

                run3.Append(fieldChar2);

                Run run4 = new Run();

                RunProperties runProperties2 = new RunProperties();
                NoProof noProof2 = new NoProof();

                runProperties2.Append(noProof2);
                Text text1 = new Text();
                text1.Text = "2";

                run4.Append(runProperties2);
                run4.Append(text1);

                Run run5 = new Run();

                RunProperties runProperties3 = new RunProperties();
                NoProof noProof3 = new NoProof();

                runProperties3.Append(noProof3);
                FieldChar fieldChar3 = new FieldChar() { FieldCharType = FieldCharValues.End };

                run5.Append(runProperties3);
                run5.Append(fieldChar3);

                paragraph1.Append(paragraphProperties1);
                paragraph1.Append(run1);
                paragraph1.Append(run2);
                paragraph1.Append(run3);
                paragraph1.Append(run4);
                paragraph1.Append(run5);

                sdtContentBlock1.Append(paragraph1);

                sdtBlock1.Append(sdtProperties1);
                sdtBlock1.Append(sdtEndCharProperties1);
                sdtBlock1.Append(sdtContentBlock1);
                return sdtBlock1;


            }
        }
        private static SdtBlock AccentBar1 {
            get {
                SdtBlock sdtBlock1 = new SdtBlock();

                SdtProperties sdtProperties1 = new SdtProperties();
                SdtId sdtId1 = new SdtId() { Val = -665018933 };

                SdtContentDocPartObject sdtContentDocPartObject1 = new SdtContentDocPartObject();
                DocPartGallery docPartGallery1 = new DocPartGallery() { Val = "Page Numbers (Top of Page)" };
                DocPartUnique docPartUnique1 = new DocPartUnique();

                sdtContentDocPartObject1.Append(docPartGallery1);
                sdtContentDocPartObject1.Append(docPartUnique1);

                sdtProperties1.Append(sdtId1);
                sdtProperties1.Append(sdtContentDocPartObject1);

                SdtEndCharProperties sdtEndCharProperties1 = new SdtEndCharProperties();

                RunProperties runProperties1 = new RunProperties();
                Color color1 = new Color() { Val = "7F7F7F", ThemeColor = ThemeColorValues.Background1, ThemeShade = "7F" };
                Spacing spacing1 = new Spacing() { Val = 60 };

                runProperties1.Append(color1);
                runProperties1.Append(spacing1);

                sdtEndCharProperties1.Append(runProperties1);

                SdtContentBlock sdtContentBlock1 = new SdtContentBlock();

                Paragraph paragraph1 = new Paragraph() { RsidParagraphAddition = "004D15AC", RsidRunAdditionDefault = "004D15AC", ParagraphId = "7C481A12", TextId = "7731C022" };

                ParagraphProperties paragraphProperties1 = new ParagraphProperties();
                ParagraphStyleId paragraphStyleId1 = new ParagraphStyleId() { Val = "Header" };

                ParagraphBorders paragraphBorders1 = new ParagraphBorders();
                BottomBorder bottomBorder1 = new BottomBorder() { Val = BorderValues.Single, Color = "D9D9D9", ThemeColor = ThemeColorValues.Background1, ThemeShade = "D9", Size = (UInt32Value)4U, Space = (UInt32Value)1U };

                paragraphBorders1.Append(bottomBorder1);

                ParagraphMarkRunProperties paragraphMarkRunProperties1 = new ParagraphMarkRunProperties();
                Bold bold1 = new Bold();
                BoldComplexScript boldComplexScript1 = new BoldComplexScript();

                paragraphMarkRunProperties1.Append(bold1);
                paragraphMarkRunProperties1.Append(boldComplexScript1);

                paragraphProperties1.Append(paragraphStyleId1);
                paragraphProperties1.Append(paragraphBorders1);
                paragraphProperties1.Append(paragraphMarkRunProperties1);

                Run run1 = new Run();
                FieldChar fieldChar1 = new FieldChar() { FieldCharType = FieldCharValues.Begin };

                run1.Append(fieldChar1);

                Run run2 = new Run();
                FieldCode fieldCode1 = new FieldCode() { Space = SpaceProcessingModeValues.Preserve };
                fieldCode1.Text = " PAGE   \\* MERGEFORMAT ";

                run2.Append(fieldCode1);

                Run run3 = new Run();
                FieldChar fieldChar2 = new FieldChar() { FieldCharType = FieldCharValues.Separate };

                run3.Append(fieldChar2);

                Run run4 = new Run();

                RunProperties runProperties2 = new RunProperties();
                Bold bold2 = new Bold();
                BoldComplexScript boldComplexScript2 = new BoldComplexScript();
                NoProof noProof1 = new NoProof();

                runProperties2.Append(bold2);
                runProperties2.Append(boldComplexScript2);
                runProperties2.Append(noProof1);
                Text text1 = new Text();
                text1.Text = "2";

                run4.Append(runProperties2);
                run4.Append(text1);

                Run run5 = new Run();

                RunProperties runProperties3 = new RunProperties();
                Bold bold3 = new Bold();
                BoldComplexScript boldComplexScript3 = new BoldComplexScript();
                NoProof noProof2 = new NoProof();

                runProperties3.Append(bold3);
                runProperties3.Append(boldComplexScript3);
                runProperties3.Append(noProof2);
                FieldChar fieldChar3 = new FieldChar() { FieldCharType = FieldCharValues.End };

                run5.Append(runProperties3);
                run5.Append(fieldChar3);

                Run run6 = new Run();

                RunProperties runProperties4 = new RunProperties();
                Bold bold4 = new Bold();
                BoldComplexScript boldComplexScript4 = new BoldComplexScript();

                runProperties4.Append(bold4);
                runProperties4.Append(boldComplexScript4);
                Text text2 = new Text() { Space = SpaceProcessingModeValues.Preserve };
                text2.Text = " | ";

                run6.Append(runProperties4);
                run6.Append(text2);

                Run run7 = new Run();

                RunProperties runProperties5 = new RunProperties();
                Color color2 = new Color() { Val = "7F7F7F", ThemeColor = ThemeColorValues.Background1, ThemeShade = "7F" };
                Spacing spacing2 = new Spacing() { Val = 60 };

                runProperties5.Append(color2);
                runProperties5.Append(spacing2);
                Text text3 = new Text();
                text3.Text = "Page";

                run7.Append(runProperties5);
                run7.Append(text3);

                paragraph1.Append(paragraphProperties1);
                paragraph1.Append(run1);
                paragraph1.Append(run2);
                paragraph1.Append(run3);
                paragraph1.Append(run4);
                paragraph1.Append(run5);
                paragraph1.Append(run6);
                paragraph1.Append(run7);

                sdtContentBlock1.Append(paragraph1);

                sdtBlock1.Append(sdtProperties1);
                sdtBlock1.Append(sdtEndCharProperties1);
                sdtBlock1.Append(sdtContentBlock1);
                return sdtBlock1;

            }
        }
        private static SdtBlock PageNumberXofY1 {
            get {
                SdtBlock sdtBlock1 = new SdtBlock();

                SdtProperties sdtProperties1 = new SdtProperties();
                SdtId sdtId1 = new SdtId() { Val = 98381352 };

                SdtContentDocPartObject sdtContentDocPartObject1 = new SdtContentDocPartObject();
                DocPartGallery docPartGallery1 = new DocPartGallery() { Val = "Page Numbers (Top of Page)" };
                DocPartUnique docPartUnique1 = new DocPartUnique();

                sdtContentDocPartObject1.Append(docPartGallery1);
                sdtContentDocPartObject1.Append(docPartUnique1);

                sdtProperties1.Append(sdtId1);
                sdtProperties1.Append(sdtContentDocPartObject1);

                SdtContentBlock sdtContentBlock1 = new SdtContentBlock();

                Paragraph paragraph1 = new Paragraph() { RsidParagraphAddition = "00E11B1E", RsidRunAdditionDefault = "00E11B1E", ParagraphId = "1296FAD3", TextId = "77777777" };

                ParagraphProperties paragraphProperties1 = new ParagraphProperties();
                ParagraphStyleId paragraphStyleId1 = new ParagraphStyleId() { Val = "Header" };

                paragraphProperties1.Append(paragraphStyleId1);

                Run run1 = new Run();
                Text text1 = new Text() { Space = SpaceProcessingModeValues.Preserve };
                text1.Text = "Page ";

                run1.Append(text1);

                Run run2 = new Run();

                RunProperties runProperties1 = new RunProperties();
                Bold bold1 = new Bold();
                BoldComplexScript boldComplexScript1 = new BoldComplexScript();
                FontSize fontSize1 = new FontSize() { Val = "24" };
                FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "24" };

                runProperties1.Append(bold1);
                runProperties1.Append(boldComplexScript1);
                runProperties1.Append(fontSize1);
                runProperties1.Append(fontSizeComplexScript1);
                FieldChar fieldChar1 = new FieldChar() { FieldCharType = FieldCharValues.Begin };

                run2.Append(runProperties1);
                run2.Append(fieldChar1);

                Run run3 = new Run();

                RunProperties runProperties2 = new RunProperties();
                Bold bold2 = new Bold();
                BoldComplexScript boldComplexScript2 = new BoldComplexScript();

                runProperties2.Append(bold2);
                runProperties2.Append(boldComplexScript2);
                FieldCode fieldCode1 = new FieldCode() { Space = SpaceProcessingModeValues.Preserve };
                fieldCode1.Text = " PAGE ";

                run3.Append(runProperties2);
                run3.Append(fieldCode1);

                Run run4 = new Run();

                RunProperties runProperties3 = new RunProperties();
                Bold bold3 = new Bold();
                BoldComplexScript boldComplexScript3 = new BoldComplexScript();
                FontSize fontSize2 = new FontSize() { Val = "24" };
                FontSizeComplexScript fontSizeComplexScript2 = new FontSizeComplexScript() { Val = "24" };

                runProperties3.Append(bold3);
                runProperties3.Append(boldComplexScript3);
                runProperties3.Append(fontSize2);
                runProperties3.Append(fontSizeComplexScript2);
                FieldChar fieldChar2 = new FieldChar() { FieldCharType = FieldCharValues.Separate };

                run4.Append(runProperties3);
                run4.Append(fieldChar2);

                Run run5 = new Run();

                RunProperties runProperties4 = new RunProperties();
                Bold bold4 = new Bold();
                BoldComplexScript boldComplexScript4 = new BoldComplexScript();
                NoProof noProof1 = new NoProof();

                runProperties4.Append(bold4);
                runProperties4.Append(boldComplexScript4);
                runProperties4.Append(noProof1);
                Text text2 = new Text();
                text2.Text = "2";

                run5.Append(runProperties4);
                run5.Append(text2);

                Run run6 = new Run();

                RunProperties runProperties5 = new RunProperties();
                Bold bold5 = new Bold();
                BoldComplexScript boldComplexScript5 = new BoldComplexScript();
                FontSize fontSize3 = new FontSize() { Val = "24" };
                FontSizeComplexScript fontSizeComplexScript3 = new FontSizeComplexScript() { Val = "24" };

                runProperties5.Append(bold5);
                runProperties5.Append(boldComplexScript5);
                runProperties5.Append(fontSize3);
                runProperties5.Append(fontSizeComplexScript3);
                FieldChar fieldChar3 = new FieldChar() { FieldCharType = FieldCharValues.End };

                run6.Append(runProperties5);
                run6.Append(fieldChar3);

                Run run7 = new Run();
                Text text3 = new Text() { Space = SpaceProcessingModeValues.Preserve };
                text3.Text = " of ";

                run7.Append(text3);

                Run run8 = new Run();

                RunProperties runProperties6 = new RunProperties();
                Bold bold6 = new Bold();
                BoldComplexScript boldComplexScript6 = new BoldComplexScript();
                FontSize fontSize4 = new FontSize() { Val = "24" };
                FontSizeComplexScript fontSizeComplexScript4 = new FontSizeComplexScript() { Val = "24" };

                runProperties6.Append(bold6);
                runProperties6.Append(boldComplexScript6);
                runProperties6.Append(fontSize4);
                runProperties6.Append(fontSizeComplexScript4);
                FieldChar fieldChar4 = new FieldChar() { FieldCharType = FieldCharValues.Begin };

                run8.Append(runProperties6);
                run8.Append(fieldChar4);

                Run run9 = new Run();

                RunProperties runProperties7 = new RunProperties();
                Bold bold7 = new Bold();
                BoldComplexScript boldComplexScript7 = new BoldComplexScript();

                runProperties7.Append(bold7);
                runProperties7.Append(boldComplexScript7);
                FieldCode fieldCode2 = new FieldCode() { Space = SpaceProcessingModeValues.Preserve };
                fieldCode2.Text = " NUMPAGES  ";

                run9.Append(runProperties7);
                run9.Append(fieldCode2);

                Run run10 = new Run();

                RunProperties runProperties8 = new RunProperties();
                Bold bold8 = new Bold();
                BoldComplexScript boldComplexScript8 = new BoldComplexScript();
                FontSize fontSize5 = new FontSize() { Val = "24" };
                FontSizeComplexScript fontSizeComplexScript5 = new FontSizeComplexScript() { Val = "24" };

                runProperties8.Append(bold8);
                runProperties8.Append(boldComplexScript8);
                runProperties8.Append(fontSize5);
                runProperties8.Append(fontSizeComplexScript5);
                FieldChar fieldChar5 = new FieldChar() { FieldCharType = FieldCharValues.Separate };

                run10.Append(runProperties8);
                run10.Append(fieldChar5);

                Run run11 = new Run();

                RunProperties runProperties9 = new RunProperties();
                Bold bold9 = new Bold();
                BoldComplexScript boldComplexScript9 = new BoldComplexScript();
                NoProof noProof2 = new NoProof();

                runProperties9.Append(bold9);
                runProperties9.Append(boldComplexScript9);
                runProperties9.Append(noProof2);
                Text text4 = new Text();
                text4.Text = "2";

                run11.Append(runProperties9);
                run11.Append(text4);

                Run run12 = new Run();

                RunProperties runProperties10 = new RunProperties();
                Bold bold10 = new Bold();
                BoldComplexScript boldComplexScript10 = new BoldComplexScript();
                FontSize fontSize6 = new FontSize() { Val = "24" };
                FontSizeComplexScript fontSizeComplexScript6 = new FontSizeComplexScript() { Val = "24" };

                runProperties10.Append(bold10);
                runProperties10.Append(boldComplexScript10);
                runProperties10.Append(fontSize6);
                runProperties10.Append(fontSizeComplexScript6);
                FieldChar fieldChar6 = new FieldChar() { FieldCharType = FieldCharValues.End };

                run12.Append(runProperties10);
                run12.Append(fieldChar6);

                paragraph1.Append(paragraphProperties1);
                paragraph1.Append(run1);
                paragraph1.Append(run2);
                paragraph1.Append(run3);
                paragraph1.Append(run4);
                paragraph1.Append(run5);
                paragraph1.Append(run6);
                paragraph1.Append(run7);
                paragraph1.Append(run8);
                paragraph1.Append(run9);
                paragraph1.Append(run10);
                paragraph1.Append(run11);
                paragraph1.Append(run12);

                sdtContentBlock1.Append(paragraph1);

                sdtBlock1.Append(sdtProperties1);
                sdtBlock1.Append(sdtContentBlock1);
                return sdtBlock1;

            }
        }
        private static SdtBlock Brackets1 {
            get {
                SdtBlock sdtBlock1 = new SdtBlock();

                SdtProperties sdtProperties1 = new SdtProperties();
                SdtId sdtId1 = new SdtId() { Val = -1490555587 };

                SdtContentDocPartObject sdtContentDocPartObject1 = new SdtContentDocPartObject();
                DocPartGallery docPartGallery1 = new DocPartGallery() { Val = "Page Numbers (Top of Page)" };
                DocPartUnique docPartUnique1 = new DocPartUnique();

                sdtContentDocPartObject1.Append(docPartGallery1);
                sdtContentDocPartObject1.Append(docPartUnique1);

                sdtProperties1.Append(sdtId1);
                sdtProperties1.Append(sdtContentDocPartObject1);

                SdtContentBlock sdtContentBlock1 = new SdtContentBlock();

                Paragraph paragraph1 = new Paragraph() { RsidParagraphAddition = "0023554E", RsidRunAdditionDefault = "0023554E", ParagraphId = "52A08644", TextId = "39BFEC02" };

                ParagraphProperties paragraphProperties1 = new ParagraphProperties();
                ParagraphStyleId paragraphStyleId1 = new ParagraphStyleId() { Val = "Header" };
                Justification justification1 = new Justification() { Val = JustificationValues.Center };

                paragraphProperties1.Append(paragraphStyleId1);
                paragraphProperties1.Append(justification1);

                Run run1 = new Run();
                Text text1 = new Text();
                text1.Text = "[";

                run1.Append(text1);

                Run run2 = new Run();
                FieldChar fieldChar1 = new FieldChar() { FieldCharType = FieldCharValues.Begin };

                run2.Append(fieldChar1);

                Run run3 = new Run();
                FieldCode fieldCode1 = new FieldCode() { Space = SpaceProcessingModeValues.Preserve };
                fieldCode1.Text = " PAGE   \\* MERGEFORMAT ";

                run3.Append(fieldCode1);

                Run run4 = new Run();
                FieldChar fieldChar2 = new FieldChar() { FieldCharType = FieldCharValues.Separate };

                run4.Append(fieldChar2);

                Run run5 = new Run();

                RunProperties runProperties1 = new RunProperties();
                NoProof noProof1 = new NoProof();

                runProperties1.Append(noProof1);
                Text text2 = new Text();
                text2.Text = "2";

                run5.Append(runProperties1);
                run5.Append(text2);

                Run run6 = new Run();

                RunProperties runProperties2 = new RunProperties();
                NoProof noProof2 = new NoProof();

                runProperties2.Append(noProof2);
                FieldChar fieldChar3 = new FieldChar() { FieldCharType = FieldCharValues.End };

                run6.Append(runProperties2);
                run6.Append(fieldChar3);

                Run run7 = new Run();
                Text text3 = new Text();
                text3.Text = "]";

                run7.Append(text3);

                paragraph1.Append(paragraphProperties1);
                paragraph1.Append(run1);
                paragraph1.Append(run2);
                paragraph1.Append(run3);
                paragraph1.Append(run4);
                paragraph1.Append(run5);
                paragraph1.Append(run6);
                paragraph1.Append(run7);

                sdtContentBlock1.Append(paragraph1);

                sdtBlock1.Append(sdtProperties1);
                sdtBlock1.Append(sdtContentBlock1);
                return sdtBlock1;

            }
        }
        private static SdtBlock Brackets2 {
            get {
                SdtBlock sdtBlock1 = new SdtBlock();

                SdtProperties sdtProperties1 = new SdtProperties();
                SdtId sdtId1 = new SdtId() { Val = 105163093 };

                SdtContentDocPartObject sdtContentDocPartObject1 = new SdtContentDocPartObject();
                DocPartGallery docPartGallery1 = new DocPartGallery() { Val = "Page Numbers (Top of Page)" };
                DocPartUnique docPartUnique1 = new DocPartUnique();

                sdtContentDocPartObject1.Append(docPartGallery1);
                sdtContentDocPartObject1.Append(docPartUnique1);

                sdtProperties1.Append(sdtId1);
                sdtProperties1.Append(sdtContentDocPartObject1);

                SdtContentBlock sdtContentBlock1 = new SdtContentBlock();

                Paragraph paragraph1 = new Paragraph() { RsidParagraphMarkRevision = "008B30DD", RsidParagraphAddition = "00E11B1E", RsidParagraphProperties = "008B30DD", RsidRunAdditionDefault = "008B30DD", ParagraphId = "16AD3494", TextId = "30352D94" };

                ParagraphProperties paragraphProperties1 = new ParagraphProperties();
                ParagraphStyleId paragraphStyleId1 = new ParagraphStyleId() { Val = "Header" };

                paragraphProperties1.Append(paragraphStyleId1);

                Run run1 = new Run();

                RunProperties runProperties1 = new RunProperties();
                NoProof noProof1 = new NoProof();

                runProperties1.Append(noProof1);

                AlternateContent alternateContent1 = new AlternateContent();

                AlternateContentChoice alternateContentChoice1 = new AlternateContentChoice() { Requires = "wpg" };

                Drawing drawing1 = new Drawing();

                Wp.Anchor anchor1 = new Wp.Anchor() { DistanceFromTop = (UInt32Value)0U, DistanceFromBottom = (UInt32Value)0U, DistanceFromLeft = (UInt32Value)114300U, DistanceFromRight = (UInt32Value)114300U, SimplePos = false, RelativeHeight = (UInt32Value)251659264U, BehindDoc = false, Locked = false, LayoutInCell = true, AllowOverlap = true, EditId = "60C49A5A", AnchorId = "4D8B63E6" };
                Wp.SimplePosition simplePosition1 = new Wp.SimplePosition() { X = 0L, Y = 0L };

                Wp.HorizontalPosition horizontalPosition1 = new Wp.HorizontalPosition() { RelativeFrom = Wp.HorizontalRelativePositionValues.Margin };
                Wp.HorizontalAlignment horizontalAlignment1 = new Wp.HorizontalAlignment();
                horizontalAlignment1.Text = "center";

                horizontalPosition1.Append(horizontalAlignment1);

                Wp.VerticalPosition verticalPosition1 = new Wp.VerticalPosition() { RelativeFrom = Wp.VerticalRelativePositionValues.TopMargin };
                Wp.VerticalAlignment verticalAlignment1 = new Wp.VerticalAlignment();
                verticalAlignment1.Text = "center";

                verticalPosition1.Append(verticalAlignment1);
                Wp.Extent extent1 = new Wp.Extent() { Cx = 5923280L, Cy = 365760L };
                Wp.EffectExtent effectExtent1 = new Wp.EffectExtent() { LeftEdge = 9525L, TopEdge = 19050L, RightEdge = 10795L, BottomEdge = 15240L };
                Wp.WrapNone wrapNone1 = new Wp.WrapNone();
                Wp.DocProperties docProperties1 = new Wp.DocProperties() { Id = (UInt32Value)1U, Name = "Group 1" };

                Wp.NonVisualGraphicFrameDrawingProperties nonVisualGraphicFrameDrawingProperties1 = new Wp.NonVisualGraphicFrameDrawingProperties();

                A.GraphicFrameLocks graphicFrameLocks1 = new A.GraphicFrameLocks();
                graphicFrameLocks1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

                nonVisualGraphicFrameDrawingProperties1.Append(graphicFrameLocks1);

                A.Graphic graphic1 = new A.Graphic();
                graphic1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

                A.GraphicData graphicData1 = new A.GraphicData() { Uri = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" };

                Wpg.WordprocessingGroup wordprocessingGroup1 = new Wpg.WordprocessingGroup();

                Wpg.NonVisualGroupDrawingShapeProperties nonVisualGroupDrawingShapeProperties1 = new Wpg.NonVisualGroupDrawingShapeProperties();
                A.GroupShapeLocks groupShapeLocks1 = new A.GroupShapeLocks();

                nonVisualGroupDrawingShapeProperties1.Append(groupShapeLocks1);

                Wpg.GroupShapeProperties groupShapeProperties1 = new Wpg.GroupShapeProperties() { BlackWhiteMode = A.BlackWhiteModeValues.Auto };

                A.TransformGroup transformGroup1 = new A.TransformGroup();
                A.Offset offset1 = new A.Offset() { X = 0L, Y = 0L };
                A.Extents extents1 = new A.Extents() { Cx = 5923280L, Cy = 365760L };
                A.ChildOffset childOffset1 = new A.ChildOffset() { X = 1778L, Y = 533L };
                A.ChildExtents childExtents1 = new A.ChildExtents() { Cx = 8698L, Cy = 365760L };

                transformGroup1.Append(offset1);
                transformGroup1.Append(extents1);
                transformGroup1.Append(childOffset1);
                transformGroup1.Append(childExtents1);

                groupShapeProperties1.Append(transformGroup1);

                Wps.WordprocessingShape wordprocessingShape1 = new Wps.WordprocessingShape();
                Wps.NonVisualDrawingProperties nonVisualDrawingProperties1 = new Wps.NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "AutoShape 2" };

                Wps.NonVisualConnectorProperties nonVisualConnectorProperties1 = new Wps.NonVisualConnectorProperties();
                A.ConnectionShapeLocks connectionShapeLocks1 = new A.ConnectionShapeLocks() { NoChangeShapeType = true };

                nonVisualConnectorProperties1.Append(connectionShapeLocks1);

                Wps.ShapeProperties shapeProperties1 = new Wps.ShapeProperties() { BlackWhiteMode = A.BlackWhiteModeValues.Auto };

                A.Transform2D transform2D1 = new A.Transform2D();
                A.Offset offset2 = new A.Offset() { X = 1778L, Y = 183413L };
                A.Extents extents2 = new A.Extents() { Cx = 8698L, Cy = 0L };

                transform2D1.Append(offset2);
                transform2D1.Append(extents2);

                A.PresetGeometry presetGeometry1 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.StraightConnector1 };
                A.AdjustValueList adjustValueList1 = new A.AdjustValueList();

                presetGeometry1.Append(adjustValueList1);
                A.NoFill noFill1 = new A.NoFill();

                A.Outline outline1 = new A.Outline() { Width = 12700 };

                A.SolidFill solidFill1 = new A.SolidFill();
                A.RgbColorModelHex rgbColorModelHex1 = new A.RgbColorModelHex() { Val = "808080" };

                solidFill1.Append(rgbColorModelHex1);
                A.Round round1 = new A.Round();
                A.HeadEnd headEnd1 = new A.HeadEnd();
                A.TailEnd tailEnd1 = new A.TailEnd();

                outline1.Append(solidFill1);
                outline1.Append(round1);
                outline1.Append(headEnd1);
                outline1.Append(tailEnd1);

                A.ShapePropertiesExtensionList shapePropertiesExtensionList1 = new A.ShapePropertiesExtensionList();

                A.ShapePropertiesExtension shapePropertiesExtension1 = new A.ShapePropertiesExtension() { Uri = "{909E8E84-426E-40DD-AFC4-6F175D3DCCD1}" };

                A14.HiddenFillProperties hiddenFillProperties1 = new A14.HiddenFillProperties();
                hiddenFillProperties1.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");
                A.NoFill noFill2 = new A.NoFill();

                hiddenFillProperties1.Append(noFill2);

                shapePropertiesExtension1.Append(hiddenFillProperties1);

                shapePropertiesExtensionList1.Append(shapePropertiesExtension1);

                shapeProperties1.Append(transform2D1);
                shapeProperties1.Append(presetGeometry1);
                shapeProperties1.Append(noFill1);
                shapeProperties1.Append(outline1);
                shapeProperties1.Append(shapePropertiesExtensionList1);
                Wps.TextBodyProperties textBodyProperties1 = new Wps.TextBodyProperties();

                wordprocessingShape1.Append(nonVisualDrawingProperties1);
                wordprocessingShape1.Append(nonVisualConnectorProperties1);
                wordprocessingShape1.Append(shapeProperties1);
                wordprocessingShape1.Append(textBodyProperties1);

                Wps.WordprocessingShape wordprocessingShape2 = new Wps.WordprocessingShape();
                Wps.NonVisualDrawingProperties nonVisualDrawingProperties2 = new Wps.NonVisualDrawingProperties() { Id = (UInt32Value)3U, Name = "AutoShape 3" };

                Wps.NonVisualDrawingShapeProperties nonVisualDrawingShapeProperties1 = new Wps.NonVisualDrawingShapeProperties();
                A.ShapeLocks shapeLocks1 = new A.ShapeLocks() { NoChangeArrowheads = true };

                nonVisualDrawingShapeProperties1.Append(shapeLocks1);

                Wps.ShapeProperties shapeProperties2 = new Wps.ShapeProperties() { BlackWhiteMode = A.BlackWhiteModeValues.Auto };

                A.Transform2D transform2D2 = new A.Transform2D();
                A.Offset offset3 = new A.Offset() { X = 5718L, Y = 533L };
                A.Extents extents3 = new A.Extents() { Cx = 792L, Cy = 365760L };

                transform2D2.Append(offset3);
                transform2D2.Append(extents3);

                A.PresetGeometry presetGeometry2 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.BracketPair };

                A.AdjustValueList adjustValueList2 = new A.AdjustValueList();
                A.ShapeGuide shapeGuide1 = new A.ShapeGuide() { Name = "adj", Formula = "val 16667" };

                adjustValueList2.Append(shapeGuide1);

                presetGeometry2.Append(adjustValueList2);

                A.SolidFill solidFill2 = new A.SolidFill();
                A.RgbColorModelHex rgbColorModelHex2 = new A.RgbColorModelHex() { Val = "FFFFFF" };

                solidFill2.Append(rgbColorModelHex2);

                A.Outline outline2 = new A.Outline() { Width = 28575 };

                A.SolidFill solidFill3 = new A.SolidFill();
                A.RgbColorModelHex rgbColorModelHex3 = new A.RgbColorModelHex() { Val = "808080" };

                solidFill3.Append(rgbColorModelHex3);
                A.Round round2 = new A.Round();
                A.HeadEnd headEnd2 = new A.HeadEnd();
                A.TailEnd tailEnd2 = new A.TailEnd();

                outline2.Append(solidFill3);
                outline2.Append(round2);
                outline2.Append(headEnd2);
                outline2.Append(tailEnd2);

                shapeProperties2.Append(transform2D2);
                shapeProperties2.Append(presetGeometry2);
                shapeProperties2.Append(solidFill2);
                shapeProperties2.Append(outline2);

                Wps.TextBoxInfo2 textBoxInfo21 = new Wps.TextBoxInfo2();

                TextBoxContent textBoxContent1 = new TextBoxContent();

                Paragraph paragraph2 = new Paragraph() { RsidParagraphAddition = "008B30DD", RsidRunAdditionDefault = "008B30DD", ParagraphId = "778440BF", TextId = "77777777" };

                ParagraphProperties paragraphProperties2 = new ParagraphProperties();
                Justification justification1 = new Justification() { Val = JustificationValues.Center };

                paragraphProperties2.Append(justification1);

                Run run2 = new Run();
                FieldChar fieldChar1 = new FieldChar() { FieldCharType = FieldCharValues.Begin };

                run2.Append(fieldChar1);

                Run run3 = new Run();
                FieldCode fieldCode1 = new FieldCode() { Space = SpaceProcessingModeValues.Preserve };
                fieldCode1.Text = " PAGE    \\* MERGEFORMAT ";

                run3.Append(fieldCode1);

                Run run4 = new Run();
                FieldChar fieldChar2 = new FieldChar() { FieldCharType = FieldCharValues.Separate };

                run4.Append(fieldChar2);

                Run run5 = new Run();

                RunProperties runProperties2 = new RunProperties();
                NoProof noProof2 = new NoProof();

                runProperties2.Append(noProof2);
                Text text1 = new Text();
                text1.Text = "2";

                run5.Append(runProperties2);
                run5.Append(text1);

                Run run6 = new Run();

                RunProperties runProperties3 = new RunProperties();
                NoProof noProof3 = new NoProof();

                runProperties3.Append(noProof3);
                FieldChar fieldChar3 = new FieldChar() { FieldCharType = FieldCharValues.End };

                run6.Append(runProperties3);
                run6.Append(fieldChar3);

                paragraph2.Append(paragraphProperties2);
                paragraph2.Append(run2);
                paragraph2.Append(run3);
                paragraph2.Append(run4);
                paragraph2.Append(run5);
                paragraph2.Append(run6);

                textBoxContent1.Append(paragraph2);

                textBoxInfo21.Append(textBoxContent1);

                Wps.TextBodyProperties textBodyProperties2 = new Wps.TextBodyProperties() { Rotation = 0, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, LeftInset = 91440, TopInset = 0, RightInset = 91440, BottomInset = 0, Anchor = A.TextAnchoringTypeValues.Top, AnchorCenter = false, UpRight = true };
                A.NoAutoFit noAutoFit1 = new A.NoAutoFit();

                textBodyProperties2.Append(noAutoFit1);

                wordprocessingShape2.Append(nonVisualDrawingProperties2);
                wordprocessingShape2.Append(nonVisualDrawingShapeProperties1);
                wordprocessingShape2.Append(shapeProperties2);
                wordprocessingShape2.Append(textBoxInfo21);
                wordprocessingShape2.Append(textBodyProperties2);

                wordprocessingGroup1.Append(nonVisualGroupDrawingShapeProperties1);
                wordprocessingGroup1.Append(groupShapeProperties1);
                wordprocessingGroup1.Append(wordprocessingShape1);
                wordprocessingGroup1.Append(wordprocessingShape2);

                graphicData1.Append(wordprocessingGroup1);

                graphic1.Append(graphicData1);

                Wp14.RelativeWidth relativeWidth1 = new Wp14.RelativeWidth() { ObjectId = Wp14.SizeRelativeHorizontallyValues.Margin };
                Wp14.PercentageWidth percentageWidth1 = new Wp14.PercentageWidth();
                percentageWidth1.Text = "100000";

                relativeWidth1.Append(percentageWidth1);

                Wp14.RelativeHeight relativeHeight1 = new Wp14.RelativeHeight() { RelativeFrom = Wp14.SizeRelativeVerticallyValues.Page };
                Wp14.PercentageHeight percentageHeight1 = new Wp14.PercentageHeight();
                percentageHeight1.Text = "0";

                relativeHeight1.Append(percentageHeight1);

                anchor1.Append(simplePosition1);
                anchor1.Append(horizontalPosition1);
                anchor1.Append(verticalPosition1);
                anchor1.Append(extent1);
                anchor1.Append(effectExtent1);
                anchor1.Append(wrapNone1);
                anchor1.Append(docProperties1);
                anchor1.Append(nonVisualGraphicFrameDrawingProperties1);
                anchor1.Append(graphic1);
                anchor1.Append(relativeWidth1);
                anchor1.Append(relativeHeight1);

                drawing1.Append(anchor1);

                alternateContentChoice1.Append(drawing1);

                AlternateContentFallback alternateContentFallback1 = new AlternateContentFallback();

                Picture picture1 = new Picture();

                V.Group group1 = new V.Group() { Id = "Group 1", Style = "position:absolute;margin-left:0;margin-top:0;width:466.4pt;height:28.8pt;z-index:251659264;mso-width-percent:1000;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:top-margin-area;mso-width-percent:1000;mso-width-relative:margin", CoordinateSize = "8698,365760", CoordinateOrigin = "1778,533", OptionalString = "_x0000_s1026" };
                group1.SetAttribute(new OpenXmlAttribute("w14", "anchorId", "http://schemas.microsoft.com/office/word/2010/wordml", "4D8B63E6"));
                group1.SetAttribute(new OpenXmlAttribute("o", "gfxdata", "urn:schemas-microsoft-com:office:office", "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\nAPMAAABbBgAAAAA=\n"));

                V.Shapetype shapetype1 = new V.Shapetype() { Id = "_x0000_t32", CoordinateSize = "21600,21600", Oned = true, Filled = false, OptionalNumber = 32, EdgePath = "m,l21600,21600e" };
                V.Path path1 = new V.Path() { AllowFill = false, ShowArrowhead = true, ConnectionPointType = Ovml.ConnectValues.None };
                Ovml.Lock lock1 = new Ovml.Lock() { Extension = V.ExtensionHandlingBehaviorValues.Edit, ShapeType = true };

                shapetype1.Append(path1);
                shapetype1.Append(lock1);
                V.Shape shape1 = new V.Shape() { Id = "AutoShape 2", Style = "position:absolute;left:1778;top:183413;width:8698;height:0;visibility:visible;mso-wrap-style:square", OptionalString = "_x0000_s1027", StrokeColor = "gray", StrokeWeight = "1pt", ConnectorType = Ovml.ConnectorValues.Straight, Type = "#_x0000_t32", EncodedPackage = "UEsDBBQABgAIAAAAIQDb4fbL7gAAAIUBAAATAAAAW0NvbnRlbnRfVHlwZXNdLnhtbHyQz07DMAyH\n70i8Q+QralM4IITa7kDhCAiNB7ASt43WOlEcyvb2pNu4IODoPz9/n1xv9vOkForiPDdwXVagiI23\njocG3rdPxR0oScgWJ8/UwIEENu3lRb09BBKV0ywNjCmFe63FjDSjlD4Q50nv44wpl3HQAc0OB9I3\nVXWrjedEnIq03oC27qjHjympx31un0wiTQLq4bS4shrAECZnMGVTvbD9QSnOhDInjzsyuiBXWQP0\nr4R18jfgnHvJr4nOknrFmJ5xzhraRtHWf3Kkpfz/yGo5S+H73hkquyhdjr3R8m2lj09svwAAAP//\nAwBQSwMEFAAGAAgAAAAhAFr0LFu/AAAAFQEAAAsAAABfcmVscy8ucmVsc2zPwWrDMAwG4Ptg72B0\nX5TuUMaI01uh19I+gLGVxCy2jGSy9e1nemrHjpL4P0nD4SetZiPRyNnCruvBUPYcYp4tXC/Htw8w\nWl0ObuVMFm6kcBhfX4Yzra62kC6xqGlKVgtLreUTUf1CyWnHhXKbTCzJ1VbKjMX5LzcTvvf9HuXR\ngPHJNKdgQU5hB+ZyK23zHztFL6w81c5zQp6m6P9TMfB3PtPWFCczVQtB9N4U2rp2HOA44NMz4y8A\nAAD//wMAUEsDBBQABgAIAAAAIQDNkUo7wgAAANoAAAAPAAAAZHJzL2Rvd25yZXYueG1sRI9Ba8JA\nFITvhf6H5RV6qxtDKSG6igiFhh7aqhdvj+wzCWbfht1Xjf76riD0OMzMN8x8ObpenSjEzrOB6SQD\nRVx723FjYLd9fylARUG22HsmAxeKsFw8PsyxtP7MP3TaSKMShGOJBlqRodQ61i05jBM/ECfv4IND\nSTI02gY8J7jrdZ5lb9phx2mhxYHWLdXHza8z0IsNn9e8kpB9V1+vu2JfIFXGPD+NqxkooVH+w/f2\nhzWQw+1KugF68QcAAP//AwBQSwECLQAUAAYACAAAACEA2+H2y+4AAACFAQAAEwAAAAAAAAAAAAAA\nAAAAAAAAW0NvbnRlbnRfVHlwZXNdLnhtbFBLAQItABQABgAIAAAAIQBa9CxbvwAAABUBAAALAAAA\nAAAAAAAAAAAAAB8BAABfcmVscy8ucmVsc1BLAQItABQABgAIAAAAIQDNkUo7wgAAANoAAAAPAAAA\nAAAAAAAAAAAAAAcCAABkcnMvZG93bnJldi54bWxQSwUGAAAAAAMAAwC3AAAA9gIAAAAA\n" };

                V.Shapetype shapetype2 = new V.Shapetype() { Id = "_x0000_t185", CoordinateSize = "21600,21600", Filled = false, OptionalNumber = 185, Adjustment = "3600", EdgePath = "m@0,nfqx0@0l0@2qy@0,21600em@1,nfqx21600@0l21600@2qy@1,21600em@0,nsqx0@0l0@2qy@0,21600l@1,21600qx21600@2l21600@0qy@1,xe" };

                V.Formulas formulas1 = new V.Formulas();
                V.Formula formula1 = new V.Formula() { Equation = "val #0" };
                V.Formula formula2 = new V.Formula() { Equation = "sum width 0 #0" };
                V.Formula formula3 = new V.Formula() { Equation = "sum height 0 #0" };
                V.Formula formula4 = new V.Formula() { Equation = "prod @0 2929 10000" };
                V.Formula formula5 = new V.Formula() { Equation = "sum width 0 @3" };
                V.Formula formula6 = new V.Formula() { Equation = "sum height 0 @3" };
                V.Formula formula7 = new V.Formula() { Equation = "val width" };
                V.Formula formula8 = new V.Formula() { Equation = "val height" };
                V.Formula formula9 = new V.Formula() { Equation = "prod width 1 2" };
                V.Formula formula10 = new V.Formula() { Equation = "prod height 1 2" };

                formulas1.Append(formula1);
                formulas1.Append(formula2);
                formulas1.Append(formula3);
                formulas1.Append(formula4);
                formulas1.Append(formula5);
                formulas1.Append(formula6);
                formulas1.Append(formula7);
                formulas1.Append(formula8);
                formulas1.Append(formula9);
                formulas1.Append(formula10);
                V.Path path2 = new V.Path() { Limo = "10800,10800", TextboxRectangle = "@3,@3,@4,@5", AllowGradientShape = true, ConnectionPointType = Ovml.ConnectValues.Custom, ConnectionPoints = "@8,0;0,@9;@8,@7;@6,@9", AllowExtrusion = false };

                V.ShapeHandles shapeHandles1 = new V.ShapeHandles();
                V.ShapeHandle shapeHandle1 = new V.ShapeHandle() { Position = "#0,topLeft", Switch = false, XRange = "0,10800" };

                shapeHandles1.Append(shapeHandle1);

                shapetype2.Append(formulas1);
                shapetype2.Append(path2);
                shapetype2.Append(shapeHandles1);

                V.Shape shape2 = new V.Shape() { Id = "AutoShape 3", Style = "position:absolute;left:5718;top:533;width:792;height:365760;visibility:visible;mso-wrap-style:square;v-text-anchor:top", OptionalString = "_x0000_s1028", Filled = true, StrokeColor = "gray", StrokeWeight = "2.25pt", Type = "#_x0000_t185", EncodedPackage = "UEsDBBQABgAIAAAAIQDb4fbL7gAAAIUBAAATAAAAW0NvbnRlbnRfVHlwZXNdLnhtbHyQz07DMAyH\n70i8Q+QralM4IITa7kDhCAiNB7ASt43WOlEcyvb2pNu4IODoPz9/n1xv9vOkForiPDdwXVagiI23\njocG3rdPxR0oScgWJ8/UwIEENu3lRb09BBKV0ywNjCmFe63FjDSjlD4Q50nv44wpl3HQAc0OB9I3\nVXWrjedEnIq03oC27qjHjympx31un0wiTQLq4bS4shrAECZnMGVTvbD9QSnOhDInjzsyuiBXWQP0\nr4R18jfgnHvJr4nOknrFmJ5xzhraRtHWf3Kkpfz/yGo5S+H73hkquyhdjr3R8m2lj09svwAAAP//\nAwBQSwMEFAAGAAgAAAAhAFr0LFu/AAAAFQEAAAsAAABfcmVscy8ucmVsc2zPwWrDMAwG4Ptg72B0\nX5TuUMaI01uh19I+gLGVxCy2jGSy9e1nemrHjpL4P0nD4SetZiPRyNnCruvBUPYcYp4tXC/Htw8w\nWl0ObuVMFm6kcBhfX4Yzra62kC6xqGlKVgtLreUTUf1CyWnHhXKbTCzJ1VbKjMX5LzcTvvf9HuXR\ngPHJNKdgQU5hB+ZyK23zHztFL6w81c5zQp6m6P9TMfB3PtPWFCczVQtB9N4U2rp2HOA44NMz4y8A\nAAD//wMAUEsDBBQABgAIAAAAIQBjussDxAAAANoAAAAPAAAAZHJzL2Rvd25yZXYueG1sRI/RasJA\nFETfhf7Dcgu+6cZai0ZXaYulpZRA1Q+4ZK9JTPZu3F01/fuuIPg4zMwZZrHqTCPO5HxlWcFomIAg\nzq2uuFCw234MpiB8QNbYWCYFf+RhtXzoLTDV9sK/dN6EQkQI+xQVlCG0qZQ+L8mgH9qWOHp76wyG\nKF0htcNLhJtGPiXJizRYcVwosaX3kvJ6czIKMvcztpPP7DR7M+vDc3081qH7Vqr/2L3OQQTqwj18\na39pBWO4Xok3QC7/AQAA//8DAFBLAQItABQABgAIAAAAIQDb4fbL7gAAAIUBAAATAAAAAAAAAAAA\nAAAAAAAAAABbQ29udGVudF9UeXBlc10ueG1sUEsBAi0AFAAGAAgAAAAhAFr0LFu/AAAAFQEAAAsA\nAAAAAAAAAAAAAAAAHwEAAF9yZWxzLy5yZWxzUEsBAi0AFAAGAAgAAAAhAGO6ywPEAAAA2gAAAA8A\nAAAAAAAAAAAAAAAABwIAAGRycy9kb3ducmV2LnhtbFBLBQYAAAAAAwADALcAAAD4AgAAAAA=\n" };

                V.TextBox textBox1 = new V.TextBox() { Inset = ",0,,0" };

                TextBoxContent textBoxContent2 = new TextBoxContent();

                Paragraph paragraph3 = new Paragraph() { RsidParagraphAddition = "008B30DD", RsidRunAdditionDefault = "008B30DD", ParagraphId = "778440BF", TextId = "77777777" };

                ParagraphProperties paragraphProperties3 = new ParagraphProperties();
                Justification justification2 = new Justification() { Val = JustificationValues.Center };

                paragraphProperties3.Append(justification2);

                Run run7 = new Run();
                FieldChar fieldChar4 = new FieldChar() { FieldCharType = FieldCharValues.Begin };

                run7.Append(fieldChar4);

                Run run8 = new Run();
                FieldCode fieldCode2 = new FieldCode() { Space = SpaceProcessingModeValues.Preserve };
                fieldCode2.Text = " PAGE    \\* MERGEFORMAT ";

                run8.Append(fieldCode2);

                Run run9 = new Run();
                FieldChar fieldChar5 = new FieldChar() { FieldCharType = FieldCharValues.Separate };

                run9.Append(fieldChar5);

                Run run10 = new Run();

                RunProperties runProperties4 = new RunProperties();
                NoProof noProof4 = new NoProof();

                runProperties4.Append(noProof4);
                Text text2 = new Text();
                text2.Text = "2";

                run10.Append(runProperties4);
                run10.Append(text2);

                Run run11 = new Run();

                RunProperties runProperties5 = new RunProperties();
                NoProof noProof5 = new NoProof();

                runProperties5.Append(noProof5);
                FieldChar fieldChar6 = new FieldChar() { FieldCharType = FieldCharValues.End };

                run11.Append(runProperties5);
                run11.Append(fieldChar6);

                paragraph3.Append(paragraphProperties3);
                paragraph3.Append(run7);
                paragraph3.Append(run8);
                paragraph3.Append(run9);
                paragraph3.Append(run10);
                paragraph3.Append(run11);

                textBoxContent2.Append(paragraph3);

                textBox1.Append(textBoxContent2);

                shape2.Append(textBox1);
                Wvml.TextWrap textWrap1 = new Wvml.TextWrap() { AnchorX = Wvml.HorizontalAnchorValues.Margin, AnchorY = Wvml.VerticalAnchorValues.Margin };

                group1.Append(shapetype1);
                group1.Append(shape1);
                group1.Append(shapetype2);
                group1.Append(shape2);
                group1.Append(textWrap1);

                picture1.Append(group1);

                alternateContentFallback1.Append(picture1);

                alternateContent1.Append(alternateContentChoice1);
                alternateContent1.Append(alternateContentFallback1);

                run1.Append(runProperties1);
                run1.Append(alternateContent1);

                paragraph1.Append(paragraphProperties1);
                paragraph1.Append(run1);

                sdtContentBlock1.Append(paragraph1);

                sdtBlock1.Append(sdtProperties1);
                sdtBlock1.Append(sdtContentBlock1);
                return sdtBlock1;

            }
        }
        private static SdtBlock Dots1 {
            get {
                SdtBlock sdtBlock1 = new SdtBlock();

                SdtProperties sdtProperties1 = new SdtProperties();
                SdtId sdtId1 = new SdtId() { Val = -451858885 };

                SdtContentDocPartObject sdtContentDocPartObject1 = new SdtContentDocPartObject();
                DocPartGallery docPartGallery1 = new DocPartGallery() { Val = "Page Numbers (Top of Page)" };
                DocPartUnique docPartUnique1 = new DocPartUnique();

                sdtContentDocPartObject1.Append(docPartGallery1);
                sdtContentDocPartObject1.Append(docPartUnique1);

                sdtProperties1.Append(sdtId1);
                sdtProperties1.Append(sdtContentDocPartObject1);

                SdtContentBlock sdtContentBlock1 = new SdtContentBlock();

                Paragraph paragraph1 = new Paragraph() { RsidParagraphAddition = "004918B9", RsidRunAdditionDefault = "004918B9", ParagraphId = "00BD54B3", TextId = "2A9D606F" };

                ParagraphProperties paragraphProperties1 = new ParagraphProperties();
                ParagraphStyleId paragraphStyleId1 = new ParagraphStyleId() { Val = "Header" };
                Justification justification1 = new Justification() { Val = JustificationValues.Center };

                paragraphProperties1.Append(paragraphStyleId1);
                paragraphProperties1.Append(justification1);

                Run run1 = new Run();

                RunProperties runProperties1 = new RunProperties();
                NoProof noProof1 = new NoProof();

                runProperties1.Append(noProof1);

                AlternateContent alternateContent1 = new AlternateContent();

                AlternateContentChoice alternateContentChoice1 = new AlternateContentChoice() { Requires = "wpg" };

                Drawing drawing1 = new Drawing();

                Wp.Inline inline1 = new Wp.Inline() { DistanceFromTop = (UInt32Value)0U, DistanceFromBottom = (UInt32Value)0U, DistanceFromLeft = (UInt32Value)0U, DistanceFromRight = (UInt32Value)0U, AnchorId = "34BA9F6C", EditId = "62CF8404" };
                Wp.Extent extent1 = new Wp.Extent() { Cx = 418465L, Cy = 221615L };
                Wp.EffectExtent effectExtent1 = new Wp.EffectExtent() { LeftEdge = 0L, TopEdge = 0L, RightEdge = 635L, BottomEdge = 0L };
                Wp.DocProperties docProperties1 = new Wp.DocProperties() { Id = (UInt32Value)4U, Name = "Group 4" };

                Wp.NonVisualGraphicFrameDrawingProperties nonVisualGraphicFrameDrawingProperties1 = new Wp.NonVisualGraphicFrameDrawingProperties();

                A.GraphicFrameLocks graphicFrameLocks1 = new A.GraphicFrameLocks();
                graphicFrameLocks1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

                nonVisualGraphicFrameDrawingProperties1.Append(graphicFrameLocks1);

                A.Graphic graphic1 = new A.Graphic();
                graphic1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

                A.GraphicData graphicData1 = new A.GraphicData() { Uri = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" };

                Wpg.WordprocessingGroup wordprocessingGroup1 = new Wpg.WordprocessingGroup();

                Wpg.NonVisualGroupDrawingShapeProperties nonVisualGroupDrawingShapeProperties1 = new Wpg.NonVisualGroupDrawingShapeProperties();
                A.GroupShapeLocks groupShapeLocks1 = new A.GroupShapeLocks();

                nonVisualGroupDrawingShapeProperties1.Append(groupShapeLocks1);

                Wpg.GroupShapeProperties groupShapeProperties1 = new Wpg.GroupShapeProperties() { BlackWhiteMode = A.BlackWhiteModeValues.Auto };

                A.TransformGroup transformGroup1 = new A.TransformGroup();
                A.Offset offset1 = new A.Offset() { X = 0L, Y = 0L };
                A.Extents extents1 = new A.Extents() { Cx = 418465L, Cy = 221615L };
                A.ChildOffset childOffset1 = new A.ChildOffset() { X = 5351L, Y = 739L };
                A.ChildExtents childExtents1 = new A.ChildExtents() { Cx = 659L, Cy = 349L };

                transformGroup1.Append(offset1);
                transformGroup1.Append(extents1);
                transformGroup1.Append(childOffset1);
                transformGroup1.Append(childExtents1);

                groupShapeProperties1.Append(transformGroup1);

                Wps.WordprocessingShape wordprocessingShape1 = new Wps.WordprocessingShape();
                Wps.NonVisualDrawingProperties nonVisualDrawingProperties1 = new Wps.NonVisualDrawingProperties() { Id = (UInt32Value)5U, Name = "Text Box 56" };

                Wps.NonVisualDrawingShapeProperties nonVisualDrawingShapeProperties1 = new Wps.NonVisualDrawingShapeProperties() { TextBox = true };
                A.ShapeLocks shapeLocks1 = new A.ShapeLocks() { NoChangeArrowheads = true };

                nonVisualDrawingShapeProperties1.Append(shapeLocks1);

                Wps.ShapeProperties shapeProperties1 = new Wps.ShapeProperties() { BlackWhiteMode = A.BlackWhiteModeValues.Auto };

                A.Transform2D transform2D1 = new A.Transform2D();
                A.Offset offset2 = new A.Offset() { X = 5351L, Y = 800L };
                A.Extents extents2 = new A.Extents() { Cx = 659L, Cy = 288L };

                transform2D1.Append(offset2);
                transform2D1.Append(extents2);

                A.PresetGeometry presetGeometry1 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
                A.AdjustValueList adjustValueList1 = new A.AdjustValueList();

                presetGeometry1.Append(adjustValueList1);
                A.NoFill noFill1 = new A.NoFill();

                A.Outline outline1 = new A.Outline();
                A.NoFill noFill2 = new A.NoFill();

                outline1.Append(noFill2);

                A.ShapePropertiesExtensionList shapePropertiesExtensionList1 = new A.ShapePropertiesExtensionList();

                A.ShapePropertiesExtension shapePropertiesExtension1 = new A.ShapePropertiesExtension() { Uri = "{909E8E84-426E-40DD-AFC4-6F175D3DCCD1}" };

                A14.HiddenFillProperties hiddenFillProperties1 = new A14.HiddenFillProperties();
                hiddenFillProperties1.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");

                A.SolidFill solidFill1 = new A.SolidFill();
                A.RgbColorModelHex rgbColorModelHex1 = new A.RgbColorModelHex() { Val = "FFFFFF" };

                solidFill1.Append(rgbColorModelHex1);

                hiddenFillProperties1.Append(solidFill1);

                shapePropertiesExtension1.Append(hiddenFillProperties1);

                A.ShapePropertiesExtension shapePropertiesExtension2 = new A.ShapePropertiesExtension() { Uri = "{91240B29-F687-4F45-9708-019B960494DF}" };

                A14.HiddenLineProperties hiddenLineProperties1 = new A14.HiddenLineProperties() { Width = 9525 };
                hiddenLineProperties1.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");

                A.SolidFill solidFill2 = new A.SolidFill();
                A.RgbColorModelHex rgbColorModelHex2 = new A.RgbColorModelHex() { Val = "000000" };

                solidFill2.Append(rgbColorModelHex2);
                A.Miter miter1 = new A.Miter() { Limit = 800000 };
                A.HeadEnd headEnd1 = new A.HeadEnd();
                A.TailEnd tailEnd1 = new A.TailEnd();

                hiddenLineProperties1.Append(solidFill2);
                hiddenLineProperties1.Append(miter1);
                hiddenLineProperties1.Append(headEnd1);
                hiddenLineProperties1.Append(tailEnd1);

                shapePropertiesExtension2.Append(hiddenLineProperties1);

                shapePropertiesExtensionList1.Append(shapePropertiesExtension1);
                shapePropertiesExtensionList1.Append(shapePropertiesExtension2);

                shapeProperties1.Append(transform2D1);
                shapeProperties1.Append(presetGeometry1);
                shapeProperties1.Append(noFill1);
                shapeProperties1.Append(outline1);
                shapeProperties1.Append(shapePropertiesExtensionList1);

                Wps.TextBoxInfo2 textBoxInfo21 = new Wps.TextBoxInfo2();

                TextBoxContent textBoxContent1 = new TextBoxContent();

                Paragraph paragraph2 = new Paragraph() { RsidParagraphAddition = "004918B9", RsidRunAdditionDefault = "004918B9", ParagraphId = "12772B1B", TextId = "77777777" };

                ParagraphProperties paragraphProperties2 = new ParagraphProperties();
                Justification justification2 = new Justification() { Val = JustificationValues.Center };

                ParagraphMarkRunProperties paragraphMarkRunProperties1 = new ParagraphMarkRunProperties();
                FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "18" };

                paragraphMarkRunProperties1.Append(fontSizeComplexScript1);

                paragraphProperties2.Append(justification2);
                paragraphProperties2.Append(paragraphMarkRunProperties1);

                Run run2 = new Run();
                FieldChar fieldChar1 = new FieldChar() { FieldCharType = FieldCharValues.Begin };

                run2.Append(fieldChar1);

                Run run3 = new Run();
                FieldCode fieldCode1 = new FieldCode() { Space = SpaceProcessingModeValues.Preserve };
                fieldCode1.Text = " PAGE    \\* MERGEFORMAT ";

                run3.Append(fieldCode1);

                Run run4 = new Run();
                FieldChar fieldChar2 = new FieldChar() { FieldCharType = FieldCharValues.Separate };

                run4.Append(fieldChar2);

                Run run5 = new Run();

                RunProperties runProperties2 = new RunProperties();
                Italic italic1 = new Italic();
                ItalicComplexScript italicComplexScript1 = new ItalicComplexScript();
                NoProof noProof2 = new NoProof();
                FontSize fontSize1 = new FontSize() { Val = "18" };
                FontSizeComplexScript fontSizeComplexScript2 = new FontSizeComplexScript() { Val = "18" };

                runProperties2.Append(italic1);
                runProperties2.Append(italicComplexScript1);
                runProperties2.Append(noProof2);
                runProperties2.Append(fontSize1);
                runProperties2.Append(fontSizeComplexScript2);
                Text text1 = new Text();
                text1.Text = "2";

                run5.Append(runProperties2);
                run5.Append(text1);

                Run run6 = new Run();

                RunProperties runProperties3 = new RunProperties();
                Italic italic2 = new Italic();
                ItalicComplexScript italicComplexScript2 = new ItalicComplexScript();
                NoProof noProof3 = new NoProof();
                FontSize fontSize2 = new FontSize() { Val = "18" };
                FontSizeComplexScript fontSizeComplexScript3 = new FontSizeComplexScript() { Val = "18" };

                runProperties3.Append(italic2);
                runProperties3.Append(italicComplexScript2);
                runProperties3.Append(noProof3);
                runProperties3.Append(fontSize2);
                runProperties3.Append(fontSizeComplexScript3);
                FieldChar fieldChar3 = new FieldChar() { FieldCharType = FieldCharValues.End };

                run6.Append(runProperties3);
                run6.Append(fieldChar3);

                paragraph2.Append(paragraphProperties2);
                paragraph2.Append(run2);
                paragraph2.Append(run3);
                paragraph2.Append(run4);
                paragraph2.Append(run5);
                paragraph2.Append(run6);

                textBoxContent1.Append(paragraph2);

                textBoxInfo21.Append(textBoxContent1);

                Wps.TextBodyProperties textBodyProperties1 = new Wps.TextBodyProperties() { Rotation = 0, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, LeftInset = 0, TopInset = 0, RightInset = 0, BottomInset = 0, Anchor = A.TextAnchoringTypeValues.Top, AnchorCenter = false, UpRight = true };
                A.NoAutoFit noAutoFit1 = new A.NoAutoFit();

                textBodyProperties1.Append(noAutoFit1);

                wordprocessingShape1.Append(nonVisualDrawingProperties1);
                wordprocessingShape1.Append(nonVisualDrawingShapeProperties1);
                wordprocessingShape1.Append(shapeProperties1);
                wordprocessingShape1.Append(textBoxInfo21);
                wordprocessingShape1.Append(textBodyProperties1);

                Wpg.GroupShape groupShape1 = new Wpg.GroupShape();
                Wpg.NonVisualDrawingProperties nonVisualDrawingProperties2 = new Wpg.NonVisualDrawingProperties() { Id = (UInt32Value)6U, Name = "Group 57" };

                Wpg.NonVisualGroupDrawingShapeProperties nonVisualGroupDrawingShapeProperties2 = new Wpg.NonVisualGroupDrawingShapeProperties();
                A.GroupShapeLocks groupShapeLocks2 = new A.GroupShapeLocks();

                nonVisualGroupDrawingShapeProperties2.Append(groupShapeLocks2);

                Wpg.GroupShapeProperties groupShapeProperties2 = new Wpg.GroupShapeProperties() { BlackWhiteMode = A.BlackWhiteModeValues.Auto };

                A.TransformGroup transformGroup2 = new A.TransformGroup();
                A.Offset offset3 = new A.Offset() { X = 5494L, Y = 739L };
                A.Extents extents3 = new A.Extents() { Cx = 372L, Cy = 72L };
                A.ChildOffset childOffset2 = new A.ChildOffset() { X = 5486L, Y = 739L };
                A.ChildExtents childExtents2 = new A.ChildExtents() { Cx = 372L, Cy = 72L };

                transformGroup2.Append(offset3);
                transformGroup2.Append(extents3);
                transformGroup2.Append(childOffset2);
                transformGroup2.Append(childExtents2);

                groupShapeProperties2.Append(transformGroup2);

                Wps.WordprocessingShape wordprocessingShape2 = new Wps.WordprocessingShape();
                Wps.NonVisualDrawingProperties nonVisualDrawingProperties3 = new Wps.NonVisualDrawingProperties() { Id = (UInt32Value)7U, Name = "Oval 58" };

                Wps.NonVisualDrawingShapeProperties nonVisualDrawingShapeProperties2 = new Wps.NonVisualDrawingShapeProperties();
                A.ShapeLocks shapeLocks2 = new A.ShapeLocks() { NoChangeArrowheads = true };

                nonVisualDrawingShapeProperties2.Append(shapeLocks2);

                Wps.ShapeProperties shapeProperties2 = new Wps.ShapeProperties() { BlackWhiteMode = A.BlackWhiteModeValues.Auto };

                A.Transform2D transform2D2 = new A.Transform2D();
                A.Offset offset4 = new A.Offset() { X = 5486L, Y = 739L };
                A.Extents extents4 = new A.Extents() { Cx = 72L, Cy = 72L };

                transform2D2.Append(offset4);
                transform2D2.Append(extents4);

                A.PresetGeometry presetGeometry2 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Ellipse };
                A.AdjustValueList adjustValueList2 = new A.AdjustValueList();

                presetGeometry2.Append(adjustValueList2);

                A.SolidFill solidFill3 = new A.SolidFill();
                A.RgbColorModelHex rgbColorModelHex3 = new A.RgbColorModelHex() { Val = "84A2C6" };

                solidFill3.Append(rgbColorModelHex3);

                A.Outline outline2 = new A.Outline();
                A.NoFill noFill3 = new A.NoFill();

                outline2.Append(noFill3);

                A.ShapePropertiesExtensionList shapePropertiesExtensionList2 = new A.ShapePropertiesExtensionList();

                A.ShapePropertiesExtension shapePropertiesExtension3 = new A.ShapePropertiesExtension() { Uri = "{91240B29-F687-4F45-9708-019B960494DF}" };

                A14.HiddenLineProperties hiddenLineProperties2 = new A14.HiddenLineProperties() { Width = 9525 };
                hiddenLineProperties2.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");

                A.SolidFill solidFill4 = new A.SolidFill();
                A.RgbColorModelHex rgbColorModelHex4 = new A.RgbColorModelHex() { Val = "000000" };

                solidFill4.Append(rgbColorModelHex4);
                A.Round round1 = new A.Round();
                A.HeadEnd headEnd2 = new A.HeadEnd();
                A.TailEnd tailEnd2 = new A.TailEnd();

                hiddenLineProperties2.Append(solidFill4);
                hiddenLineProperties2.Append(round1);
                hiddenLineProperties2.Append(headEnd2);
                hiddenLineProperties2.Append(tailEnd2);

                shapePropertiesExtension3.Append(hiddenLineProperties2);

                shapePropertiesExtensionList2.Append(shapePropertiesExtension3);

                shapeProperties2.Append(transform2D2);
                shapeProperties2.Append(presetGeometry2);
                shapeProperties2.Append(solidFill3);
                shapeProperties2.Append(outline2);
                shapeProperties2.Append(shapePropertiesExtensionList2);

                Wps.TextBodyProperties textBodyProperties2 = new Wps.TextBodyProperties() { Rotation = 0, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, LeftInset = 91440, TopInset = 45720, RightInset = 91440, BottomInset = 45720, Anchor = A.TextAnchoringTypeValues.Top, AnchorCenter = false, UpRight = true };
                A.NoAutoFit noAutoFit2 = new A.NoAutoFit();

                textBodyProperties2.Append(noAutoFit2);

                wordprocessingShape2.Append(nonVisualDrawingProperties3);
                wordprocessingShape2.Append(nonVisualDrawingShapeProperties2);
                wordprocessingShape2.Append(shapeProperties2);
                wordprocessingShape2.Append(textBodyProperties2);

                Wps.WordprocessingShape wordprocessingShape3 = new Wps.WordprocessingShape();
                Wps.NonVisualDrawingProperties nonVisualDrawingProperties4 = new Wps.NonVisualDrawingProperties() { Id = (UInt32Value)8U, Name = "Oval 59" };

                Wps.NonVisualDrawingShapeProperties nonVisualDrawingShapeProperties3 = new Wps.NonVisualDrawingShapeProperties();
                A.ShapeLocks shapeLocks3 = new A.ShapeLocks() { NoChangeArrowheads = true };

                nonVisualDrawingShapeProperties3.Append(shapeLocks3);

                Wps.ShapeProperties shapeProperties3 = new Wps.ShapeProperties() { BlackWhiteMode = A.BlackWhiteModeValues.Auto };

                A.Transform2D transform2D3 = new A.Transform2D();
                A.Offset offset5 = new A.Offset() { X = 5636L, Y = 739L };
                A.Extents extents5 = new A.Extents() { Cx = 72L, Cy = 72L };

                transform2D3.Append(offset5);
                transform2D3.Append(extents5);

                A.PresetGeometry presetGeometry3 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Ellipse };
                A.AdjustValueList adjustValueList3 = new A.AdjustValueList();

                presetGeometry3.Append(adjustValueList3);

                A.SolidFill solidFill5 = new A.SolidFill();
                A.RgbColorModelHex rgbColorModelHex5 = new A.RgbColorModelHex() { Val = "84A2C6" };

                solidFill5.Append(rgbColorModelHex5);

                A.Outline outline3 = new A.Outline();
                A.NoFill noFill4 = new A.NoFill();

                outline3.Append(noFill4);

                A.ShapePropertiesExtensionList shapePropertiesExtensionList3 = new A.ShapePropertiesExtensionList();

                A.ShapePropertiesExtension shapePropertiesExtension4 = new A.ShapePropertiesExtension() { Uri = "{91240B29-F687-4F45-9708-019B960494DF}" };

                A14.HiddenLineProperties hiddenLineProperties3 = new A14.HiddenLineProperties() { Width = 9525 };
                hiddenLineProperties3.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");

                A.SolidFill solidFill6 = new A.SolidFill();
                A.RgbColorModelHex rgbColorModelHex6 = new A.RgbColorModelHex() { Val = "000000" };

                solidFill6.Append(rgbColorModelHex6);
                A.Round round2 = new A.Round();
                A.HeadEnd headEnd3 = new A.HeadEnd();
                A.TailEnd tailEnd3 = new A.TailEnd();

                hiddenLineProperties3.Append(solidFill6);
                hiddenLineProperties3.Append(round2);
                hiddenLineProperties3.Append(headEnd3);
                hiddenLineProperties3.Append(tailEnd3);

                shapePropertiesExtension4.Append(hiddenLineProperties3);

                shapePropertiesExtensionList3.Append(shapePropertiesExtension4);

                shapeProperties3.Append(transform2D3);
                shapeProperties3.Append(presetGeometry3);
                shapeProperties3.Append(solidFill5);
                shapeProperties3.Append(outline3);
                shapeProperties3.Append(shapePropertiesExtensionList3);

                Wps.TextBodyProperties textBodyProperties3 = new Wps.TextBodyProperties() { Rotation = 0, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, LeftInset = 91440, TopInset = 45720, RightInset = 91440, BottomInset = 45720, Anchor = A.TextAnchoringTypeValues.Top, AnchorCenter = false, UpRight = true };
                A.NoAutoFit noAutoFit3 = new A.NoAutoFit();

                textBodyProperties3.Append(noAutoFit3);

                wordprocessingShape3.Append(nonVisualDrawingProperties4);
                wordprocessingShape3.Append(nonVisualDrawingShapeProperties3);
                wordprocessingShape3.Append(shapeProperties3);
                wordprocessingShape3.Append(textBodyProperties3);

                Wps.WordprocessingShape wordprocessingShape4 = new Wps.WordprocessingShape();
                Wps.NonVisualDrawingProperties nonVisualDrawingProperties5 = new Wps.NonVisualDrawingProperties() { Id = (UInt32Value)9U, Name = "Oval 60" };

                Wps.NonVisualDrawingShapeProperties nonVisualDrawingShapeProperties4 = new Wps.NonVisualDrawingShapeProperties();
                A.ShapeLocks shapeLocks4 = new A.ShapeLocks() { NoChangeArrowheads = true };

                nonVisualDrawingShapeProperties4.Append(shapeLocks4);

                Wps.ShapeProperties shapeProperties4 = new Wps.ShapeProperties() { BlackWhiteMode = A.BlackWhiteModeValues.Auto };

                A.Transform2D transform2D4 = new A.Transform2D();
                A.Offset offset6 = new A.Offset() { X = 5786L, Y = 739L };
                A.Extents extents6 = new A.Extents() { Cx = 72L, Cy = 72L };

                transform2D4.Append(offset6);
                transform2D4.Append(extents6);

                A.PresetGeometry presetGeometry4 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Ellipse };
                A.AdjustValueList adjustValueList4 = new A.AdjustValueList();

                presetGeometry4.Append(adjustValueList4);

                A.SolidFill solidFill7 = new A.SolidFill();
                A.RgbColorModelHex rgbColorModelHex7 = new A.RgbColorModelHex() { Val = "84A2C6" };

                solidFill7.Append(rgbColorModelHex7);

                A.Outline outline4 = new A.Outline();
                A.NoFill noFill5 = new A.NoFill();

                outline4.Append(noFill5);

                A.ShapePropertiesExtensionList shapePropertiesExtensionList4 = new A.ShapePropertiesExtensionList();

                A.ShapePropertiesExtension shapePropertiesExtension5 = new A.ShapePropertiesExtension() { Uri = "{91240B29-F687-4F45-9708-019B960494DF}" };

                A14.HiddenLineProperties hiddenLineProperties4 = new A14.HiddenLineProperties() { Width = 9525 };
                hiddenLineProperties4.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");

                A.SolidFill solidFill8 = new A.SolidFill();
                A.RgbColorModelHex rgbColorModelHex8 = new A.RgbColorModelHex() { Val = "000000" };

                solidFill8.Append(rgbColorModelHex8);
                A.Round round3 = new A.Round();
                A.HeadEnd headEnd4 = new A.HeadEnd();
                A.TailEnd tailEnd4 = new A.TailEnd();

                hiddenLineProperties4.Append(solidFill8);
                hiddenLineProperties4.Append(round3);
                hiddenLineProperties4.Append(headEnd4);
                hiddenLineProperties4.Append(tailEnd4);

                shapePropertiesExtension5.Append(hiddenLineProperties4);

                shapePropertiesExtensionList4.Append(shapePropertiesExtension5);

                shapeProperties4.Append(transform2D4);
                shapeProperties4.Append(presetGeometry4);
                shapeProperties4.Append(solidFill7);
                shapeProperties4.Append(outline4);
                shapeProperties4.Append(shapePropertiesExtensionList4);

                Wps.TextBodyProperties textBodyProperties4 = new Wps.TextBodyProperties() { Rotation = 0, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, LeftInset = 91440, TopInset = 45720, RightInset = 91440, BottomInset = 45720, Anchor = A.TextAnchoringTypeValues.Top, AnchorCenter = false, UpRight = true };
                A.NoAutoFit noAutoFit4 = new A.NoAutoFit();

                textBodyProperties4.Append(noAutoFit4);

                wordprocessingShape4.Append(nonVisualDrawingProperties5);
                wordprocessingShape4.Append(nonVisualDrawingShapeProperties4);
                wordprocessingShape4.Append(shapeProperties4);
                wordprocessingShape4.Append(textBodyProperties4);

                groupShape1.Append(nonVisualDrawingProperties2);
                groupShape1.Append(nonVisualGroupDrawingShapeProperties2);
                groupShape1.Append(groupShapeProperties2);
                groupShape1.Append(wordprocessingShape2);
                groupShape1.Append(wordprocessingShape3);
                groupShape1.Append(wordprocessingShape4);

                wordprocessingGroup1.Append(nonVisualGroupDrawingShapeProperties1);
                wordprocessingGroup1.Append(groupShapeProperties1);
                wordprocessingGroup1.Append(wordprocessingShape1);
                wordprocessingGroup1.Append(groupShape1);

                graphicData1.Append(wordprocessingGroup1);

                graphic1.Append(graphicData1);

                inline1.Append(extent1);
                inline1.Append(effectExtent1);
                inline1.Append(docProperties1);
                inline1.Append(nonVisualGraphicFrameDrawingProperties1);
                inline1.Append(graphic1);

                drawing1.Append(inline1);

                alternateContentChoice1.Append(drawing1);

                AlternateContentFallback alternateContentFallback1 = new AlternateContentFallback();

                Picture picture1 = new Picture();

                V.Group group1 = new V.Group() { Id = "Group 4", Style = "width:32.95pt;height:17.45pt;mso-position-horizontal-relative:char;mso-position-vertical-relative:line", CoordinateSize = "659,349", CoordinateOrigin = "5351,739", OptionalString = "_x0000_s1026" };
                group1.SetAttribute(new OpenXmlAttribute("w14", "anchorId", "http://schemas.microsoft.com/office/word/2010/wordml", "34BA9F6C"));
                group1.SetAttribute(new OpenXmlAttribute("o", "gfxdata", "urn:schemas-microsoft-com:office:office", "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\n"));

                V.Shapetype shapetype1 = new V.Shapetype() { Id = "_x0000_t202", CoordinateSize = "21600,21600", OptionalNumber = 202, EdgePath = "m,l,21600r21600,l21600,xe" };
                V.Stroke stroke1 = new V.Stroke() { JoinStyle = V.StrokeJoinStyleValues.Miter };
                V.Path path1 = new V.Path() { AllowGradientShape = true, ConnectionPointType = Ovml.ConnectValues.Rectangle };

                shapetype1.Append(stroke1);
                shapetype1.Append(path1);

                V.Shape shape1 = new V.Shape() { Id = "Text Box 56", Style = "position:absolute;left:5351;top:800;width:659;height:288;visibility:visible;mso-wrap-style:square;v-text-anchor:top", OptionalString = "_x0000_s1027", Filled = false, Stroked = false, Type = "#_x0000_t202", EncodedPackage = "UEsDBBQABgAIAAAAIQDb4fbL7gAAAIUBAAATAAAAW0NvbnRlbnRfVHlwZXNdLnhtbHyQz07DMAyH\n70i8Q+QralM4IITa7kDhCAiNB7ASt43WOlEcyvb2pNu4IODoPz9/n1xv9vOkForiPDdwXVagiI23\njocG3rdPxR0oScgWJ8/UwIEENu3lRb09BBKV0ywNjCmFe63FjDSjlD4Q50nv44wpl3HQAc0OB9I3\nVXWrjedEnIq03oC27qjHjympx31un0wiTQLq4bS4shrAECZnMGVTvbD9QSnOhDInjzsyuiBXWQP0\nr4R18jfgnHvJr4nOknrFmJ5xzhraRtHWf3Kkpfz/yGo5S+H73hkquyhdjr3R8m2lj09svwAAAP//\nAwBQSwMEFAAGAAgAAAAhAFr0LFu/AAAAFQEAAAsAAABfcmVscy8ucmVsc2zPwWrDMAwG4Ptg72B0\nX5TuUMaI01uh19I+gLGVxCy2jGSy9e1nemrHjpL4P0nD4SetZiPRyNnCruvBUPYcYp4tXC/Htw8w\nWl0ObuVMFm6kcBhfX4Yzra62kC6xqGlKVgtLreUTUf1CyWnHhXKbTCzJ1VbKjMX5LzcTvvf9HuXR\ngPHJNKdgQU5hB+ZyK23zHztFL6w81c5zQp6m6P9TMfB3PtPWFCczVQtB9N4U2rp2HOA44NMz4y8A\nAAD//wMAUEsDBBQABgAIAAAAIQBXZw8vwwAAANoAAAAPAAAAZHJzL2Rvd25yZXYueG1sRI9Ba8JA\nFITvBf/D8gRvdWNBqdFVRCwUhGKMB4/P7DNZzL6N2a3Gf98VCh6HmfmGmS87W4sbtd44VjAaJiCI\nC6cNlwoO+df7JwgfkDXWjknBgzwsF723Oaba3Tmj2z6UIkLYp6igCqFJpfRFRRb90DXE0Tu71mKI\nsi2lbvEe4baWH0kykRYNx4UKG1pXVFz2v1bB6sjZxlx/TrvsnJk8nya8nVyUGvS71QxEoC68wv/t\nb61gDM8r8QbIxR8AAAD//wMAUEsBAi0AFAAGAAgAAAAhANvh9svuAAAAhQEAABMAAAAAAAAAAAAA\nAAAAAAAAAFtDb250ZW50X1R5cGVzXS54bWxQSwECLQAUAAYACAAAACEAWvQsW78AAAAVAQAACwAA\nAAAAAAAAAAAAAAAfAQAAX3JlbHMvLnJlbHNQSwECLQAUAAYACAAAACEAV2cPL8MAAADaAAAADwAA\nAAAAAAAAAAAAAAAHAgAAZHJzL2Rvd25yZXYueG1sUEsFBgAAAAADAAMAtwAAAPcCAAAAAA==\n" };

                V.TextBox textBox1 = new V.TextBox() { Inset = "0,0,0,0" };

                TextBoxContent textBoxContent2 = new TextBoxContent();

                Paragraph paragraph3 = new Paragraph() { RsidParagraphAddition = "004918B9", RsidRunAdditionDefault = "004918B9", ParagraphId = "12772B1B", TextId = "77777777" };

                ParagraphProperties paragraphProperties3 = new ParagraphProperties();
                Justification justification3 = new Justification() { Val = JustificationValues.Center };

                ParagraphMarkRunProperties paragraphMarkRunProperties2 = new ParagraphMarkRunProperties();
                FontSizeComplexScript fontSizeComplexScript4 = new FontSizeComplexScript() { Val = "18" };

                paragraphMarkRunProperties2.Append(fontSizeComplexScript4);

                paragraphProperties3.Append(justification3);
                paragraphProperties3.Append(paragraphMarkRunProperties2);

                Run run7 = new Run();
                FieldChar fieldChar4 = new FieldChar() { FieldCharType = FieldCharValues.Begin };

                run7.Append(fieldChar4);

                Run run8 = new Run();
                FieldCode fieldCode2 = new FieldCode() { Space = SpaceProcessingModeValues.Preserve };
                fieldCode2.Text = " PAGE    \\* MERGEFORMAT ";

                run8.Append(fieldCode2);

                Run run9 = new Run();
                FieldChar fieldChar5 = new FieldChar() { FieldCharType = FieldCharValues.Separate };

                run9.Append(fieldChar5);

                Run run10 = new Run();

                RunProperties runProperties4 = new RunProperties();
                Italic italic3 = new Italic();
                ItalicComplexScript italicComplexScript3 = new ItalicComplexScript();
                NoProof noProof4 = new NoProof();
                FontSize fontSize3 = new FontSize() { Val = "18" };
                FontSizeComplexScript fontSizeComplexScript5 = new FontSizeComplexScript() { Val = "18" };

                runProperties4.Append(italic3);
                runProperties4.Append(italicComplexScript3);
                runProperties4.Append(noProof4);
                runProperties4.Append(fontSize3);
                runProperties4.Append(fontSizeComplexScript5);
                Text text2 = new Text();
                text2.Text = "2";

                run10.Append(runProperties4);
                run10.Append(text2);

                Run run11 = new Run();

                RunProperties runProperties5 = new RunProperties();
                Italic italic4 = new Italic();
                ItalicComplexScript italicComplexScript4 = new ItalicComplexScript();
                NoProof noProof5 = new NoProof();
                FontSize fontSize4 = new FontSize() { Val = "18" };
                FontSizeComplexScript fontSizeComplexScript6 = new FontSizeComplexScript() { Val = "18" };

                runProperties5.Append(italic4);
                runProperties5.Append(italicComplexScript4);
                runProperties5.Append(noProof5);
                runProperties5.Append(fontSize4);
                runProperties5.Append(fontSizeComplexScript6);
                FieldChar fieldChar6 = new FieldChar() { FieldCharType = FieldCharValues.End };

                run11.Append(runProperties5);
                run11.Append(fieldChar6);

                paragraph3.Append(paragraphProperties3);
                paragraph3.Append(run7);
                paragraph3.Append(run8);
                paragraph3.Append(run9);
                paragraph3.Append(run10);
                paragraph3.Append(run11);

                textBoxContent2.Append(paragraph3);

                textBox1.Append(textBoxContent2);

                shape1.Append(textBox1);

                V.Group group2 = new V.Group() { Id = "Group 57", Style = "position:absolute;left:5494;top:739;width:372;height:72", CoordinateSize = "372,72", CoordinateOrigin = "5486,739", OptionalString = "_x0000_s1028" };
                group2.SetAttribute(new OpenXmlAttribute("o", "gfxdata", "urn:schemas-microsoft-com:office:office", "UEsDBBQABgAIAAAAIQDb4fbL7gAAAIUBAAATAAAAW0NvbnRlbnRfVHlwZXNdLnhtbHyQz07DMAyH\n70i8Q+QralM4IITa7kDhCAiNB7ASt43WOlEcyvb2pNu4IODoPz9/n1xv9vOkForiPDdwXVagiI23\njocG3rdPxR0oScgWJ8/UwIEENu3lRb09BBKV0ywNjCmFe63FjDSjlD4Q50nv44wpl3HQAc0OB9I3\nVXWrjedEnIq03oC27qjHjympx31un0wiTQLq4bS4shrAECZnMGVTvbD9QSnOhDInjzsyuiBXWQP0\nr4R18jfgnHvJr4nOknrFmJ5xzhraRtHWf3Kkpfz/yGo5S+H73hkquyhdjr3R8m2lj09svwAAAP//\nAwBQSwMEFAAGAAgAAAAhAFr0LFu/AAAAFQEAAAsAAABfcmVscy8ucmVsc2zPwWrDMAwG4Ptg72B0\nX5TuUMaI01uh19I+gLGVxCy2jGSy9e1nemrHjpL4P0nD4SetZiPRyNnCruvBUPYcYp4tXC/Htw8w\nWl0ObuVMFm6kcBhfX4Yzra62kC6xqGlKVgtLreUTUf1CyWnHhXKbTCzJ1VbKjMX5LzcTvvf9HuXR\ngPHJNKdgQU5hB+ZyK23zHztFL6w81c5zQp6m6P9TMfB3PtPWFCczVQtB9N4U2rp2HOA44NMz4y8A\nAAD//wMAUEsDBBQABgAIAAAAIQDNCx1fwgAAANoAAAAPAAAAZHJzL2Rvd25yZXYueG1sRI9Bi8Iw\nFITvC/6H8ARva1plRapRRFQ8iLAqiLdH82yLzUtpYlv/vVkQ9jjMzDfMfNmZUjRUu8KygngYgSBO\nrS44U3A5b7+nIJxH1lhaJgUvcrBc9L7mmGjb8i81J5+JAGGXoILc+yqR0qU5GXRDWxEH725rgz7I\nOpO6xjbATSlHUTSRBgsOCzlWtM4pfZyeRsGuxXY1jjfN4XFfv27nn+P1EJNSg363moHw1Pn/8Ke9\n1wom8Hcl3AC5eAMAAP//AwBQSwECLQAUAAYACAAAACEA2+H2y+4AAACFAQAAEwAAAAAAAAAAAAAA\nAAAAAAAAW0NvbnRlbnRfVHlwZXNdLnhtbFBLAQItABQABgAIAAAAIQBa9CxbvwAAABUBAAALAAAA\nAAAAAAAAAAAAAB8BAABfcmVscy8ucmVsc1BLAQItABQABgAIAAAAIQDNCx1fwgAAANoAAAAPAAAA\nAAAAAAAAAAAAAAcCAABkcnMvZG93bnJldi54bWxQSwUGAAAAAAMAAwC3AAAA9gIAAAAA\n"));

                V.Oval oval1 = new V.Oval() { Id = "Oval 58", Style = "position:absolute;left:5486;top:739;width:72;height:72;visibility:visible;mso-wrap-style:square;v-text-anchor:top", OptionalString = "_x0000_s1029", FillColor = "#84a2c6", Stroked = false };
                oval1.SetAttribute(new OpenXmlAttribute("o", "gfxdata", "urn:schemas-microsoft-com:office:office", "UEsDBBQABgAIAAAAIQDb4fbL7gAAAIUBAAATAAAAW0NvbnRlbnRfVHlwZXNdLnhtbHyQz07DMAyH\n70i8Q+QralM4IITa7kDhCAiNB7ASt43WOlEcyvb2pNu4IODoPz9/n1xv9vOkForiPDdwXVagiI23\njocG3rdPxR0oScgWJ8/UwIEENu3lRb09BBKV0ywNjCmFe63FjDSjlD4Q50nv44wpl3HQAc0OB9I3\nVXWrjedEnIq03oC27qjHjympx31un0wiTQLq4bS4shrAECZnMGVTvbD9QSnOhDInjzsyuiBXWQP0\nr4R18jfgnHvJr4nOknrFmJ5xzhraRtHWf3Kkpfz/yGo5S+H73hkquyhdjr3R8m2lj09svwAAAP//\nAwBQSwMEFAAGAAgAAAAhAFr0LFu/AAAAFQEAAAsAAABfcmVscy8ucmVsc2zPwWrDMAwG4Ptg72B0\nX5TuUMaI01uh19I+gLGVxCy2jGSy9e1nemrHjpL4P0nD4SetZiPRyNnCruvBUPYcYp4tXC/Htw8w\nWl0ObuVMFm6kcBhfX4Yzra62kC6xqGlKVgtLreUTUf1CyWnHhXKbTCzJ1VbKjMX5LzcTvvf9HuXR\ngPHJNKdgQU5hB+ZyK23zHztFL6w81c5zQp6m6P9TMfB3PtPWFCczVQtB9N4U2rp2HOA44NMz4y8A\nAAD//wMAUEsDBBQABgAIAAAAIQCVU+0kvgAAANoAAAAPAAAAZHJzL2Rvd25yZXYueG1sRI9BD8FA\nFITvEv9h8yRubDkgZQkS4qo4uD3dp2103zbdVfXvrUTiOJmZbzKLVWtK0VDtCssKRsMIBHFqdcGZ\ngvNpN5iBcB5ZY2mZFLzJwWrZ7Sww1vbFR2oSn4kAYRejgtz7KpbSpTkZdENbEQfvbmuDPsg6k7rG\nV4CbUo6jaCINFhwWcqxom1P6SJ5GQbG3o8tukxzdtZls5bq8bezlplS/167nIDy1/h/+tQ9awRS+\nV8INkMsPAAAA//8DAFBLAQItABQABgAIAAAAIQDb4fbL7gAAAIUBAAATAAAAAAAAAAAAAAAAAAAA\nAABbQ29udGVudF9UeXBlc10ueG1sUEsBAi0AFAAGAAgAAAAhAFr0LFu/AAAAFQEAAAsAAAAAAAAA\nAAAAAAAAHwEAAF9yZWxzLy5yZWxzUEsBAi0AFAAGAAgAAAAhAJVT7SS+AAAA2gAAAA8AAAAAAAAA\nAAAAAAAABwIAAGRycy9kb3ducmV2LnhtbFBLBQYAAAAAAwADALcAAADyAgAAAAA=\n"));

                V.Oval oval2 = new V.Oval() { Id = "Oval 59", Style = "position:absolute;left:5636;top:739;width:72;height:72;visibility:visible;mso-wrap-style:square;v-text-anchor:top", OptionalString = "_x0000_s1030", FillColor = "#84a2c6", Stroked = false };
                oval2.SetAttribute(new OpenXmlAttribute("o", "gfxdata", "urn:schemas-microsoft-com:office:office", "UEsDBBQABgAIAAAAIQDb4fbL7gAAAIUBAAATAAAAW0NvbnRlbnRfVHlwZXNdLnhtbHyQz07DMAyH\n70i8Q+QralM4IITa7kDhCAiNB7ASt43WOlEcyvb2pNu4IODoPz9/n1xv9vOkForiPDdwXVagiI23\njocG3rdPxR0oScgWJ8/UwIEENu3lRb09BBKV0ywNjCmFe63FjDSjlD4Q50nv44wpl3HQAc0OB9I3\nVXWrjedEnIq03oC27qjHjympx31un0wiTQLq4bS4shrAECZnMGVTvbD9QSnOhDInjzsyuiBXWQP0\nr4R18jfgnHvJr4nOknrFmJ5xzhraRtHWf3Kkpfz/yGo5S+H73hkquyhdjr3R8m2lj09svwAAAP//\nAwBQSwMEFAAGAAgAAAAhAFr0LFu/AAAAFQEAAAsAAABfcmVscy8ucmVsc2zPwWrDMAwG4Ptg72B0\nX5TuUMaI01uh19I+gLGVxCy2jGSy9e1nemrHjpL4P0nD4SetZiPRyNnCruvBUPYcYp4tXC/Htw8w\nWl0ObuVMFm6kcBhfX4Yzra62kC6xqGlKVgtLreUTUf1CyWnHhXKbTCzJ1VbKjMX5LzcTvvf9HuXR\ngPHJNKdgQU5hB+ZyK23zHztFL6w81c5zQp6m6P9TMfB3PtPWFCczVQtB9N4U2rp2HOA44NMz4y8A\nAAD//wMAUEsDBBQABgAIAAAAIQDkzHlWuwAAANoAAAAPAAAAZHJzL2Rvd25yZXYueG1sRE+9CsIw\nEN4F3yGc4KapDiLVWKqguFp1cDubsy02l9LEWt/eDILjx/e/TnpTi45aV1lWMJtGIIhzqysuFFzO\n+8kShPPIGmvLpOBDDpLNcLDGWNs3n6jLfCFCCLsYFZTeN7GULi/JoJvahjhwD9sa9AG2hdQtvkO4\nqeU8ihbSYMWhocSGdiXlz+xlFFQHO7vut9nJ3brFTqb1fWuvd6XGoz5dgfDU+7/45z5qBWFruBJu\ngNx8AQAA//8DAFBLAQItABQABgAIAAAAIQDb4fbL7gAAAIUBAAATAAAAAAAAAAAAAAAAAAAAAABb\nQ29udGVudF9UeXBlc10ueG1sUEsBAi0AFAAGAAgAAAAhAFr0LFu/AAAAFQEAAAsAAAAAAAAAAAAA\nAAAAHwEAAF9yZWxzLy5yZWxzUEsBAi0AFAAGAAgAAAAhAOTMeVa7AAAA2gAAAA8AAAAAAAAAAAAA\nAAAABwIAAGRycy9kb3ducmV2LnhtbFBLBQYAAAAAAwADALcAAADvAgAAAAA=\n"));

                V.Oval oval3 = new V.Oval() { Id = "Oval 60", Style = "position:absolute;left:5786;top:739;width:72;height:72;visibility:visible;mso-wrap-style:square;v-text-anchor:top", OptionalString = "_x0000_s1031", FillColor = "#84a2c6", Stroked = false };
                oval3.SetAttribute(new OpenXmlAttribute("o", "gfxdata", "urn:schemas-microsoft-com:office:office", "UEsDBBQABgAIAAAAIQDb4fbL7gAAAIUBAAATAAAAW0NvbnRlbnRfVHlwZXNdLnhtbHyQz07DMAyH\n70i8Q+QralM4IITa7kDhCAiNB7ASt43WOlEcyvb2pNu4IODoPz9/n1xv9vOkForiPDdwXVagiI23\njocG3rdPxR0oScgWJ8/UwIEENu3lRb09BBKV0ywNjCmFe63FjDSjlD4Q50nv44wpl3HQAc0OB9I3\nVXWrjedEnIq03oC27qjHjympx31un0wiTQLq4bS4shrAECZnMGVTvbD9QSnOhDInjzsyuiBXWQP0\nr4R18jfgnHvJr4nOknrFmJ5xzhraRtHWf3Kkpfz/yGo5S+H73hkquyhdjr3R8m2lj09svwAAAP//\nAwBQSwMEFAAGAAgAAAAhAFr0LFu/AAAAFQEAAAsAAABfcmVscy8ucmVsc2zPwWrDMAwG4Ptg72B0\nX5TuUMaI01uh19I+gLGVxCy2jGSy9e1nemrHjpL4P0nD4SetZiPRyNnCruvBUPYcYp4tXC/Htw8w\nWl0ObuVMFm6kcBhfX4Yzra62kC6xqGlKVgtLreUTUf1CyWnHhXKbTCzJ1VbKjMX5LzcTvvf9HuXR\ngPHJNKdgQU5hB+ZyK23zHztFL6w81c5zQp6m6P9TMfB3PtPWFCczVQtB9N4U2rp2HOA44NMz4y8A\nAAD//wMAUEsDBBQABgAIAAAAIQCLgNzNvgAAANoAAAAPAAAAZHJzL2Rvd25yZXYueG1sRI9BD8FA\nFITvEv9h8yRubDkIZQkS4qo4uD3dp2103zbdVfXvrUTiOJmZbzKLVWtK0VDtCssKRsMIBHFqdcGZ\ngvNpN5iCcB5ZY2mZFLzJwWrZ7Sww1vbFR2oSn4kAYRejgtz7KpbSpTkZdENbEQfvbmuDPsg6k7rG\nV4CbUo6jaCINFhwWcqxom1P6SJ5GQbG3o8tukxzdtZls5bq8bezlplS/167nIDy1/h/+tQ9awQy+\nV8INkMsPAAAA//8DAFBLAQItABQABgAIAAAAIQDb4fbL7gAAAIUBAAATAAAAAAAAAAAAAAAAAAAA\nAABbQ29udGVudF9UeXBlc10ueG1sUEsBAi0AFAAGAAgAAAAhAFr0LFu/AAAAFQEAAAsAAAAAAAAA\nAAAAAAAAHwEAAF9yZWxzLy5yZWxzUEsBAi0AFAAGAAgAAAAhAIuA3M2+AAAA2gAAAA8AAAAAAAAA\nAAAAAAAABwIAAGRycy9kb3ducmV2LnhtbFBLBQYAAAAAAwADALcAAADyAgAAAAA=\n"));

                group2.Append(oval1);
                group2.Append(oval2);
                group2.Append(oval3);
                Wvml.AnchorLock anchorLock1 = new Wvml.AnchorLock();

                group1.Append(shapetype1);
                group1.Append(shape1);
                group1.Append(group2);
                group1.Append(anchorLock1);

                picture1.Append(group1);

                alternateContentFallback1.Append(picture1);

                alternateContent1.Append(alternateContentChoice1);
                alternateContent1.Append(alternateContentFallback1);

                run1.Append(runProperties1);
                run1.Append(alternateContent1);

                paragraph1.Append(paragraphProperties1);
                paragraph1.Append(run1);

                sdtContentBlock1.Append(paragraph1);

                sdtBlock1.Append(sdtProperties1);
                sdtBlock1.Append(sdtContentBlock1);
                return sdtBlock1;

            }
        }
        private static SdtBlock LargeItalics1 {
            get {
                SdtBlock sdtBlock1 = new SdtBlock();

                SdtProperties sdtProperties1 = new SdtProperties();

                RunProperties runProperties1 = new RunProperties();
                RunFonts runFonts1 = new RunFonts() { EastAsiaTheme = ThemeFontValues.MinorEastAsia };

                runProperties1.Append(runFonts1);
                SdtId sdtId1 = new SdtId() { Val = 142853570 };

                SdtContentDocPartObject sdtContentDocPartObject1 = new SdtContentDocPartObject();
                DocPartGallery docPartGallery1 = new DocPartGallery() { Val = "Page Numbers (Top of Page)" };
                DocPartUnique docPartUnique1 = new DocPartUnique();

                sdtContentDocPartObject1.Append(docPartGallery1);
                sdtContentDocPartObject1.Append(docPartUnique1);

                sdtProperties1.Append(runProperties1);
                sdtProperties1.Append(sdtId1);
                sdtProperties1.Append(sdtContentDocPartObject1);

                SdtEndCharProperties sdtEndCharProperties1 = new SdtEndCharProperties();

                RunProperties runProperties2 = new RunProperties();
                RunFonts runFonts2 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
                Italic italic1 = new Italic();
                ItalicComplexScript italicComplexScript1 = new ItalicComplexScript();
                Color color1 = new Color() { Val = "BFBFBF", ThemeColor = ThemeColorValues.Background1, ThemeShade = "BF" };
                FontSize fontSize1 = new FontSize() { Val = "72" };
                FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "72" };

                runProperties2.Append(runFonts2);
                runProperties2.Append(italic1);
                runProperties2.Append(italicComplexScript1);
                runProperties2.Append(color1);
                runProperties2.Append(fontSize1);
                runProperties2.Append(fontSizeComplexScript1);

                sdtEndCharProperties1.Append(runProperties2);

                SdtContentBlock sdtContentBlock1 = new SdtContentBlock();

                Paragraph paragraph1 = new Paragraph() { RsidParagraphMarkRevision = "004918B9", RsidParagraphAddition = "00E11B1E", RsidParagraphProperties = "004918B9", RsidRunAdditionDefault = "00F07E2D", ParagraphId = "16AD3494", TextId = "554CF436" };

                ParagraphProperties paragraphProperties1 = new ParagraphProperties();
                ParagraphStyleId paragraphStyleId1 = new ParagraphStyleId() { Val = "Header" };

                paragraphProperties1.Append(paragraphStyleId1);

                Run run1 = new Run();

                RunProperties runProperties3 = new RunProperties();
                RunFonts runFonts3 = new RunFonts() { EastAsiaTheme = ThemeFontValues.MinorEastAsia };

                runProperties3.Append(runFonts3);
                FieldChar fieldChar1 = new FieldChar() { FieldCharType = FieldCharValues.Begin };

                run1.Append(runProperties3);
                run1.Append(fieldChar1);

                Run run2 = new Run();
                FieldCode fieldCode1 = new FieldCode() { Space = SpaceProcessingModeValues.Preserve };
                fieldCode1.Text = " PAGE    \\* MERGEFORMAT ";

                run2.Append(fieldCode1);

                Run run3 = new Run();

                RunProperties runProperties4 = new RunProperties();
                RunFonts runFonts4 = new RunFonts() { EastAsiaTheme = ThemeFontValues.MinorEastAsia };

                runProperties4.Append(runFonts4);
                FieldChar fieldChar2 = new FieldChar() { FieldCharType = FieldCharValues.Separate };

                run3.Append(runProperties4);
                run3.Append(fieldChar2);

                Run run4 = new Run();

                RunProperties runProperties5 = new RunProperties();
                RunFonts runFonts5 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
                Italic italic2 = new Italic();
                ItalicComplexScript italicComplexScript2 = new ItalicComplexScript();
                NoProof noProof1 = new NoProof();
                Color color2 = new Color() { Val = "BFBFBF", ThemeColor = ThemeColorValues.Background1, ThemeShade = "BF" };
                Spacing spacing1 = new Spacing() { Val = -40 };
                FontSize fontSize2 = new FontSize() { Val = "72" };
                FontSizeComplexScript fontSizeComplexScript2 = new FontSizeComplexScript() { Val = "72" };

                runProperties5.Append(runFonts5);
                runProperties5.Append(italic2);
                runProperties5.Append(italicComplexScript2);
                runProperties5.Append(noProof1);
                runProperties5.Append(color2);
                runProperties5.Append(spacing1);
                runProperties5.Append(fontSize2);
                runProperties5.Append(fontSizeComplexScript2);
                Text text1 = new Text();
                text1.Text = "2";

                run4.Append(runProperties5);
                run4.Append(text1);

                Run run5 = new Run();

                RunProperties runProperties6 = new RunProperties();
                RunFonts runFonts6 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
                Italic italic3 = new Italic();
                ItalicComplexScript italicComplexScript3 = new ItalicComplexScript();
                NoProof noProof2 = new NoProof();
                Color color3 = new Color() { Val = "BFBFBF", ThemeColor = ThemeColorValues.Background1, ThemeShade = "BF" };
                Spacing spacing2 = new Spacing() { Val = -40 };
                FontSize fontSize3 = new FontSize() { Val = "72" };
                FontSizeComplexScript fontSizeComplexScript3 = new FontSizeComplexScript() { Val = "72" };

                runProperties6.Append(runFonts6);
                runProperties6.Append(italic3);
                runProperties6.Append(italicComplexScript3);
                runProperties6.Append(noProof2);
                runProperties6.Append(color3);
                runProperties6.Append(spacing2);
                runProperties6.Append(fontSize3);
                runProperties6.Append(fontSizeComplexScript3);
                FieldChar fieldChar3 = new FieldChar() { FieldCharType = FieldCharValues.End };

                run5.Append(runProperties6);
                run5.Append(fieldChar3);

                Run run6 = new Run();

                RunProperties runProperties7 = new RunProperties();
                RunFonts runFonts7 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
                Italic italic4 = new Italic();
                ItalicComplexScript italicComplexScript4 = new ItalicComplexScript();
                Color color4 = new Color() { Val = "BFBFBF", ThemeColor = ThemeColorValues.Background1, ThemeShade = "BF" };
                FontSize fontSize4 = new FontSize() { Val = "72" };
                FontSizeComplexScript fontSizeComplexScript4 = new FontSizeComplexScript() { Val = "72" };

                runProperties7.Append(runFonts7);
                runProperties7.Append(italic4);
                runProperties7.Append(italicComplexScript4);
                runProperties7.Append(color4);
                runProperties7.Append(fontSize4);
                runProperties7.Append(fontSizeComplexScript4);
                Text text2 = new Text();
                text2.Text = ":";

                run6.Append(runProperties7);
                run6.Append(text2);

                paragraph1.Append(paragraphProperties1);
                paragraph1.Append(run1);
                paragraph1.Append(run2);
                paragraph1.Append(run3);
                paragraph1.Append(run4);
                paragraph1.Append(run5);
                paragraph1.Append(run6);

                sdtContentBlock1.Append(paragraph1);

                sdtBlock1.Append(sdtProperties1);
                sdtBlock1.Append(sdtEndCharProperties1);
                sdtBlock1.Append(sdtContentBlock1);
                return sdtBlock1;

            }
        }
        private static SdtBlock Roman1 {
            get {
                SdtBlock sdtBlock1 = new SdtBlock();

                SdtProperties sdtProperties1 = new SdtProperties();
                SdtId sdtId1 = new SdtId() { Val = -628620079 };

                SdtContentDocPartObject sdtContentDocPartObject1 = new SdtContentDocPartObject();
                DocPartGallery docPartGallery1 = new DocPartGallery() { Val = "Page Numbers (Top of Page)" };
                DocPartUnique docPartUnique1 = new DocPartUnique();

                sdtContentDocPartObject1.Append(docPartGallery1);
                sdtContentDocPartObject1.Append(docPartUnique1);

                sdtProperties1.Append(sdtId1);
                sdtProperties1.Append(sdtContentDocPartObject1);

                SdtEndCharProperties sdtEndCharProperties1 = new SdtEndCharProperties();

                RunProperties runProperties1 = new RunProperties();
                NoProof noProof1 = new NoProof();
                Color color1 = new Color() { Val = "4472C4", ThemeColor = ThemeColorValues.Accent1 };
                FontSize fontSize1 = new FontSize() { Val = "28" };
                FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "28" };

                runProperties1.Append(noProof1);
                runProperties1.Append(color1);
                runProperties1.Append(fontSize1);
                runProperties1.Append(fontSizeComplexScript1);

                sdtEndCharProperties1.Append(runProperties1);

                SdtContentBlock sdtContentBlock1 = new SdtContentBlock();

                Paragraph paragraph1 = new Paragraph() { RsidParagraphAddition = "00405F46", RsidRunAdditionDefault = "00405F46", ParagraphId = "734066F7", TextId = "09DFE444" };

                ParagraphProperties paragraphProperties1 = new ParagraphProperties();
                ParagraphStyleId paragraphStyleId1 = new ParagraphStyleId() { Val = "Header" };

                ParagraphMarkRunProperties paragraphMarkRunProperties1 = new ParagraphMarkRunProperties();
                Color color2 = new Color() { Val = "4472C4", ThemeColor = ThemeColorValues.Accent1 };
                FontSize fontSize2 = new FontSize() { Val = "28" };
                FontSizeComplexScript fontSizeComplexScript2 = new FontSizeComplexScript() { Val = "28" };

                paragraphMarkRunProperties1.Append(color2);
                paragraphMarkRunProperties1.Append(fontSize2);
                paragraphMarkRunProperties1.Append(fontSizeComplexScript2);

                paragraphProperties1.Append(paragraphStyleId1);
                paragraphProperties1.Append(paragraphMarkRunProperties1);

                Run run1 = new Run();
                FieldChar fieldChar1 = new FieldChar() { FieldCharType = FieldCharValues.Begin };

                run1.Append(fieldChar1);

                Run run2 = new Run();
                FieldCode fieldCode1 = new FieldCode() { Space = SpaceProcessingModeValues.Preserve };
                fieldCode1.Text = " PAGE  \\* ROMAN  \\* MERGEFORMAT ";

                run2.Append(fieldCode1);

                Run run3 = new Run();
                FieldChar fieldChar2 = new FieldChar() { FieldCharType = FieldCharValues.Separate };

                run3.Append(fieldChar2);

                Run run4 = new Run();

                RunProperties runProperties2 = new RunProperties();
                NoProof noProof2 = new NoProof();
                Color color3 = new Color() { Val = "4472C4", ThemeColor = ThemeColorValues.Accent1 };
                FontSize fontSize3 = new FontSize() { Val = "28" };
                FontSizeComplexScript fontSizeComplexScript3 = new FontSizeComplexScript() { Val = "28" };

                runProperties2.Append(noProof2);
                runProperties2.Append(color3);
                runProperties2.Append(fontSize3);
                runProperties2.Append(fontSizeComplexScript3);
                Text text1 = new Text();
                text1.Text = "I";

                run4.Append(runProperties2);
                run4.Append(text1);

                Run run5 = new Run();

                RunProperties runProperties3 = new RunProperties();
                NoProof noProof3 = new NoProof();
                Color color4 = new Color() { Val = "4472C4", ThemeColor = ThemeColorValues.Accent1 };
                FontSize fontSize4 = new FontSize() { Val = "28" };
                FontSizeComplexScript fontSizeComplexScript4 = new FontSizeComplexScript() { Val = "28" };

                runProperties3.Append(noProof3);
                runProperties3.Append(color4);
                runProperties3.Append(fontSize4);
                runProperties3.Append(fontSizeComplexScript4);
                Text text2 = new Text();
                text2.Text = "I";

                run5.Append(runProperties3);
                run5.Append(text2);

                Run run6 = new Run();

                RunProperties runProperties4 = new RunProperties();
                NoProof noProof4 = new NoProof();
                Color color5 = new Color() { Val = "4472C4", ThemeColor = ThemeColorValues.Accent1 };
                FontSize fontSize5 = new FontSize() { Val = "28" };
                FontSizeComplexScript fontSizeComplexScript5 = new FontSizeComplexScript() { Val = "28" };

                runProperties4.Append(noProof4);
                runProperties4.Append(color5);
                runProperties4.Append(fontSize5);
                runProperties4.Append(fontSizeComplexScript5);
                FieldChar fieldChar3 = new FieldChar() { FieldCharType = FieldCharValues.End };

                run6.Append(runProperties4);
                run6.Append(fieldChar3);

                paragraph1.Append(paragraphProperties1);
                paragraph1.Append(run1);
                paragraph1.Append(run2);
                paragraph1.Append(run3);
                paragraph1.Append(run4);
                paragraph1.Append(run5);
                paragraph1.Append(run6);

                sdtContentBlock1.Append(paragraph1);

                sdtBlock1.Append(sdtProperties1);
                sdtBlock1.Append(sdtEndCharProperties1);
                sdtBlock1.Append(sdtContentBlock1);
                return sdtBlock1;

            }
        }
        private static SdtBlock Tildes1 {
            get {
                SdtBlock sdtBlock1 = new SdtBlock();

                SdtProperties sdtProperties1 = new SdtProperties();

                RunProperties runProperties1 = new RunProperties();
                RunFonts runFonts1 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
                FontSize fontSize1 = new FontSize() { Val = "28" };
                FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "28" };

                runProperties1.Append(runFonts1);
                runProperties1.Append(fontSize1);
                runProperties1.Append(fontSizeComplexScript1);
                SdtId sdtId1 = new SdtId() { Val = -1235152154 };

                SdtContentDocPartObject sdtContentDocPartObject1 = new SdtContentDocPartObject();
                DocPartGallery docPartGallery1 = new DocPartGallery() { Val = "Page Numbers (Top of Page)" };
                DocPartUnique docPartUnique1 = new DocPartUnique();

                sdtContentDocPartObject1.Append(docPartGallery1);
                sdtContentDocPartObject1.Append(docPartUnique1);

                sdtProperties1.Append(runProperties1);
                sdtProperties1.Append(sdtId1);
                sdtProperties1.Append(sdtContentDocPartObject1);

                SdtContentBlock sdtContentBlock1 = new SdtContentBlock();

                Paragraph paragraph1 = new Paragraph() { RsidParagraphAddition = "008A693D", RsidRunAdditionDefault = "008A693D", ParagraphId = "15F6120A", TextId = "6160D751" };

                ParagraphProperties paragraphProperties1 = new ParagraphProperties();
                ParagraphStyleId paragraphStyleId1 = new ParagraphStyleId() { Val = "Header" };
                Justification justification1 = new Justification() { Val = JustificationValues.Center };

                ParagraphMarkRunProperties paragraphMarkRunProperties1 = new ParagraphMarkRunProperties();
                RunFonts runFonts2 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
                FontSize fontSize2 = new FontSize() { Val = "28" };
                FontSizeComplexScript fontSizeComplexScript2 = new FontSizeComplexScript() { Val = "28" };

                paragraphMarkRunProperties1.Append(runFonts2);
                paragraphMarkRunProperties1.Append(fontSize2);
                paragraphMarkRunProperties1.Append(fontSizeComplexScript2);

                paragraphProperties1.Append(paragraphStyleId1);
                paragraphProperties1.Append(justification1);
                paragraphProperties1.Append(paragraphMarkRunProperties1);

                Run run1 = new Run();

                RunProperties runProperties2 = new RunProperties();
                RunFonts runFonts3 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
                FontSize fontSize3 = new FontSize() { Val = "28" };
                FontSizeComplexScript fontSizeComplexScript3 = new FontSizeComplexScript() { Val = "28" };

                runProperties2.Append(runFonts3);
                runProperties2.Append(fontSize3);
                runProperties2.Append(fontSizeComplexScript3);
                Text text1 = new Text() { Space = SpaceProcessingModeValues.Preserve };
                text1.Text = "~ ";

                run1.Append(runProperties2);
                run1.Append(text1);

                Run run2 = new Run();

                RunProperties runProperties3 = new RunProperties();
                RunFonts runFonts4 = new RunFonts() { ComplexScript = "Times New Roman", EastAsiaTheme = ThemeFontValues.MinorEastAsia };

                runProperties3.Append(runFonts4);
                FieldChar fieldChar1 = new FieldChar() { FieldCharType = FieldCharValues.Begin };

                run2.Append(runProperties3);
                run2.Append(fieldChar1);

                Run run3 = new Run();
                FieldCode fieldCode1 = new FieldCode() { Space = SpaceProcessingModeValues.Preserve };
                fieldCode1.Text = " PAGE    \\* MERGEFORMAT ";

                run3.Append(fieldCode1);

                Run run4 = new Run();

                RunProperties runProperties4 = new RunProperties();
                RunFonts runFonts5 = new RunFonts() { ComplexScript = "Times New Roman", EastAsiaTheme = ThemeFontValues.MinorEastAsia };

                runProperties4.Append(runFonts5);
                FieldChar fieldChar2 = new FieldChar() { FieldCharType = FieldCharValues.Separate };

                run4.Append(runProperties4);
                run4.Append(fieldChar2);

                Run run5 = new Run();

                RunProperties runProperties5 = new RunProperties();
                RunFonts runFonts6 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
                NoProof noProof1 = new NoProof();
                FontSize fontSize4 = new FontSize() { Val = "28" };
                FontSizeComplexScript fontSizeComplexScript4 = new FontSizeComplexScript() { Val = "28" };

                runProperties5.Append(runFonts6);
                runProperties5.Append(noProof1);
                runProperties5.Append(fontSize4);
                runProperties5.Append(fontSizeComplexScript4);
                Text text2 = new Text();
                text2.Text = "2";

                run5.Append(runProperties5);
                run5.Append(text2);

                Run run6 = new Run();

                RunProperties runProperties6 = new RunProperties();
                RunFonts runFonts7 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
                NoProof noProof2 = new NoProof();
                FontSize fontSize5 = new FontSize() { Val = "28" };
                FontSizeComplexScript fontSizeComplexScript5 = new FontSizeComplexScript() { Val = "28" };

                runProperties6.Append(runFonts7);
                runProperties6.Append(noProof2);
                runProperties6.Append(fontSize5);
                runProperties6.Append(fontSizeComplexScript5);
                FieldChar fieldChar3 = new FieldChar() { FieldCharType = FieldCharValues.End };

                run6.Append(runProperties6);
                run6.Append(fieldChar3);

                Run run7 = new Run();

                RunProperties runProperties7 = new RunProperties();
                RunFonts runFonts8 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
                FontSize fontSize6 = new FontSize() { Val = "28" };
                FontSizeComplexScript fontSizeComplexScript6 = new FontSizeComplexScript() { Val = "28" };

                runProperties7.Append(runFonts8);
                runProperties7.Append(fontSize6);
                runProperties7.Append(fontSizeComplexScript6);
                Text text3 = new Text() { Space = SpaceProcessingModeValues.Preserve };
                text3.Text = " ~";

                run7.Append(runProperties7);
                run7.Append(text3);

                paragraph1.Append(paragraphProperties1);
                paragraph1.Append(run1);
                paragraph1.Append(run2);
                paragraph1.Append(run3);
                paragraph1.Append(run4);
                paragraph1.Append(run5);
                paragraph1.Append(run6);
                paragraph1.Append(run7);

                sdtContentBlock1.Append(paragraph1);

                sdtBlock1.Append(sdtProperties1);
                sdtBlock1.Append(sdtContentBlock1);
                return sdtBlock1;

            }
        }
        private static SdtBlock FooterTwoBars1 {
            get {
                SdtBlock sdtBlock1 = new SdtBlock();

                SdtProperties sdtProperties1 = new SdtProperties();
                SdtId sdtId1 = new SdtId() { Val = -249269937 };

                SdtContentDocPartObject sdtContentDocPartObject1 = new SdtContentDocPartObject();
                DocPartGallery docPartGallery1 = new DocPartGallery() { Val = "Page Numbers (Bottom of Page)" };
                DocPartUnique docPartUnique1 = new DocPartUnique();

                sdtContentDocPartObject1.Append(docPartGallery1);
                sdtContentDocPartObject1.Append(docPartUnique1);

                sdtProperties1.Append(sdtId1);
                sdtProperties1.Append(sdtContentDocPartObject1);

                SdtContentBlock sdtContentBlock1 = new SdtContentBlock();

                Paragraph paragraph1 = new Paragraph() { RsidParagraphMarkRevision = "008F4CDC", RsidParagraphAddition = "00E11B1E", RsidParagraphProperties = "008F4CDC", RsidRunAdditionDefault = "0031701D", ParagraphId = "6DDE6F19", TextId = "5A6BE73D" };

                ParagraphProperties paragraphProperties1 = new ParagraphProperties();
                ParagraphStyleId paragraphStyleId1 = new ParagraphStyleId() { Val = "Footer" };

                paragraphProperties1.Append(paragraphStyleId1);

                Run run1 = new Run();

                RunProperties runProperties1 = new RunProperties();
                RunFonts runFonts1 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
                NoProof noProof1 = new NoProof();
                FontSize fontSize1 = new FontSize() { Val = "28" };
                FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "28" };

                runProperties1.Append(runFonts1);
                runProperties1.Append(noProof1);
                runProperties1.Append(fontSize1);
                runProperties1.Append(fontSizeComplexScript1);

                AlternateContent alternateContent1 = new AlternateContent();

                AlternateContentChoice alternateContentChoice1 = new AlternateContentChoice() { Requires = "wps" };

                Drawing drawing1 = new Drawing();

                Wp.Anchor anchor1 = new Wp.Anchor() { DistanceFromTop = (UInt32Value)0U, DistanceFromBottom = (UInt32Value)0U, DistanceFromLeft = (UInt32Value)114300U, DistanceFromRight = (UInt32Value)114300U, SimplePos = false, RelativeHeight = (UInt32Value)251659264U, BehindDoc = false, Locked = false, LayoutInCell = true, AllowOverlap = true, EditId = "549845CB", AnchorId = "624E7F1E" };
                Wp.SimplePosition simplePosition1 = new Wp.SimplePosition() { X = 0L, Y = 0L };

                Wp.HorizontalPosition horizontalPosition1 = new Wp.HorizontalPosition() { RelativeFrom = Wp.HorizontalRelativePositionValues.LeftMargin };
                Wp.HorizontalAlignment horizontalAlignment1 = new Wp.HorizontalAlignment();
                horizontalAlignment1.Text = "center";

                horizontalPosition1.Append(horizontalAlignment1);

                Wp.VerticalPosition verticalPosition1 = new Wp.VerticalPosition() { RelativeFrom = Wp.VerticalRelativePositionValues.BottomMargin };
                Wp.VerticalAlignment verticalAlignment1 = new Wp.VerticalAlignment();
                verticalAlignment1.Text = "center";

                verticalPosition1.Append(verticalAlignment1);
                Wp.Extent extent1 = new Wp.Extent() { Cx = 512445L, Cy = 441325L };
                Wp.EffectExtent effectExtent1 = new Wp.EffectExtent() { LeftEdge = 0L, TopEdge = 0L, RightEdge = 1905L, BottomEdge = 0L };
                Wp.WrapNone wrapNone1 = new Wp.WrapNone();
                Wp.DocProperties docProperties1 = new Wp.DocProperties() { Id = (UInt32Value)34U, Name = "Flowchart: Alternate Process 34" };

                Wp.NonVisualGraphicFrameDrawingProperties nonVisualGraphicFrameDrawingProperties1 = new Wp.NonVisualGraphicFrameDrawingProperties();

                A.GraphicFrameLocks graphicFrameLocks1 = new A.GraphicFrameLocks();
                graphicFrameLocks1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

                nonVisualGraphicFrameDrawingProperties1.Append(graphicFrameLocks1);

                A.Graphic graphic1 = new A.Graphic();
                graphic1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

                A.GraphicData graphicData1 = new A.GraphicData() { Uri = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape" };

                Wps.WordprocessingShape wordprocessingShape1 = new Wps.WordprocessingShape();

                Wps.NonVisualDrawingShapeProperties nonVisualDrawingShapeProperties1 = new Wps.NonVisualDrawingShapeProperties();
                A.ShapeLocks shapeLocks1 = new A.ShapeLocks() { NoChangeArrowheads = true };

                nonVisualDrawingShapeProperties1.Append(shapeLocks1);

                Wps.ShapeProperties shapeProperties1 = new Wps.ShapeProperties() { BlackWhiteMode = A.BlackWhiteModeValues.Auto };

                A.Transform2D transform2D1 = new A.Transform2D();
                A.Offset offset1 = new A.Offset() { X = 0L, Y = 0L };
                A.Extents extents1 = new A.Extents() { Cx = 512445L, Cy = 441325L };

                transform2D1.Append(offset1);
                transform2D1.Append(extents1);

                A.PresetGeometry presetGeometry1 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.FlowChartAlternateProcess };
                A.AdjustValueList adjustValueList1 = new A.AdjustValueList();

                presetGeometry1.Append(adjustValueList1);
                A.NoFill noFill1 = new A.NoFill();

                A.Outline outline1 = new A.Outline();
                A.NoFill noFill2 = new A.NoFill();

                outline1.Append(noFill2);

                A.ShapePropertiesExtensionList shapePropertiesExtensionList1 = new A.ShapePropertiesExtensionList();

                A.ShapePropertiesExtension shapePropertiesExtension1 = new A.ShapePropertiesExtension() { Uri = "{909E8E84-426E-40DD-AFC4-6F175D3DCCD1}" };

                A14.HiddenFillProperties hiddenFillProperties1 = new A14.HiddenFillProperties();
                hiddenFillProperties1.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");

                A.SolidFill solidFill1 = new A.SolidFill();
                A.RgbColorModelHex rgbColorModelHex1 = new A.RgbColorModelHex() { Val = "5C83B4" };

                solidFill1.Append(rgbColorModelHex1);

                hiddenFillProperties1.Append(solidFill1);

                shapePropertiesExtension1.Append(hiddenFillProperties1);

                A.ShapePropertiesExtension shapePropertiesExtension2 = new A.ShapePropertiesExtension() { Uri = "{91240B29-F687-4F45-9708-019B960494DF}" };

                A14.HiddenLineProperties hiddenLineProperties1 = new A14.HiddenLineProperties() { Width = 9525 };
                hiddenLineProperties1.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");

                A.SolidFill solidFill2 = new A.SolidFill();
                A.RgbColorModelHex rgbColorModelHex2 = new A.RgbColorModelHex() { Val = "737373" };

                solidFill2.Append(rgbColorModelHex2);
                A.Miter miter1 = new A.Miter() { Limit = 800000 };
                A.HeadEnd headEnd1 = new A.HeadEnd();
                A.TailEnd tailEnd1 = new A.TailEnd();

                hiddenLineProperties1.Append(solidFill2);
                hiddenLineProperties1.Append(miter1);
                hiddenLineProperties1.Append(headEnd1);
                hiddenLineProperties1.Append(tailEnd1);

                shapePropertiesExtension2.Append(hiddenLineProperties1);

                shapePropertiesExtensionList1.Append(shapePropertiesExtension1);
                shapePropertiesExtensionList1.Append(shapePropertiesExtension2);

                shapeProperties1.Append(transform2D1);
                shapeProperties1.Append(presetGeometry1);
                shapeProperties1.Append(noFill1);
                shapeProperties1.Append(outline1);
                shapeProperties1.Append(shapePropertiesExtensionList1);

                Wps.TextBoxInfo2 textBoxInfo21 = new Wps.TextBoxInfo2();

                TextBoxContent textBoxContent1 = new TextBoxContent();

                Paragraph paragraph2 = new Paragraph() { RsidParagraphAddition = "0031701D", RsidRunAdditionDefault = "0031701D", ParagraphId = "786F8CD1", TextId = "77777777" };

                ParagraphProperties paragraphProperties2 = new ParagraphProperties();
                ParagraphStyleId paragraphStyleId2 = new ParagraphStyleId() { Val = "Footer" };

                ParagraphBorders paragraphBorders1 = new ParagraphBorders();
                TopBorder topBorder1 = new TopBorder() { Val = BorderValues.Single, Color = "A5A5A5", ThemeColor = ThemeColorValues.Accent3, Size = (UInt32Value)12U, Space = (UInt32Value)1U };
                BottomBorder bottomBorder1 = new BottomBorder() { Val = BorderValues.Single, Color = "A5A5A5", ThemeColor = ThemeColorValues.Accent3, Size = (UInt32Value)48U, Space = (UInt32Value)1U };

                paragraphBorders1.Append(topBorder1);
                paragraphBorders1.Append(bottomBorder1);
                Justification justification1 = new Justification() { Val = JustificationValues.Center };

                ParagraphMarkRunProperties paragraphMarkRunProperties1 = new ParagraphMarkRunProperties();
                FontSize fontSize2 = new FontSize() { Val = "28" };
                FontSizeComplexScript fontSizeComplexScript2 = new FontSizeComplexScript() { Val = "28" };

                paragraphMarkRunProperties1.Append(fontSize2);
                paragraphMarkRunProperties1.Append(fontSizeComplexScript2);

                paragraphProperties2.Append(paragraphStyleId2);
                paragraphProperties2.Append(paragraphBorders1);
                paragraphProperties2.Append(justification1);
                paragraphProperties2.Append(paragraphMarkRunProperties1);

                Run run2 = new Run();
                FieldChar fieldChar1 = new FieldChar() { FieldCharType = FieldCharValues.Begin };

                run2.Append(fieldChar1);

                Run run3 = new Run();
                FieldCode fieldCode1 = new FieldCode() { Space = SpaceProcessingModeValues.Preserve };
                fieldCode1.Text = " PAGE    \\* MERGEFORMAT ";

                run3.Append(fieldCode1);

                Run run4 = new Run();
                FieldChar fieldChar2 = new FieldChar() { FieldCharType = FieldCharValues.Separate };

                run4.Append(fieldChar2);

                Run run5 = new Run();

                RunProperties runProperties2 = new RunProperties();
                NoProof noProof2 = new NoProof();
                FontSize fontSize3 = new FontSize() { Val = "28" };
                FontSizeComplexScript fontSizeComplexScript3 = new FontSizeComplexScript() { Val = "28" };

                runProperties2.Append(noProof2);
                runProperties2.Append(fontSize3);
                runProperties2.Append(fontSizeComplexScript3);
                Text text1 = new Text();
                text1.Text = "2";

                run5.Append(runProperties2);
                run5.Append(text1);

                Run run6 = new Run();

                RunProperties runProperties3 = new RunProperties();
                NoProof noProof3 = new NoProof();
                FontSize fontSize4 = new FontSize() { Val = "28" };
                FontSizeComplexScript fontSizeComplexScript4 = new FontSizeComplexScript() { Val = "28" };

                runProperties3.Append(noProof3);
                runProperties3.Append(fontSize4);
                runProperties3.Append(fontSizeComplexScript4);
                FieldChar fieldChar3 = new FieldChar() { FieldCharType = FieldCharValues.End };

                run6.Append(runProperties3);
                run6.Append(fieldChar3);

                paragraph2.Append(paragraphProperties2);
                paragraph2.Append(run2);
                paragraph2.Append(run3);
                paragraph2.Append(run4);
                paragraph2.Append(run5);
                paragraph2.Append(run6);

                textBoxContent1.Append(paragraph2);

                textBoxInfo21.Append(textBoxContent1);

                Wps.TextBodyProperties textBodyProperties1 = new Wps.TextBodyProperties() { Rotation = 0, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, LeftInset = 91440, TopInset = 45720, RightInset = 91440, BottomInset = 45720, Anchor = A.TextAnchoringTypeValues.Top, AnchorCenter = false, UpRight = true };
                A.NoAutoFit noAutoFit1 = new A.NoAutoFit();

                textBodyProperties1.Append(noAutoFit1);

                wordprocessingShape1.Append(nonVisualDrawingShapeProperties1);
                wordprocessingShape1.Append(shapeProperties1);
                wordprocessingShape1.Append(textBoxInfo21);
                wordprocessingShape1.Append(textBodyProperties1);

                graphicData1.Append(wordprocessingShape1);

                graphic1.Append(graphicData1);

                Wp14.RelativeWidth relativeWidth1 = new Wp14.RelativeWidth() { ObjectId = Wp14.SizeRelativeHorizontallyValues.Page };
                Wp14.PercentageWidth percentageWidth1 = new Wp14.PercentageWidth();
                percentageWidth1.Text = "0";

                relativeWidth1.Append(percentageWidth1);

                Wp14.RelativeHeight relativeHeight1 = new Wp14.RelativeHeight() { RelativeFrom = Wp14.SizeRelativeVerticallyValues.Page };
                Wp14.PercentageHeight percentageHeight1 = new Wp14.PercentageHeight();
                percentageHeight1.Text = "0";

                relativeHeight1.Append(percentageHeight1);

                anchor1.Append(simplePosition1);
                anchor1.Append(horizontalPosition1);
                anchor1.Append(verticalPosition1);
                anchor1.Append(extent1);
                anchor1.Append(effectExtent1);
                anchor1.Append(wrapNone1);
                anchor1.Append(docProperties1);
                anchor1.Append(nonVisualGraphicFrameDrawingProperties1);
                anchor1.Append(graphic1);
                anchor1.Append(relativeWidth1);
                anchor1.Append(relativeHeight1);

                drawing1.Append(anchor1);

                alternateContentChoice1.Append(drawing1);

                AlternateContentFallback alternateContentFallback1 = new AlternateContentFallback();

                Picture picture1 = new Picture();

                V.Shapetype shapetype1 = new V.Shapetype() { Id = "_x0000_t176", CoordinateSize = "21600,21600", OptionalNumber = 176, Adjustment = "2700", EdgePath = "m@0,qx0@0l0@2qy@0,21600l@1,21600qx21600@2l21600@0qy@1,xe" };
                shapetype1.SetAttribute(new OpenXmlAttribute("w14", "anchorId", "http://schemas.microsoft.com/office/word/2010/wordml", "624E7F1E"));
                V.Stroke stroke1 = new V.Stroke() { JoinStyle = V.StrokeJoinStyleValues.Miter };

                V.Formulas formulas1 = new V.Formulas();
                V.Formula formula1 = new V.Formula() { Equation = "val #0" };
                V.Formula formula2 = new V.Formula() { Equation = "sum width 0 #0" };
                V.Formula formula3 = new V.Formula() { Equation = "sum height 0 #0" };
                V.Formula formula4 = new V.Formula() { Equation = "prod @0 2929 10000" };
                V.Formula formula5 = new V.Formula() { Equation = "sum width 0 @3" };
                V.Formula formula6 = new V.Formula() { Equation = "sum height 0 @3" };
                V.Formula formula7 = new V.Formula() { Equation = "val width" };
                V.Formula formula8 = new V.Formula() { Equation = "val height" };
                V.Formula formula9 = new V.Formula() { Equation = "prod width 1 2" };
                V.Formula formula10 = new V.Formula() { Equation = "prod height 1 2" };

                formulas1.Append(formula1);
                formulas1.Append(formula2);
                formulas1.Append(formula3);
                formulas1.Append(formula4);
                formulas1.Append(formula5);
                formulas1.Append(formula6);
                formulas1.Append(formula7);
                formulas1.Append(formula8);
                formulas1.Append(formula9);
                formulas1.Append(formula10);
                V.Path path1 = new V.Path() { Limo = "10800,10800", TextboxRectangle = "@3,@3,@4,@5", AllowGradientShape = true, ConnectionPointType = Ovml.ConnectValues.Custom, ConnectionPoints = "@8,0;0,@9;@8,@7;@6,@9" };

                shapetype1.Append(stroke1);
                shapetype1.Append(formulas1);
                shapetype1.Append(path1);

                V.Shape shape1 = new V.Shape() { Id = "Flowchart: Alternate Process 34", Style = "position:absolute;margin-left:0;margin-top:0;width:40.35pt;height:34.75pt;z-index:251659264;visibility:visible;mso-wrap-style:square;mso-width-percent:0;mso-height-percent:0;mso-wrap-distance-left:9pt;mso-wrap-distance-top:0;mso-wrap-distance-right:9pt;mso-wrap-distance-bottom:0;mso-position-horizontal:center;mso-position-horizontal-relative:left-margin-area;mso-position-vertical:center;mso-position-vertical-relative:bottom-margin-area;mso-width-percent:0;mso-height-percent:0;mso-width-relative:page;mso-height-relative:page;v-text-anchor:top", OptionalString = "_x0000_s1026", Filled = false, FillColor = "#5c83b4", Stroked = false, StrokeColor = "#737373", Type = "#_x0000_t176", EncodedPackage = "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\nAAAABAAEAPMAAABEBQAAAAA=\n" };

                V.TextBox textBox1 = new V.TextBox();

                TextBoxContent textBoxContent2 = new TextBoxContent();

                Paragraph paragraph3 = new Paragraph() { RsidParagraphAddition = "0031701D", RsidRunAdditionDefault = "0031701D", ParagraphId = "786F8CD1", TextId = "77777777" };

                ParagraphProperties paragraphProperties3 = new ParagraphProperties();
                ParagraphStyleId paragraphStyleId3 = new ParagraphStyleId() { Val = "Footer" };

                ParagraphBorders paragraphBorders2 = new ParagraphBorders();
                TopBorder topBorder2 = new TopBorder() { Val = BorderValues.Single, Color = "A5A5A5", ThemeColor = ThemeColorValues.Accent3, Size = (UInt32Value)12U, Space = (UInt32Value)1U };
                BottomBorder bottomBorder2 = new BottomBorder() { Val = BorderValues.Single, Color = "A5A5A5", ThemeColor = ThemeColorValues.Accent3, Size = (UInt32Value)48U, Space = (UInt32Value)1U };

                paragraphBorders2.Append(topBorder2);
                paragraphBorders2.Append(bottomBorder2);
                Justification justification2 = new Justification() { Val = JustificationValues.Center };

                ParagraphMarkRunProperties paragraphMarkRunProperties2 = new ParagraphMarkRunProperties();
                FontSize fontSize5 = new FontSize() { Val = "28" };
                FontSizeComplexScript fontSizeComplexScript5 = new FontSizeComplexScript() { Val = "28" };

                paragraphMarkRunProperties2.Append(fontSize5);
                paragraphMarkRunProperties2.Append(fontSizeComplexScript5);

                paragraphProperties3.Append(paragraphStyleId3);
                paragraphProperties3.Append(paragraphBorders2);
                paragraphProperties3.Append(justification2);
                paragraphProperties3.Append(paragraphMarkRunProperties2);

                Run run7 = new Run();
                FieldChar fieldChar4 = new FieldChar() { FieldCharType = FieldCharValues.Begin };

                run7.Append(fieldChar4);

                Run run8 = new Run();
                FieldCode fieldCode2 = new FieldCode() { Space = SpaceProcessingModeValues.Preserve };
                fieldCode2.Text = " PAGE    \\* MERGEFORMAT ";

                run8.Append(fieldCode2);

                Run run9 = new Run();
                FieldChar fieldChar5 = new FieldChar() { FieldCharType = FieldCharValues.Separate };

                run9.Append(fieldChar5);

                Run run10 = new Run();

                RunProperties runProperties4 = new RunProperties();
                NoProof noProof4 = new NoProof();
                FontSize fontSize6 = new FontSize() { Val = "28" };
                FontSizeComplexScript fontSizeComplexScript6 = new FontSizeComplexScript() { Val = "28" };

                runProperties4.Append(noProof4);
                runProperties4.Append(fontSize6);
                runProperties4.Append(fontSizeComplexScript6);
                Text text2 = new Text();
                text2.Text = "2";

                run10.Append(runProperties4);
                run10.Append(text2);

                Run run11 = new Run();

                RunProperties runProperties5 = new RunProperties();
                NoProof noProof5 = new NoProof();
                FontSize fontSize7 = new FontSize() { Val = "28" };
                FontSizeComplexScript fontSizeComplexScript7 = new FontSizeComplexScript() { Val = "28" };

                runProperties5.Append(noProof5);
                runProperties5.Append(fontSize7);
                runProperties5.Append(fontSizeComplexScript7);
                FieldChar fieldChar6 = new FieldChar() { FieldCharType = FieldCharValues.End };

                run11.Append(runProperties5);
                run11.Append(fieldChar6);

                paragraph3.Append(paragraphProperties3);
                paragraph3.Append(run7);
                paragraph3.Append(run8);
                paragraph3.Append(run9);
                paragraph3.Append(run10);
                paragraph3.Append(run11);

                textBoxContent2.Append(paragraph3);

                textBox1.Append(textBoxContent2);
                Wvml.TextWrap textWrap1 = new Wvml.TextWrap() { AnchorX = Wvml.HorizontalAnchorValues.Margin, AnchorY = Wvml.VerticalAnchorValues.Margin };

                shape1.Append(textBox1);
                shape1.Append(textWrap1);

                picture1.Append(shapetype1);
                picture1.Append(shape1);

                alternateContentFallback1.Append(picture1);

                alternateContent1.Append(alternateContentChoice1);
                alternateContent1.Append(alternateContentFallback1);

                run1.Append(runProperties1);
                run1.Append(alternateContent1);

                paragraph1.Append(paragraphProperties1);
                paragraph1.Append(run1);

                sdtContentBlock1.Append(paragraph1);

                sdtBlock1.Append(sdtProperties1);
                sdtBlock1.Append(sdtContentBlock1);
                return sdtBlock1;

            }
        }
        private static SdtBlock TopLine1 {
            get {
                SdtBlock sdtBlock1 = new SdtBlock();

                SdtProperties sdtProperties1 = new SdtProperties();
                SdtId sdtId1 = new SdtId() { Val = 1362552636 };

                SdtContentDocPartObject sdtContentDocPartObject1 = new SdtContentDocPartObject();
                DocPartGallery docPartGallery1 = new DocPartGallery() { Val = "Page Numbers (Bottom of Page)" };
                DocPartUnique docPartUnique1 = new DocPartUnique();

                sdtContentDocPartObject1.Append(docPartGallery1);
                sdtContentDocPartObject1.Append(docPartUnique1);

                sdtProperties1.Append(sdtId1);
                sdtProperties1.Append(sdtContentDocPartObject1);

                SdtContentBlock sdtContentBlock1 = new SdtContentBlock();

                Paragraph paragraph1 = new Paragraph() { RsidParagraphMarkRevision = "008F4CDC", RsidParagraphAddition = "00E11B1E", RsidParagraphProperties = "008F4CDC", RsidRunAdditionDefault = "008F4CDC", ParagraphId = "6DDE6F19", TextId = "4285F883" };

                ParagraphProperties paragraphProperties1 = new ParagraphProperties();
                ParagraphStyleId paragraphStyleId1 = new ParagraphStyleId() { Val = "Footer" };

                paragraphProperties1.Append(paragraphStyleId1);

                Run run1 = new Run();

                RunProperties runProperties1 = new RunProperties();
                NoProof noProof1 = new NoProof();

                runProperties1.Append(noProof1);

                AlternateContent alternateContent1 = new AlternateContent();

                AlternateContentChoice alternateContentChoice1 = new AlternateContentChoice() { Requires = "wps" };

                Drawing drawing1 = new Drawing();

                Wp.Anchor anchor1 = new Wp.Anchor() { DistanceFromTop = (UInt32Value)0U, DistanceFromBottom = (UInt32Value)0U, DistanceFromLeft = (UInt32Value)114300U, DistanceFromRight = (UInt32Value)114300U, SimplePos = false, RelativeHeight = (UInt32Value)251659264U, BehindDoc = false, Locked = false, LayoutInCell = true, AllowOverlap = true, EditId = "034C1953", AnchorId = "5EF200D0" };
                Wp.SimplePosition simplePosition1 = new Wp.SimplePosition() { X = 0L, Y = 0L };

                Wp.HorizontalPosition horizontalPosition1 = new Wp.HorizontalPosition() { RelativeFrom = Wp.HorizontalRelativePositionValues.LeftMargin };
                Wp.HorizontalAlignment horizontalAlignment1 = new Wp.HorizontalAlignment();
                horizontalAlignment1.Text = "center";

                horizontalPosition1.Append(horizontalAlignment1);

                Wp.VerticalPosition verticalPosition1 = new Wp.VerticalPosition() { RelativeFrom = Wp.VerticalRelativePositionValues.BottomMargin };
                Wp.VerticalAlignment verticalAlignment1 = new Wp.VerticalAlignment();
                verticalAlignment1.Text = "center";

                verticalPosition1.Append(verticalAlignment1);
                Wp.Extent extent1 = new Wp.Extent() { Cx = 565785L, Cy = 191770L };
                Wp.EffectExtent effectExtent1 = new Wp.EffectExtent() { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L };
                Wp.WrapNone wrapNone1 = new Wp.WrapNone();
                Wp.DocProperties docProperties1 = new Wp.DocProperties() { Id = (UInt32Value)33U, Name = "Rectangle 33" };

                Wp.NonVisualGraphicFrameDrawingProperties nonVisualGraphicFrameDrawingProperties1 = new Wp.NonVisualGraphicFrameDrawingProperties();

                A.GraphicFrameLocks graphicFrameLocks1 = new A.GraphicFrameLocks();
                graphicFrameLocks1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

                nonVisualGraphicFrameDrawingProperties1.Append(graphicFrameLocks1);

                A.Graphic graphic1 = new A.Graphic();
                graphic1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

                A.GraphicData graphicData1 = new A.GraphicData() { Uri = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape" };

                Wps.WordprocessingShape wordprocessingShape1 = new Wps.WordprocessingShape();

                Wps.NonVisualDrawingShapeProperties nonVisualDrawingShapeProperties1 = new Wps.NonVisualDrawingShapeProperties();
                A.ShapeLocks shapeLocks1 = new A.ShapeLocks() { NoChangeArrowheads = true };

                nonVisualDrawingShapeProperties1.Append(shapeLocks1);

                Wps.ShapeProperties shapeProperties1 = new Wps.ShapeProperties() { BlackWhiteMode = A.BlackWhiteModeValues.Auto };

                A.Transform2D transform2D1 = new A.Transform2D() { Rotation = 10800000, HorizontalFlip = true };
                A.Offset offset1 = new A.Offset() { X = 0L, Y = 0L };
                A.Extents extents1 = new A.Extents() { Cx = 565785L, Cy = 191770L };

                transform2D1.Append(offset1);
                transform2D1.Append(extents1);

                A.PresetGeometry presetGeometry1 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
                A.AdjustValueList adjustValueList1 = new A.AdjustValueList();

                presetGeometry1.Append(adjustValueList1);
                A.NoFill noFill1 = new A.NoFill();

                A.Outline outline1 = new A.Outline();
                A.NoFill noFill2 = new A.NoFill();

                outline1.Append(noFill2);

                A.ShapePropertiesExtensionList shapePropertiesExtensionList1 = new A.ShapePropertiesExtensionList();

                A.ShapePropertiesExtension shapePropertiesExtension1 = new A.ShapePropertiesExtension() { Uri = "{909E8E84-426E-40DD-AFC4-6F175D3DCCD1}" };

                A14.HiddenFillProperties hiddenFillProperties1 = new A14.HiddenFillProperties();
                hiddenFillProperties1.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");

                A.SolidFill solidFill1 = new A.SolidFill();
                A.RgbColorModelHex rgbColorModelHex1 = new A.RgbColorModelHex() { Val = "C0504D" };

                solidFill1.Append(rgbColorModelHex1);

                hiddenFillProperties1.Append(solidFill1);

                shapePropertiesExtension1.Append(hiddenFillProperties1);

                A.ShapePropertiesExtension shapePropertiesExtension2 = new A.ShapePropertiesExtension() { Uri = "{91240B29-F687-4F45-9708-019B960494DF}" };

                A14.HiddenLineProperties hiddenLineProperties1 = new A14.HiddenLineProperties() { Width = 28575 };
                hiddenLineProperties1.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");

                A.SolidFill solidFill2 = new A.SolidFill();
                A.RgbColorModelHex rgbColorModelHex2 = new A.RgbColorModelHex() { Val = "5C83B4" };

                solidFill2.Append(rgbColorModelHex2);
                A.Miter miter1 = new A.Miter() { Limit = 800000 };
                A.HeadEnd headEnd1 = new A.HeadEnd();
                A.TailEnd tailEnd1 = new A.TailEnd();

                hiddenLineProperties1.Append(solidFill2);
                hiddenLineProperties1.Append(miter1);
                hiddenLineProperties1.Append(headEnd1);
                hiddenLineProperties1.Append(tailEnd1);

                shapePropertiesExtension2.Append(hiddenLineProperties1);

                shapePropertiesExtensionList1.Append(shapePropertiesExtension1);
                shapePropertiesExtensionList1.Append(shapePropertiesExtension2);

                shapeProperties1.Append(transform2D1);
                shapeProperties1.Append(presetGeometry1);
                shapeProperties1.Append(noFill1);
                shapeProperties1.Append(outline1);
                shapeProperties1.Append(shapePropertiesExtensionList1);

                Wps.TextBoxInfo2 textBoxInfo21 = new Wps.TextBoxInfo2();

                TextBoxContent textBoxContent1 = new TextBoxContent();

                Paragraph paragraph2 = new Paragraph() { RsidParagraphAddition = "008F4CDC", RsidRunAdditionDefault = "008F4CDC", ParagraphId = "051608B4", TextId = "77777777" };

                ParagraphProperties paragraphProperties2 = new ParagraphProperties();

                ParagraphBorders paragraphBorders1 = new ParagraphBorders();
                TopBorder topBorder1 = new TopBorder() { Val = BorderValues.Single, Color = "7F7F7F", ThemeColor = ThemeColorValues.Background1, ThemeShade = "7F", Size = (UInt32Value)4U, Space = (UInt32Value)1U };

                paragraphBorders1.Append(topBorder1);
                Justification justification1 = new Justification() { Val = JustificationValues.Center };

                ParagraphMarkRunProperties paragraphMarkRunProperties1 = new ParagraphMarkRunProperties();
                Color color1 = new Color() { Val = "ED7D31", ThemeColor = ThemeColorValues.Accent2 };

                paragraphMarkRunProperties1.Append(color1);

                paragraphProperties2.Append(paragraphBorders1);
                paragraphProperties2.Append(justification1);
                paragraphProperties2.Append(paragraphMarkRunProperties1);

                Run run2 = new Run();
                FieldChar fieldChar1 = new FieldChar() { FieldCharType = FieldCharValues.Begin };

                run2.Append(fieldChar1);

                Run run3 = new Run();
                FieldCode fieldCode1 = new FieldCode() { Space = SpaceProcessingModeValues.Preserve };
                fieldCode1.Text = " PAGE   \\* MERGEFORMAT ";

                run3.Append(fieldCode1);

                Run run4 = new Run();
                FieldChar fieldChar2 = new FieldChar() { FieldCharType = FieldCharValues.Separate };

                run4.Append(fieldChar2);

                Run run5 = new Run();

                RunProperties runProperties2 = new RunProperties();
                NoProof noProof2 = new NoProof();
                Color color2 = new Color() { Val = "ED7D31", ThemeColor = ThemeColorValues.Accent2 };

                runProperties2.Append(noProof2);
                runProperties2.Append(color2);
                Text text1 = new Text();
                text1.Text = "2";

                run5.Append(runProperties2);
                run5.Append(text1);

                Run run6 = new Run();

                RunProperties runProperties3 = new RunProperties();
                NoProof noProof3 = new NoProof();
                Color color3 = new Color() { Val = "ED7D31", ThemeColor = ThemeColorValues.Accent2 };

                runProperties3.Append(noProof3);
                runProperties3.Append(color3);
                FieldChar fieldChar3 = new FieldChar() { FieldCharType = FieldCharValues.End };

                run6.Append(runProperties3);
                run6.Append(fieldChar3);

                paragraph2.Append(paragraphProperties2);
                paragraph2.Append(run2);
                paragraph2.Append(run3);
                paragraph2.Append(run4);
                paragraph2.Append(run5);
                paragraph2.Append(run6);

                textBoxContent1.Append(paragraph2);

                textBoxInfo21.Append(textBoxContent1);

                Wps.TextBodyProperties textBodyProperties1 = new Wps.TextBodyProperties() { Rotation = 0, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, LeftInset = 91440, TopInset = 0, RightInset = 91440, BottomInset = 0, Anchor = A.TextAnchoringTypeValues.Top, AnchorCenter = false, UpRight = true };
                A.NoAutoFit noAutoFit1 = new A.NoAutoFit();

                textBodyProperties1.Append(noAutoFit1);

                wordprocessingShape1.Append(nonVisualDrawingShapeProperties1);
                wordprocessingShape1.Append(shapeProperties1);
                wordprocessingShape1.Append(textBoxInfo21);
                wordprocessingShape1.Append(textBodyProperties1);

                graphicData1.Append(wordprocessingShape1);

                graphic1.Append(graphicData1);

                Wp14.RelativeWidth relativeWidth1 = new Wp14.RelativeWidth() { ObjectId = Wp14.SizeRelativeHorizontallyValues.Page };
                Wp14.PercentageWidth percentageWidth1 = new Wp14.PercentageWidth();
                percentageWidth1.Text = "0";

                relativeWidth1.Append(percentageWidth1);

                Wp14.RelativeHeight relativeHeight1 = new Wp14.RelativeHeight() { RelativeFrom = Wp14.SizeRelativeVerticallyValues.BottomMargin };
                Wp14.PercentageHeight percentageHeight1 = new Wp14.PercentageHeight();
                percentageHeight1.Text = "0";

                relativeHeight1.Append(percentageHeight1);

                anchor1.Append(simplePosition1);
                anchor1.Append(horizontalPosition1);
                anchor1.Append(verticalPosition1);
                anchor1.Append(extent1);
                anchor1.Append(effectExtent1);
                anchor1.Append(wrapNone1);
                anchor1.Append(docProperties1);
                anchor1.Append(nonVisualGraphicFrameDrawingProperties1);
                anchor1.Append(graphic1);
                anchor1.Append(relativeWidth1);
                anchor1.Append(relativeHeight1);

                drawing1.Append(anchor1);

                alternateContentChoice1.Append(drawing1);

                AlternateContentFallback alternateContentFallback1 = new AlternateContentFallback();

                Picture picture1 = new Picture();

                V.Rectangle rectangle1 = new V.Rectangle() { Id = "Rectangle 33", Style = "position:absolute;margin-left:0;margin-top:0;width:44.55pt;height:15.1pt;rotation:180;flip:x;z-index:251659264;visibility:visible;mso-wrap-style:square;mso-width-percent:0;mso-height-percent:0;mso-wrap-distance-left:9pt;mso-wrap-distance-top:0;mso-wrap-distance-right:9pt;mso-wrap-distance-bottom:0;mso-position-horizontal:center;mso-position-horizontal-relative:left-margin-area;mso-position-vertical:center;mso-position-vertical-relative:bottom-margin-area;mso-width-percent:0;mso-height-percent:0;mso-width-relative:page;mso-height-relative:bottom-margin-area;v-text-anchor:top", OptionalString = "_x0000_s1026", Filled = false, FillColor = "#c0504d", Stroked = false, StrokeColor = "#5c83b4", StrokeWeight = "2.25pt" };
                rectangle1.SetAttribute(new OpenXmlAttribute("w14", "anchorId", "http://schemas.microsoft.com/office/word/2010/wordml", "5EF200D0"));
                rectangle1.SetAttribute(new OpenXmlAttribute("o", "gfxdata", "urn:schemas-microsoft-com:office:office", "UEsDBBQABgAIAAAAIQC2gziS/gAAAOEBAAATAAAAW0NvbnRlbnRfVHlwZXNdLnhtbJSRQU7DMBBF\n90jcwfIWJU67QAgl6YK0S0CoHGBkTxKLZGx5TGhvj5O2G0SRWNoz/78nu9wcxkFMGNg6quQqL6RA\n0s5Y6ir5vt9lD1JwBDIwOMJKHpHlpr69KfdHjyxSmriSfYz+USnWPY7AufNIadK6MEJMx9ApD/oD\nOlTrorhX2lFEilmcO2RdNtjC5xDF9pCuTyYBB5bi6bQ4syoJ3g9WQ0ymaiLzg5KdCXlKLjvcW893\nSUOqXwnz5DrgnHtJTxOsQfEKIT7DmDSUCaxw7Rqn8787ZsmRM9e2VmPeBN4uqYvTtW7jvijg9N/y\nJsXecLq0q+WD6m8AAAD//wMAUEsDBBQABgAIAAAAIQA4/SH/1gAAAJQBAAALAAAAX3JlbHMvLnJl\nbHOkkMFqwzAMhu+DvYPRfXGawxijTi+j0GvpHsDYimMaW0Yy2fr2M4PBMnrbUb/Q94l/f/hMi1qR\nJVI2sOt6UJgd+ZiDgffL8ekFlFSbvV0oo4EbChzGx4f9GRdb25HMsYhqlCwG5lrLq9biZkxWOiqY\n22YiTra2kYMu1l1tQD30/bPm3wwYN0x18gb45AdQl1tp5j/sFB2T0FQ7R0nTNEV3j6o9feQzro1i\nOWA14Fm+Q8a1a8+Bvu/d/dMb2JY5uiPbhG/ktn4cqGU/er3pcvwCAAD//wMAUEsDBBQABgAIAAAA\nIQBG31uP5QEAAKcDAAAOAAAAZHJzL2Uyb0RvYy54bWysU9tu2zAMfR+wfxD0vtgumiY14hRFi24D\nugvQ9QNkWYqFyaJGKbGzrx+lZEm3vg3zg0BS1CHPIb26mQbLdgqDAdfwalZyppyEzrhNw5+/Pbxb\nchaicJ2w4FTD9yrwm/XbN6vR1+oCerCdQkYgLtSjb3gfo6+LIsheDSLMwCtHlxpwEJFc3BQdipHQ\nB1tclOVVMQJ2HkGqECh6f7jk64yvtZLxi9ZBRWYbTr3FfGI+23QW65WoNyh8b+SxDfEPXQzCOCp6\ngroXUbAtmldQg5EIAXScSRgK0NpIlTkQm6r8i81TL7zKXEic4E8yhf8HKz/vnvxXTK0H/wjye2AO\n7nrhNuoWEcZeiY7KVUmoYvShPj1ITqCnrB0/QUejFdsIWYNJ48AQSOuqXJbp40xb4z8knFSJaLMp\nz2B/moGaIpMUnF/NF8s5Z5KuqutqscgzKkSdUNNjjyG+VzCwZDQcacQZVOweQ0xdnlNSuoMHY20e\ns3V/BCgxRTKrRCTtTKjj1E6UncwWuj3xy0yIAm051esBf3I20sY0PPzYClSc2Y+ONLquLi/TimWH\nDHwZbX9HhZME0fDI2cG8i4d13Ho0mz6Jluk4uCU9tcmUzt0c+6VtyEyPm5vW7aWfs87/1/oXAAAA\n//8DAFBLAwQUAAYACAAAACEAI+V68dsAAAADAQAADwAAAGRycy9kb3ducmV2LnhtbEyPT0vDQBDF\n70K/wzIFb3bTVqSmmRQRBPFPo1U8b7PTJJidjdltG799Ry96GXi8x3u/yVaDa9WB+tB4RphOElDE\npbcNVwjvb3cXC1AhGram9UwI3xRglY/OMpNaf+RXOmxipaSEQ2oQ6hi7VOtQ1uRMmPiOWLyd752J\nIvtK294cpdy1epYkV9qZhmWhNh3d1lR+bvYOwX98Pdpi7Z61LtZP5f3l/OWhYMTz8XCzBBVpiH9h\n+MEXdMiFaev3bINqEeSR+HvFW1xPQW0R5skMdJ7p/+z5CQAA//8DAFBLAQItABQABgAIAAAAIQC2\ngziS/gAAAOEBAAATAAAAAAAAAAAAAAAAAAAAAABbQ29udGVudF9UeXBlc10ueG1sUEsBAi0AFAAG\nAAgAAAAhADj9If/WAAAAlAEAAAsAAAAAAAAAAAAAAAAALwEAAF9yZWxzLy5yZWxzUEsBAi0AFAAG\nAAgAAAAhAEbfW4/lAQAApwMAAA4AAAAAAAAAAAAAAAAALgIAAGRycy9lMm9Eb2MueG1sUEsBAi0A\nFAAGAAgAAAAhACPlevHbAAAAAwEAAA8AAAAAAAAAAAAAAAAAPwQAAGRycy9kb3ducmV2LnhtbFBL\nBQYAAAAABAAEAPMAAABHBQAAAAA=\n"));

                V.TextBox textBox1 = new V.TextBox() { Inset = ",0,,0" };

                TextBoxContent textBoxContent2 = new TextBoxContent();

                Paragraph paragraph3 = new Paragraph() { RsidParagraphAddition = "008F4CDC", RsidRunAdditionDefault = "008F4CDC", ParagraphId = "051608B4", TextId = "77777777" };

                ParagraphProperties paragraphProperties3 = new ParagraphProperties();

                ParagraphBorders paragraphBorders2 = new ParagraphBorders();
                TopBorder topBorder2 = new TopBorder() { Val = BorderValues.Single, Color = "7F7F7F", ThemeColor = ThemeColorValues.Background1, ThemeShade = "7F", Size = (UInt32Value)4U, Space = (UInt32Value)1U };

                paragraphBorders2.Append(topBorder2);
                Justification justification2 = new Justification() { Val = JustificationValues.Center };

                ParagraphMarkRunProperties paragraphMarkRunProperties2 = new ParagraphMarkRunProperties();
                Color color4 = new Color() { Val = "ED7D31", ThemeColor = ThemeColorValues.Accent2 };

                paragraphMarkRunProperties2.Append(color4);

                paragraphProperties3.Append(paragraphBorders2);
                paragraphProperties3.Append(justification2);
                paragraphProperties3.Append(paragraphMarkRunProperties2);

                Run run7 = new Run();
                FieldChar fieldChar4 = new FieldChar() { FieldCharType = FieldCharValues.Begin };

                run7.Append(fieldChar4);

                Run run8 = new Run();
                FieldCode fieldCode2 = new FieldCode() { Space = SpaceProcessingModeValues.Preserve };
                fieldCode2.Text = " PAGE   \\* MERGEFORMAT ";

                run8.Append(fieldCode2);

                Run run9 = new Run();
                FieldChar fieldChar5 = new FieldChar() { FieldCharType = FieldCharValues.Separate };

                run9.Append(fieldChar5);

                Run run10 = new Run();

                RunProperties runProperties4 = new RunProperties();
                NoProof noProof4 = new NoProof();
                Color color5 = new Color() { Val = "ED7D31", ThemeColor = ThemeColorValues.Accent2 };

                runProperties4.Append(noProof4);
                runProperties4.Append(color5);
                Text text2 = new Text();
                text2.Text = "2";

                run10.Append(runProperties4);
                run10.Append(text2);

                Run run11 = new Run();

                RunProperties runProperties5 = new RunProperties();
                NoProof noProof5 = new NoProof();
                Color color6 = new Color() { Val = "ED7D31", ThemeColor = ThemeColorValues.Accent2 };

                runProperties5.Append(noProof5);
                runProperties5.Append(color6);
                FieldChar fieldChar6 = new FieldChar() { FieldCharType = FieldCharValues.End };

                run11.Append(runProperties5);
                run11.Append(fieldChar6);

                paragraph3.Append(paragraphProperties3);
                paragraph3.Append(run7);
                paragraph3.Append(run8);
                paragraph3.Append(run9);
                paragraph3.Append(run10);
                paragraph3.Append(run11);

                textBoxContent2.Append(paragraph3);

                textBox1.Append(textBoxContent2);
                Wvml.TextWrap textWrap1 = new Wvml.TextWrap() { AnchorX = Wvml.HorizontalAnchorValues.Margin, AnchorY = Wvml.VerticalAnchorValues.Margin };

                rectangle1.Append(textBox1);
                rectangle1.Append(textWrap1);

                picture1.Append(rectangle1);

                alternateContentFallback1.Append(picture1);

                alternateContent1.Append(alternateContentChoice1);
                alternateContent1.Append(alternateContentFallback1);

                run1.Append(runProperties1);
                run1.Append(alternateContent1);

                paragraph1.Append(paragraphProperties1);
                paragraph1.Append(run1);

                sdtContentBlock1.Append(paragraph1);

                sdtBlock1.Append(sdtProperties1);
                sdtBlock1.Append(sdtContentBlock1);
                return sdtBlock1;

            }
        }
        private static SdtBlock Tab1 {
            get {
                SdtBlock sdtBlock1 = new SdtBlock();

                SdtProperties sdtProperties1 = new SdtProperties();
                SdtId sdtId1 = new SdtId() { Val = 1176225630 };

                SdtContentDocPartObject sdtContentDocPartObject1 = new SdtContentDocPartObject();
                DocPartGallery docPartGallery1 = new DocPartGallery() { Val = "Page Numbers (Bottom of Page)" };
                DocPartUnique docPartUnique1 = new DocPartUnique();

                sdtContentDocPartObject1.Append(docPartGallery1);
                sdtContentDocPartObject1.Append(docPartUnique1);

                sdtProperties1.Append(sdtId1);
                sdtProperties1.Append(sdtContentDocPartObject1);

                SdtContentBlock sdtContentBlock1 = new SdtContentBlock();

                Paragraph paragraph1 = new Paragraph() { RsidParagraphMarkRevision = "00354B38", RsidParagraphAddition = "00E11B1E", RsidParagraphProperties = "00354B38", RsidRunAdditionDefault = "00354B38", ParagraphId = "6DDE6F19", TextId = "66CB5F24" };

                ParagraphProperties paragraphProperties1 = new ParagraphProperties();
                ParagraphStyleId paragraphStyleId1 = new ParagraphStyleId() { Val = "Footer" };

                paragraphProperties1.Append(paragraphStyleId1);

                Run run1 = new Run();

                RunProperties runProperties1 = new RunProperties();
                NoProof noProof1 = new NoProof();

                runProperties1.Append(noProof1);

                AlternateContent alternateContent1 = new AlternateContent();

                AlternateContentChoice alternateContentChoice1 = new AlternateContentChoice() { Requires = "wpg" };

                Drawing drawing1 = new Drawing();

                Wp.Anchor anchor1 = new Wp.Anchor() { DistanceFromTop = (UInt32Value)0U, DistanceFromBottom = (UInt32Value)0U, DistanceFromLeft = (UInt32Value)114300U, DistanceFromRight = (UInt32Value)114300U, SimplePos = false, RelativeHeight = (UInt32Value)251659264U, BehindDoc = false, Locked = false, LayoutInCell = true, AllowOverlap = true, EditId = "253D75B2", AnchorId = "74BB3F01" };
                Wp.SimplePosition simplePosition1 = new Wp.SimplePosition() { X = 0L, Y = 0L };

                Wp.HorizontalPosition horizontalPosition1 = new Wp.HorizontalPosition() { RelativeFrom = Wp.HorizontalRelativePositionValues.Page };
                Wp.HorizontalAlignment horizontalAlignment1 = new Wp.HorizontalAlignment();
                horizontalAlignment1.Text = "center";

                horizontalPosition1.Append(horizontalAlignment1);

                Wp.VerticalPosition verticalPosition1 = new Wp.VerticalPosition() { RelativeFrom = Wp.VerticalRelativePositionValues.BottomMargin };
                Wp.VerticalAlignment verticalAlignment1 = new Wp.VerticalAlignment();
                verticalAlignment1.Text = "center";

                verticalPosition1.Append(verticalAlignment1);
                Wp.Extent extent1 = new Wp.Extent() { Cx = 7753350L, Cy = 190500L };
                Wp.EffectExtent effectExtent1 = new Wp.EffectExtent() { LeftEdge = 9525L, TopEdge = 9525L, RightEdge = 9525L, BottomEdge = 0L };
                Wp.WrapNone wrapNone1 = new Wp.WrapNone();
                Wp.DocProperties docProperties1 = new Wp.DocProperties() { Id = (UInt32Value)27U, Name = "Group 27" };

                Wp.NonVisualGraphicFrameDrawingProperties nonVisualGraphicFrameDrawingProperties1 = new Wp.NonVisualGraphicFrameDrawingProperties();

                A.GraphicFrameLocks graphicFrameLocks1 = new A.GraphicFrameLocks();
                graphicFrameLocks1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

                nonVisualGraphicFrameDrawingProperties1.Append(graphicFrameLocks1);

                A.Graphic graphic1 = new A.Graphic();
                graphic1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

                A.GraphicData graphicData1 = new A.GraphicData() { Uri = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" };

                Wpg.WordprocessingGroup wordprocessingGroup1 = new Wpg.WordprocessingGroup();

                Wpg.NonVisualGroupDrawingShapeProperties nonVisualGroupDrawingShapeProperties1 = new Wpg.NonVisualGroupDrawingShapeProperties();
                A.GroupShapeLocks groupShapeLocks1 = new A.GroupShapeLocks();

                nonVisualGroupDrawingShapeProperties1.Append(groupShapeLocks1);

                Wpg.GroupShapeProperties groupShapeProperties1 = new Wpg.GroupShapeProperties() { BlackWhiteMode = A.BlackWhiteModeValues.Auto };

                A.TransformGroup transformGroup1 = new A.TransformGroup();
                A.Offset offset1 = new A.Offset() { X = 0L, Y = 0L };
                A.Extents extents1 = new A.Extents() { Cx = 7753350L, Cy = 190500L };
                A.ChildOffset childOffset1 = new A.ChildOffset() { X = -8L, Y = 14978L };
                A.ChildExtents childExtents1 = new A.ChildExtents() { Cx = 12255L, Cy = 300L };

                transformGroup1.Append(offset1);
                transformGroup1.Append(extents1);
                transformGroup1.Append(childOffset1);
                transformGroup1.Append(childExtents1);

                groupShapeProperties1.Append(transformGroup1);

                Wps.WordprocessingShape wordprocessingShape1 = new Wps.WordprocessingShape();
                Wps.NonVisualDrawingProperties nonVisualDrawingProperties1 = new Wps.NonVisualDrawingProperties() { Id = (UInt32Value)28U, Name = "Text Box 25" };

                Wps.NonVisualDrawingShapeProperties nonVisualDrawingShapeProperties1 = new Wps.NonVisualDrawingShapeProperties() { TextBox = true };
                A.ShapeLocks shapeLocks1 = new A.ShapeLocks() { NoChangeArrowheads = true };

                nonVisualDrawingShapeProperties1.Append(shapeLocks1);

                Wps.ShapeProperties shapeProperties1 = new Wps.ShapeProperties() { BlackWhiteMode = A.BlackWhiteModeValues.Auto };

                A.Transform2D transform2D1 = new A.Transform2D();
                A.Offset offset2 = new A.Offset() { X = 782L, Y = 14990L };
                A.Extents extents2 = new A.Extents() { Cx = 659L, Cy = 288L };

                transform2D1.Append(offset2);
                transform2D1.Append(extents2);

                A.PresetGeometry presetGeometry1 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
                A.AdjustValueList adjustValueList1 = new A.AdjustValueList();

                presetGeometry1.Append(adjustValueList1);
                A.NoFill noFill1 = new A.NoFill();

                A.Outline outline1 = new A.Outline();
                A.NoFill noFill2 = new A.NoFill();

                outline1.Append(noFill2);

                A.ShapePropertiesExtensionList shapePropertiesExtensionList1 = new A.ShapePropertiesExtensionList();

                A.ShapePropertiesExtension shapePropertiesExtension1 = new A.ShapePropertiesExtension() { Uri = "{909E8E84-426E-40DD-AFC4-6F175D3DCCD1}" };

                A14.HiddenFillProperties hiddenFillProperties1 = new A14.HiddenFillProperties();
                hiddenFillProperties1.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");

                A.SolidFill solidFill1 = new A.SolidFill();
                A.RgbColorModelHex rgbColorModelHex1 = new A.RgbColorModelHex() { Val = "FFFFFF" };

                solidFill1.Append(rgbColorModelHex1);

                hiddenFillProperties1.Append(solidFill1);

                shapePropertiesExtension1.Append(hiddenFillProperties1);

                A.ShapePropertiesExtension shapePropertiesExtension2 = new A.ShapePropertiesExtension() { Uri = "{91240B29-F687-4F45-9708-019B960494DF}" };

                A14.HiddenLineProperties hiddenLineProperties1 = new A14.HiddenLineProperties() { Width = 9525 };
                hiddenLineProperties1.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");

                A.SolidFill solidFill2 = new A.SolidFill();
                A.RgbColorModelHex rgbColorModelHex2 = new A.RgbColorModelHex() { Val = "000000" };

                solidFill2.Append(rgbColorModelHex2);
                A.Miter miter1 = new A.Miter() { Limit = 800000 };
                A.HeadEnd headEnd1 = new A.HeadEnd();
                A.TailEnd tailEnd1 = new A.TailEnd();

                hiddenLineProperties1.Append(solidFill2);
                hiddenLineProperties1.Append(miter1);
                hiddenLineProperties1.Append(headEnd1);
                hiddenLineProperties1.Append(tailEnd1);

                shapePropertiesExtension2.Append(hiddenLineProperties1);

                shapePropertiesExtensionList1.Append(shapePropertiesExtension1);
                shapePropertiesExtensionList1.Append(shapePropertiesExtension2);

                shapeProperties1.Append(transform2D1);
                shapeProperties1.Append(presetGeometry1);
                shapeProperties1.Append(noFill1);
                shapeProperties1.Append(outline1);
                shapeProperties1.Append(shapePropertiesExtensionList1);

                Wps.TextBoxInfo2 textBoxInfo21 = new Wps.TextBoxInfo2();

                TextBoxContent textBoxContent1 = new TextBoxContent();

                Paragraph paragraph2 = new Paragraph() { RsidParagraphAddition = "00354B38", RsidRunAdditionDefault = "00354B38", ParagraphId = "4E6CF7E2", TextId = "77777777" };

                ParagraphProperties paragraphProperties2 = new ParagraphProperties();
                Justification justification1 = new Justification() { Val = JustificationValues.Center };

                paragraphProperties2.Append(justification1);

                Run run2 = new Run();
                FieldChar fieldChar1 = new FieldChar() { FieldCharType = FieldCharValues.Begin };

                run2.Append(fieldChar1);

                Run run3 = new Run();
                FieldCode fieldCode1 = new FieldCode() { Space = SpaceProcessingModeValues.Preserve };
                fieldCode1.Text = " PAGE    \\* MERGEFORMAT ";

                run3.Append(fieldCode1);

                Run run4 = new Run();
                FieldChar fieldChar2 = new FieldChar() { FieldCharType = FieldCharValues.Separate };

                run4.Append(fieldChar2);

                Run run5 = new Run();

                RunProperties runProperties2 = new RunProperties();
                NoProof noProof2 = new NoProof();
                Color color1 = new Color() { Val = "8C8C8C", ThemeColor = ThemeColorValues.Background1, ThemeShade = "8C" };

                runProperties2.Append(noProof2);
                runProperties2.Append(color1);
                Text text1 = new Text();
                text1.Text = "2";

                run5.Append(runProperties2);
                run5.Append(text1);

                Run run6 = new Run();

                RunProperties runProperties3 = new RunProperties();
                NoProof noProof3 = new NoProof();
                Color color2 = new Color() { Val = "8C8C8C", ThemeColor = ThemeColorValues.Background1, ThemeShade = "8C" };

                runProperties3.Append(noProof3);
                runProperties3.Append(color2);
                FieldChar fieldChar3 = new FieldChar() { FieldCharType = FieldCharValues.End };

                run6.Append(runProperties3);
                run6.Append(fieldChar3);

                paragraph2.Append(paragraphProperties2);
                paragraph2.Append(run2);
                paragraph2.Append(run3);
                paragraph2.Append(run4);
                paragraph2.Append(run5);
                paragraph2.Append(run6);

                textBoxContent1.Append(paragraph2);

                textBoxInfo21.Append(textBoxContent1);

                Wps.TextBodyProperties textBodyProperties1 = new Wps.TextBodyProperties() { Rotation = 0, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, LeftInset = 0, TopInset = 0, RightInset = 0, BottomInset = 0, Anchor = A.TextAnchoringTypeValues.Top, AnchorCenter = false, UpRight = true };
                A.NoAutoFit noAutoFit1 = new A.NoAutoFit();

                textBodyProperties1.Append(noAutoFit1);

                wordprocessingShape1.Append(nonVisualDrawingProperties1);
                wordprocessingShape1.Append(nonVisualDrawingShapeProperties1);
                wordprocessingShape1.Append(shapeProperties1);
                wordprocessingShape1.Append(textBoxInfo21);
                wordprocessingShape1.Append(textBodyProperties1);

                Wpg.GroupShape groupShape1 = new Wpg.GroupShape();
                Wpg.NonVisualDrawingProperties nonVisualDrawingProperties2 = new Wpg.NonVisualDrawingProperties() { Id = (UInt32Value)29U, Name = "Group 31" };

                Wpg.NonVisualGroupDrawingShapeProperties nonVisualGroupDrawingShapeProperties2 = new Wpg.NonVisualGroupDrawingShapeProperties();
                A.GroupShapeLocks groupShapeLocks2 = new A.GroupShapeLocks();

                nonVisualGroupDrawingShapeProperties2.Append(groupShapeLocks2);

                Wpg.GroupShapeProperties groupShapeProperties2 = new Wpg.GroupShapeProperties() { BlackWhiteMode = A.BlackWhiteModeValues.Auto };

                A.TransformGroup transformGroup2 = new A.TransformGroup();
                A.Offset offset3 = new A.Offset() { X = -8L, Y = 14978L };
                A.Extents extents3 = new A.Extents() { Cx = 12255L, Cy = 230L };
                A.ChildOffset childOffset2 = new A.ChildOffset() { X = -8L, Y = 14978L };
                A.ChildExtents childExtents2 = new A.ChildExtents() { Cx = 12255L, Cy = 230L };

                transformGroup2.Append(offset3);
                transformGroup2.Append(extents3);
                transformGroup2.Append(childOffset2);
                transformGroup2.Append(childExtents2);

                groupShapeProperties2.Append(transformGroup2);

                Wps.WordprocessingShape wordprocessingShape2 = new Wps.WordprocessingShape();
                Wps.NonVisualDrawingProperties nonVisualDrawingProperties3 = new Wps.NonVisualDrawingProperties() { Id = (UInt32Value)30U, Name = "AutoShape 27" };

                Wps.NonVisualConnectorProperties nonVisualConnectorProperties1 = new Wps.NonVisualConnectorProperties();
                A.ConnectionShapeLocks connectionShapeLocks1 = new A.ConnectionShapeLocks() { NoChangeShapeType = true };

                nonVisualConnectorProperties1.Append(connectionShapeLocks1);

                Wps.ShapeProperties shapeProperties2 = new Wps.ShapeProperties() { BlackWhiteMode = A.BlackWhiteModeValues.Auto };

                A.Transform2D transform2D2 = new A.Transform2D() { VerticalFlip = true };
                A.Offset offset4 = new A.Offset() { X = -8L, Y = 14978L };
                A.Extents extents4 = new A.Extents() { Cx = 1260L, Cy = 230L };

                transform2D2.Append(offset4);
                transform2D2.Append(extents4);

                A.PresetGeometry presetGeometry2 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.BentConnector3 };

                A.AdjustValueList adjustValueList2 = new A.AdjustValueList();
                A.ShapeGuide shapeGuide1 = new A.ShapeGuide() { Name = "adj1", Formula = "val 50000" };

                adjustValueList2.Append(shapeGuide1);

                presetGeometry2.Append(adjustValueList2);
                A.NoFill noFill3 = new A.NoFill();

                A.Outline outline2 = new A.Outline() { Width = 9525 };

                A.SolidFill solidFill3 = new A.SolidFill();
                A.RgbColorModelHex rgbColorModelHex3 = new A.RgbColorModelHex() { Val = "A5A5A5" };

                solidFill3.Append(rgbColorModelHex3);
                A.Miter miter2 = new A.Miter() { Limit = 800000 };
                A.HeadEnd headEnd2 = new A.HeadEnd();
                A.TailEnd tailEnd2 = new A.TailEnd();

                outline2.Append(solidFill3);
                outline2.Append(miter2);
                outline2.Append(headEnd2);
                outline2.Append(tailEnd2);

                A.ShapePropertiesExtensionList shapePropertiesExtensionList2 = new A.ShapePropertiesExtensionList();

                A.ShapePropertiesExtension shapePropertiesExtension3 = new A.ShapePropertiesExtension() { Uri = "{909E8E84-426E-40DD-AFC4-6F175D3DCCD1}" };

                A14.HiddenFillProperties hiddenFillProperties2 = new A14.HiddenFillProperties();
                hiddenFillProperties2.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");
                A.NoFill noFill4 = new A.NoFill();

                hiddenFillProperties2.Append(noFill4);

                shapePropertiesExtension3.Append(hiddenFillProperties2);

                shapePropertiesExtensionList2.Append(shapePropertiesExtension3);

                shapeProperties2.Append(transform2D2);
                shapeProperties2.Append(presetGeometry2);
                shapeProperties2.Append(noFill3);
                shapeProperties2.Append(outline2);
                shapeProperties2.Append(shapePropertiesExtensionList2);
                Wps.TextBodyProperties textBodyProperties2 = new Wps.TextBodyProperties();

                wordprocessingShape2.Append(nonVisualDrawingProperties3);
                wordprocessingShape2.Append(nonVisualConnectorProperties1);
                wordprocessingShape2.Append(shapeProperties2);
                wordprocessingShape2.Append(textBodyProperties2);

                Wps.WordprocessingShape wordprocessingShape3 = new Wps.WordprocessingShape();
                Wps.NonVisualDrawingProperties nonVisualDrawingProperties4 = new Wps.NonVisualDrawingProperties() { Id = (UInt32Value)31U, Name = "AutoShape 28" };

                Wps.NonVisualConnectorProperties nonVisualConnectorProperties2 = new Wps.NonVisualConnectorProperties();
                A.ConnectionShapeLocks connectionShapeLocks2 = new A.ConnectionShapeLocks() { NoChangeShapeType = true };

                nonVisualConnectorProperties2.Append(connectionShapeLocks2);

                Wps.ShapeProperties shapeProperties3 = new Wps.ShapeProperties() { BlackWhiteMode = A.BlackWhiteModeValues.Auto };

                A.Transform2D transform2D3 = new A.Transform2D() { Rotation = 10800000 };
                A.Offset offset5 = new A.Offset() { X = 1252L, Y = 14978L };
                A.Extents extents5 = new A.Extents() { Cx = 10995L, Cy = 230L };

                transform2D3.Append(offset5);
                transform2D3.Append(extents5);

                A.PresetGeometry presetGeometry3 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.BentConnector3 };

                A.AdjustValueList adjustValueList3 = new A.AdjustValueList();
                A.ShapeGuide shapeGuide2 = new A.ShapeGuide() { Name = "adj1", Formula = "val 96778" };

                adjustValueList3.Append(shapeGuide2);

                presetGeometry3.Append(adjustValueList3);
                A.NoFill noFill5 = new A.NoFill();

                A.Outline outline3 = new A.Outline() { Width = 9525 };

                A.SolidFill solidFill4 = new A.SolidFill();
                A.RgbColorModelHex rgbColorModelHex4 = new A.RgbColorModelHex() { Val = "A5A5A5" };

                solidFill4.Append(rgbColorModelHex4);
                A.Miter miter3 = new A.Miter() { Limit = 800000 };
                A.HeadEnd headEnd3 = new A.HeadEnd();
                A.TailEnd tailEnd3 = new A.TailEnd();

                outline3.Append(solidFill4);
                outline3.Append(miter3);
                outline3.Append(headEnd3);
                outline3.Append(tailEnd3);

                A.ShapePropertiesExtensionList shapePropertiesExtensionList3 = new A.ShapePropertiesExtensionList();

                A.ShapePropertiesExtension shapePropertiesExtension4 = new A.ShapePropertiesExtension() { Uri = "{909E8E84-426E-40DD-AFC4-6F175D3DCCD1}" };

                A14.HiddenFillProperties hiddenFillProperties3 = new A14.HiddenFillProperties();
                hiddenFillProperties3.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");
                A.NoFill noFill6 = new A.NoFill();

                hiddenFillProperties3.Append(noFill6);

                shapePropertiesExtension4.Append(hiddenFillProperties3);

                shapePropertiesExtensionList3.Append(shapePropertiesExtension4);

                shapeProperties3.Append(transform2D3);
                shapeProperties3.Append(presetGeometry3);
                shapeProperties3.Append(noFill5);
                shapeProperties3.Append(outline3);
                shapeProperties3.Append(shapePropertiesExtensionList3);
                Wps.TextBodyProperties textBodyProperties3 = new Wps.TextBodyProperties();

                wordprocessingShape3.Append(nonVisualDrawingProperties4);
                wordprocessingShape3.Append(nonVisualConnectorProperties2);
                wordprocessingShape3.Append(shapeProperties3);
                wordprocessingShape3.Append(textBodyProperties3);

                groupShape1.Append(nonVisualDrawingProperties2);
                groupShape1.Append(nonVisualGroupDrawingShapeProperties2);
                groupShape1.Append(groupShapeProperties2);
                groupShape1.Append(wordprocessingShape2);
                groupShape1.Append(wordprocessingShape3);

                wordprocessingGroup1.Append(nonVisualGroupDrawingShapeProperties1);
                wordprocessingGroup1.Append(groupShapeProperties1);
                wordprocessingGroup1.Append(wordprocessingShape1);
                wordprocessingGroup1.Append(groupShape1);

                graphicData1.Append(wordprocessingGroup1);

                graphic1.Append(graphicData1);

                Wp14.RelativeWidth relativeWidth1 = new Wp14.RelativeWidth() { ObjectId = Wp14.SizeRelativeHorizontallyValues.Page };
                Wp14.PercentageWidth percentageWidth1 = new Wp14.PercentageWidth();
                percentageWidth1.Text = "100000";

                relativeWidth1.Append(percentageWidth1);

                Wp14.RelativeHeight relativeHeight1 = new Wp14.RelativeHeight() { RelativeFrom = Wp14.SizeRelativeVerticallyValues.Page };
                Wp14.PercentageHeight percentageHeight1 = new Wp14.PercentageHeight();
                percentageHeight1.Text = "0";

                relativeHeight1.Append(percentageHeight1);

                anchor1.Append(simplePosition1);
                anchor1.Append(horizontalPosition1);
                anchor1.Append(verticalPosition1);
                anchor1.Append(extent1);
                anchor1.Append(effectExtent1);
                anchor1.Append(wrapNone1);
                anchor1.Append(docProperties1);
                anchor1.Append(nonVisualGraphicFrameDrawingProperties1);
                anchor1.Append(graphic1);
                anchor1.Append(relativeWidth1);
                anchor1.Append(relativeHeight1);

                drawing1.Append(anchor1);

                alternateContentChoice1.Append(drawing1);

                AlternateContentFallback alternateContentFallback1 = new AlternateContentFallback();

                Picture picture1 = new Picture();

                V.Group group1 = new V.Group() { Id = "Group 27", Style = "position:absolute;margin-left:0;margin-top:0;width:610.5pt;height:15pt;z-index:251659264;mso-width-percent:1000;mso-position-horizontal:center;mso-position-horizontal-relative:page;mso-position-vertical:center;mso-position-vertical-relative:bottom-margin-area;mso-width-percent:1000", CoordinateSize = "12255,300", CoordinateOrigin = "-8,14978", OptionalString = "_x0000_s1026" };
                group1.SetAttribute(new OpenXmlAttribute("w14", "anchorId", "http://schemas.microsoft.com/office/word/2010/wordml", "74BB3F01"));
                group1.SetAttribute(new OpenXmlAttribute("o", "gfxdata", "urn:schemas-microsoft-com:office:office", "UEsDBBQABgAIAAAAIQC2gziS/gAAAOEBAAATAAAAW0NvbnRlbnRfVHlwZXNdLnhtbJSRQU7DMBBF\n90jcwfIWJU67QAgl6YK0S0CoHGBkTxKLZGx5TGhvj5O2G0SRWNoz/78nu9wcxkFMGNg6quQqL6RA\n0s5Y6ir5vt9lD1JwBDIwOMJKHpHlpr69KfdHjyxSmriSfYz+USnWPY7AufNIadK6MEJMx9ApD/oD\nOlTrorhX2lFEilmcO2RdNtjC5xDF9pCuTyYBB5bi6bQ4syoJ3g9WQ0ymaiLzg5KdCXlKLjvcW893\nSUOqXwnz5DrgnHtJTxOsQfEKIT7DmDSUCaxw7Rqn8787ZsmRM9e2VmPeBN4uqYvTtW7jvijg9N/y\nJsXecLq0q+WD6m8AAAD//wMAUEsDBBQABgAIAAAAIQA4/SH/1gAAAJQBAAALAAAAX3JlbHMvLnJl\nbHOkkMFqwzAMhu+DvYPRfXGawxijTi+j0GvpHsDYimMaW0Yy2fr2M4PBMnrbUb/Q94l/f/hMi1qR\nJVI2sOt6UJgd+ZiDgffL8ekFlFSbvV0oo4EbChzGx4f9GRdb25HMsYhqlCwG5lrLq9biZkxWOiqY\n22YiTra2kYMu1l1tQD30/bPm3wwYN0x18gb45AdQl1tp5j/sFB2T0FQ7R0nTNEV3j6o9feQzro1i\nOWA14Fm+Q8a1a8+Bvu/d/dMb2JY5uiPbhG/ktn4cqGU/er3pcvwCAAD//wMAUEsDBBQABgAIAAAA\nIQAJyj91YwMAAG8KAAAOAAAAZHJzL2Uyb0RvYy54bWzUVm1v0zAQ/o7Ef7D8naVNydpGy6bRvQhp\nwKQNvruJ8wKJHWx3Sfn13Nlu2nUwpMFAqFJ1sX3n556755Kjk76pyR1XupIioeODESVcpDKrRJHQ\nj7cXr2aUaMNExmopeELXXNOT45cvjro25qEsZZ1xRSCI0HHXJrQ0po2DQKclb5g+kC0XsJlL1TAD\nj6oIMsU6iN7UQTgaHQadVFmrZMq1htUzt0mPbfw856n5kOeaG1InFLAZ+6/s/xL/g+MjFheKtWWV\nehjsCSgaVgm4dAh1xgwjK1U9CNVUqZJa5uYglU0g87xKuc0BshmP9rK5VHLV2lyKuCvagSagdo+n\nJ4dN399dqvamvVYOPZhXMv2igZega4t4dx+fC3eYLLt3MoN6spWRNvE+Vw2GgJRIb/ldD/zy3pAU\nFqfTaDKJoAwp7I3no2jkC5CWUCV0w3bBvdfz6czVJi3Pvfc4DKPI+U6cY8Bid62F6qFh6aGX9JYu\n/Xt03ZSs5bYKGum4VqTKEhoCUsEaoOAW03sjexJGCBlvh2NIKTE9rEM6liHtmCVCLkomCn6qlOxK\nzjLAN0ZPyGJwdXE0BvkV1dNZuCFt7vncEH4YzR1h4czSORDG4lZpc8llQ9BIqAKlWJTs7kobBLM9\ngmUV8qKqa1hncS3uLcBBXLHgEa9Dbvpl78lYymwNaSjpxAfDAoxSqm+UdCC8hOqvK6Y4JfVbAVSg\nSjeG2hjLjcFECq4JNZQ4c2GcmletqooSIjuyhTyFzswrmwry6lB4nNAdCNO3szN3SgukudJaAZKJ\nrc6+GlDrf0otD9t+U8Gdpg8nT1SLdxyK/w/UAtA9pVgWqygSTnfkshBuAqW98BNo0Ik9fbtuQWr3\nZOJcsLY/lwnJ66r9tGkKP5seY/vQT6d9zrZq8IJZcmEWUgjQjVSTrXRQG0Xmk2XZ5zEleVPDO+WO\n1QQG3jC5rNAe1xnpEjqPYKxgUC3rKkMR2gdVLBe1IhA0oacR/uwE2TvWVAbernXVJHSGV/sGwqFz\nLjKrZsOq2tk/FrLTDQ4EZNoL5y/MVxDdw46xQ8xPyefqGDunxiPPF5Lt22YcRsOg3bydBpmO5nP/\nbnqezpkfTt2dUKX/t3O2k8f2k/2qsRLwX2D42bT7bE9tvxOPvwMAAP//AwBQSwMEFAAGAAgAAAAh\nAPAtuOTbAAAABQEAAA8AAABkcnMvZG93bnJldi54bWxMj8FOwzAQRO9I/QdrkbhRuykCFOJUgMoN\nhChpy9GNlzhqvA62m4a/x+UCl5FGs5p5WyxG27EBfWgdSZhNBTCk2umWGgnV+9PlLbAQFWnVOUIJ\n3xhgUU7OCpVrd6Q3HFaxYamEQq4kmBj7nPNQG7QqTF2PlLJP562KyfqGa6+Oqdx2PBPimlvVUlow\nqsdHg/V+dbASspv1VVh+9K8PL+uvzfC8rYxvKikvzsf7O2ARx/h3DCf8hA5lYtq5A+nAOgnpkfir\npyzLZsnvJMyFAF4W/D99+QMAAP//AwBQSwECLQAUAAYACAAAACEAtoM4kv4AAADhAQAAEwAAAAAA\nAAAAAAAAAAAAAAAAW0NvbnRlbnRfVHlwZXNdLnhtbFBLAQItABQABgAIAAAAIQA4/SH/1gAAAJQB\nAAALAAAAAAAAAAAAAAAAAC8BAABfcmVscy8ucmVsc1BLAQItABQABgAIAAAAIQAJyj91YwMAAG8K\nAAAOAAAAAAAAAAAAAAAAAC4CAABkcnMvZTJvRG9jLnhtbFBLAQItABQABgAIAAAAIQDwLbjk2wAA\nAAUBAAAPAAAAAAAAAAAAAAAAAL0FAABkcnMvZG93bnJldi54bWxQSwUGAAAAAAQABADzAAAAxQYA\nAAAA\n"));

                V.Shapetype shapetype1 = new V.Shapetype() { Id = "_x0000_t202", CoordinateSize = "21600,21600", OptionalNumber = 202, EdgePath = "m,l,21600r21600,l21600,xe" };
                V.Stroke stroke1 = new V.Stroke() { JoinStyle = V.StrokeJoinStyleValues.Miter };
                V.Path path1 = new V.Path() { AllowGradientShape = true, ConnectionPointType = Ovml.ConnectValues.Rectangle };

                shapetype1.Append(stroke1);
                shapetype1.Append(path1);

                V.Shape shape1 = new V.Shape() { Id = "Text Box 25", Style = "position:absolute;left:782;top:14990;width:659;height:288;visibility:visible;mso-wrap-style:square;v-text-anchor:top", OptionalString = "_x0000_s1027", Filled = false, Stroked = false, Type = "#_x0000_t202", EncodedPackage = "UEsDBBQABgAIAAAAIQDb4fbL7gAAAIUBAAATAAAAW0NvbnRlbnRfVHlwZXNdLnhtbHyQz07DMAyH\n70i8Q+QralM4IITa7kDhCAiNB7ASt43WOlEcyvb2pNu4IODoPz9/n1xv9vOkForiPDdwXVagiI23\njocG3rdPxR0oScgWJ8/UwIEENu3lRb09BBKV0ywNjCmFe63FjDSjlD4Q50nv44wpl3HQAc0OB9I3\nVXWrjedEnIq03oC27qjHjympx31un0wiTQLq4bS4shrAECZnMGVTvbD9QSnOhDInjzsyuiBXWQP0\nr4R18jfgnHvJr4nOknrFmJ5xzhraRtHWf3Kkpfz/yGo5S+H73hkquyhdjr3R8m2lj09svwAAAP//\nAwBQSwMEFAAGAAgAAAAhAFr0LFu/AAAAFQEAAAsAAABfcmVscy8ucmVsc2zPwWrDMAwG4Ptg72B0\nX5TuUMaI01uh19I+gLGVxCy2jGSy9e1nemrHjpL4P0nD4SetZiPRyNnCruvBUPYcYp4tXC/Htw8w\nWl0ObuVMFm6kcBhfX4Yzra62kC6xqGlKVgtLreUTUf1CyWnHhXKbTCzJ1VbKjMX5LzcTvvf9HuXR\ngPHJNKdgQU5hB+ZyK23zHztFL6w81c5zQp6m6P9TMfB3PtPWFCczVQtB9N4U2rp2HOA44NMz4y8A\nAAD//wMAUEsDBBQABgAIAAAAIQAlI2QMwAAAANsAAAAPAAAAZHJzL2Rvd25yZXYueG1sRE9Ni8Iw\nEL0L+x/CCN5sqgfRrlFkWWFBEGs97HG2GdtgM6lNVuu/NwfB4+N9L9e9bcSNOm8cK5gkKQji0mnD\nlYJTsR3PQfiArLFxTAoe5GG9+hgsMdPuzjndjqESMYR9hgrqENpMSl/WZNEnriWO3Nl1FkOEXSV1\nh/cYbhs5TdOZtGg4NtTY0ldN5eX4bxVsfjn/Ntf93yE/56YoFinvZhelRsN+8wkiUB/e4pf7RyuY\nxrHxS/wBcvUEAAD//wMAUEsBAi0AFAAGAAgAAAAhANvh9svuAAAAhQEAABMAAAAAAAAAAAAAAAAA\nAAAAAFtDb250ZW50X1R5cGVzXS54bWxQSwECLQAUAAYACAAAACEAWvQsW78AAAAVAQAACwAAAAAA\nAAAAAAAAAAAfAQAAX3JlbHMvLnJlbHNQSwECLQAUAAYACAAAACEAJSNkDMAAAADbAAAADwAAAAAA\nAAAAAAAAAAAHAgAAZHJzL2Rvd25yZXYueG1sUEsFBgAAAAADAAMAtwAAAPQCAAAAAA==\n" };

                V.TextBox textBox1 = new V.TextBox() { Inset = "0,0,0,0" };

                TextBoxContent textBoxContent2 = new TextBoxContent();

                Paragraph paragraph3 = new Paragraph() { RsidParagraphAddition = "00354B38", RsidRunAdditionDefault = "00354B38", ParagraphId = "4E6CF7E2", TextId = "77777777" };

                ParagraphProperties paragraphProperties3 = new ParagraphProperties();
                Justification justification2 = new Justification() { Val = JustificationValues.Center };

                paragraphProperties3.Append(justification2);

                Run run7 = new Run();
                FieldChar fieldChar4 = new FieldChar() { FieldCharType = FieldCharValues.Begin };

                run7.Append(fieldChar4);

                Run run8 = new Run();
                FieldCode fieldCode2 = new FieldCode() { Space = SpaceProcessingModeValues.Preserve };
                fieldCode2.Text = " PAGE    \\* MERGEFORMAT ";

                run8.Append(fieldCode2);

                Run run9 = new Run();
                FieldChar fieldChar5 = new FieldChar() { FieldCharType = FieldCharValues.Separate };

                run9.Append(fieldChar5);

                Run run10 = new Run();

                RunProperties runProperties4 = new RunProperties();
                NoProof noProof4 = new NoProof();
                Color color3 = new Color() { Val = "8C8C8C", ThemeColor = ThemeColorValues.Background1, ThemeShade = "8C" };

                runProperties4.Append(noProof4);
                runProperties4.Append(color3);
                Text text2 = new Text();
                text2.Text = "2";

                run10.Append(runProperties4);
                run10.Append(text2);

                Run run11 = new Run();

                RunProperties runProperties5 = new RunProperties();
                NoProof noProof5 = new NoProof();
                Color color4 = new Color() { Val = "8C8C8C", ThemeColor = ThemeColorValues.Background1, ThemeShade = "8C" };

                runProperties5.Append(noProof5);
                runProperties5.Append(color4);
                FieldChar fieldChar6 = new FieldChar() { FieldCharType = FieldCharValues.End };

                run11.Append(runProperties5);
                run11.Append(fieldChar6);

                paragraph3.Append(paragraphProperties3);
                paragraph3.Append(run7);
                paragraph3.Append(run8);
                paragraph3.Append(run9);
                paragraph3.Append(run10);
                paragraph3.Append(run11);

                textBoxContent2.Append(paragraph3);

                textBox1.Append(textBoxContent2);

                shape1.Append(textBox1);

                V.Group group2 = new V.Group() { Id = "Group 31", Style = "position:absolute;left:-8;top:14978;width:12255;height:230", CoordinateSize = "12255,230", CoordinateOrigin = "-8,14978", OptionalString = "_x0000_s1028" };
                group2.SetAttribute(new OpenXmlAttribute("o", "gfxdata", "urn:schemas-microsoft-com:office:office", "UEsDBBQABgAIAAAAIQDb4fbL7gAAAIUBAAATAAAAW0NvbnRlbnRfVHlwZXNdLnhtbHyQz07DMAyH\n70i8Q+QralM4IITa7kDhCAiNB7ASt43WOlEcyvb2pNu4IODoPz9/n1xv9vOkForiPDdwXVagiI23\njocG3rdPxR0oScgWJ8/UwIEENu3lRb09BBKV0ywNjCmFe63FjDSjlD4Q50nv44wpl3HQAc0OB9I3\nVXWrjedEnIq03oC27qjHjympx31un0wiTQLq4bS4shrAECZnMGVTvbD9QSnOhDInjzsyuiBXWQP0\nr4R18jfgnHvJr4nOknrFmJ5xzhraRtHWf3Kkpfz/yGo5S+H73hkquyhdjr3R8m2lj09svwAAAP//\nAwBQSwMEFAAGAAgAAAAhAFr0LFu/AAAAFQEAAAsAAABfcmVscy8ucmVsc2zPwWrDMAwG4Ptg72B0\nX5TuUMaI01uh19I+gLGVxCy2jGSy9e1nemrHjpL4P0nD4SetZiPRyNnCruvBUPYcYp4tXC/Htw8w\nWl0ObuVMFm6kcBhfX4Yzra62kC6xqGlKVgtLreUTUf1CyWnHhXKbTCzJ1VbKjMX5LzcTvvf9HuXR\ngPHJNKdgQU5hB+ZyK23zHztFL6w81c5zQp6m6P9TMfB3PtPWFCczVQtB9N4U2rp2HOA44NMz4y8A\nAAD//wMAUEsDBBQABgAIAAAAIQCya6E6xAAAANsAAAAPAAAAZHJzL2Rvd25yZXYueG1sRI9Bi8Iw\nFITvwv6H8Ba8aVoXZa1GEdkVDyKoC+Lt0TzbYvNSmmxb/70RBI/DzHzDzJedKUVDtSssK4iHEQji\n1OqCMwV/p9/BNwjnkTWWlknBnRwsFx+9OSbatnyg5ugzESDsElSQe18lUro0J4NuaCvi4F1tbdAH\nWWdS19gGuCnlKIom0mDBYSHHitY5pbfjv1GwabFdfcU/ze52Xd8vp/H+vItJqf5nt5qB8NT5d/jV\n3moFoyk8v4QfIBcPAAAA//8DAFBLAQItABQABgAIAAAAIQDb4fbL7gAAAIUBAAATAAAAAAAAAAAA\nAAAAAAAAAABbQ29udGVudF9UeXBlc10ueG1sUEsBAi0AFAAGAAgAAAAhAFr0LFu/AAAAFQEAAAsA\nAAAAAAAAAAAAAAAAHwEAAF9yZWxzLy5yZWxzUEsBAi0AFAAGAAgAAAAhALJroTrEAAAA2wAAAA8A\nAAAAAAAAAAAAAAAABwIAAGRycy9kb3ducmV2LnhtbFBLBQYAAAAAAwADALcAAAD4AgAAAAA=\n"));

                V.Shapetype shapetype2 = new V.Shapetype() { Id = "_x0000_t34", CoordinateSize = "21600,21600", Oned = true, Filled = false, OptionalNumber = 34, Adjustment = "10800", EdgePath = "m,l@0,0@0,21600,21600,21600e" };
                V.Stroke stroke2 = new V.Stroke() { JoinStyle = V.StrokeJoinStyleValues.Miter };

                V.Formulas formulas1 = new V.Formulas();
                V.Formula formula1 = new V.Formula() { Equation = "val #0" };

                formulas1.Append(formula1);
                V.Path path2 = new V.Path() { AllowFill = false, ShowArrowhead = true, ConnectionPointType = Ovml.ConnectValues.None };

                V.ShapeHandles shapeHandles1 = new V.ShapeHandles();
                V.ShapeHandle shapeHandle1 = new V.ShapeHandle() { Position = "#0,center" };

                shapeHandles1.Append(shapeHandle1);
                Ovml.Lock lock1 = new Ovml.Lock() { Extension = V.ExtensionHandlingBehaviorValues.Edit, ShapeType = true };

                shapetype2.Append(stroke2);
                shapetype2.Append(formulas1);
                shapetype2.Append(path2);
                shapetype2.Append(shapeHandles1);
                shapetype2.Append(lock1);
                V.Shape shape2 = new V.Shape() { Id = "AutoShape 27", Style = "position:absolute;left:-8;top:14978;width:1260;height:230;flip:y;visibility:visible;mso-wrap-style:square", OptionalString = "_x0000_s1029", StrokeColor = "#a5a5a5", ConnectorType = Ovml.ConnectorValues.Elbow, Type = "#_x0000_t34", EncodedPackage = "UEsDBBQABgAIAAAAIQDb4fbL7gAAAIUBAAATAAAAW0NvbnRlbnRfVHlwZXNdLnhtbHyQz07DMAyH\n70i8Q+QralM4IITa7kDhCAiNB7ASt43WOlEcyvb2pNu4IODoPz9/n1xv9vOkForiPDdwXVagiI23\njocG3rdPxR0oScgWJ8/UwIEENu3lRb09BBKV0ywNjCmFe63FjDSjlD4Q50nv44wpl3HQAc0OB9I3\nVXWrjedEnIq03oC27qjHjympx31un0wiTQLq4bS4shrAECZnMGVTvbD9QSnOhDInjzsyuiBXWQP0\nr4R18jfgnHvJr4nOknrFmJ5xzhraRtHWf3Kkpfz/yGo5S+H73hkquyhdjr3R8m2lj09svwAAAP//\nAwBQSwMEFAAGAAgAAAAhAFr0LFu/AAAAFQEAAAsAAABfcmVscy8ucmVsc2zPwWrDMAwG4Ptg72B0\nX5TuUMaI01uh19I+gLGVxCy2jGSy9e1nemrHjpL4P0nD4SetZiPRyNnCruvBUPYcYp4tXC/Htw8w\nWl0ObuVMFm6kcBhfX4Yzra62kC6xqGlKVgtLreUTUf1CyWnHhXKbTCzJ1VbKjMX5LzcTvvf9HuXR\ngPHJNKdgQU5hB+ZyK23zHztFL6w81c5zQp6m6P9TMfB3PtPWFCczVQtB9N4U2rp2HOA44NMz4y8A\nAAD//wMAUEsDBBQABgAIAAAAIQBMgYtEwgAAANsAAAAPAAAAZHJzL2Rvd25yZXYueG1sRE/Pa8Iw\nFL4L+x/CG+xSbKrDMTqjyEC6yw7rWtjx2bw1Zc1LaaJW//rlIHj8+H6vt5PtxYlG3zlWsEgzEMSN\n0x23Cqrv/fwVhA/IGnvHpOBCHrabh9kac+3O/EWnMrQihrDPUYEJYcil9I0hiz51A3Hkft1oMUQ4\ntlKPeI7htpfLLHuRFjuODQYHejfU/JVHqyDxmayb1Y8pkuLzcNU1VztbKPX0OO3eQASawl18c39o\nBc9xffwSf4Dc/AMAAP//AwBQSwECLQAUAAYACAAAACEA2+H2y+4AAACFAQAAEwAAAAAAAAAAAAAA\nAAAAAAAAW0NvbnRlbnRfVHlwZXNdLnhtbFBLAQItABQABgAIAAAAIQBa9CxbvwAAABUBAAALAAAA\nAAAAAAAAAAAAAB8BAABfcmVscy8ucmVsc1BLAQItABQABgAIAAAAIQBMgYtEwgAAANsAAAAPAAAA\nAAAAAAAAAAAAAAcCAABkcnMvZG93bnJldi54bWxQSwUGAAAAAAMAAwC3AAAA9gIAAAAA\n" };
                V.Shape shape3 = new V.Shape() { Id = "AutoShape 28", Style = "position:absolute;left:1252;top:14978;width:10995;height:230;rotation:180;visibility:visible;mso-wrap-style:square", OptionalString = "_x0000_s1030", StrokeColor = "#a5a5a5", ConnectorType = Ovml.ConnectorValues.Elbow, Type = "#_x0000_t34", Adjustment = "20904", EncodedPackage = "UEsDBBQABgAIAAAAIQDb4fbL7gAAAIUBAAATAAAAW0NvbnRlbnRfVHlwZXNdLnhtbHyQz07DMAyH\n70i8Q+QralM4IITa7kDhCAiNB7ASt43WOlEcyvb2pNu4IODoPz9/n1xv9vOkForiPDdwXVagiI23\njocG3rdPxR0oScgWJ8/UwIEENu3lRb09BBKV0ywNjCmFe63FjDSjlD4Q50nv44wpl3HQAc0OB9I3\nVXWrjedEnIq03oC27qjHjympx31un0wiTQLq4bS4shrAECZnMGVTvbD9QSnOhDInjzsyuiBXWQP0\nr4R18jfgnHvJr4nOknrFmJ5xzhraRtHWf3Kkpfz/yGo5S+H73hkquyhdjr3R8m2lj09svwAAAP//\nAwBQSwMEFAAGAAgAAAAhAFr0LFu/AAAAFQEAAAsAAABfcmVscy8ucmVsc2zPwWrDMAwG4Ptg72B0\nX5TuUMaI01uh19I+gLGVxCy2jGSy9e1nemrHjpL4P0nD4SetZiPRyNnCruvBUPYcYp4tXC/Htw8w\nWl0ObuVMFm6kcBhfX4Yzra62kC6xqGlKVgtLreUTUf1CyWnHhXKbTCzJ1VbKjMX5LzcTvvf9HuXR\ngPHJNKdgQU5hB+ZyK23zHztFL6w81c5zQp6m6P9TMfB3PtPWFCczVQtB9N4U2rp2HOA44NMz4y8A\nAAD//wMAUEsDBBQABgAIAAAAIQBoEb+gxQAAANsAAAAPAAAAZHJzL2Rvd25yZXYueG1sRI9Ba8JA\nFITvQv/D8gpepG6MUEp0lWBQBClU68XbI/tMYrJvQ3aN8d93C4Ueh5n5hlmuB9OInjpXWVYwm0Yg\niHOrKy4UnL+3bx8gnEfW2FgmBU9ysF69jJaYaPvgI/UnX4gAYZeggtL7NpHS5SUZdFPbEgfvajuD\nPsiukLrDR4CbRsZR9C4NVhwWSmxpU1Jen+5Gwedxd64v8p7FQ5VObnjILrevTKnx65AuQHga/H/4\nr73XCuYz+P0SfoBc/QAAAP//AwBQSwECLQAUAAYACAAAACEA2+H2y+4AAACFAQAAEwAAAAAAAAAA\nAAAAAAAAAAAAW0NvbnRlbnRfVHlwZXNdLnhtbFBLAQItABQABgAIAAAAIQBa9CxbvwAAABUBAAAL\nAAAAAAAAAAAAAAAAAB8BAABfcmVscy8ucmVsc1BLAQItABQABgAIAAAAIQBoEb+gxQAAANsAAAAP\nAAAAAAAAAAAAAAAAAAcCAABkcnMvZG93bnJldi54bWxQSwUGAAAAAAMAAwC3AAAA+QIAAAAA\n" };

                group2.Append(shapetype2);
                group2.Append(shape2);
                group2.Append(shape3);
                Wvml.TextWrap textWrap1 = new Wvml.TextWrap() { AnchorX = Wvml.HorizontalAnchorValues.Page, AnchorY = Wvml.VerticalAnchorValues.Margin };

                group1.Append(shapetype1);
                group1.Append(shape1);
                group1.Append(group2);
                group1.Append(textWrap1);

                picture1.Append(group1);

                alternateContentFallback1.Append(picture1);

                alternateContent1.Append(alternateContentChoice1);
                alternateContent1.Append(alternateContentFallback1);

                run1.Append(runProperties1);
                run1.Append(alternateContent1);

                paragraph1.Append(paragraphProperties1);
                paragraph1.Append(run1);

                sdtContentBlock1.Append(paragraph1);

                sdtBlock1.Append(sdtProperties1);
                sdtBlock1.Append(sdtContentBlock1);
                return sdtBlock1;

            }
        }
        private static SdtBlock ThickLine1 {
            get {
                SdtBlock sdtBlock1 = new SdtBlock();

                SdtProperties sdtProperties1 = new SdtProperties();
                SdtId sdtId1 = new SdtId() { Val = 1056356717 };

                SdtContentDocPartObject sdtContentDocPartObject1 = new SdtContentDocPartObject();
                DocPartGallery docPartGallery1 = new DocPartGallery() { Val = "Page Numbers (Bottom of Page)" };
                DocPartUnique docPartUnique1 = new DocPartUnique();

                sdtContentDocPartObject1.Append(docPartGallery1);
                sdtContentDocPartObject1.Append(docPartUnique1);

                sdtProperties1.Append(sdtId1);
                sdtProperties1.Append(sdtContentDocPartObject1);

                SdtEndCharProperties sdtEndCharProperties1 = new SdtEndCharProperties();

                RunProperties runProperties1 = new RunProperties();
                NoProof noProof1 = new NoProof();

                runProperties1.Append(noProof1);

                sdtEndCharProperties1.Append(runProperties1);

                SdtContentBlock sdtContentBlock1 = new SdtContentBlock();

                Paragraph paragraph1 = new Paragraph() { RsidParagraphAddition = "00B41B51", RsidRunAdditionDefault = "00B41B51", ParagraphId = "2615380B", TextId = "66A693A7" };

                ParagraphProperties paragraphProperties1 = new ParagraphProperties();
                ParagraphStyleId paragraphStyleId1 = new ParagraphStyleId() { Val = "Footer" };
                Justification justification1 = new Justification() { Val = JustificationValues.Center };

                paragraphProperties1.Append(paragraphStyleId1);
                paragraphProperties1.Append(justification1);

                Run run1 = new Run();

                RunProperties runProperties2 = new RunProperties();
                NoProof noProof2 = new NoProof();

                runProperties2.Append(noProof2);

                AlternateContent alternateContent1 = new AlternateContent();

                AlternateContentChoice alternateContentChoice1 = new AlternateContentChoice() { Requires = "wps" };

                Drawing drawing1 = new Drawing();

                Wp.Inline inline1 = new Wp.Inline() { DistanceFromTop = (UInt32Value)0U, DistanceFromBottom = (UInt32Value)0U, DistanceFromLeft = (UInt32Value)0U, DistanceFromRight = (UInt32Value)0U, AnchorId = "3081DCFF", EditId = "5004DB72" };
                Wp.Extent extent1 = new Wp.Extent() { Cx = 5467350L, Cy = 54610L };
                Wp.EffectExtent effectExtent1 = new Wp.EffectExtent() { LeftEdge = 9525L, TopEdge = 19050L, RightEdge = 9525L, BottomEdge = 12065L };
                Wp.DocProperties docProperties1 = new Wp.DocProperties() { Id = (UInt32Value)26U, Name = "Flowchart: Decision 26" };

                Wp.NonVisualGraphicFrameDrawingProperties nonVisualGraphicFrameDrawingProperties1 = new Wp.NonVisualGraphicFrameDrawingProperties();

                A.GraphicFrameLocks graphicFrameLocks1 = new A.GraphicFrameLocks();
                graphicFrameLocks1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

                nonVisualGraphicFrameDrawingProperties1.Append(graphicFrameLocks1);

                A.Graphic graphic1 = new A.Graphic();
                graphic1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

                A.GraphicData graphicData1 = new A.GraphicData() { Uri = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape" };

                Wps.WordprocessingShape wordprocessingShape1 = new Wps.WordprocessingShape();

                Wps.NonVisualDrawingShapeProperties nonVisualDrawingShapeProperties1 = new Wps.NonVisualDrawingShapeProperties();
                A.ShapeLocks shapeLocks1 = new A.ShapeLocks() { NoChangeArrowheads = true };

                nonVisualDrawingShapeProperties1.Append(shapeLocks1);

                Wps.ShapeProperties shapeProperties1 = new Wps.ShapeProperties() { BlackWhiteMode = A.BlackWhiteModeValues.Auto };

                A.Transform2D transform2D1 = new A.Transform2D();
                A.Offset offset1 = new A.Offset() { X = 0L, Y = 0L };
                A.Extents extents1 = new A.Extents() { Cx = 5467350L, Cy = 54610L };

                transform2D1.Append(offset1);
                transform2D1.Append(extents1);

                A.PresetGeometry presetGeometry1 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.FlowChartDecision };
                A.AdjustValueList adjustValueList1 = new A.AdjustValueList();

                presetGeometry1.Append(adjustValueList1);

                A.SolidFill solidFill1 = new A.SolidFill();
                A.RgbColorModelHex rgbColorModelHex1 = new A.RgbColorModelHex() { Val = "000000" };

                solidFill1.Append(rgbColorModelHex1);

                A.Outline outline1 = new A.Outline() { Width = 9525 };

                A.SolidFill solidFill2 = new A.SolidFill();
                A.RgbColorModelHex rgbColorModelHex2 = new A.RgbColorModelHex() { Val = "000000" };

                solidFill2.Append(rgbColorModelHex2);
                A.Miter miter1 = new A.Miter() { Limit = 800000 };
                A.HeadEnd headEnd1 = new A.HeadEnd();
                A.TailEnd tailEnd1 = new A.TailEnd();

                outline1.Append(solidFill2);
                outline1.Append(miter1);
                outline1.Append(headEnd1);
                outline1.Append(tailEnd1);

                shapeProperties1.Append(transform2D1);
                shapeProperties1.Append(presetGeometry1);
                shapeProperties1.Append(solidFill1);
                shapeProperties1.Append(outline1);

                Wps.TextBodyProperties textBodyProperties1 = new Wps.TextBodyProperties() { Rotation = 0, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, LeftInset = 91440, TopInset = 45720, RightInset = 91440, BottomInset = 45720, Anchor = A.TextAnchoringTypeValues.Top, AnchorCenter = false, UpRight = true };
                A.NoAutoFit noAutoFit1 = new A.NoAutoFit();

                textBodyProperties1.Append(noAutoFit1);

                wordprocessingShape1.Append(nonVisualDrawingShapeProperties1);
                wordprocessingShape1.Append(shapeProperties1);
                wordprocessingShape1.Append(textBodyProperties1);

                graphicData1.Append(wordprocessingShape1);

                graphic1.Append(graphicData1);

                inline1.Append(extent1);
                inline1.Append(effectExtent1);
                inline1.Append(docProperties1);
                inline1.Append(nonVisualGraphicFrameDrawingProperties1);
                inline1.Append(graphic1);

                drawing1.Append(inline1);

                alternateContentChoice1.Append(drawing1);

                AlternateContentFallback alternateContentFallback1 = new AlternateContentFallback();

                Picture picture1 = new Picture();

                V.Shapetype shapetype1 = new V.Shapetype() { Id = "_x0000_t110", CoordinateSize = "21600,21600", OptionalNumber = 110, EdgePath = "m10800,l,10800,10800,21600,21600,10800xe" };
                shapetype1.SetAttribute(new OpenXmlAttribute("w14", "anchorId", "http://schemas.microsoft.com/office/word/2010/wordml", "1F9B523E"));
                V.Stroke stroke1 = new V.Stroke() { JoinStyle = V.StrokeJoinStyleValues.Miter };
                V.Path path1 = new V.Path() { TextboxRectangle = "5400,5400,16200,16200", AllowGradientShape = true, ConnectionPointType = Ovml.ConnectValues.Rectangle };

                shapetype1.Append(stroke1);
                shapetype1.Append(path1);

                V.Shape shape1 = new V.Shape() { Id = "Flowchart: Decision 26", Style = "width:430.5pt;height:4.3pt;visibility:visible;mso-wrap-style:square;mso-left-percent:-10001;mso-top-percent:-10001;mso-position-horizontal:absolute;mso-position-horizontal-relative:char;mso-position-vertical:absolute;mso-position-vertical-relative:line;mso-left-percent:-10001;mso-top-percent:-10001;v-text-anchor:top", OptionalString = "_x0000_s1026", FillColor = "black", Type = "#_x0000_t110", EncodedPackage = "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\nAAA=\n" };
                Wvml.AnchorLock anchorLock1 = new Wvml.AnchorLock();

                shape1.Append(anchorLock1);

                picture1.Append(shapetype1);
                picture1.Append(shape1);

                alternateContentFallback1.Append(picture1);

                alternateContent1.Append(alternateContentChoice1);
                alternateContent1.Append(alternateContentFallback1);

                run1.Append(runProperties2);
                run1.Append(alternateContent1);

                paragraph1.Append(paragraphProperties1);
                paragraph1.Append(run1);

                Paragraph paragraph2 = new Paragraph() { RsidParagraphAddition = "00B41B51", RsidRunAdditionDefault = "00B41B51", ParagraphId = "1713C6FB", TextId = "34095FA2" };

                ParagraphProperties paragraphProperties2 = new ParagraphProperties();
                ParagraphStyleId paragraphStyleId2 = new ParagraphStyleId() { Val = "Footer" };
                Justification justification2 = new Justification() { Val = JustificationValues.Center };

                paragraphProperties2.Append(paragraphStyleId2);
                paragraphProperties2.Append(justification2);

                Run run2 = new Run();
                FieldChar fieldChar1 = new FieldChar() { FieldCharType = FieldCharValues.Begin };

                run2.Append(fieldChar1);

                Run run3 = new Run();
                FieldCode fieldCode1 = new FieldCode() { Space = SpaceProcessingModeValues.Preserve };
                fieldCode1.Text = " PAGE    \\* MERGEFORMAT ";

                run3.Append(fieldCode1);

                Run run4 = new Run();
                FieldChar fieldChar2 = new FieldChar() { FieldCharType = FieldCharValues.Separate };

                run4.Append(fieldChar2);

                Run run5 = new Run();

                RunProperties runProperties3 = new RunProperties();
                NoProof noProof3 = new NoProof();

                runProperties3.Append(noProof3);
                Text text1 = new Text();
                text1.Text = "2";

                run5.Append(runProperties3);
                run5.Append(text1);

                Run run6 = new Run();

                RunProperties runProperties4 = new RunProperties();
                NoProof noProof4 = new NoProof();

                runProperties4.Append(noProof4);
                FieldChar fieldChar3 = new FieldChar() { FieldCharType = FieldCharValues.End };

                run6.Append(runProperties4);
                run6.Append(fieldChar3);

                paragraph2.Append(paragraphProperties2);
                paragraph2.Append(run2);
                paragraph2.Append(run3);
                paragraph2.Append(run4);
                paragraph2.Append(run5);
                paragraph2.Append(run6);

                sdtContentBlock1.Append(paragraph1);
                sdtContentBlock1.Append(paragraph2);

                sdtBlock1.Append(sdtProperties1);
                sdtBlock1.Append(sdtEndCharProperties1);
                sdtBlock1.Append(sdtContentBlock1);
                return sdtBlock1;

            }
        }
        private static SdtBlock RoundedRectangle1 {
            get {
                SdtBlock sdtBlock1 = new SdtBlock();

                SdtProperties sdtProperties1 = new SdtProperties();
                SdtId sdtId1 = new SdtId() { Val = -1945291954 };

                SdtContentDocPartObject sdtContentDocPartObject1 = new SdtContentDocPartObject();
                DocPartGallery docPartGallery1 = new DocPartGallery() { Val = "Page Numbers (Top of Page)" };
                DocPartUnique docPartUnique1 = new DocPartUnique();

                sdtContentDocPartObject1.Append(docPartGallery1);
                sdtContentDocPartObject1.Append(docPartUnique1);

                sdtProperties1.Append(sdtId1);
                sdtProperties1.Append(sdtContentDocPartObject1);

                SdtContentBlock sdtContentBlock1 = new SdtContentBlock();

                Paragraph paragraph1 = new Paragraph() { RsidParagraphAddition = "009D383E", RsidRunAdditionDefault = "009D383E", ParagraphId = "2427BF06", TextId = "68F62570" };

                ParagraphProperties paragraphProperties1 = new ParagraphProperties();
                ParagraphStyleId paragraphStyleId1 = new ParagraphStyleId() { Val = "Header" };
                Indentation indentation1 = new Indentation() { Left = "-864" };

                paragraphProperties1.Append(paragraphStyleId1);
                paragraphProperties1.Append(indentation1);

                Run run1 = new Run();

                RunProperties runProperties1 = new RunProperties();
                NoProof noProof1 = new NoProof();

                runProperties1.Append(noProof1);

                AlternateContent alternateContent1 = new AlternateContent();

                AlternateContentChoice alternateContentChoice1 = new AlternateContentChoice() { Requires = "wpg" };

                Drawing drawing1 = new Drawing();

                Wp.Inline inline1 = new Wp.Inline() { DistanceFromTop = (UInt32Value)0U, DistanceFromBottom = (UInt32Value)0U, DistanceFromLeft = (UInt32Value)0U, DistanceFromRight = (UInt32Value)0U, AnchorId = "553944E0", EditId = "4F1437F3" };
                Wp.Extent extent1 = new Wp.Extent() { Cx = 548640L, Cy = 237490L };
                Wp.EffectExtent effectExtent1 = new Wp.EffectExtent() { LeftEdge = 9525L, TopEdge = 9525L, RightEdge = 13335L, BottomEdge = 10160L };
                Wp.DocProperties docProperties1 = new Wp.DocProperties() { Id = (UInt32Value)22U, Name = "Group 22" };

                Wp.NonVisualGraphicFrameDrawingProperties nonVisualGraphicFrameDrawingProperties1 = new Wp.NonVisualGraphicFrameDrawingProperties();

                A.GraphicFrameLocks graphicFrameLocks1 = new A.GraphicFrameLocks();
                graphicFrameLocks1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

                nonVisualGraphicFrameDrawingProperties1.Append(graphicFrameLocks1);

                A.Graphic graphic1 = new A.Graphic();
                graphic1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

                A.GraphicData graphicData1 = new A.GraphicData() { Uri = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" };

                Wpg.WordprocessingGroup wordprocessingGroup1 = new Wpg.WordprocessingGroup();

                Wpg.NonVisualGroupDrawingShapeProperties nonVisualGroupDrawingShapeProperties1 = new Wpg.NonVisualGroupDrawingShapeProperties();
                A.GroupShapeLocks groupShapeLocks1 = new A.GroupShapeLocks();

                nonVisualGroupDrawingShapeProperties1.Append(groupShapeLocks1);

                Wpg.GroupShapeProperties groupShapeProperties1 = new Wpg.GroupShapeProperties() { BlackWhiteMode = A.BlackWhiteModeValues.Auto };

                A.TransformGroup transformGroup1 = new A.TransformGroup();
                A.Offset offset1 = new A.Offset() { X = 0L, Y = 0L };
                A.Extents extents1 = new A.Extents() { Cx = 548640L, Cy = 237490L };
                A.ChildOffset childOffset1 = new A.ChildOffset() { X = 614L, Y = 660L };
                A.ChildExtents childExtents1 = new A.ChildExtents() { Cx = 864L, Cy = 374L };

                transformGroup1.Append(offset1);
                transformGroup1.Append(extents1);
                transformGroup1.Append(childOffset1);
                transformGroup1.Append(childExtents1);

                groupShapeProperties1.Append(transformGroup1);

                Wps.WordprocessingShape wordprocessingShape1 = new Wps.WordprocessingShape();
                Wps.NonVisualDrawingProperties nonVisualDrawingProperties1 = new Wps.NonVisualDrawingProperties() { Id = (UInt32Value)23U, Name = "AutoShape 42" };

                Wps.NonVisualDrawingShapeProperties nonVisualDrawingShapeProperties1 = new Wps.NonVisualDrawingShapeProperties();
                A.ShapeLocks shapeLocks1 = new A.ShapeLocks() { NoChangeArrowheads = true };

                nonVisualDrawingShapeProperties1.Append(shapeLocks1);

                Wps.ShapeProperties shapeProperties1 = new Wps.ShapeProperties() { BlackWhiteMode = A.BlackWhiteModeValues.Auto };

                A.Transform2D transform2D1 = new A.Transform2D() { Rotation = -5400000 };
                A.Offset offset2 = new A.Offset() { X = 859L, Y = 415L };
                A.Extents extents2 = new A.Extents() { Cx = 374L, Cy = 864L };

                transform2D1.Append(offset2);
                transform2D1.Append(extents2);

                A.PresetGeometry presetGeometry1 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.RoundRectangle };

                A.AdjustValueList adjustValueList1 = new A.AdjustValueList();
                A.ShapeGuide shapeGuide1 = new A.ShapeGuide() { Name = "adj", Formula = "val 16667" };

                adjustValueList1.Append(shapeGuide1);

                presetGeometry1.Append(adjustValueList1);

                A.SolidFill solidFill1 = new A.SolidFill();
                A.RgbColorModelHex rgbColorModelHex1 = new A.RgbColorModelHex() { Val = "FFFFFF" };

                solidFill1.Append(rgbColorModelHex1);

                A.Outline outline1 = new A.Outline() { Width = 9525 };

                A.SolidFill solidFill2 = new A.SolidFill();
                A.RgbColorModelHex rgbColorModelHex2 = new A.RgbColorModelHex() { Val = "E4BE84" };

                solidFill2.Append(rgbColorModelHex2);
                A.Round round1 = new A.Round();
                A.HeadEnd headEnd1 = new A.HeadEnd();
                A.TailEnd tailEnd1 = new A.TailEnd();

                outline1.Append(solidFill2);
                outline1.Append(round1);
                outline1.Append(headEnd1);
                outline1.Append(tailEnd1);

                shapeProperties1.Append(transform2D1);
                shapeProperties1.Append(presetGeometry1);
                shapeProperties1.Append(solidFill1);
                shapeProperties1.Append(outline1);

                Wps.TextBodyProperties textBodyProperties1 = new Wps.TextBodyProperties() { Rotation = 0, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, LeftInset = 91440, TopInset = 45720, RightInset = 91440, BottomInset = 45720, Anchor = A.TextAnchoringTypeValues.Top, AnchorCenter = false, UpRight = true };
                A.NoAutoFit noAutoFit1 = new A.NoAutoFit();

                textBodyProperties1.Append(noAutoFit1);

                wordprocessingShape1.Append(nonVisualDrawingProperties1);
                wordprocessingShape1.Append(nonVisualDrawingShapeProperties1);
                wordprocessingShape1.Append(shapeProperties1);
                wordprocessingShape1.Append(textBodyProperties1);

                Wps.WordprocessingShape wordprocessingShape2 = new Wps.WordprocessingShape();
                Wps.NonVisualDrawingProperties nonVisualDrawingProperties2 = new Wps.NonVisualDrawingProperties() { Id = (UInt32Value)24U, Name = "AutoShape 43" };

                Wps.NonVisualDrawingShapeProperties nonVisualDrawingShapeProperties2 = new Wps.NonVisualDrawingShapeProperties();
                A.ShapeLocks shapeLocks2 = new A.ShapeLocks() { NoChangeArrowheads = true };

                nonVisualDrawingShapeProperties2.Append(shapeLocks2);

                Wps.ShapeProperties shapeProperties2 = new Wps.ShapeProperties() { BlackWhiteMode = A.BlackWhiteModeValues.Auto };

                A.Transform2D transform2D2 = new A.Transform2D() { Rotation = -5400000 };
                A.Offset offset3 = new A.Offset() { X = 898L, Y = 451L };
                A.Extents extents3 = new A.Extents() { Cx = 296L, Cy = 792L };

                transform2D2.Append(offset3);
                transform2D2.Append(extents3);

                A.PresetGeometry presetGeometry2 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.RoundRectangle };

                A.AdjustValueList adjustValueList2 = new A.AdjustValueList();
                A.ShapeGuide shapeGuide2 = new A.ShapeGuide() { Name = "adj", Formula = "val 16667" };

                adjustValueList2.Append(shapeGuide2);

                presetGeometry2.Append(adjustValueList2);

                A.SolidFill solidFill3 = new A.SolidFill();
                A.RgbColorModelHex rgbColorModelHex3 = new A.RgbColorModelHex() { Val = "E4BE84" };

                solidFill3.Append(rgbColorModelHex3);

                A.Outline outline2 = new A.Outline() { Width = 9525 };

                A.SolidFill solidFill4 = new A.SolidFill();
                A.RgbColorModelHex rgbColorModelHex4 = new A.RgbColorModelHex() { Val = "E4BE84" };

                solidFill4.Append(rgbColorModelHex4);
                A.Round round2 = new A.Round();
                A.HeadEnd headEnd2 = new A.HeadEnd();
                A.TailEnd tailEnd2 = new A.TailEnd();

                outline2.Append(solidFill4);
                outline2.Append(round2);
                outline2.Append(headEnd2);
                outline2.Append(tailEnd2);

                shapeProperties2.Append(transform2D2);
                shapeProperties2.Append(presetGeometry2);
                shapeProperties2.Append(solidFill3);
                shapeProperties2.Append(outline2);

                Wps.TextBodyProperties textBodyProperties2 = new Wps.TextBodyProperties() { Rotation = 0, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, LeftInset = 91440, TopInset = 45720, RightInset = 91440, BottomInset = 45720, Anchor = A.TextAnchoringTypeValues.Top, AnchorCenter = false, UpRight = true };
                A.NoAutoFit noAutoFit2 = new A.NoAutoFit();

                textBodyProperties2.Append(noAutoFit2);

                wordprocessingShape2.Append(nonVisualDrawingProperties2);
                wordprocessingShape2.Append(nonVisualDrawingShapeProperties2);
                wordprocessingShape2.Append(shapeProperties2);
                wordprocessingShape2.Append(textBodyProperties2);

                Wps.WordprocessingShape wordprocessingShape3 = new Wps.WordprocessingShape();
                Wps.NonVisualDrawingProperties nonVisualDrawingProperties3 = new Wps.NonVisualDrawingProperties() { Id = (UInt32Value)25U, Name = "Text Box 44" };

                Wps.NonVisualDrawingShapeProperties nonVisualDrawingShapeProperties3 = new Wps.NonVisualDrawingShapeProperties() { TextBox = true };
                A.ShapeLocks shapeLocks3 = new A.ShapeLocks() { NoChangeArrowheads = true };

                nonVisualDrawingShapeProperties3.Append(shapeLocks3);

                Wps.ShapeProperties shapeProperties3 = new Wps.ShapeProperties() { BlackWhiteMode = A.BlackWhiteModeValues.Auto };

                A.Transform2D transform2D3 = new A.Transform2D();
                A.Offset offset4 = new A.Offset() { X = 732L, Y = 716L };
                A.Extents extents4 = new A.Extents() { Cx = 659L, Cy = 288L };

                transform2D3.Append(offset4);
                transform2D3.Append(extents4);

                A.PresetGeometry presetGeometry3 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
                A.AdjustValueList adjustValueList3 = new A.AdjustValueList();

                presetGeometry3.Append(adjustValueList3);
                A.NoFill noFill1 = new A.NoFill();

                A.Outline outline3 = new A.Outline();
                A.NoFill noFill2 = new A.NoFill();

                outline3.Append(noFill2);

                A.ShapePropertiesExtensionList shapePropertiesExtensionList1 = new A.ShapePropertiesExtensionList();

                A.ShapePropertiesExtension shapePropertiesExtension1 = new A.ShapePropertiesExtension() { Uri = "{909E8E84-426E-40DD-AFC4-6F175D3DCCD1}" };

                A14.HiddenFillProperties hiddenFillProperties1 = new A14.HiddenFillProperties();
                hiddenFillProperties1.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");

                A.SolidFill solidFill5 = new A.SolidFill();
                A.RgbColorModelHex rgbColorModelHex5 = new A.RgbColorModelHex() { Val = "FFFFFF" };

                solidFill5.Append(rgbColorModelHex5);

                hiddenFillProperties1.Append(solidFill5);

                shapePropertiesExtension1.Append(hiddenFillProperties1);

                A.ShapePropertiesExtension shapePropertiesExtension2 = new A.ShapePropertiesExtension() { Uri = "{91240B29-F687-4F45-9708-019B960494DF}" };

                A14.HiddenLineProperties hiddenLineProperties1 = new A14.HiddenLineProperties() { Width = 9525 };
                hiddenLineProperties1.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");

                A.SolidFill solidFill6 = new A.SolidFill();
                A.RgbColorModelHex rgbColorModelHex6 = new A.RgbColorModelHex() { Val = "000000" };

                solidFill6.Append(rgbColorModelHex6);
                A.Miter miter1 = new A.Miter() { Limit = 800000 };
                A.HeadEnd headEnd3 = new A.HeadEnd();
                A.TailEnd tailEnd3 = new A.TailEnd();

                hiddenLineProperties1.Append(solidFill6);
                hiddenLineProperties1.Append(miter1);
                hiddenLineProperties1.Append(headEnd3);
                hiddenLineProperties1.Append(tailEnd3);

                shapePropertiesExtension2.Append(hiddenLineProperties1);

                shapePropertiesExtensionList1.Append(shapePropertiesExtension1);
                shapePropertiesExtensionList1.Append(shapePropertiesExtension2);

                shapeProperties3.Append(transform2D3);
                shapeProperties3.Append(presetGeometry3);
                shapeProperties3.Append(noFill1);
                shapeProperties3.Append(outline3);
                shapeProperties3.Append(shapePropertiesExtensionList1);

                Wps.TextBoxInfo2 textBoxInfo21 = new Wps.TextBoxInfo2();

                TextBoxContent textBoxContent1 = new TextBoxContent();

                Paragraph paragraph2 = new Paragraph() { RsidParagraphAddition = "009D383E", RsidRunAdditionDefault = "009D383E", ParagraphId = "4B129D42", TextId = "77777777" };

                ParagraphProperties paragraphProperties2 = new ParagraphProperties();
                Justification justification1 = new Justification() { Val = JustificationValues.Right };

                paragraphProperties2.Append(justification1);

                Run run2 = new Run();
                FieldChar fieldChar1 = new FieldChar() { FieldCharType = FieldCharValues.Begin };

                run2.Append(fieldChar1);

                Run run3 = new Run();
                FieldCode fieldCode1 = new FieldCode() { Space = SpaceProcessingModeValues.Preserve };
                fieldCode1.Text = " PAGE    \\* MERGEFORMAT ";

                run3.Append(fieldCode1);

                Run run4 = new Run();
                FieldChar fieldChar2 = new FieldChar() { FieldCharType = FieldCharValues.Separate };

                run4.Append(fieldChar2);

                Run run5 = new Run();

                RunProperties runProperties2 = new RunProperties();
                Bold bold1 = new Bold();
                BoldComplexScript boldComplexScript1 = new BoldComplexScript();
                NoProof noProof2 = new NoProof();
                Color color1 = new Color() { Val = "FFFFFF", ThemeColor = ThemeColorValues.Background1 };

                runProperties2.Append(bold1);
                runProperties2.Append(boldComplexScript1);
                runProperties2.Append(noProof2);
                runProperties2.Append(color1);
                Text text1 = new Text();
                text1.Text = "2";

                run5.Append(runProperties2);
                run5.Append(text1);

                Run run6 = new Run();

                RunProperties runProperties3 = new RunProperties();
                Bold bold2 = new Bold();
                BoldComplexScript boldComplexScript2 = new BoldComplexScript();
                NoProof noProof3 = new NoProof();
                Color color2 = new Color() { Val = "FFFFFF", ThemeColor = ThemeColorValues.Background1 };

                runProperties3.Append(bold2);
                runProperties3.Append(boldComplexScript2);
                runProperties3.Append(noProof3);
                runProperties3.Append(color2);
                FieldChar fieldChar3 = new FieldChar() { FieldCharType = FieldCharValues.End };

                run6.Append(runProperties3);
                run6.Append(fieldChar3);

                paragraph2.Append(paragraphProperties2);
                paragraph2.Append(run2);
                paragraph2.Append(run3);
                paragraph2.Append(run4);
                paragraph2.Append(run5);
                paragraph2.Append(run6);

                textBoxContent1.Append(paragraph2);

                textBoxInfo21.Append(textBoxContent1);

                Wps.TextBodyProperties textBodyProperties3 = new Wps.TextBodyProperties() { Rotation = 0, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, LeftInset = 0, TopInset = 0, RightInset = 0, BottomInset = 0, Anchor = A.TextAnchoringTypeValues.Top, AnchorCenter = false, UpRight = true };
                A.NoAutoFit noAutoFit3 = new A.NoAutoFit();

                textBodyProperties3.Append(noAutoFit3);

                wordprocessingShape3.Append(nonVisualDrawingProperties3);
                wordprocessingShape3.Append(nonVisualDrawingShapeProperties3);
                wordprocessingShape3.Append(shapeProperties3);
                wordprocessingShape3.Append(textBoxInfo21);
                wordprocessingShape3.Append(textBodyProperties3);

                wordprocessingGroup1.Append(nonVisualGroupDrawingShapeProperties1);
                wordprocessingGroup1.Append(groupShapeProperties1);
                wordprocessingGroup1.Append(wordprocessingShape1);
                wordprocessingGroup1.Append(wordprocessingShape2);
                wordprocessingGroup1.Append(wordprocessingShape3);

                graphicData1.Append(wordprocessingGroup1);

                graphic1.Append(graphicData1);

                inline1.Append(extent1);
                inline1.Append(effectExtent1);
                inline1.Append(docProperties1);
                inline1.Append(nonVisualGraphicFrameDrawingProperties1);
                inline1.Append(graphic1);

                drawing1.Append(inline1);

                alternateContentChoice1.Append(drawing1);

                AlternateContentFallback alternateContentFallback1 = new AlternateContentFallback();

                Picture picture1 = new Picture();

                V.Group group1 = new V.Group() { Id = "Group 22", Style = "width:43.2pt;height:18.7pt;mso-position-horizontal-relative:char;mso-position-vertical-relative:line", CoordinateSize = "864,374", CoordinateOrigin = "614,660", OptionalString = "_x0000_s1026" };
                group1.SetAttribute(new OpenXmlAttribute("w14", "anchorId", "http://schemas.microsoft.com/office/word/2010/wordml", "553944E0"));
                group1.SetAttribute(new OpenXmlAttribute("o", "gfxdata", "urn:schemas-microsoft-com:office:office", "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\n"));

                V.RoundRectangle roundRectangle1 = new V.RoundRectangle() { Id = "AutoShape 42", Style = "position:absolute;left:859;top:415;width:374;height:864;rotation:-90;visibility:visible;mso-wrap-style:square;v-text-anchor:top", OptionalString = "_x0000_s1027", StrokeColor = "#e4be84", ArcSize = "10923f" };
                roundRectangle1.SetAttribute(new OpenXmlAttribute("o", "gfxdata", "urn:schemas-microsoft-com:office:office", "UEsDBBQABgAIAAAAIQDb4fbL7gAAAIUBAAATAAAAW0NvbnRlbnRfVHlwZXNdLnhtbHyQz07DMAyH\n70i8Q+QralM4IITa7kDhCAiNB7ASt43WOlEcyvb2pNu4IODoPz9/n1xv9vOkForiPDdwXVagiI23\njocG3rdPxR0oScgWJ8/UwIEENu3lRb09BBKV0ywNjCmFe63FjDSjlD4Q50nv44wpl3HQAc0OB9I3\nVXWrjedEnIq03oC27qjHjympx31un0wiTQLq4bS4shrAECZnMGVTvbD9QSnOhDInjzsyuiBXWQP0\nr4R18jfgnHvJr4nOknrFmJ5xzhraRtHWf3Kkpfz/yGo5S+H73hkquyhdjr3R8m2lj09svwAAAP//\nAwBQSwMEFAAGAAgAAAAhAFr0LFu/AAAAFQEAAAsAAABfcmVscy8ucmVsc2zPwWrDMAwG4Ptg72B0\nX5TuUMaI01uh19I+gLGVxCy2jGSy9e1nemrHjpL4P0nD4SetZiPRyNnCruvBUPYcYp4tXC/Htw8w\nWl0ObuVMFm6kcBhfX4Yzra62kC6xqGlKVgtLreUTUf1CyWnHhXKbTCzJ1VbKjMX5LzcTvvf9HuXR\ngPHJNKdgQU5hB+ZyK23zHztFL6w81c5zQp6m6P9TMfB3PtPWFCczVQtB9N4U2rp2HOA44NMz4y8A\nAAD//wMAUEsDBBQABgAIAAAAIQCr/nuhxAAAANsAAAAPAAAAZHJzL2Rvd25yZXYueG1sRI9Ba4NA\nFITvhfyH5QVykbjWQgnGTQgBoYeA1PbQ48N9UYn7VtyNmv76bqHQ4zAz3zD5cTG9mGh0nWUFz3EC\ngri2uuNGwedHsd2BcB5ZY2+ZFDzIwfGwesox03bmd5oq34gAYZehgtb7IZPS1S0ZdLEdiIN3taNB\nH+TYSD3iHOCml2mSvEqDHYeFFgc6t1TfqrtRoNPHTkZl0X9HRTndv3x1mYtKqc16Oe1BeFr8f/iv\n/aYVpC/w+yX8AHn4AQAA//8DAFBLAQItABQABgAIAAAAIQDb4fbL7gAAAIUBAAATAAAAAAAAAAAA\nAAAAAAAAAABbQ29udGVudF9UeXBlc10ueG1sUEsBAi0AFAAGAAgAAAAhAFr0LFu/AAAAFQEAAAsA\nAAAAAAAAAAAAAAAAHwEAAF9yZWxzLy5yZWxzUEsBAi0AFAAGAAgAAAAhAKv+e6HEAAAA2wAAAA8A\nAAAAAAAAAAAAAAAABwIAAGRycy9kb3ducmV2LnhtbFBLBQYAAAAAAwADALcAAAD4AgAAAAA=\n"));

                V.RoundRectangle roundRectangle2 = new V.RoundRectangle() { Id = "AutoShape 43", Style = "position:absolute;left:898;top:451;width:296;height:792;rotation:-90;visibility:visible;mso-wrap-style:square;v-text-anchor:top", OptionalString = "_x0000_s1028", FillColor = "#e4be84", StrokeColor = "#e4be84", ArcSize = "10923f" };
                roundRectangle2.SetAttribute(new OpenXmlAttribute("o", "gfxdata", "urn:schemas-microsoft-com:office:office", "UEsDBBQABgAIAAAAIQDb4fbL7gAAAIUBAAATAAAAW0NvbnRlbnRfVHlwZXNdLnhtbHyQz07DMAyH\n70i8Q+QralM4IITa7kDhCAiNB7ASt43WOlEcyvb2pNu4IODoPz9/n1xv9vOkForiPDdwXVagiI23\njocG3rdPxR0oScgWJ8/UwIEENu3lRb09BBKV0ywNjCmFe63FjDSjlD4Q50nv44wpl3HQAc0OB9I3\nVXWrjedEnIq03oC27qjHjympx31un0wiTQLq4bS4shrAECZnMGVTvbD9QSnOhDInjzsyuiBXWQP0\nr4R18jfgnHvJr4nOknrFmJ5xzhraRtHWf3Kkpfz/yGo5S+H73hkquyhdjr3R8m2lj09svwAAAP//\nAwBQSwMEFAAGAAgAAAAhAFr0LFu/AAAAFQEAAAsAAABfcmVscy8ucmVsc2zPwWrDMAwG4Ptg72B0\nX5TuUMaI01uh19I+gLGVxCy2jGSy9e1nemrHjpL4P0nD4SetZiPRyNnCruvBUPYcYp4tXC/Htw8w\nWl0ObuVMFm6kcBhfX4Yzra62kC6xqGlKVgtLreUTUf1CyWnHhXKbTCzJ1VbKjMX5LzcTvvf9HuXR\ngPHJNKdgQU5hB+ZyK23zHztFL6w81c5zQp6m6P9TMfB3PtPWFCczVQtB9N4U2rp2HOA44NMz4y8A\nAAD//wMAUEsDBBQABgAIAAAAIQATn8OnxQAAANsAAAAPAAAAZHJzL2Rvd25yZXYueG1sRI9BawIx\nFITvhf6H8ApepJtV2lq2RhFB8FbUInt83Tw3q5uXJYm69dc3BaHHYWa+Yabz3rbiQj40jhWMshwE\nceV0w7WCr93q+R1EiMgaW8ek4IcCzGePD1MstLvyhi7bWIsE4VCgAhNjV0gZKkMWQ+Y64uQdnLcY\nk/S11B6vCW5bOc7zN2mx4bRgsKOloeq0PVsFn6Usl6/l92SzyP3tMNrfaGiOSg2e+sUHiEh9/A/f\n22utYPwCf1/SD5CzXwAAAP//AwBQSwECLQAUAAYACAAAACEA2+H2y+4AAACFAQAAEwAAAAAAAAAA\nAAAAAAAAAAAAW0NvbnRlbnRfVHlwZXNdLnhtbFBLAQItABQABgAIAAAAIQBa9CxbvwAAABUBAAAL\nAAAAAAAAAAAAAAAAAB8BAABfcmVscy8ucmVsc1BLAQItABQABgAIAAAAIQATn8OnxQAAANsAAAAP\nAAAAAAAAAAAAAAAAAAcCAABkcnMvZG93bnJldi54bWxQSwUGAAAAAAMAAwC3AAAA+QIAAAAA\n"));

                V.Shapetype shapetype1 = new V.Shapetype() { Id = "_x0000_t202", CoordinateSize = "21600,21600", OptionalNumber = 202, EdgePath = "m,l,21600r21600,l21600,xe" };
                V.Stroke stroke1 = new V.Stroke() { JoinStyle = V.StrokeJoinStyleValues.Miter };
                V.Path path1 = new V.Path() { AllowGradientShape = true, ConnectionPointType = Ovml.ConnectValues.Rectangle };

                shapetype1.Append(stroke1);
                shapetype1.Append(path1);

                V.Shape shape1 = new V.Shape() { Id = "Text Box 44", Style = "position:absolute;left:732;top:716;width:659;height:288;visibility:visible;mso-wrap-style:square;v-text-anchor:top", OptionalString = "_x0000_s1029", Filled = false, Stroked = false, Type = "#_x0000_t202", EncodedPackage = "UEsDBBQABgAIAAAAIQDb4fbL7gAAAIUBAAATAAAAW0NvbnRlbnRfVHlwZXNdLnhtbHyQz07DMAyH\n70i8Q+QralM4IITa7kDhCAiNB7ASt43WOlEcyvb2pNu4IODoPz9/n1xv9vOkForiPDdwXVagiI23\njocG3rdPxR0oScgWJ8/UwIEENu3lRb09BBKV0ywNjCmFe63FjDSjlD4Q50nv44wpl3HQAc0OB9I3\nVXWrjedEnIq03oC27qjHjympx31un0wiTQLq4bS4shrAECZnMGVTvbD9QSnOhDInjzsyuiBXWQP0\nr4R18jfgnHvJr4nOknrFmJ5xzhraRtHWf3Kkpfz/yGo5S+H73hkquyhdjr3R8m2lj09svwAAAP//\nAwBQSwMEFAAGAAgAAAAhAFr0LFu/AAAAFQEAAAsAAABfcmVscy8ucmVsc2zPwWrDMAwG4Ptg72B0\nX5TuUMaI01uh19I+gLGVxCy2jGSy9e1nemrHjpL4P0nD4SetZiPRyNnCruvBUPYcYp4tXC/Htw8w\nWl0ObuVMFm6kcBhfX4Yzra62kC6xqGlKVgtLreUTUf1CyWnHhXKbTCzJ1VbKjMX5LzcTvvf9HuXR\ngPHJNKdgQU5hB+ZyK23zHztFL6w81c5zQp6m6P9TMfB3PtPWFCczVQtB9N4U2rp2HOA44NMz4y8A\nAAD//wMAUEsDBBQABgAIAAAAIQDLIsuSwwAAANsAAAAPAAAAZHJzL2Rvd25yZXYueG1sRI9Ba8JA\nFITvQv/D8gredKOg2OgqUhQKQjGmB4/P7DNZzL5Ns1uN/74rCB6HmfmGWaw6W4srtd44VjAaJiCI\nC6cNlwp+8u1gBsIHZI21Y1JwJw+r5Vtvgal2N87oegiliBD2KSqoQmhSKX1RkUU/dA1x9M6utRii\nbEupW7xFuK3lOEmm0qLhuFBhQ58VFZfDn1WwPnK2Mb/fp312zkyefyS8m16U6r936zmIQF14hZ/t\nL61gPIHHl/gD5PIfAAD//wMAUEsBAi0AFAAGAAgAAAAhANvh9svuAAAAhQEAABMAAAAAAAAAAAAA\nAAAAAAAAAFtDb250ZW50X1R5cGVzXS54bWxQSwECLQAUAAYACAAAACEAWvQsW78AAAAVAQAACwAA\nAAAAAAAAAAAAAAAfAQAAX3JlbHMvLnJlbHNQSwECLQAUAAYACAAAACEAyyLLksMAAADbAAAADwAA\nAAAAAAAAAAAAAAAHAgAAZHJzL2Rvd25yZXYueG1sUEsFBgAAAAADAAMAtwAAAPcCAAAAAA==\n" };

                V.TextBox textBox1 = new V.TextBox() { Inset = "0,0,0,0" };

                TextBoxContent textBoxContent2 = new TextBoxContent();

                Paragraph paragraph3 = new Paragraph() { RsidParagraphAddition = "009D383E", RsidRunAdditionDefault = "009D383E", ParagraphId = "4B129D42", TextId = "77777777" };

                ParagraphProperties paragraphProperties3 = new ParagraphProperties();
                Justification justification2 = new Justification() { Val = JustificationValues.Right };

                paragraphProperties3.Append(justification2);

                Run run7 = new Run();
                FieldChar fieldChar4 = new FieldChar() { FieldCharType = FieldCharValues.Begin };

                run7.Append(fieldChar4);

                Run run8 = new Run();
                FieldCode fieldCode2 = new FieldCode() { Space = SpaceProcessingModeValues.Preserve };
                fieldCode2.Text = " PAGE    \\* MERGEFORMAT ";

                run8.Append(fieldCode2);

                Run run9 = new Run();
                FieldChar fieldChar5 = new FieldChar() { FieldCharType = FieldCharValues.Separate };

                run9.Append(fieldChar5);

                Run run10 = new Run();

                RunProperties runProperties4 = new RunProperties();
                Bold bold3 = new Bold();
                BoldComplexScript boldComplexScript3 = new BoldComplexScript();
                NoProof noProof4 = new NoProof();
                Color color3 = new Color() { Val = "FFFFFF", ThemeColor = ThemeColorValues.Background1 };

                runProperties4.Append(bold3);
                runProperties4.Append(boldComplexScript3);
                runProperties4.Append(noProof4);
                runProperties4.Append(color3);
                Text text2 = new Text();
                text2.Text = "2";

                run10.Append(runProperties4);
                run10.Append(text2);

                Run run11 = new Run();

                RunProperties runProperties5 = new RunProperties();
                Bold bold4 = new Bold();
                BoldComplexScript boldComplexScript4 = new BoldComplexScript();
                NoProof noProof5 = new NoProof();
                Color color4 = new Color() { Val = "FFFFFF", ThemeColor = ThemeColorValues.Background1 };

                runProperties5.Append(bold4);
                runProperties5.Append(boldComplexScript4);
                runProperties5.Append(noProof5);
                runProperties5.Append(color4);
                FieldChar fieldChar6 = new FieldChar() { FieldCharType = FieldCharValues.End };

                run11.Append(runProperties5);
                run11.Append(fieldChar6);

                paragraph3.Append(paragraphProperties3);
                paragraph3.Append(run7);
                paragraph3.Append(run8);
                paragraph3.Append(run9);
                paragraph3.Append(run10);
                paragraph3.Append(run11);

                textBoxContent2.Append(paragraph3);

                textBox1.Append(textBoxContent2);

                shape1.Append(textBox1);
                Wvml.AnchorLock anchorLock1 = new Wvml.AnchorLock();

                group1.Append(roundRectangle1);
                group1.Append(roundRectangle2);
                group1.Append(shapetype1);
                group1.Append(shape1);
                group1.Append(anchorLock1);

                picture1.Append(group1);

                alternateContentFallback1.Append(picture1);

                alternateContent1.Append(alternateContentChoice1);
                alternateContent1.Append(alternateContentFallback1);

                run1.Append(runProperties1);
                run1.Append(alternateContent1);

                paragraph1.Append(paragraphProperties1);
                paragraph1.Append(run1);

                sdtContentBlock1.Append(paragraph1);

                sdtBlock1.Append(sdtProperties1);
                sdtBlock1.Append(sdtContentBlock1);
                return sdtBlock1;

            }
        }
        private static SdtBlock Circle1 {
            get {
                SdtBlock sdtBlock1 = new SdtBlock();

                SdtProperties sdtProperties1 = new SdtProperties();
                SdtId sdtId1 = new SdtId() { Val = -259830808 };

                SdtContentDocPartObject sdtContentDocPartObject1 = new SdtContentDocPartObject();
                DocPartGallery docPartGallery1 = new DocPartGallery() { Val = "Page Numbers (Top of Page)" };
                DocPartUnique docPartUnique1 = new DocPartUnique();

                sdtContentDocPartObject1.Append(docPartGallery1);
                sdtContentDocPartObject1.Append(docPartUnique1);

                sdtProperties1.Append(sdtId1);
                sdtProperties1.Append(sdtContentDocPartObject1);

                SdtContentBlock sdtContentBlock1 = new SdtContentBlock();

                Paragraph paragraph1 = new Paragraph() { RsidParagraphMarkRevision = "000455E6", RsidParagraphAddition = "00E11B1E", RsidParagraphProperties = "000455E6", RsidRunAdditionDefault = "000455E6", ParagraphId = "16AD3494", TextId = "561B8176" };

                ParagraphProperties paragraphProperties1 = new ParagraphProperties();
                ParagraphStyleId paragraphStyleId1 = new ParagraphStyleId() { Val = "Header" };

                paragraphProperties1.Append(paragraphStyleId1);

                Run run1 = new Run();

                RunProperties runProperties1 = new RunProperties();
                NoProof noProof1 = new NoProof();

                runProperties1.Append(noProof1);

                AlternateContent alternateContent1 = new AlternateContent();

                AlternateContentChoice alternateContentChoice1 = new AlternateContentChoice() { Requires = "wps" };

                Drawing drawing1 = new Drawing();

                Wp.Anchor anchor1 = new Wp.Anchor() { DistanceFromTop = (UInt32Value)0U, DistanceFromBottom = (UInt32Value)0U, DistanceFromLeft = (UInt32Value)114300U, DistanceFromRight = (UInt32Value)114300U, SimplePos = false, RelativeHeight = (UInt32Value)251659264U, BehindDoc = false, Locked = false, LayoutInCell = false, AllowOverlap = true, EditId = "60FAEEAB", AnchorId = "46D32B50" };
                Wp.SimplePosition simplePosition1 = new Wp.SimplePosition() { X = 0L, Y = 0L };

                Wp.HorizontalPosition horizontalPosition1 = new Wp.HorizontalPosition() { RelativeFrom = Wp.HorizontalRelativePositionValues.Margin };
                Wp.HorizontalAlignment horizontalAlignment1 = new Wp.HorizontalAlignment();
                horizontalAlignment1.Text = "center";

                horizontalPosition1.Append(horizontalAlignment1);

                Wp.VerticalPosition verticalPosition1 = new Wp.VerticalPosition() { RelativeFrom = Wp.VerticalRelativePositionValues.TopMargin };
                Wp.VerticalAlignment verticalAlignment1 = new Wp.VerticalAlignment();
                verticalAlignment1.Text = "center";

                verticalPosition1.Append(verticalAlignment1);
                Wp.Extent extent1 = new Wp.Extent() { Cx = 626745L, Cy = 626745L };
                Wp.EffectExtent effectExtent1 = new Wp.EffectExtent() { LeftEdge = 0L, TopEdge = 0L, RightEdge = 1905L, BottomEdge = 1905L };
                Wp.WrapNone wrapNone1 = new Wp.WrapNone();
                Wp.DocProperties docProperties1 = new Wp.DocProperties() { Id = (UInt32Value)21U, Name = "Oval 21" };

                Wp.NonVisualGraphicFrameDrawingProperties nonVisualGraphicFrameDrawingProperties1 = new Wp.NonVisualGraphicFrameDrawingProperties();

                A.GraphicFrameLocks graphicFrameLocks1 = new A.GraphicFrameLocks();
                graphicFrameLocks1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

                nonVisualGraphicFrameDrawingProperties1.Append(graphicFrameLocks1);

                A.Graphic graphic1 = new A.Graphic();
                graphic1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

                A.GraphicData graphicData1 = new A.GraphicData() { Uri = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape" };

                Wps.WordprocessingShape wordprocessingShape1 = new Wps.WordprocessingShape();

                Wps.NonVisualDrawingShapeProperties nonVisualDrawingShapeProperties1 = new Wps.NonVisualDrawingShapeProperties();
                A.ShapeLocks shapeLocks1 = new A.ShapeLocks() { NoChangeArrowheads = true };

                nonVisualDrawingShapeProperties1.Append(shapeLocks1);

                Wps.ShapeProperties shapeProperties1 = new Wps.ShapeProperties() { BlackWhiteMode = A.BlackWhiteModeValues.Auto };

                A.Transform2D transform2D1 = new A.Transform2D();
                A.Offset offset1 = new A.Offset() { X = 0L, Y = 0L };
                A.Extents extents1 = new A.Extents() { Cx = 626745L, Cy = 626745L };

                transform2D1.Append(offset1);
                transform2D1.Append(extents1);

                A.PresetGeometry presetGeometry1 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Ellipse };
                A.AdjustValueList adjustValueList1 = new A.AdjustValueList();

                presetGeometry1.Append(adjustValueList1);

                A.SolidFill solidFill1 = new A.SolidFill();
                A.RgbColorModelHex rgbColorModelHex1 = new A.RgbColorModelHex() { Val = "40618B" };

                solidFill1.Append(rgbColorModelHex1);

                A.Outline outline1 = new A.Outline();
                A.NoFill noFill1 = new A.NoFill();

                outline1.Append(noFill1);

                A.ShapePropertiesExtensionList shapePropertiesExtensionList1 = new A.ShapePropertiesExtensionList();

                A.ShapePropertiesExtension shapePropertiesExtension1 = new A.ShapePropertiesExtension() { Uri = "{91240B29-F687-4F45-9708-019B960494DF}" };

                A14.HiddenLineProperties hiddenLineProperties1 = new A14.HiddenLineProperties() { Width = 9525 };
                hiddenLineProperties1.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");

                A.SolidFill solidFill2 = new A.SolidFill();
                A.RgbColorModelHex rgbColorModelHex2 = new A.RgbColorModelHex() { Val = "000000" };

                solidFill2.Append(rgbColorModelHex2);
                A.Round round1 = new A.Round();
                A.HeadEnd headEnd1 = new A.HeadEnd();
                A.TailEnd tailEnd1 = new A.TailEnd();

                hiddenLineProperties1.Append(solidFill2);
                hiddenLineProperties1.Append(round1);
                hiddenLineProperties1.Append(headEnd1);
                hiddenLineProperties1.Append(tailEnd1);

                shapePropertiesExtension1.Append(hiddenLineProperties1);

                shapePropertiesExtensionList1.Append(shapePropertiesExtension1);

                shapeProperties1.Append(transform2D1);
                shapeProperties1.Append(presetGeometry1);
                shapeProperties1.Append(solidFill1);
                shapeProperties1.Append(outline1);
                shapeProperties1.Append(shapePropertiesExtensionList1);

                Wps.TextBoxInfo2 textBoxInfo21 = new Wps.TextBoxInfo2();

                TextBoxContent textBoxContent1 = new TextBoxContent();

                Paragraph paragraph2 = new Paragraph() { RsidParagraphAddition = "000455E6", RsidRunAdditionDefault = "000455E6", ParagraphId = "2E317082", TextId = "77777777" };

                ParagraphProperties paragraphProperties2 = new ParagraphProperties();
                ParagraphStyleId paragraphStyleId2 = new ParagraphStyleId() { Val = "Footer" };
                Justification justification1 = new Justification() { Val = JustificationValues.Center };

                ParagraphMarkRunProperties paragraphMarkRunProperties1 = new ParagraphMarkRunProperties();
                Bold bold1 = new Bold();
                BoldComplexScript boldComplexScript1 = new BoldComplexScript();
                Color color1 = new Color() { Val = "FFFFFF", ThemeColor = ThemeColorValues.Background1 };
                FontSize fontSize1 = new FontSize() { Val = "32" };
                FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "32" };

                paragraphMarkRunProperties1.Append(bold1);
                paragraphMarkRunProperties1.Append(boldComplexScript1);
                paragraphMarkRunProperties1.Append(color1);
                paragraphMarkRunProperties1.Append(fontSize1);
                paragraphMarkRunProperties1.Append(fontSizeComplexScript1);

                paragraphProperties2.Append(paragraphStyleId2);
                paragraphProperties2.Append(justification1);
                paragraphProperties2.Append(paragraphMarkRunProperties1);

                Run run2 = new Run();
                FieldChar fieldChar1 = new FieldChar() { FieldCharType = FieldCharValues.Begin };

                run2.Append(fieldChar1);

                Run run3 = new Run();
                FieldCode fieldCode1 = new FieldCode() { Space = SpaceProcessingModeValues.Preserve };
                fieldCode1.Text = " PAGE    \\* MERGEFORMAT ";

                run3.Append(fieldCode1);

                Run run4 = new Run();
                FieldChar fieldChar2 = new FieldChar() { FieldCharType = FieldCharValues.Separate };

                run4.Append(fieldChar2);

                Run run5 = new Run();

                RunProperties runProperties2 = new RunProperties();
                Bold bold2 = new Bold();
                BoldComplexScript boldComplexScript2 = new BoldComplexScript();
                NoProof noProof2 = new NoProof();
                Color color2 = new Color() { Val = "FFFFFF", ThemeColor = ThemeColorValues.Background1 };
                FontSize fontSize2 = new FontSize() { Val = "32" };
                FontSizeComplexScript fontSizeComplexScript2 = new FontSizeComplexScript() { Val = "32" };

                runProperties2.Append(bold2);
                runProperties2.Append(boldComplexScript2);
                runProperties2.Append(noProof2);
                runProperties2.Append(color2);
                runProperties2.Append(fontSize2);
                runProperties2.Append(fontSizeComplexScript2);
                Text text1 = new Text();
                text1.Text = "2";

                run5.Append(runProperties2);
                run5.Append(text1);

                Run run6 = new Run();

                RunProperties runProperties3 = new RunProperties();
                Bold bold3 = new Bold();
                BoldComplexScript boldComplexScript3 = new BoldComplexScript();
                NoProof noProof3 = new NoProof();
                Color color3 = new Color() { Val = "FFFFFF", ThemeColor = ThemeColorValues.Background1 };
                FontSize fontSize3 = new FontSize() { Val = "32" };
                FontSizeComplexScript fontSizeComplexScript3 = new FontSizeComplexScript() { Val = "32" };

                runProperties3.Append(bold3);
                runProperties3.Append(boldComplexScript3);
                runProperties3.Append(noProof3);
                runProperties3.Append(color3);
                runProperties3.Append(fontSize3);
                runProperties3.Append(fontSizeComplexScript3);
                FieldChar fieldChar3 = new FieldChar() { FieldCharType = FieldCharValues.End };

                run6.Append(runProperties3);
                run6.Append(fieldChar3);

                paragraph2.Append(paragraphProperties2);
                paragraph2.Append(run2);
                paragraph2.Append(run3);
                paragraph2.Append(run4);
                paragraph2.Append(run5);
                paragraph2.Append(run6);

                textBoxContent1.Append(paragraph2);

                textBoxInfo21.Append(textBoxContent1);

                Wps.TextBodyProperties textBodyProperties1 = new Wps.TextBodyProperties() { Rotation = 0, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, LeftInset = 91440, TopInset = 45720, RightInset = 91440, BottomInset = 45720, Anchor = A.TextAnchoringTypeValues.Center, AnchorCenter = false, UpRight = true };
                A.NoAutoFit noAutoFit1 = new A.NoAutoFit();

                textBodyProperties1.Append(noAutoFit1);

                wordprocessingShape1.Append(nonVisualDrawingShapeProperties1);
                wordprocessingShape1.Append(shapeProperties1);
                wordprocessingShape1.Append(textBoxInfo21);
                wordprocessingShape1.Append(textBodyProperties1);

                graphicData1.Append(wordprocessingShape1);

                graphic1.Append(graphicData1);

                Wp14.RelativeWidth relativeWidth1 = new Wp14.RelativeWidth() { ObjectId = Wp14.SizeRelativeHorizontallyValues.Page };
                Wp14.PercentageWidth percentageWidth1 = new Wp14.PercentageWidth();
                percentageWidth1.Text = "0";

                relativeWidth1.Append(percentageWidth1);

                Wp14.RelativeHeight relativeHeight1 = new Wp14.RelativeHeight() { RelativeFrom = Wp14.SizeRelativeVerticallyValues.Page };
                Wp14.PercentageHeight percentageHeight1 = new Wp14.PercentageHeight();
                percentageHeight1.Text = "0";

                relativeHeight1.Append(percentageHeight1);

                anchor1.Append(simplePosition1);
                anchor1.Append(horizontalPosition1);
                anchor1.Append(verticalPosition1);
                anchor1.Append(extent1);
                anchor1.Append(effectExtent1);
                anchor1.Append(wrapNone1);
                anchor1.Append(docProperties1);
                anchor1.Append(nonVisualGraphicFrameDrawingProperties1);
                anchor1.Append(graphic1);
                anchor1.Append(relativeWidth1);
                anchor1.Append(relativeHeight1);

                drawing1.Append(anchor1);

                alternateContentChoice1.Append(drawing1);

                AlternateContentFallback alternateContentFallback1 = new AlternateContentFallback();

                Picture picture1 = new Picture();

                V.Oval oval1 = new V.Oval() { Id = "Oval 21", Style = "position:absolute;margin-left:0;margin-top:0;width:49.35pt;height:49.35pt;z-index:251659264;visibility:visible;mso-wrap-style:square;mso-width-percent:0;mso-height-percent:0;mso-wrap-distance-left:9pt;mso-wrap-distance-top:0;mso-wrap-distance-right:9pt;mso-wrap-distance-bottom:0;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:top-margin-area;mso-width-percent:0;mso-height-percent:0;mso-width-relative:page;mso-height-relative:page;v-text-anchor:middle", OptionalString = "_x0000_s1026", AllowInCell = false, FillColor = "#40618b", Stroked = false };
                oval1.SetAttribute(new OpenXmlAttribute("w14", "anchorId", "http://schemas.microsoft.com/office/word/2010/wordml", "46D32B50"));
                oval1.SetAttribute(new OpenXmlAttribute("o", "gfxdata", "urn:schemas-microsoft-com:office:office", "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\nZXYueG1sUEsFBgAAAAAEAAQA8wAAAE8FAAAAAA==\n"));

                V.TextBox textBox1 = new V.TextBox();

                TextBoxContent textBoxContent2 = new TextBoxContent();

                Paragraph paragraph3 = new Paragraph() { RsidParagraphAddition = "000455E6", RsidRunAdditionDefault = "000455E6", ParagraphId = "2E317082", TextId = "77777777" };

                ParagraphProperties paragraphProperties3 = new ParagraphProperties();
                ParagraphStyleId paragraphStyleId3 = new ParagraphStyleId() { Val = "Footer" };
                Justification justification2 = new Justification() { Val = JustificationValues.Center };

                ParagraphMarkRunProperties paragraphMarkRunProperties2 = new ParagraphMarkRunProperties();
                Bold bold4 = new Bold();
                BoldComplexScript boldComplexScript4 = new BoldComplexScript();
                Color color4 = new Color() { Val = "FFFFFF", ThemeColor = ThemeColorValues.Background1 };
                FontSize fontSize4 = new FontSize() { Val = "32" };
                FontSizeComplexScript fontSizeComplexScript4 = new FontSizeComplexScript() { Val = "32" };

                paragraphMarkRunProperties2.Append(bold4);
                paragraphMarkRunProperties2.Append(boldComplexScript4);
                paragraphMarkRunProperties2.Append(color4);
                paragraphMarkRunProperties2.Append(fontSize4);
                paragraphMarkRunProperties2.Append(fontSizeComplexScript4);

                paragraphProperties3.Append(paragraphStyleId3);
                paragraphProperties3.Append(justification2);
                paragraphProperties3.Append(paragraphMarkRunProperties2);

                Run run7 = new Run();
                FieldChar fieldChar4 = new FieldChar() { FieldCharType = FieldCharValues.Begin };

                run7.Append(fieldChar4);

                Run run8 = new Run();
                FieldCode fieldCode2 = new FieldCode() { Space = SpaceProcessingModeValues.Preserve };
                fieldCode2.Text = " PAGE    \\* MERGEFORMAT ";

                run8.Append(fieldCode2);

                Run run9 = new Run();
                FieldChar fieldChar5 = new FieldChar() { FieldCharType = FieldCharValues.Separate };

                run9.Append(fieldChar5);

                Run run10 = new Run();

                RunProperties runProperties4 = new RunProperties();
                Bold bold5 = new Bold();
                BoldComplexScript boldComplexScript5 = new BoldComplexScript();
                NoProof noProof4 = new NoProof();
                Color color5 = new Color() { Val = "FFFFFF", ThemeColor = ThemeColorValues.Background1 };
                FontSize fontSize5 = new FontSize() { Val = "32" };
                FontSizeComplexScript fontSizeComplexScript5 = new FontSizeComplexScript() { Val = "32" };

                runProperties4.Append(bold5);
                runProperties4.Append(boldComplexScript5);
                runProperties4.Append(noProof4);
                runProperties4.Append(color5);
                runProperties4.Append(fontSize5);
                runProperties4.Append(fontSizeComplexScript5);
                Text text2 = new Text();
                text2.Text = "2";

                run10.Append(runProperties4);
                run10.Append(text2);

                Run run11 = new Run();

                RunProperties runProperties5 = new RunProperties();
                Bold bold6 = new Bold();
                BoldComplexScript boldComplexScript6 = new BoldComplexScript();
                NoProof noProof5 = new NoProof();
                Color color6 = new Color() { Val = "FFFFFF", ThemeColor = ThemeColorValues.Background1 };
                FontSize fontSize6 = new FontSize() { Val = "32" };
                FontSizeComplexScript fontSizeComplexScript6 = new FontSizeComplexScript() { Val = "32" };

                runProperties5.Append(bold6);
                runProperties5.Append(boldComplexScript6);
                runProperties5.Append(noProof5);
                runProperties5.Append(color6);
                runProperties5.Append(fontSize6);
                runProperties5.Append(fontSizeComplexScript6);
                FieldChar fieldChar6 = new FieldChar() { FieldCharType = FieldCharValues.End };

                run11.Append(runProperties5);
                run11.Append(fieldChar6);

                paragraph3.Append(paragraphProperties3);
                paragraph3.Append(run7);
                paragraph3.Append(run8);
                paragraph3.Append(run9);
                paragraph3.Append(run10);
                paragraph3.Append(run11);

                textBoxContent2.Append(paragraph3);

                textBox1.Append(textBoxContent2);
                Wvml.TextWrap textWrap1 = new Wvml.TextWrap() { AnchorX = Wvml.HorizontalAnchorValues.Margin, AnchorY = Wvml.VerticalAnchorValues.Margin };

                oval1.Append(textBox1);
                oval1.Append(textWrap1);

                picture1.Append(oval1);

                alternateContentFallback1.Append(picture1);

                alternateContent1.Append(alternateContentChoice1);
                alternateContent1.Append(alternateContentFallback1);

                run1.Append(runProperties1);
                run1.Append(alternateContent1);

                paragraph1.Append(paragraphProperties1);
                paragraph1.Append(run1);

                sdtContentBlock1.Append(paragraph1);

                sdtBlock1.Append(sdtProperties1);
                sdtBlock1.Append(sdtContentBlock1);
                return sdtBlock1;

            }
        }
        private static SdtBlock VeryLarge1 {
            get {
                SdtBlock sdtBlock1 = new SdtBlock();

                SdtProperties sdtProperties1 = new SdtProperties();
                SdtId sdtId1 = new SdtId() { Val = 866250882 };

                SdtContentDocPartObject sdtContentDocPartObject1 = new SdtContentDocPartObject();
                DocPartGallery docPartGallery1 = new DocPartGallery() { Val = "Page Numbers (Top of Page)" };
                DocPartUnique docPartUnique1 = new DocPartUnique();

                sdtContentDocPartObject1.Append(docPartGallery1);
                sdtContentDocPartObject1.Append(docPartUnique1);

                sdtProperties1.Append(sdtId1);
                sdtProperties1.Append(sdtContentDocPartObject1);

                SdtContentBlock sdtContentBlock1 = new SdtContentBlock();

                Paragraph paragraph1 = new Paragraph() { RsidParagraphMarkRevision = "007B0E26", RsidParagraphAddition = "00E11B1E", RsidParagraphProperties = "007B0E26", RsidRunAdditionDefault = "007B0E26", ParagraphId = "16AD3494", TextId = "23E16AF3" };

                ParagraphProperties paragraphProperties1 = new ParagraphProperties();
                ParagraphStyleId paragraphStyleId1 = new ParagraphStyleId() { Val = "Header" };

                paragraphProperties1.Append(paragraphStyleId1);

                Run run1 = new Run();

                RunProperties runProperties1 = new RunProperties();
                NoProof noProof1 = new NoProof();

                runProperties1.Append(noProof1);

                AlternateContent alternateContent1 = new AlternateContent();

                AlternateContentChoice alternateContentChoice1 = new AlternateContentChoice() { Requires = "wps" };

                Drawing drawing1 = new Drawing();

                Wp.Anchor anchor1 = new Wp.Anchor() { DistanceFromTop = (UInt32Value)0U, DistanceFromBottom = (UInt32Value)0U, DistanceFromLeft = (UInt32Value)114300U, DistanceFromRight = (UInt32Value)114300U, SimplePos = false, RelativeHeight = (UInt32Value)251659264U, BehindDoc = false, Locked = false, LayoutInCell = false, AllowOverlap = true, EditId = "6A814248", AnchorId = "5EC28866" };
                Wp.SimplePosition simplePosition1 = new Wp.SimplePosition() { X = 0L, Y = 0L };

                AlternateContent alternateContent2 = new AlternateContent();

                AlternateContentChoice alternateContentChoice2 = new AlternateContentChoice() { Requires = "wp14" };

                Wp.HorizontalPosition horizontalPosition1 = new Wp.HorizontalPosition() { RelativeFrom = Wp.HorizontalRelativePositionValues.Margin };
                Wp14.PercentagePositionHeightOffset percentagePositionHeightOffset1 = new Wp14.PercentagePositionHeightOffset();
                percentagePositionHeightOffset1.Text = "80000";

                horizontalPosition1.Append(percentagePositionHeightOffset1);

                alternateContentChoice2.Append(horizontalPosition1);

                AlternateContentFallback alternateContentFallback1 = new AlternateContentFallback();

                Wp.HorizontalPosition horizontalPosition2 = new Wp.HorizontalPosition() { RelativeFrom = Wp.HorizontalRelativePositionValues.Page };
                Wp.PositionOffset positionOffset1 = new Wp.PositionOffset();
                positionOffset1.Text = "5669280";

                horizontalPosition2.Append(positionOffset1);

                alternateContentFallback1.Append(horizontalPosition2);

                alternateContent2.Append(alternateContentChoice2);
                alternateContent2.Append(alternateContentFallback1);

                Wp.VerticalPosition verticalPosition1 = new Wp.VerticalPosition() { RelativeFrom = Wp.VerticalRelativePositionValues.Page };
                Wp.PositionOffset positionOffset2 = new Wp.PositionOffset();
                positionOffset2.Text = "365760";

                verticalPosition1.Append(positionOffset2);
                Wp.Extent extent1 = new Wp.Extent() { Cx = 1811655L, Cy = 1346835L };
                Wp.EffectExtent effectExtent1 = new Wp.EffectExtent() { LeftEdge = 0L, TopEdge = 3810L, RightEdge = 0L, BottomEdge = 1905L };
                Wp.WrapNone wrapNone1 = new Wp.WrapNone();
                Wp.DocProperties docProperties1 = new Wp.DocProperties() { Id = (UInt32Value)20U, Name = "Rectangle 20" };

                Wp.NonVisualGraphicFrameDrawingProperties nonVisualGraphicFrameDrawingProperties1 = new Wp.NonVisualGraphicFrameDrawingProperties();

                A.GraphicFrameLocks graphicFrameLocks1 = new A.GraphicFrameLocks();
                graphicFrameLocks1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

                nonVisualGraphicFrameDrawingProperties1.Append(graphicFrameLocks1);

                A.Graphic graphic1 = new A.Graphic();
                graphic1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

                A.GraphicData graphicData1 = new A.GraphicData() { Uri = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape" };

                Wps.WordprocessingShape wordprocessingShape1 = new Wps.WordprocessingShape();

                Wps.NonVisualDrawingShapeProperties nonVisualDrawingShapeProperties1 = new Wps.NonVisualDrawingShapeProperties();
                A.ShapeLocks shapeLocks1 = new A.ShapeLocks() { NoChangeArrowheads = true };

                nonVisualDrawingShapeProperties1.Append(shapeLocks1);

                Wps.ShapeProperties shapeProperties1 = new Wps.ShapeProperties() { BlackWhiteMode = A.BlackWhiteModeValues.Auto };

                A.Transform2D transform2D1 = new A.Transform2D();
                A.Offset offset1 = new A.Offset() { X = 0L, Y = 0L };
                A.Extents extents1 = new A.Extents() { Cx = 1811655L, Cy = 1346835L };

                transform2D1.Append(offset1);
                transform2D1.Append(extents1);

                A.PresetGeometry presetGeometry1 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
                A.AdjustValueList adjustValueList1 = new A.AdjustValueList();

                presetGeometry1.Append(adjustValueList1);

                A.SolidFill solidFill1 = new A.SolidFill();
                A.RgbColorModelHex rgbColorModelHex1 = new A.RgbColorModelHex() { Val = "FFFFFF" };

                solidFill1.Append(rgbColorModelHex1);

                A.Outline outline1 = new A.Outline();
                A.NoFill noFill1 = new A.NoFill();

                outline1.Append(noFill1);

                A.ShapePropertiesExtensionList shapePropertiesExtensionList1 = new A.ShapePropertiesExtensionList();

                A.ShapePropertiesExtension shapePropertiesExtension1 = new A.ShapePropertiesExtension() { Uri = "{91240B29-F687-4F45-9708-019B960494DF}" };

                A14.HiddenLineProperties hiddenLineProperties1 = new A14.HiddenLineProperties() { Width = 9525 };
                hiddenLineProperties1.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");

                A.SolidFill solidFill2 = new A.SolidFill();
                A.RgbColorModelHex rgbColorModelHex2 = new A.RgbColorModelHex() { Val = "000000" };

                solidFill2.Append(rgbColorModelHex2);
                A.Miter miter1 = new A.Miter() { Limit = 800000 };
                A.HeadEnd headEnd1 = new A.HeadEnd();
                A.TailEnd tailEnd1 = new A.TailEnd();

                hiddenLineProperties1.Append(solidFill2);
                hiddenLineProperties1.Append(miter1);
                hiddenLineProperties1.Append(headEnd1);
                hiddenLineProperties1.Append(tailEnd1);

                shapePropertiesExtension1.Append(hiddenLineProperties1);

                shapePropertiesExtensionList1.Append(shapePropertiesExtension1);

                shapeProperties1.Append(transform2D1);
                shapeProperties1.Append(presetGeometry1);
                shapeProperties1.Append(solidFill1);
                shapeProperties1.Append(outline1);
                shapeProperties1.Append(shapePropertiesExtensionList1);

                Wps.TextBoxInfo2 textBoxInfo21 = new Wps.TextBoxInfo2();

                TextBoxContent textBoxContent1 = new TextBoxContent();

                Paragraph paragraph2 = new Paragraph() { RsidParagraphAddition = "007B0E26", RsidRunAdditionDefault = "007B0E26", ParagraphId = "67F79515", TextId = "77777777" };

                ParagraphProperties paragraphProperties2 = new ParagraphProperties();
                Justification justification1 = new Justification() { Val = JustificationValues.Right };

                ParagraphMarkRunProperties paragraphMarkRunProperties1 = new ParagraphMarkRunProperties();
                Color color1 = new Color() { Val = "A6A6A6", ThemeColor = ThemeColorValues.Background1, ThemeShade = "A6" };
                FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "144" };

                paragraphMarkRunProperties1.Append(color1);
                paragraphMarkRunProperties1.Append(fontSizeComplexScript1);

                paragraphProperties2.Append(justification1);
                paragraphProperties2.Append(paragraphMarkRunProperties1);

                Run run2 = new Run();
                FieldChar fieldChar1 = new FieldChar() { FieldCharType = FieldCharValues.Begin };

                run2.Append(fieldChar1);

                Run run3 = new Run();
                FieldCode fieldCode1 = new FieldCode() { Space = SpaceProcessingModeValues.Preserve };
                fieldCode1.Text = " PAGE    \\* MERGEFORMAT ";

                run3.Append(fieldCode1);

                Run run4 = new Run();
                FieldChar fieldChar2 = new FieldChar() { FieldCharType = FieldCharValues.Separate };

                run4.Append(fieldChar2);

                Run run5 = new Run();

                RunProperties runProperties2 = new RunProperties();
                NoProof noProof2 = new NoProof();
                Color color2 = new Color() { Val = "A6A6A6", ThemeColor = ThemeColorValues.Background1, ThemeShade = "A6" };
                FontSize fontSize1 = new FontSize() { Val = "144" };
                FontSizeComplexScript fontSizeComplexScript2 = new FontSizeComplexScript() { Val = "144" };

                runProperties2.Append(noProof2);
                runProperties2.Append(color2);
                runProperties2.Append(fontSize1);
                runProperties2.Append(fontSizeComplexScript2);
                Text text1 = new Text();
                text1.Text = "2";

                run5.Append(runProperties2);
                run5.Append(text1);

                Run run6 = new Run();

                RunProperties runProperties3 = new RunProperties();
                NoProof noProof3 = new NoProof();
                Color color3 = new Color() { Val = "A6A6A6", ThemeColor = ThemeColorValues.Background1, ThemeShade = "A6" };
                FontSize fontSize2 = new FontSize() { Val = "144" };
                FontSizeComplexScript fontSizeComplexScript3 = new FontSizeComplexScript() { Val = "144" };

                runProperties3.Append(noProof3);
                runProperties3.Append(color3);
                runProperties3.Append(fontSize2);
                runProperties3.Append(fontSizeComplexScript3);
                FieldChar fieldChar3 = new FieldChar() { FieldCharType = FieldCharValues.End };

                run6.Append(runProperties3);
                run6.Append(fieldChar3);

                paragraph2.Append(paragraphProperties2);
                paragraph2.Append(run2);
                paragraph2.Append(run3);
                paragraph2.Append(run4);
                paragraph2.Append(run5);
                paragraph2.Append(run6);

                textBoxContent1.Append(paragraph2);

                textBoxInfo21.Append(textBoxContent1);

                Wps.TextBodyProperties textBodyProperties1 = new Wps.TextBodyProperties() { Rotation = 0, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, LeftInset = 91440, TopInset = 45720, RightInset = 91440, BottomInset = 45720, Anchor = A.TextAnchoringTypeValues.Top, AnchorCenter = false, UpRight = true };
                A.NoAutoFit noAutoFit1 = new A.NoAutoFit();

                textBodyProperties1.Append(noAutoFit1);

                wordprocessingShape1.Append(nonVisualDrawingShapeProperties1);
                wordprocessingShape1.Append(shapeProperties1);
                wordprocessingShape1.Append(textBoxInfo21);
                wordprocessingShape1.Append(textBodyProperties1);

                graphicData1.Append(wordprocessingShape1);

                graphic1.Append(graphicData1);

                Wp14.RelativeWidth relativeWidth1 = new Wp14.RelativeWidth() { ObjectId = Wp14.SizeRelativeHorizontallyValues.Page };
                Wp14.PercentageWidth percentageWidth1 = new Wp14.PercentageWidth();
                percentageWidth1.Text = "0";

                relativeWidth1.Append(percentageWidth1);

                Wp14.RelativeHeight relativeHeight1 = new Wp14.RelativeHeight() { RelativeFrom = Wp14.SizeRelativeVerticallyValues.Page };
                Wp14.PercentageHeight percentageHeight1 = new Wp14.PercentageHeight();
                percentageHeight1.Text = "0";

                relativeHeight1.Append(percentageHeight1);

                anchor1.Append(simplePosition1);
                anchor1.Append(alternateContent2);
                anchor1.Append(verticalPosition1);
                anchor1.Append(extent1);
                anchor1.Append(effectExtent1);
                anchor1.Append(wrapNone1);
                anchor1.Append(docProperties1);
                anchor1.Append(nonVisualGraphicFrameDrawingProperties1);
                anchor1.Append(graphic1);
                anchor1.Append(relativeWidth1);
                anchor1.Append(relativeHeight1);

                drawing1.Append(anchor1);

                alternateContentChoice1.Append(drawing1);

                AlternateContentFallback alternateContentFallback2 = new AlternateContentFallback();

                Picture picture1 = new Picture();

                V.Rectangle rectangle1 = new V.Rectangle() { Id = "Rectangle 20", Style = "position:absolute;margin-left:0;margin-top:28.8pt;width:142.65pt;height:106.05pt;z-index:251659264;visibility:visible;mso-wrap-style:square;mso-width-percent:0;mso-height-percent:0;mso-left-percent:800;mso-wrap-distance-left:9pt;mso-wrap-distance-top:0;mso-wrap-distance-right:9pt;mso-wrap-distance-bottom:0;mso-position-horizontal-relative:margin;mso-position-vertical:absolute;mso-position-vertical-relative:page;mso-width-percent:0;mso-height-percent:0;mso-left-percent:800;mso-width-relative:page;mso-height-relative:page;v-text-anchor:top", OptionalString = "_x0000_s1026", AllowInCell = false, Stroked = false };
                rectangle1.SetAttribute(new OpenXmlAttribute("w14", "anchorId", "http://schemas.microsoft.com/office/word/2010/wordml", "5EC28866"));
                rectangle1.SetAttribute(new OpenXmlAttribute("o", "gfxdata", "urn:schemas-microsoft-com:office:office", "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\n"));

                V.TextBox textBox1 = new V.TextBox();

                TextBoxContent textBoxContent2 = new TextBoxContent();

                Paragraph paragraph3 = new Paragraph() { RsidParagraphAddition = "007B0E26", RsidRunAdditionDefault = "007B0E26", ParagraphId = "67F79515", TextId = "77777777" };

                ParagraphProperties paragraphProperties3 = new ParagraphProperties();
                Justification justification2 = new Justification() { Val = JustificationValues.Right };

                ParagraphMarkRunProperties paragraphMarkRunProperties2 = new ParagraphMarkRunProperties();
                Color color4 = new Color() { Val = "A6A6A6", ThemeColor = ThemeColorValues.Background1, ThemeShade = "A6" };
                FontSizeComplexScript fontSizeComplexScript4 = new FontSizeComplexScript() { Val = "144" };

                paragraphMarkRunProperties2.Append(color4);
                paragraphMarkRunProperties2.Append(fontSizeComplexScript4);

                paragraphProperties3.Append(justification2);
                paragraphProperties3.Append(paragraphMarkRunProperties2);

                Run run7 = new Run();
                FieldChar fieldChar4 = new FieldChar() { FieldCharType = FieldCharValues.Begin };

                run7.Append(fieldChar4);

                Run run8 = new Run();
                FieldCode fieldCode2 = new FieldCode() { Space = SpaceProcessingModeValues.Preserve };
                fieldCode2.Text = " PAGE    \\* MERGEFORMAT ";

                run8.Append(fieldCode2);

                Run run9 = new Run();
                FieldChar fieldChar5 = new FieldChar() { FieldCharType = FieldCharValues.Separate };

                run9.Append(fieldChar5);

                Run run10 = new Run();

                RunProperties runProperties4 = new RunProperties();
                NoProof noProof4 = new NoProof();
                Color color5 = new Color() { Val = "A6A6A6", ThemeColor = ThemeColorValues.Background1, ThemeShade = "A6" };
                FontSize fontSize3 = new FontSize() { Val = "144" };
                FontSizeComplexScript fontSizeComplexScript5 = new FontSizeComplexScript() { Val = "144" };

                runProperties4.Append(noProof4);
                runProperties4.Append(color5);
                runProperties4.Append(fontSize3);
                runProperties4.Append(fontSizeComplexScript5);
                Text text2 = new Text();
                text2.Text = "2";

                run10.Append(runProperties4);
                run10.Append(text2);

                Run run11 = new Run();

                RunProperties runProperties5 = new RunProperties();
                NoProof noProof5 = new NoProof();
                Color color6 = new Color() { Val = "A6A6A6", ThemeColor = ThemeColorValues.Background1, ThemeShade = "A6" };
                FontSize fontSize4 = new FontSize() { Val = "144" };
                FontSizeComplexScript fontSizeComplexScript6 = new FontSizeComplexScript() { Val = "144" };

                runProperties5.Append(noProof5);
                runProperties5.Append(color6);
                runProperties5.Append(fontSize4);
                runProperties5.Append(fontSizeComplexScript6);
                FieldChar fieldChar6 = new FieldChar() { FieldCharType = FieldCharValues.End };

                run11.Append(runProperties5);
                run11.Append(fieldChar6);

                paragraph3.Append(paragraphProperties3);
                paragraph3.Append(run7);
                paragraph3.Append(run8);
                paragraph3.Append(run9);
                paragraph3.Append(run10);
                paragraph3.Append(run11);

                textBoxContent2.Append(paragraph3);

                textBox1.Append(textBoxContent2);
                Wvml.TextWrap textWrap1 = new Wvml.TextWrap() { AnchorX = Wvml.HorizontalAnchorValues.Margin, AnchorY = Wvml.VerticalAnchorValues.Page };

                rectangle1.Append(textBox1);
                rectangle1.Append(textWrap1);

                picture1.Append(rectangle1);

                alternateContentFallback2.Append(picture1);

                alternateContent1.Append(alternateContentChoice1);
                alternateContent1.Append(alternateContentFallback2);

                run1.Append(runProperties1);
                run1.Append(alternateContent1);

                paragraph1.Append(paragraphProperties1);
                paragraph1.Append(run1);

                sdtContentBlock1.Append(paragraph1);

                sdtBlock1.Append(sdtProperties1);
                sdtBlock1.Append(sdtContentBlock1);
                return sdtBlock1;

            }
        }
        private static SdtBlock ThinLine1 {
            get {
                SdtBlock sdtBlock1 = new SdtBlock();

                SdtProperties sdtProperties1 = new SdtProperties();
                SdtId sdtId1 = new SdtId() { Val = -1309477069 };

                SdtContentDocPartObject sdtContentDocPartObject1 = new SdtContentDocPartObject();
                DocPartGallery docPartGallery1 = new DocPartGallery() { Val = "Page Numbers (Bottom of Page)" };
                DocPartUnique docPartUnique1 = new DocPartUnique();

                sdtContentDocPartObject1.Append(docPartGallery1);
                sdtContentDocPartObject1.Append(docPartUnique1);

                sdtProperties1.Append(sdtId1);
                sdtProperties1.Append(sdtContentDocPartObject1);

                SdtEndCharProperties sdtEndCharProperties1 = new SdtEndCharProperties();

                RunProperties runProperties1 = new RunProperties();
                NoProof noProof1 = new NoProof();

                runProperties1.Append(noProof1);

                sdtEndCharProperties1.Append(runProperties1);

                SdtContentBlock sdtContentBlock1 = new SdtContentBlock();

                Paragraph paragraph1 = new Paragraph() { RsidParagraphAddition = "00C023DB", RsidRunAdditionDefault = "00C023DB", ParagraphId = "4D8BBF20", TextId = "39377FA2" };

                ParagraphProperties paragraphProperties1 = new ParagraphProperties();
                ParagraphStyleId paragraphStyleId1 = new ParagraphStyleId() { Val = "Footer" };
                Justification justification1 = new Justification() { Val = JustificationValues.Center };

                paragraphProperties1.Append(paragraphStyleId1);
                paragraphProperties1.Append(justification1);

                Run run1 = new Run();

                RunProperties runProperties2 = new RunProperties();
                NoProof noProof2 = new NoProof();

                runProperties2.Append(noProof2);

                AlternateContent alternateContent1 = new AlternateContent();

                AlternateContentChoice alternateContentChoice1 = new AlternateContentChoice() { Requires = "wps" };

                Drawing drawing1 = new Drawing();

                Wp.Inline inline1 = new Wp.Inline() { DistanceFromTop = (UInt32Value)0U, DistanceFromBottom = (UInt32Value)0U, DistanceFromLeft = (UInt32Value)0U, DistanceFromRight = (UInt32Value)0U, AnchorId = "2AAFC29A", EditId = "3A08D08E" };
                Wp.Extent extent1 = new Wp.Extent() { Cx = 5467350L, Cy = 45085L };
                Wp.EffectExtent effectExtent1 = new Wp.EffectExtent() { LeftEdge = 9525L, TopEdge = 9525L, RightEdge = 0L, BottomEdge = 2540L };
                Wp.DocProperties docProperties1 = new Wp.DocProperties() { Id = (UInt32Value)1U, Name = "Flowchart: Decision 1", Description = "Light horizontal" };

                Wp.NonVisualGraphicFrameDrawingProperties nonVisualGraphicFrameDrawingProperties1 = new Wp.NonVisualGraphicFrameDrawingProperties();

                A.GraphicFrameLocks graphicFrameLocks1 = new A.GraphicFrameLocks();
                graphicFrameLocks1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

                nonVisualGraphicFrameDrawingProperties1.Append(graphicFrameLocks1);

                A.Graphic graphic1 = new A.Graphic();
                graphic1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

                A.GraphicData graphicData1 = new A.GraphicData() { Uri = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape" };

                Wps.WordprocessingShape wordprocessingShape1 = new Wps.WordprocessingShape();

                Wps.NonVisualDrawingShapeProperties nonVisualDrawingShapeProperties1 = new Wps.NonVisualDrawingShapeProperties();
                A.ShapeLocks shapeLocks1 = new A.ShapeLocks() { NoChangeArrowheads = true };

                nonVisualDrawingShapeProperties1.Append(shapeLocks1);

                Wps.ShapeProperties shapeProperties1 = new Wps.ShapeProperties() { BlackWhiteMode = A.BlackWhiteModeValues.Auto };

                A.Transform2D transform2D1 = new A.Transform2D() { VerticalFlip = true };
                A.Offset offset1 = new A.Offset() { X = 0L, Y = 0L };
                A.Extents extents1 = new A.Extents() { Cx = 5467350L, Cy = 45085L };

                transform2D1.Append(offset1);
                transform2D1.Append(extents1);

                A.PresetGeometry presetGeometry1 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.FlowChartDecision };
                A.AdjustValueList adjustValueList1 = new A.AdjustValueList();

                presetGeometry1.Append(adjustValueList1);

                A.PatternFill patternFill1 = new A.PatternFill() { Preset = A.PresetPatternValues.LightHorizontal };

                A.ForegroundColor foregroundColor1 = new A.ForegroundColor();
                A.RgbColorModelHex rgbColorModelHex1 = new A.RgbColorModelHex() { Val = "000000" };

                foregroundColor1.Append(rgbColorModelHex1);

                A.BackgroundColor backgroundColor1 = new A.BackgroundColor();
                A.RgbColorModelHex rgbColorModelHex2 = new A.RgbColorModelHex() { Val = "FFFFFF" };

                backgroundColor1.Append(rgbColorModelHex2);

                patternFill1.Append(foregroundColor1);
                patternFill1.Append(backgroundColor1);

                A.Outline outline1 = new A.Outline();
                A.NoFill noFill1 = new A.NoFill();

                outline1.Append(noFill1);

                A.ShapePropertiesExtensionList shapePropertiesExtensionList1 = new A.ShapePropertiesExtensionList();

                A.ShapePropertiesExtension shapePropertiesExtension1 = new A.ShapePropertiesExtension() { Uri = "{91240B29-F687-4F45-9708-019B960494DF}" };

                A14.HiddenLineProperties hiddenLineProperties1 = new A14.HiddenLineProperties() { Width = 9525 };
                hiddenLineProperties1.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");

                A.SolidFill solidFill1 = new A.SolidFill();
                A.RgbColorModelHex rgbColorModelHex3 = new A.RgbColorModelHex() { Val = "000000" };

                solidFill1.Append(rgbColorModelHex3);
                A.Miter miter1 = new A.Miter() { Limit = 800000 };
                A.HeadEnd headEnd1 = new A.HeadEnd();
                A.TailEnd tailEnd1 = new A.TailEnd();

                hiddenLineProperties1.Append(solidFill1);
                hiddenLineProperties1.Append(miter1);
                hiddenLineProperties1.Append(headEnd1);
                hiddenLineProperties1.Append(tailEnd1);

                shapePropertiesExtension1.Append(hiddenLineProperties1);

                shapePropertiesExtensionList1.Append(shapePropertiesExtension1);

                shapeProperties1.Append(transform2D1);
                shapeProperties1.Append(presetGeometry1);
                shapeProperties1.Append(patternFill1);
                shapeProperties1.Append(outline1);
                shapeProperties1.Append(shapePropertiesExtensionList1);

                Wps.TextBodyProperties textBodyProperties1 = new Wps.TextBodyProperties() { Rotation = 0, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, LeftInset = 91440, TopInset = 45720, RightInset = 91440, BottomInset = 45720, Anchor = A.TextAnchoringTypeValues.Top, AnchorCenter = false, UpRight = true };
                A.NoAutoFit noAutoFit1 = new A.NoAutoFit();

                textBodyProperties1.Append(noAutoFit1);

                wordprocessingShape1.Append(nonVisualDrawingShapeProperties1);
                wordprocessingShape1.Append(shapeProperties1);
                wordprocessingShape1.Append(textBodyProperties1);

                graphicData1.Append(wordprocessingShape1);

                graphic1.Append(graphicData1);

                inline1.Append(extent1);
                inline1.Append(effectExtent1);
                inline1.Append(docProperties1);
                inline1.Append(nonVisualGraphicFrameDrawingProperties1);
                inline1.Append(graphic1);

                drawing1.Append(inline1);

                alternateContentChoice1.Append(drawing1);

                AlternateContentFallback alternateContentFallback1 = new AlternateContentFallback();

                Picture picture1 = new Picture();

                V.Shapetype shapetype1 = new V.Shapetype() { Id = "_x0000_t110", CoordinateSize = "21600,21600", OptionalNumber = 110, EdgePath = "m10800,l,10800,10800,21600,21600,10800xe" };
                shapetype1.SetAttribute(new OpenXmlAttribute("w14", "anchorId", "http://schemas.microsoft.com/office/word/2010/wordml", "2F56B1C1"));
                V.Stroke stroke1 = new V.Stroke() { JoinStyle = V.StrokeJoinStyleValues.Miter };
                V.Path path1 = new V.Path() { TextboxRectangle = "5400,5400,16200,16200", AllowGradientShape = true, ConnectionPointType = Ovml.ConnectValues.Rectangle };

                shapetype1.Append(stroke1);
                shapetype1.Append(path1);

                V.Shape shape1 = new V.Shape() { Id = "Flowchart: Decision 1", Style = "width:430.5pt;height:3.55pt;flip:y;visibility:visible;mso-wrap-style:square;mso-left-percent:-10001;mso-top-percent:-10001;mso-position-horizontal:absolute;mso-position-horizontal-relative:char;mso-position-vertical:absolute;mso-position-vertical-relative:line;mso-left-percent:-10001;mso-top-percent:-10001;v-text-anchor:top", Alternate = "Light horizontal", OptionalString = "_x0000_s1026", FillColor = "black", Stroked = false, Type = "#_x0000_t110", EncodedPackage = "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\n" };
                V.Fill fill1 = new V.Fill() { Type = V.FillTypeValues.Pattern, Title = "", RelationshipId = "rId1" };
                Wvml.AnchorLock anchorLock1 = new Wvml.AnchorLock();

                shape1.Append(fill1);
                shape1.Append(anchorLock1);

                picture1.Append(shapetype1);
                picture1.Append(shape1);

                alternateContentFallback1.Append(picture1);

                alternateContent1.Append(alternateContentChoice1);
                alternateContent1.Append(alternateContentFallback1);

                run1.Append(runProperties2);
                run1.Append(alternateContent1);

                paragraph1.Append(paragraphProperties1);
                paragraph1.Append(run1);

                Paragraph paragraph2 = new Paragraph() { RsidParagraphAddition = "00C023DB", RsidRunAdditionDefault = "00C023DB", ParagraphId = "053018A7", TextId = "47B9AD3B" };

                ParagraphProperties paragraphProperties2 = new ParagraphProperties();
                ParagraphStyleId paragraphStyleId2 = new ParagraphStyleId() { Val = "Footer" };
                Justification justification2 = new Justification() { Val = JustificationValues.Center };

                paragraphProperties2.Append(paragraphStyleId2);
                paragraphProperties2.Append(justification2);

                Run run2 = new Run();
                FieldChar fieldChar1 = new FieldChar() { FieldCharType = FieldCharValues.Begin };

                run2.Append(fieldChar1);

                Run run3 = new Run();
                FieldCode fieldCode1 = new FieldCode() { Space = SpaceProcessingModeValues.Preserve };
                fieldCode1.Text = " PAGE    \\* MERGEFORMAT ";

                run3.Append(fieldCode1);

                Run run4 = new Run();
                FieldChar fieldChar2 = new FieldChar() { FieldCharType = FieldCharValues.Separate };

                run4.Append(fieldChar2);

                Run run5 = new Run();

                RunProperties runProperties3 = new RunProperties();
                NoProof noProof3 = new NoProof();

                runProperties3.Append(noProof3);
                Text text1 = new Text();
                text1.Text = "2";

                run5.Append(runProperties3);
                run5.Append(text1);

                Run run6 = new Run();

                RunProperties runProperties4 = new RunProperties();
                NoProof noProof4 = new NoProof();

                runProperties4.Append(noProof4);
                FieldChar fieldChar3 = new FieldChar() { FieldCharType = FieldCharValues.End };

                run6.Append(runProperties4);
                run6.Append(fieldChar3);

                paragraph2.Append(paragraphProperties2);
                paragraph2.Append(run2);
                paragraph2.Append(run3);
                paragraph2.Append(run4);
                paragraph2.Append(run5);
                paragraph2.Append(run6);

                sdtContentBlock1.Append(paragraph1);
                sdtContentBlock1.Append(paragraph2);

                sdtBlock1.Append(sdtProperties1);
                sdtBlock1.Append(sdtEndCharProperties1);
                sdtBlock1.Append(sdtContentBlock1);
                return sdtBlock1;

            }
        }

        private static SdtBlock VerticalOutline1 {
            get {
                SdtBlock sdtBlock1 = new SdtBlock();

                SdtProperties sdtProperties1 = new SdtProperties();
                SdtId sdtId1 = new SdtId() { Val = -753432457 };

                SdtContentDocPartObject sdtContentDocPartObject1 = new SdtContentDocPartObject();
                DocPartGallery docPartGallery1 = new DocPartGallery() { Val = "Page Numbers (Top of Page)" };
                DocPartUnique docPartUnique1 = new DocPartUnique();

                sdtContentDocPartObject1.Append(docPartGallery1);
                sdtContentDocPartObject1.Append(docPartUnique1);

                sdtProperties1.Append(sdtId1);
                sdtProperties1.Append(sdtContentDocPartObject1);

                SdtContentBlock sdtContentBlock1 = new SdtContentBlock();

                Paragraph paragraph1 = new Paragraph() { RsidParagraphMarkRevision = "00E9532F", RsidParagraphAddition = "00E11B1E", RsidParagraphProperties = "00E9532F", RsidRunAdditionDefault = "00E9532F", ParagraphId = "16AD3494", TextId = "12EB5799" };

                ParagraphProperties paragraphProperties1 = new ParagraphProperties();
                ParagraphStyleId paragraphStyleId1 = new ParagraphStyleId() { Val = "Header" };

                paragraphProperties1.Append(paragraphStyleId1);

                Run run1 = new Run();

                RunProperties runProperties1 = new RunProperties();
                NoProof noProof1 = new NoProof();

                runProperties1.Append(noProof1);

                AlternateContent alternateContent1 = new AlternateContent();

                AlternateContentChoice alternateContentChoice1 = new AlternateContentChoice() { Requires = "wpg" };

                Drawing drawing1 = new Drawing();

                Wp.Anchor anchor1 = new Wp.Anchor() { DistanceFromTop = (UInt32Value)0U, DistanceFromBottom = (UInt32Value)0U, DistanceFromLeft = (UInt32Value)114300U, DistanceFromRight = (UInt32Value)114300U, SimplePos = false, RelativeHeight = (UInt32Value)251659264U, BehindDoc = false, Locked = false, LayoutInCell = false, AllowOverlap = true, EditId = "5402C24E", AnchorId = "5BD30002" };
                Wp.SimplePosition simplePosition1 = new Wp.SimplePosition() { X = 0L, Y = 0L };

                Wp.HorizontalPosition horizontalPosition1 = new Wp.HorizontalPosition() { RelativeFrom = Wp.HorizontalRelativePositionValues.LeftMargin };
                Wp.HorizontalAlignment horizontalAlignment1 = new Wp.HorizontalAlignment();
                horizontalAlignment1.Text = "right";

                horizontalPosition1.Append(horizontalAlignment1);

                Wp.VerticalPosition verticalPosition1 = new Wp.VerticalPosition() { RelativeFrom = Wp.VerticalRelativePositionValues.Margin };
                Wp.VerticalAlignment verticalAlignment1 = new Wp.VerticalAlignment();
                verticalAlignment1.Text = "top";

                verticalPosition1.Append(verticalAlignment1);
                Wp.Extent extent1 = new Wp.Extent() { Cx = 904875L, Cy = 1902460L };
                Wp.EffectExtent effectExtent1 = new Wp.EffectExtent() { LeftEdge = 11430L, TopEdge = 9525L, RightEdge = 0L, BottomEdge = 2540L };
                Wp.WrapNone wrapNone1 = new Wp.WrapNone();
                Wp.DocProperties docProperties1 = new Wp.DocProperties() { Id = (UInt32Value)10U, Name = "Group 10" };

                Wp.NonVisualGraphicFrameDrawingProperties nonVisualGraphicFrameDrawingProperties1 = new Wp.NonVisualGraphicFrameDrawingProperties();

                A.GraphicFrameLocks graphicFrameLocks1 = new A.GraphicFrameLocks();
                graphicFrameLocks1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

                nonVisualGraphicFrameDrawingProperties1.Append(graphicFrameLocks1);

                A.Graphic graphic1 = new A.Graphic();
                graphic1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

                A.GraphicData graphicData1 = new A.GraphicData() { Uri = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" };

                Wpg.WordprocessingGroup wordprocessingGroup1 = new Wpg.WordprocessingGroup();

                Wpg.NonVisualGroupDrawingShapeProperties nonVisualGroupDrawingShapeProperties1 = new Wpg.NonVisualGroupDrawingShapeProperties();
                A.GroupShapeLocks groupShapeLocks1 = new A.GroupShapeLocks();

                nonVisualGroupDrawingShapeProperties1.Append(groupShapeLocks1);

                Wpg.GroupShapeProperties groupShapeProperties1 = new Wpg.GroupShapeProperties() { BlackWhiteMode = A.BlackWhiteModeValues.Auto };

                A.TransformGroup transformGroup1 = new A.TransformGroup() { VerticalFlip = true };
                A.Offset offset1 = new A.Offset() { X = 0L, Y = 0L };
                A.Extents extents1 = new A.Extents() { Cx = 904875L, Cy = 1902460L };
                A.ChildOffset childOffset1 = new A.ChildOffset() { X = 13L, Y = 11415L };
                A.ChildExtents childExtents1 = new A.ChildExtents() { Cx = 1425L, Cy = 2996L };

                transformGroup1.Append(offset1);
                transformGroup1.Append(extents1);
                transformGroup1.Append(childOffset1);
                transformGroup1.Append(childExtents1);

                groupShapeProperties1.Append(transformGroup1);

                Wpg.GroupShape groupShape1 = new Wpg.GroupShape();
                Wpg.NonVisualDrawingProperties nonVisualDrawingProperties1 = new Wpg.NonVisualDrawingProperties() { Id = (UInt32Value)11U, Name = "Group 7" };

                Wpg.NonVisualGroupDrawingShapeProperties nonVisualGroupDrawingShapeProperties2 = new Wpg.NonVisualGroupDrawingShapeProperties();
                A.GroupShapeLocks groupShapeLocks2 = new A.GroupShapeLocks();

                nonVisualGroupDrawingShapeProperties2.Append(groupShapeLocks2);

                Wpg.GroupShapeProperties groupShapeProperties2 = new Wpg.GroupShapeProperties() { BlackWhiteMode = A.BlackWhiteModeValues.Auto };

                A.TransformGroup transformGroup2 = new A.TransformGroup() { VerticalFlip = true };
                A.Offset offset2 = new A.Offset() { X = 13L, Y = 14340L };
                A.Extents extents2 = new A.Extents() { Cx = 1410L, Cy = 71L };
                A.ChildOffset childOffset2 = new A.ChildOffset() { X = -83L, Y = 540L };
                A.ChildExtents childExtents2 = new A.ChildExtents() { Cx = 1218L, Cy = 71L };

                transformGroup2.Append(offset2);
                transformGroup2.Append(extents2);
                transformGroup2.Append(childOffset2);
                transformGroup2.Append(childExtents2);

                groupShapeProperties2.Append(transformGroup2);

                Wps.WordprocessingShape wordprocessingShape1 = new Wps.WordprocessingShape();
                Wps.NonVisualDrawingProperties nonVisualDrawingProperties2 = new Wps.NonVisualDrawingProperties() { Id = (UInt32Value)12U, Name = "Rectangle 8" };

                Wps.NonVisualDrawingShapeProperties nonVisualDrawingShapeProperties1 = new Wps.NonVisualDrawingShapeProperties();
                A.ShapeLocks shapeLocks1 = new A.ShapeLocks() { NoChangeArrowheads = true };

                nonVisualDrawingShapeProperties1.Append(shapeLocks1);

                Wps.ShapeProperties shapeProperties1 = new Wps.ShapeProperties() { BlackWhiteMode = A.BlackWhiteModeValues.Auto };

                A.Transform2D transform2D1 = new A.Transform2D();
                A.Offset offset3 = new A.Offset() { X = 678L, Y = 540L };
                A.Extents extents3 = new A.Extents() { Cx = 457L, Cy = 71L };

                transform2D1.Append(offset3);
                transform2D1.Append(extents3);

                A.PresetGeometry presetGeometry1 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
                A.AdjustValueList adjustValueList1 = new A.AdjustValueList();

                presetGeometry1.Append(adjustValueList1);

                A.SolidFill solidFill1 = new A.SolidFill();
                A.RgbColorModelHex rgbColorModelHex1 = new A.RgbColorModelHex() { Val = "5F497A" };

                solidFill1.Append(rgbColorModelHex1);

                A.Outline outline1 = new A.Outline() { Width = 9525 };

                A.SolidFill solidFill2 = new A.SolidFill();
                A.RgbColorModelHex rgbColorModelHex2 = new A.RgbColorModelHex() { Val = "5F497A" };

                solidFill2.Append(rgbColorModelHex2);
                A.Miter miter1 = new A.Miter() { Limit = 800000 };
                A.HeadEnd headEnd1 = new A.HeadEnd();
                A.TailEnd tailEnd1 = new A.TailEnd();

                outline1.Append(solidFill2);
                outline1.Append(miter1);
                outline1.Append(headEnd1);
                outline1.Append(tailEnd1);

                shapeProperties1.Append(transform2D1);
                shapeProperties1.Append(presetGeometry1);
                shapeProperties1.Append(solidFill1);
                shapeProperties1.Append(outline1);

                Wps.TextBodyProperties textBodyProperties1 = new Wps.TextBodyProperties() { Rotation = 0, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, LeftInset = 91440, TopInset = 45720, RightInset = 91440, BottomInset = 45720, Anchor = A.TextAnchoringTypeValues.Top, AnchorCenter = false, UpRight = true };
                A.NoAutoFit noAutoFit1 = new A.NoAutoFit();

                textBodyProperties1.Append(noAutoFit1);

                wordprocessingShape1.Append(nonVisualDrawingProperties2);
                wordprocessingShape1.Append(nonVisualDrawingShapeProperties1);
                wordprocessingShape1.Append(shapeProperties1);
                wordprocessingShape1.Append(textBodyProperties1);

                Wps.WordprocessingShape wordprocessingShape2 = new Wps.WordprocessingShape();
                Wps.NonVisualDrawingProperties nonVisualDrawingProperties3 = new Wps.NonVisualDrawingProperties() { Id = (UInt32Value)13U, Name = "AutoShape 9" };

                Wps.NonVisualConnectorProperties nonVisualConnectorProperties1 = new Wps.NonVisualConnectorProperties();
                A.ConnectionShapeLocks connectionShapeLocks1 = new A.ConnectionShapeLocks() { NoChangeShapeType = true };

                nonVisualConnectorProperties1.Append(connectionShapeLocks1);

                Wps.ShapeProperties shapeProperties2 = new Wps.ShapeProperties() { BlackWhiteMode = A.BlackWhiteModeValues.Auto };

                A.Transform2D transform2D2 = new A.Transform2D() { HorizontalFlip = true };
                A.Offset offset4 = new A.Offset() { X = -83L, Y = 540L };
                A.Extents extents4 = new A.Extents() { Cx = 761L, Cy = 0L };

                transform2D2.Append(offset4);
                transform2D2.Append(extents4);

                A.PresetGeometry presetGeometry2 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.StraightConnector1 };
                A.AdjustValueList adjustValueList2 = new A.AdjustValueList();

                presetGeometry2.Append(adjustValueList2);
                A.NoFill noFill1 = new A.NoFill();

                A.Outline outline2 = new A.Outline() { Width = 9525 };

                A.SolidFill solidFill3 = new A.SolidFill();
                A.RgbColorModelHex rgbColorModelHex3 = new A.RgbColorModelHex() { Val = "5F497A" };

                solidFill3.Append(rgbColorModelHex3);
                A.Round round1 = new A.Round();
                A.HeadEnd headEnd2 = new A.HeadEnd();
                A.TailEnd tailEnd2 = new A.TailEnd();

                outline2.Append(solidFill3);
                outline2.Append(round1);
                outline2.Append(headEnd2);
                outline2.Append(tailEnd2);

                A.ShapePropertiesExtensionList shapePropertiesExtensionList1 = new A.ShapePropertiesExtensionList();

                A.ShapePropertiesExtension shapePropertiesExtension1 = new A.ShapePropertiesExtension() { Uri = "{909E8E84-426E-40DD-AFC4-6F175D3DCCD1}" };

                A14.HiddenFillProperties hiddenFillProperties1 = new A14.HiddenFillProperties();
                hiddenFillProperties1.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");
                A.NoFill noFill2 = new A.NoFill();

                hiddenFillProperties1.Append(noFill2);

                shapePropertiesExtension1.Append(hiddenFillProperties1);

                shapePropertiesExtensionList1.Append(shapePropertiesExtension1);

                shapeProperties2.Append(transform2D2);
                shapeProperties2.Append(presetGeometry2);
                shapeProperties2.Append(noFill1);
                shapeProperties2.Append(outline2);
                shapeProperties2.Append(shapePropertiesExtensionList1);
                Wps.TextBodyProperties textBodyProperties2 = new Wps.TextBodyProperties();

                wordprocessingShape2.Append(nonVisualDrawingProperties3);
                wordprocessingShape2.Append(nonVisualConnectorProperties1);
                wordprocessingShape2.Append(shapeProperties2);
                wordprocessingShape2.Append(textBodyProperties2);

                groupShape1.Append(nonVisualDrawingProperties1);
                groupShape1.Append(nonVisualGroupDrawingShapeProperties2);
                groupShape1.Append(groupShapeProperties2);
                groupShape1.Append(wordprocessingShape1);
                groupShape1.Append(wordprocessingShape2);

                Wps.WordprocessingShape wordprocessingShape3 = new Wps.WordprocessingShape();
                Wps.NonVisualDrawingProperties nonVisualDrawingProperties4 = new Wps.NonVisualDrawingProperties() { Id = (UInt32Value)14U, Name = "Rectangle 10" };

                Wps.NonVisualDrawingShapeProperties nonVisualDrawingShapeProperties2 = new Wps.NonVisualDrawingShapeProperties();
                A.ShapeLocks shapeLocks2 = new A.ShapeLocks() { NoChangeArrowheads = true };

                nonVisualDrawingShapeProperties2.Append(shapeLocks2);

                Wps.ShapeProperties shapeProperties3 = new Wps.ShapeProperties() { BlackWhiteMode = A.BlackWhiteModeValues.Auto };

                A.Transform2D transform2D3 = new A.Transform2D();
                A.Offset offset5 = new A.Offset() { X = 405L, Y = 11415L };
                A.Extents extents5 = new A.Extents() { Cx = 1033L, Cy = 2805L };

                transform2D3.Append(offset5);
                transform2D3.Append(extents5);

                A.PresetGeometry presetGeometry3 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
                A.AdjustValueList adjustValueList3 = new A.AdjustValueList();

                presetGeometry3.Append(adjustValueList3);

                A.SolidFill solidFill4 = new A.SolidFill();
                A.RgbColorModelHex rgbColorModelHex4 = new A.RgbColorModelHex() { Val = "FFFFFF" };

                solidFill4.Append(rgbColorModelHex4);

                A.Outline outline3 = new A.Outline();
                A.NoFill noFill3 = new A.NoFill();

                outline3.Append(noFill3);

                A.ShapePropertiesExtensionList shapePropertiesExtensionList2 = new A.ShapePropertiesExtensionList();

                A.ShapePropertiesExtension shapePropertiesExtension2 = new A.ShapePropertiesExtension() { Uri = "{91240B29-F687-4F45-9708-019B960494DF}" };

                A14.HiddenLineProperties hiddenLineProperties1 = new A14.HiddenLineProperties() { Width = 9525 };
                hiddenLineProperties1.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");

                A.SolidFill solidFill5 = new A.SolidFill();
                A.RgbColorModelHex rgbColorModelHex5 = new A.RgbColorModelHex() { Val = "000000" };

                solidFill5.Append(rgbColorModelHex5);
                A.Miter miter2 = new A.Miter() { Limit = 800000 };
                A.HeadEnd headEnd3 = new A.HeadEnd();
                A.TailEnd tailEnd3 = new A.TailEnd();

                hiddenLineProperties1.Append(solidFill5);
                hiddenLineProperties1.Append(miter2);
                hiddenLineProperties1.Append(headEnd3);
                hiddenLineProperties1.Append(tailEnd3);

                shapePropertiesExtension2.Append(hiddenLineProperties1);

                shapePropertiesExtensionList2.Append(shapePropertiesExtension2);

                shapeProperties3.Append(transform2D3);
                shapeProperties3.Append(presetGeometry3);
                shapeProperties3.Append(solidFill4);
                shapeProperties3.Append(outline3);
                shapeProperties3.Append(shapePropertiesExtensionList2);

                Wps.TextBoxInfo2 textBoxInfo21 = new Wps.TextBoxInfo2();

                TextBoxContent textBoxContent1 = new TextBoxContent();

                Paragraph paragraph2 = new Paragraph() { RsidParagraphAddition = "00E9532F", RsidRunAdditionDefault = "00E9532F", ParagraphId = "48B42751", TextId = "77777777" };

                ParagraphProperties paragraphProperties2 = new ParagraphProperties();
                ParagraphStyleId paragraphStyleId2 = new ParagraphStyleId() { Val = "NoSpacing" };

                paragraphProperties2.Append(paragraphStyleId2);

                Run run2 = new Run();
                FieldChar fieldChar1 = new FieldChar() { FieldCharType = FieldCharValues.Begin };

                run2.Append(fieldChar1);

                Run run3 = new Run();
                FieldCode fieldCode1 = new FieldCode() { Space = SpaceProcessingModeValues.Preserve };
                fieldCode1.Text = " PAGE    \\* MERGEFORMAT ";

                run3.Append(fieldCode1);

                Run run4 = new Run();
                FieldChar fieldChar2 = new FieldChar() { FieldCharType = FieldCharValues.Separate };

                run4.Append(fieldChar2);

                Run run5 = new Run();

                RunProperties runProperties2 = new RunProperties();
                Bold bold1 = new Bold();
                BoldComplexScript boldComplexScript1 = new BoldComplexScript();
                NoProof noProof2 = new NoProof();
                Color color1 = new Color() { Val = "BF8F00", ThemeColor = ThemeColorValues.Accent4, ThemeShade = "BF" };
                FontSize fontSize1 = new FontSize() { Val = "52" };
                FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "52" };

                runProperties2.Append(bold1);
                runProperties2.Append(boldComplexScript1);
                runProperties2.Append(noProof2);
                runProperties2.Append(color1);
                runProperties2.Append(fontSize1);
                runProperties2.Append(fontSizeComplexScript1);
                Text text1 = new Text();
                text1.Text = "2";

                run5.Append(runProperties2);
                run5.Append(text1);

                Run run6 = new Run();

                RunProperties runProperties3 = new RunProperties();
                Bold bold2 = new Bold();
                BoldComplexScript boldComplexScript2 = new BoldComplexScript();
                NoProof noProof3 = new NoProof();
                Color color2 = new Color() { Val = "BF8F00", ThemeColor = ThemeColorValues.Accent4, ThemeShade = "BF" };
                FontSize fontSize2 = new FontSize() { Val = "52" };
                FontSizeComplexScript fontSizeComplexScript2 = new FontSizeComplexScript() { Val = "52" };

                runProperties3.Append(bold2);
                runProperties3.Append(boldComplexScript2);
                runProperties3.Append(noProof3);
                runProperties3.Append(color2);
                runProperties3.Append(fontSize2);
                runProperties3.Append(fontSizeComplexScript2);
                FieldChar fieldChar3 = new FieldChar() { FieldCharType = FieldCharValues.End };

                run6.Append(runProperties3);
                run6.Append(fieldChar3);

                paragraph2.Append(paragraphProperties2);
                paragraph2.Append(run2);
                paragraph2.Append(run3);
                paragraph2.Append(run4);
                paragraph2.Append(run5);
                paragraph2.Append(run6);

                textBoxContent1.Append(paragraph2);

                textBoxInfo21.Append(textBoxContent1);

                Wps.TextBodyProperties textBodyProperties3 = new Wps.TextBodyProperties() { Rotation = 0, Vertical = A.TextVerticalValues.Vertical, Wrap = A.TextWrappingValues.Square, LeftInset = 0, TopInset = 0, RightInset = 0, BottomInset = 0, Anchor = A.TextAnchoringTypeValues.Top, AnchorCenter = false, UpRight = true };
                A.NoAutoFit noAutoFit2 = new A.NoAutoFit();

                textBodyProperties3.Append(noAutoFit2);

                wordprocessingShape3.Append(nonVisualDrawingProperties4);
                wordprocessingShape3.Append(nonVisualDrawingShapeProperties2);
                wordprocessingShape3.Append(shapeProperties3);
                wordprocessingShape3.Append(textBoxInfo21);
                wordprocessingShape3.Append(textBodyProperties3);

                wordprocessingGroup1.Append(nonVisualGroupDrawingShapeProperties1);
                wordprocessingGroup1.Append(groupShapeProperties1);
                wordprocessingGroup1.Append(groupShape1);
                wordprocessingGroup1.Append(wordprocessingShape3);

                graphicData1.Append(wordprocessingGroup1);

                graphic1.Append(graphicData1);

                Wp14.RelativeWidth relativeWidth1 = new Wp14.RelativeWidth() { ObjectId = Wp14.SizeRelativeHorizontallyValues.LeftMargin };
                Wp14.PercentageWidth percentageWidth1 = new Wp14.PercentageWidth();
                percentageWidth1.Text = "100000";

                relativeWidth1.Append(percentageWidth1);

                Wp14.RelativeHeight relativeHeight1 = new Wp14.RelativeHeight() { RelativeFrom = Wp14.SizeRelativeVerticallyValues.Page };
                Wp14.PercentageHeight percentageHeight1 = new Wp14.PercentageHeight();
                percentageHeight1.Text = "0";

                relativeHeight1.Append(percentageHeight1);

                anchor1.Append(simplePosition1);
                anchor1.Append(horizontalPosition1);
                anchor1.Append(verticalPosition1);
                anchor1.Append(extent1);
                anchor1.Append(effectExtent1);
                anchor1.Append(wrapNone1);
                anchor1.Append(docProperties1);
                anchor1.Append(nonVisualGraphicFrameDrawingProperties1);
                anchor1.Append(graphic1);
                anchor1.Append(relativeWidth1);
                anchor1.Append(relativeHeight1);

                drawing1.Append(anchor1);

                alternateContentChoice1.Append(drawing1);

                AlternateContentFallback alternateContentFallback1 = new AlternateContentFallback();

                Picture picture1 = new Picture();

                V.Group group1 = new V.Group() { Id = "Group 10", Style = "position:absolute;margin-left:20.05pt;margin-top:0;width:71.25pt;height:149.8pt;flip:y;z-index:251659264;mso-width-percent:1000;mso-position-horizontal:right;mso-position-horizontal-relative:left-margin-area;mso-position-vertical:top;mso-position-vertical-relative:margin;mso-width-percent:1000;mso-width-relative:left-margin-area", CoordinateSize = "1425,2996", CoordinateOrigin = "13,11415", OptionalString = "_x0000_s1026", AllowInCell = false };
                group1.SetAttribute(new OpenXmlAttribute("w14", "anchorId", "http://schemas.microsoft.com/office/word/2010/wordml", "5BD30002"));
                group1.SetAttribute(new OpenXmlAttribute("o", "gfxdata", "urn:schemas-microsoft-com:office:office", "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\n"));

                V.Group group2 = new V.Group() { Id = "Group 7", Style = "position:absolute;left:13;top:14340;width:1410;height:71;flip:y", CoordinateSize = "1218,71", CoordinateOrigin = "-83,540", OptionalString = "_x0000_s1027" };
                group2.SetAttribute(new OpenXmlAttribute("o", "gfxdata", "urn:schemas-microsoft-com:office:office", "UEsDBBQABgAIAAAAIQDb4fbL7gAAAIUBAAATAAAAW0NvbnRlbnRfVHlwZXNdLnhtbHyQz07DMAyH\n70i8Q+QralM4IITa7kDhCAiNB7ASt43WOlEcyvb2pNu4IODoPz9/n1xv9vOkForiPDdwXVagiI23\njocG3rdPxR0oScgWJ8/UwIEENu3lRb09BBKV0ywNjCmFe63FjDSjlD4Q50nv44wpl3HQAc0OB9I3\nVXWrjedEnIq03oC27qjHjympx31un0wiTQLq4bS4shrAECZnMGVTvbD9QSnOhDInjzsyuiBXWQP0\nr4R18jfgnHvJr4nOknrFmJ5xzhraRtHWf3Kkpfz/yGo5S+H73hkquyhdjr3R8m2lj09svwAAAP//\nAwBQSwMEFAAGAAgAAAAhAFr0LFu/AAAAFQEAAAsAAABfcmVscy8ucmVsc2zPwWrDMAwG4Ptg72B0\nX5TuUMaI01uh19I+gLGVxCy2jGSy9e1nemrHjpL4P0nD4SetZiPRyNnCruvBUPYcYp4tXC/Htw8w\nWl0ObuVMFm6kcBhfX4Yzra62kC6xqGlKVgtLreUTUf1CyWnHhXKbTCzJ1VbKjMX5LzcTvvf9HuXR\ngPHJNKdgQU5hB+ZyK23zHztFL6w81c5zQp6m6P9TMfB3PtPWFCczVQtB9N4U2rp2HOA44NMz4y8A\nAAD//wMAUEsDBBQABgAIAAAAIQAgpgOrvwAAANsAAAAPAAAAZHJzL2Rvd25yZXYueG1sRE9Li8Iw\nEL4L/ocwgjdNlbJINYoIioiXrQ88Ds3YBptJaaLWf79ZWNjbfHzPWaw6W4sXtd44VjAZJyCIC6cN\nlwrOp+1oBsIHZI21Y1LwIQ+rZb+3wEy7N3/TKw+liCHsM1RQhdBkUvqiIot+7BriyN1dazFE2JZS\nt/iO4baW0yT5khYNx4YKG9pUVDzyp1VwWZuU0uvtcEwKor2Wt11uUqWGg249BxGoC//iP/dex/kT\n+P0lHiCXPwAAAP//AwBQSwECLQAUAAYACAAAACEA2+H2y+4AAACFAQAAEwAAAAAAAAAAAAAAAAAA\nAAAAW0NvbnRlbnRfVHlwZXNdLnhtbFBLAQItABQABgAIAAAAIQBa9CxbvwAAABUBAAALAAAAAAAA\nAAAAAAAAAB8BAABfcmVscy8ucmVsc1BLAQItABQABgAIAAAAIQAgpgOrvwAAANsAAAAPAAAAAAAA\nAAAAAAAAAAcCAABkcnMvZG93bnJldi54bWxQSwUGAAAAAAMAAwC3AAAA8wIAAAAA\n"));

                V.Rectangle rectangle1 = new V.Rectangle() { Id = "Rectangle 8", Style = "position:absolute;left:678;top:540;width:457;height:71;visibility:visible;mso-wrap-style:square;v-text-anchor:top", OptionalString = "_x0000_s1028", FillColor = "#5f497a", StrokeColor = "#5f497a" };
                rectangle1.SetAttribute(new OpenXmlAttribute("o", "gfxdata", "urn:schemas-microsoft-com:office:office", "UEsDBBQABgAIAAAAIQDb4fbL7gAAAIUBAAATAAAAW0NvbnRlbnRfVHlwZXNdLnhtbHyQz07DMAyH\n70i8Q+QralM4IITa7kDhCAiNB7ASt43WOlEcyvb2pNu4IODoPz9/n1xv9vOkForiPDdwXVagiI23\njocG3rdPxR0oScgWJ8/UwIEENu3lRb09BBKV0ywNjCmFe63FjDSjlD4Q50nv44wpl3HQAc0OB9I3\nVXWrjedEnIq03oC27qjHjympx31un0wiTQLq4bS4shrAECZnMGVTvbD9QSnOhDInjzsyuiBXWQP0\nr4R18jfgnHvJr4nOknrFmJ5xzhraRtHWf3Kkpfz/yGo5S+H73hkquyhdjr3R8m2lj09svwAAAP//\nAwBQSwMEFAAGAAgAAAAhAFr0LFu/AAAAFQEAAAsAAABfcmVscy8ucmVsc2zPwWrDMAwG4Ptg72B0\nX5TuUMaI01uh19I+gLGVxCy2jGSy9e1nemrHjpL4P0nD4SetZiPRyNnCruvBUPYcYp4tXC/Htw8w\nWl0ObuVMFm6kcBhfX4Yzra62kC6xqGlKVgtLreUTUf1CyWnHhXKbTCzJ1VbKjMX5LzcTvvf9HuXR\ngPHJNKdgQU5hB+ZyK23zHztFL6w81c5zQp6m6P9TMfB3PtPWFCczVQtB9N4U2rp2HOA44NMz4y8A\nAAD//wMAUEsDBBQABgAIAAAAIQA5FpHkwgAAANsAAAAPAAAAZHJzL2Rvd25yZXYueG1sRE9La8JA\nEL4X+h+WEXqrGz2UNrqKlCqFiuKLXofsNInNzsTsNsZ/7woFb/PxPWc87VylWmp8KWxg0E9AEWdi\nS84N7Hfz51dQPiBbrITJwIU8TCePD2NMrZx5Q+025CqGsE/RQBFCnWrts4Ic+r7UxJH7kcZhiLDJ\ntW3wHMNdpYdJ8qIdlhwbCqzpvaDsd/vnDBzlW9rDStbL5YmSj+NssX77Whjz1OtmI1CBunAX/7s/\nbZw/hNsv8QA9uQIAAP//AwBQSwECLQAUAAYACAAAACEA2+H2y+4AAACFAQAAEwAAAAAAAAAAAAAA\nAAAAAAAAW0NvbnRlbnRfVHlwZXNdLnhtbFBLAQItABQABgAIAAAAIQBa9CxbvwAAABUBAAALAAAA\nAAAAAAAAAAAAAB8BAABfcmVscy8ucmVsc1BLAQItABQABgAIAAAAIQA5FpHkwgAAANsAAAAPAAAA\nAAAAAAAAAAAAAAcCAABkcnMvZG93bnJldi54bWxQSwUGAAAAAAMAAwC3AAAA9gIAAAAA\n"));

                V.Shapetype shapetype1 = new V.Shapetype() { Id = "_x0000_t32", CoordinateSize = "21600,21600", Oned = true, Filled = false, OptionalNumber = 32, EdgePath = "m,l21600,21600e" };
                V.Path path1 = new V.Path() { AllowFill = false, ShowArrowhead = true, ConnectionPointType = Ovml.ConnectValues.None };
                Ovml.Lock lock1 = new Ovml.Lock() { Extension = V.ExtensionHandlingBehaviorValues.Edit, ShapeType = true };

                shapetype1.Append(path1);
                shapetype1.Append(lock1);
                V.Shape shape1 = new V.Shape() { Id = "AutoShape 9", Style = "position:absolute;left:-83;top:540;width:761;height:0;flip:x;visibility:visible;mso-wrap-style:square", OptionalString = "_x0000_s1029", StrokeColor = "#5f497a", ConnectorType = Ovml.ConnectorValues.Straight, Type = "#_x0000_t32", EncodedPackage = "UEsDBBQABgAIAAAAIQDb4fbL7gAAAIUBAAATAAAAW0NvbnRlbnRfVHlwZXNdLnhtbHyQz07DMAyH\n70i8Q+QralM4IITa7kDhCAiNB7ASt43WOlEcyvb2pNu4IODoPz9/n1xv9vOkForiPDdwXVagiI23\njocG3rdPxR0oScgWJ8/UwIEENu3lRb09BBKV0ywNjCmFe63FjDSjlD4Q50nv44wpl3HQAc0OB9I3\nVXWrjedEnIq03oC27qjHjympx31un0wiTQLq4bS4shrAECZnMGVTvbD9QSnOhDInjzsyuiBXWQP0\nr4R18jfgnHvJr4nOknrFmJ5xzhraRtHWf3Kkpfz/yGo5S+H73hkquyhdjr3R8m2lj09svwAAAP//\nAwBQSwMEFAAGAAgAAAAhAFr0LFu/AAAAFQEAAAsAAABfcmVscy8ucmVsc2zPwWrDMAwG4Ptg72B0\nX5TuUMaI01uh19I+gLGVxCy2jGSy9e1nemrHjpL4P0nD4SetZiPRyNnCruvBUPYcYp4tXC/Htw8w\nWl0ObuVMFm6kcBhfX4Yzra62kC6xqGlKVgtLreUTUf1CyWnHhXKbTCzJ1VbKjMX5LzcTvvf9HuXR\ngPHJNKdgQU5hB+ZyK23zHztFL6w81c5zQp6m6P9TMfB3PtPWFCczVQtB9N4U2rp2HOA44NMz4y8A\nAAD//wMAUEsDBBQABgAIAAAAIQDtLREpwQAAANsAAAAPAAAAZHJzL2Rvd25yZXYueG1sRE/basJA\nEH0X/IdlCn0R3TSCSHSVEijkrW30A8bsmESzszG7ubRf3y0U+jaHc539cTKNGKhztWUFL6sIBHFh\ndc2lgvPpbbkF4TyyxsYyKfgiB8fDfLbHRNuRP2nIfSlCCLsEFVTet4mUrqjIoFvZljhwV9sZ9AF2\npdQdjiHcNDKOoo00WHNoqLCltKLinvdGgV1kj1Re+NZP3228Lq4f71k+KvX8NL3uQHia/L/4z53p\nMH8Nv7+EA+ThBwAA//8DAFBLAQItABQABgAIAAAAIQDb4fbL7gAAAIUBAAATAAAAAAAAAAAAAAAA\nAAAAAABbQ29udGVudF9UeXBlc10ueG1sUEsBAi0AFAAGAAgAAAAhAFr0LFu/AAAAFQEAAAsAAAAA\nAAAAAAAAAAAAHwEAAF9yZWxzLy5yZWxzUEsBAi0AFAAGAAgAAAAhAO0tESnBAAAA2wAAAA8AAAAA\nAAAAAAAAAAAABwIAAGRycy9kb3ducmV2LnhtbFBLBQYAAAAAAwADALcAAAD1AgAAAAA=\n" };

                group2.Append(rectangle1);
                group2.Append(shapetype1);
                group2.Append(shape1);

                V.Rectangle rectangle2 = new V.Rectangle() { Id = "Rectangle 10", Style = "position:absolute;left:405;top:11415;width:1033;height:2805;visibility:visible;mso-wrap-style:square;v-text-anchor:top", OptionalString = "_x0000_s1030", Stroked = false };
                rectangle2.SetAttribute(new OpenXmlAttribute("o", "gfxdata", "urn:schemas-microsoft-com:office:office", "UEsDBBQABgAIAAAAIQDb4fbL7gAAAIUBAAATAAAAW0NvbnRlbnRfVHlwZXNdLnhtbHyQz07DMAyH\n70i8Q+QralM4IITa7kDhCAiNB7ASt43WOlEcyvb2pNu4IODoPz9/n1xv9vOkForiPDdwXVagiI23\njocG3rdPxR0oScgWJ8/UwIEENu3lRb09BBKV0ywNjCmFe63FjDSjlD4Q50nv44wpl3HQAc0OB9I3\nVXWrjedEnIq03oC27qjHjympx31un0wiTQLq4bS4shrAECZnMGVTvbD9QSnOhDInjzsyuiBXWQP0\nr4R18jfgnHvJr4nOknrFmJ5xzhraRtHWf3Kkpfz/yGo5S+H73hkquyhdjr3R8m2lj09svwAAAP//\nAwBQSwMEFAAGAAgAAAAhAFr0LFu/AAAAFQEAAAsAAABfcmVscy8ucmVsc2zPwWrDMAwG4Ptg72B0\nX5TuUMaI01uh19I+gLGVxCy2jGSy9e1nemrHjpL4P0nD4SetZiPRyNnCruvBUPYcYp4tXC/Htw8w\nWl0ObuVMFm6kcBhfX4Yzra62kC6xqGlKVgtLreUTUf1CyWnHhXKbTCzJ1VbKjMX5LzcTvvf9HuXR\ngPHJNKdgQU5hB+ZyK23zHztFL6w81c5zQp6m6P9TMfB3PtPWFCczVQtB9N4U2rp2HOA44NMz4y8A\nAAD//wMAUEsDBBQABgAIAAAAIQCBFKNDwQAAANsAAAAPAAAAZHJzL2Rvd25yZXYueG1sRE9Ni8Iw\nEL0v+B/CCN7WVBG3VKOoIHjwsl1BvY3N2BabSWlSW//9RljY2zze5yzXvanEkxpXWlYwGUcgiDOr\nS84VnH72nzEI55E1VpZJwYscrFeDjyUm2nb8Tc/U5yKEsEtQQeF9nUjpsoIMurGtiQN3t41BH2CT\nS91gF8JNJadRNJcGSw4NBda0Kyh7pK1R8DVr023Vxe093h39+SbN5XadKjUa9psFCE+9/xf/uQ86\nzJ/B+5dwgFz9AgAA//8DAFBLAQItABQABgAIAAAAIQDb4fbL7gAAAIUBAAATAAAAAAAAAAAAAAAA\nAAAAAABbQ29udGVudF9UeXBlc10ueG1sUEsBAi0AFAAGAAgAAAAhAFr0LFu/AAAAFQEAAAsAAAAA\nAAAAAAAAAAAAHwEAAF9yZWxzLy5yZWxzUEsBAi0AFAAGAAgAAAAhAIEUo0PBAAAA2wAAAA8AAAAA\nAAAAAAAAAAAABwIAAGRycy9kb3ducmV2LnhtbFBLBQYAAAAAAwADALcAAAD1AgAAAAA=\n"));

                V.TextBox textBox1 = new V.TextBox() { Style = "layout-flow:vertical", Inset = "0,0,0,0" };

                TextBoxContent textBoxContent2 = new TextBoxContent();

                Paragraph paragraph3 = new Paragraph() { RsidParagraphAddition = "00E9532F", RsidRunAdditionDefault = "00E9532F", ParagraphId = "48B42751", TextId = "77777777" };

                ParagraphProperties paragraphProperties3 = new ParagraphProperties();
                ParagraphStyleId paragraphStyleId3 = new ParagraphStyleId() { Val = "NoSpacing" };

                paragraphProperties3.Append(paragraphStyleId3);

                Run run7 = new Run();
                FieldChar fieldChar4 = new FieldChar() { FieldCharType = FieldCharValues.Begin };

                run7.Append(fieldChar4);

                Run run8 = new Run();
                FieldCode fieldCode2 = new FieldCode() { Space = SpaceProcessingModeValues.Preserve };
                fieldCode2.Text = " PAGE    \\* MERGEFORMAT ";

                run8.Append(fieldCode2);

                Run run9 = new Run();
                FieldChar fieldChar5 = new FieldChar() { FieldCharType = FieldCharValues.Separate };

                run9.Append(fieldChar5);

                Run run10 = new Run();

                RunProperties runProperties4 = new RunProperties();
                Bold bold3 = new Bold();
                BoldComplexScript boldComplexScript3 = new BoldComplexScript();
                NoProof noProof4 = new NoProof();
                Color color3 = new Color() { Val = "BF8F00", ThemeColor = ThemeColorValues.Accent4, ThemeShade = "BF" };
                FontSize fontSize3 = new FontSize() { Val = "52" };
                FontSizeComplexScript fontSizeComplexScript3 = new FontSizeComplexScript() { Val = "52" };

                runProperties4.Append(bold3);
                runProperties4.Append(boldComplexScript3);
                runProperties4.Append(noProof4);
                runProperties4.Append(color3);
                runProperties4.Append(fontSize3);
                runProperties4.Append(fontSizeComplexScript3);
                Text text2 = new Text();
                text2.Text = "2";

                run10.Append(runProperties4);
                run10.Append(text2);

                Run run11 = new Run();

                RunProperties runProperties5 = new RunProperties();
                Bold bold4 = new Bold();
                BoldComplexScript boldComplexScript4 = new BoldComplexScript();
                NoProof noProof5 = new NoProof();
                Color color4 = new Color() { Val = "BF8F00", ThemeColor = ThemeColorValues.Accent4, ThemeShade = "BF" };
                FontSize fontSize4 = new FontSize() { Val = "52" };
                FontSizeComplexScript fontSizeComplexScript4 = new FontSizeComplexScript() { Val = "52" };

                runProperties5.Append(bold4);
                runProperties5.Append(boldComplexScript4);
                runProperties5.Append(noProof5);
                runProperties5.Append(color4);
                runProperties5.Append(fontSize4);
                runProperties5.Append(fontSizeComplexScript4);
                FieldChar fieldChar6 = new FieldChar() { FieldCharType = FieldCharValues.End };

                run11.Append(runProperties5);
                run11.Append(fieldChar6);

                paragraph3.Append(paragraphProperties3);
                paragraph3.Append(run7);
                paragraph3.Append(run8);
                paragraph3.Append(run9);
                paragraph3.Append(run10);
                paragraph3.Append(run11);

                textBoxContent2.Append(paragraph3);

                textBox1.Append(textBoxContent2);

                rectangle2.Append(textBox1);
                Wvml.TextWrap textWrap1 = new Wvml.TextWrap() { AnchorX = Wvml.HorizontalAnchorValues.Margin, AnchorY = Wvml.VerticalAnchorValues.Margin };

                group1.Append(group2);
                group1.Append(rectangle2);
                group1.Append(textWrap1);

                picture1.Append(group1);

                alternateContentFallback1.Append(picture1);

                alternateContent1.Append(alternateContentChoice1);
                alternateContent1.Append(alternateContentFallback1);

                run1.Append(runProperties1);
                run1.Append(alternateContent1);

                paragraph1.Append(paragraphProperties1);
                paragraph1.Append(run1);

                sdtContentBlock1.Append(paragraph1);

                sdtBlock1.Append(sdtProperties1);
                sdtBlock1.Append(sdtContentBlock1);
                return sdtBlock1;

            }
        }
        private static SdtBlock VerticalOutline2 {
            get {
                SdtBlock sdtBlock1 = new SdtBlock();

                SdtProperties sdtProperties1 = new SdtProperties();
                SdtId sdtId1 = new SdtId() { Val = -1315946405 };

                SdtContentDocPartObject sdtContentDocPartObject1 = new SdtContentDocPartObject();
                DocPartGallery docPartGallery1 = new DocPartGallery() { Val = "Page Numbers (Top of Page)" };
                DocPartUnique docPartUnique1 = new DocPartUnique();

                sdtContentDocPartObject1.Append(docPartGallery1);
                sdtContentDocPartObject1.Append(docPartUnique1);

                sdtProperties1.Append(sdtId1);
                sdtProperties1.Append(sdtContentDocPartObject1);

                SdtContentBlock sdtContentBlock1 = new SdtContentBlock();

                Paragraph paragraph1 = new Paragraph() { RsidParagraphMarkRevision = "002C0D80", RsidParagraphAddition = "00E11B1E", RsidParagraphProperties = "002C0D80", RsidRunAdditionDefault = "002C0D80", ParagraphId = "16AD3494", TextId = "179C0DC2" };

                ParagraphProperties paragraphProperties1 = new ParagraphProperties();
                ParagraphStyleId paragraphStyleId1 = new ParagraphStyleId() { Val = "Header" };

                paragraphProperties1.Append(paragraphStyleId1);

                Run run1 = new Run();

                RunProperties runProperties1 = new RunProperties();
                NoProof noProof1 = new NoProof();

                runProperties1.Append(noProof1);

                AlternateContent alternateContent1 = new AlternateContent();

                AlternateContentChoice alternateContentChoice1 = new AlternateContentChoice() { Requires = "wpg" };

                Drawing drawing1 = new Drawing();

                Wp.Anchor anchor1 = new Wp.Anchor() { DistanceFromTop = (UInt32Value)0U, DistanceFromBottom = (UInt32Value)0U, DistanceFromLeft = (UInt32Value)114300U, DistanceFromRight = (UInt32Value)114300U, SimplePos = false, RelativeHeight = (UInt32Value)251659264U, BehindDoc = false, Locked = false, LayoutInCell = false, AllowOverlap = true, EditId = "0E69DD95", AnchorId = "028B62BF" };
                Wp.SimplePosition simplePosition1 = new Wp.SimplePosition() { X = 0L, Y = 0L };

                Wp.HorizontalPosition horizontalPosition1 = new Wp.HorizontalPosition() { RelativeFrom = Wp.HorizontalRelativePositionValues.RightMargin };
                Wp.HorizontalAlignment horizontalAlignment1 = new Wp.HorizontalAlignment();
                horizontalAlignment1.Text = "left";

                horizontalPosition1.Append(horizontalAlignment1);

                Wp.VerticalPosition verticalPosition1 = new Wp.VerticalPosition() { RelativeFrom = Wp.VerticalRelativePositionValues.Margin };
                Wp.VerticalAlignment verticalAlignment1 = new Wp.VerticalAlignment();
                verticalAlignment1.Text = "top";

                verticalPosition1.Append(verticalAlignment1);
                Wp.Extent extent1 = new Wp.Extent() { Cx = 902335L, Cy = 1902460L };
                Wp.EffectExtent effectExtent1 = new Wp.EffectExtent() { LeftEdge = 0L, TopEdge = 9525L, RightEdge = 12065L, BottomEdge = 2540L };
                Wp.WrapNone wrapNone1 = new Wp.WrapNone();
                Wp.DocProperties docProperties1 = new Wp.DocProperties() { Id = (UInt32Value)15U, Name = "Group 15" };

                Wp.NonVisualGraphicFrameDrawingProperties nonVisualGraphicFrameDrawingProperties1 = new Wp.NonVisualGraphicFrameDrawingProperties();

                A.GraphicFrameLocks graphicFrameLocks1 = new A.GraphicFrameLocks();
                graphicFrameLocks1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

                nonVisualGraphicFrameDrawingProperties1.Append(graphicFrameLocks1);

                A.Graphic graphic1 = new A.Graphic();
                graphic1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

                A.GraphicData graphicData1 = new A.GraphicData() { Uri = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" };

                Wpg.WordprocessingGroup wordprocessingGroup1 = new Wpg.WordprocessingGroup();

                Wpg.NonVisualGroupDrawingShapeProperties nonVisualGroupDrawingShapeProperties1 = new Wpg.NonVisualGroupDrawingShapeProperties();
                A.GroupShapeLocks groupShapeLocks1 = new A.GroupShapeLocks();

                nonVisualGroupDrawingShapeProperties1.Append(groupShapeLocks1);

                Wpg.GroupShapeProperties groupShapeProperties1 = new Wpg.GroupShapeProperties() { BlackWhiteMode = A.BlackWhiteModeValues.Auto };

                A.TransformGroup transformGroup1 = new A.TransformGroup() { HorizontalFlip = true, VerticalFlip = true };
                A.Offset offset1 = new A.Offset() { X = 0L, Y = 0L };
                A.Extents extents1 = new A.Extents() { Cx = 902335L, Cy = 1902460L };
                A.ChildOffset childOffset1 = new A.ChildOffset() { X = 13L, Y = 11415L };
                A.ChildExtents childExtents1 = new A.ChildExtents() { Cx = 1425L, Cy = 2996L };

                transformGroup1.Append(offset1);
                transformGroup1.Append(extents1);
                transformGroup1.Append(childOffset1);
                transformGroup1.Append(childExtents1);

                groupShapeProperties1.Append(transformGroup1);

                Wpg.GroupShape groupShape1 = new Wpg.GroupShape();
                Wpg.NonVisualDrawingProperties nonVisualDrawingProperties1 = new Wpg.NonVisualDrawingProperties() { Id = (UInt32Value)16U, Name = "Group 7" };

                Wpg.NonVisualGroupDrawingShapeProperties nonVisualGroupDrawingShapeProperties2 = new Wpg.NonVisualGroupDrawingShapeProperties();
                A.GroupShapeLocks groupShapeLocks2 = new A.GroupShapeLocks();

                nonVisualGroupDrawingShapeProperties2.Append(groupShapeLocks2);

                Wpg.GroupShapeProperties groupShapeProperties2 = new Wpg.GroupShapeProperties() { BlackWhiteMode = A.BlackWhiteModeValues.Auto };

                A.TransformGroup transformGroup2 = new A.TransformGroup() { VerticalFlip = true };
                A.Offset offset2 = new A.Offset() { X = 13L, Y = 14340L };
                A.Extents extents2 = new A.Extents() { Cx = 1410L, Cy = 71L };
                A.ChildOffset childOffset2 = new A.ChildOffset() { X = -83L, Y = 540L };
                A.ChildExtents childExtents2 = new A.ChildExtents() { Cx = 1218L, Cy = 71L };

                transformGroup2.Append(offset2);
                transformGroup2.Append(extents2);
                transformGroup2.Append(childOffset2);
                transformGroup2.Append(childExtents2);

                groupShapeProperties2.Append(transformGroup2);

                Wps.WordprocessingShape wordprocessingShape1 = new Wps.WordprocessingShape();
                Wps.NonVisualDrawingProperties nonVisualDrawingProperties2 = new Wps.NonVisualDrawingProperties() { Id = (UInt32Value)17U, Name = "Rectangle 8" };

                Wps.NonVisualDrawingShapeProperties nonVisualDrawingShapeProperties1 = new Wps.NonVisualDrawingShapeProperties();
                A.ShapeLocks shapeLocks1 = new A.ShapeLocks() { NoChangeArrowheads = true };

                nonVisualDrawingShapeProperties1.Append(shapeLocks1);

                Wps.ShapeProperties shapeProperties1 = new Wps.ShapeProperties() { BlackWhiteMode = A.BlackWhiteModeValues.Auto };

                A.Transform2D transform2D1 = new A.Transform2D();
                A.Offset offset3 = new A.Offset() { X = 678L, Y = 540L };
                A.Extents extents3 = new A.Extents() { Cx = 457L, Cy = 71L };

                transform2D1.Append(offset3);
                transform2D1.Append(extents3);

                A.PresetGeometry presetGeometry1 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
                A.AdjustValueList adjustValueList1 = new A.AdjustValueList();

                presetGeometry1.Append(adjustValueList1);

                A.SolidFill solidFill1 = new A.SolidFill();
                A.RgbColorModelHex rgbColorModelHex1 = new A.RgbColorModelHex() { Val = "5F497A" };

                solidFill1.Append(rgbColorModelHex1);

                A.Outline outline1 = new A.Outline() { Width = 9525 };

                A.SolidFill solidFill2 = new A.SolidFill();
                A.RgbColorModelHex rgbColorModelHex2 = new A.RgbColorModelHex() { Val = "5F497A" };

                solidFill2.Append(rgbColorModelHex2);
                A.Miter miter1 = new A.Miter() { Limit = 800000 };
                A.HeadEnd headEnd1 = new A.HeadEnd();
                A.TailEnd tailEnd1 = new A.TailEnd();

                outline1.Append(solidFill2);
                outline1.Append(miter1);
                outline1.Append(headEnd1);
                outline1.Append(tailEnd1);

                shapeProperties1.Append(transform2D1);
                shapeProperties1.Append(presetGeometry1);
                shapeProperties1.Append(solidFill1);
                shapeProperties1.Append(outline1);

                Wps.TextBodyProperties textBodyProperties1 = new Wps.TextBodyProperties() { Rotation = 0, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, LeftInset = 91440, TopInset = 45720, RightInset = 91440, BottomInset = 45720, Anchor = A.TextAnchoringTypeValues.Top, AnchorCenter = false, UpRight = true };
                A.NoAutoFit noAutoFit1 = new A.NoAutoFit();

                textBodyProperties1.Append(noAutoFit1);

                wordprocessingShape1.Append(nonVisualDrawingProperties2);
                wordprocessingShape1.Append(nonVisualDrawingShapeProperties1);
                wordprocessingShape1.Append(shapeProperties1);
                wordprocessingShape1.Append(textBodyProperties1);

                Wps.WordprocessingShape wordprocessingShape2 = new Wps.WordprocessingShape();
                Wps.NonVisualDrawingProperties nonVisualDrawingProperties3 = new Wps.NonVisualDrawingProperties() { Id = (UInt32Value)18U, Name = "AutoShape 9" };

                Wps.NonVisualConnectorProperties nonVisualConnectorProperties1 = new Wps.NonVisualConnectorProperties();
                A.ConnectionShapeLocks connectionShapeLocks1 = new A.ConnectionShapeLocks() { NoChangeShapeType = true };

                nonVisualConnectorProperties1.Append(connectionShapeLocks1);

                Wps.ShapeProperties shapeProperties2 = new Wps.ShapeProperties() { BlackWhiteMode = A.BlackWhiteModeValues.Auto };

                A.Transform2D transform2D2 = new A.Transform2D() { HorizontalFlip = true };
                A.Offset offset4 = new A.Offset() { X = -83L, Y = 540L };
                A.Extents extents4 = new A.Extents() { Cx = 761L, Cy = 0L };

                transform2D2.Append(offset4);
                transform2D2.Append(extents4);

                A.PresetGeometry presetGeometry2 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.StraightConnector1 };
                A.AdjustValueList adjustValueList2 = new A.AdjustValueList();

                presetGeometry2.Append(adjustValueList2);
                A.NoFill noFill1 = new A.NoFill();

                A.Outline outline2 = new A.Outline() { Width = 9525 };

                A.SolidFill solidFill3 = new A.SolidFill();
                A.RgbColorModelHex rgbColorModelHex3 = new A.RgbColorModelHex() { Val = "5F497A" };

                solidFill3.Append(rgbColorModelHex3);
                A.Round round1 = new A.Round();
                A.HeadEnd headEnd2 = new A.HeadEnd();
                A.TailEnd tailEnd2 = new A.TailEnd();

                outline2.Append(solidFill3);
                outline2.Append(round1);
                outline2.Append(headEnd2);
                outline2.Append(tailEnd2);

                A.ShapePropertiesExtensionList shapePropertiesExtensionList1 = new A.ShapePropertiesExtensionList();

                A.ShapePropertiesExtension shapePropertiesExtension1 = new A.ShapePropertiesExtension() { Uri = "{909E8E84-426E-40DD-AFC4-6F175D3DCCD1}" };

                A14.HiddenFillProperties hiddenFillProperties1 = new A14.HiddenFillProperties();
                hiddenFillProperties1.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");
                A.NoFill noFill2 = new A.NoFill();

                hiddenFillProperties1.Append(noFill2);

                shapePropertiesExtension1.Append(hiddenFillProperties1);

                shapePropertiesExtensionList1.Append(shapePropertiesExtension1);

                shapeProperties2.Append(transform2D2);
                shapeProperties2.Append(presetGeometry2);
                shapeProperties2.Append(noFill1);
                shapeProperties2.Append(outline2);
                shapeProperties2.Append(shapePropertiesExtensionList1);
                Wps.TextBodyProperties textBodyProperties2 = new Wps.TextBodyProperties();

                wordprocessingShape2.Append(nonVisualDrawingProperties3);
                wordprocessingShape2.Append(nonVisualConnectorProperties1);
                wordprocessingShape2.Append(shapeProperties2);
                wordprocessingShape2.Append(textBodyProperties2);

                groupShape1.Append(nonVisualDrawingProperties1);
                groupShape1.Append(nonVisualGroupDrawingShapeProperties2);
                groupShape1.Append(groupShapeProperties2);
                groupShape1.Append(wordprocessingShape1);
                groupShape1.Append(wordprocessingShape2);

                Wps.WordprocessingShape wordprocessingShape3 = new Wps.WordprocessingShape();
                Wps.NonVisualDrawingProperties nonVisualDrawingProperties4 = new Wps.NonVisualDrawingProperties() { Id = (UInt32Value)19U, Name = "Rectangle 10" };

                Wps.NonVisualDrawingShapeProperties nonVisualDrawingShapeProperties2 = new Wps.NonVisualDrawingShapeProperties();
                A.ShapeLocks shapeLocks2 = new A.ShapeLocks() { NoChangeArrowheads = true };

                nonVisualDrawingShapeProperties2.Append(shapeLocks2);

                Wps.ShapeProperties shapeProperties3 = new Wps.ShapeProperties() { BlackWhiteMode = A.BlackWhiteModeValues.Auto };

                A.Transform2D transform2D3 = new A.Transform2D();
                A.Offset offset5 = new A.Offset() { X = 405L, Y = 11415L };
                A.Extents extents5 = new A.Extents() { Cx = 1033L, Cy = 2805L };

                transform2D3.Append(offset5);
                transform2D3.Append(extents5);

                A.PresetGeometry presetGeometry3 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
                A.AdjustValueList adjustValueList3 = new A.AdjustValueList();

                presetGeometry3.Append(adjustValueList3);

                A.SolidFill solidFill4 = new A.SolidFill();
                A.RgbColorModelHex rgbColorModelHex4 = new A.RgbColorModelHex() { Val = "FFFFFF" };

                solidFill4.Append(rgbColorModelHex4);

                A.Outline outline3 = new A.Outline();
                A.NoFill noFill3 = new A.NoFill();

                outline3.Append(noFill3);

                A.ShapePropertiesExtensionList shapePropertiesExtensionList2 = new A.ShapePropertiesExtensionList();

                A.ShapePropertiesExtension shapePropertiesExtension2 = new A.ShapePropertiesExtension() { Uri = "{91240B29-F687-4F45-9708-019B960494DF}" };

                A14.HiddenLineProperties hiddenLineProperties1 = new A14.HiddenLineProperties() { Width = 9525 };
                hiddenLineProperties1.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");

                A.SolidFill solidFill5 = new A.SolidFill();
                A.RgbColorModelHex rgbColorModelHex5 = new A.RgbColorModelHex() { Val = "000000" };

                solidFill5.Append(rgbColorModelHex5);
                A.Miter miter2 = new A.Miter() { Limit = 800000 };
                A.HeadEnd headEnd3 = new A.HeadEnd();
                A.TailEnd tailEnd3 = new A.TailEnd();

                hiddenLineProperties1.Append(solidFill5);
                hiddenLineProperties1.Append(miter2);
                hiddenLineProperties1.Append(headEnd3);
                hiddenLineProperties1.Append(tailEnd3);

                shapePropertiesExtension2.Append(hiddenLineProperties1);

                shapePropertiesExtensionList2.Append(shapePropertiesExtension2);

                shapeProperties3.Append(transform2D3);
                shapeProperties3.Append(presetGeometry3);
                shapeProperties3.Append(solidFill4);
                shapeProperties3.Append(outline3);
                shapeProperties3.Append(shapePropertiesExtensionList2);

                Wps.TextBoxInfo2 textBoxInfo21 = new Wps.TextBoxInfo2();

                TextBoxContent textBoxContent1 = new TextBoxContent();

                Paragraph paragraph2 = new Paragraph() { RsidParagraphAddition = "002C0D80", RsidRunAdditionDefault = "002C0D80", ParagraphId = "3857EF7C", TextId = "77777777" };

                ParagraphProperties paragraphProperties2 = new ParagraphProperties();
                ParagraphStyleId paragraphStyleId2 = new ParagraphStyleId() { Val = "NoSpacing" };
                Justification justification1 = new Justification() { Val = JustificationValues.Right };

                paragraphProperties2.Append(paragraphStyleId2);
                paragraphProperties2.Append(justification1);

                Run run2 = new Run();
                FieldChar fieldChar1 = new FieldChar() { FieldCharType = FieldCharValues.Begin };

                run2.Append(fieldChar1);

                Run run3 = new Run();
                FieldCode fieldCode1 = new FieldCode() { Space = SpaceProcessingModeValues.Preserve };
                fieldCode1.Text = " PAGE    \\* MERGEFORMAT ";

                run3.Append(fieldCode1);

                Run run4 = new Run();
                FieldChar fieldChar2 = new FieldChar() { FieldCharType = FieldCharValues.Separate };

                run4.Append(fieldChar2);

                Run run5 = new Run();

                RunProperties runProperties2 = new RunProperties();
                Bold bold1 = new Bold();
                BoldComplexScript boldComplexScript1 = new BoldComplexScript();
                NoProof noProof2 = new NoProof();
                Color color1 = new Color() { Val = "BF8F00", ThemeColor = ThemeColorValues.Accent4, ThemeShade = "BF" };
                FontSize fontSize1 = new FontSize() { Val = "52" };
                FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "52" };

                runProperties2.Append(bold1);
                runProperties2.Append(boldComplexScript1);
                runProperties2.Append(noProof2);
                runProperties2.Append(color1);
                runProperties2.Append(fontSize1);
                runProperties2.Append(fontSizeComplexScript1);
                Text text1 = new Text();
                text1.Text = "2";

                run5.Append(runProperties2);
                run5.Append(text1);

                Run run6 = new Run();

                RunProperties runProperties3 = new RunProperties();
                Bold bold2 = new Bold();
                BoldComplexScript boldComplexScript2 = new BoldComplexScript();
                NoProof noProof3 = new NoProof();
                Color color2 = new Color() { Val = "BF8F00", ThemeColor = ThemeColorValues.Accent4, ThemeShade = "BF" };
                FontSize fontSize2 = new FontSize() { Val = "52" };
                FontSizeComplexScript fontSizeComplexScript2 = new FontSizeComplexScript() { Val = "52" };

                runProperties3.Append(bold2);
                runProperties3.Append(boldComplexScript2);
                runProperties3.Append(noProof3);
                runProperties3.Append(color2);
                runProperties3.Append(fontSize2);
                runProperties3.Append(fontSizeComplexScript2);
                FieldChar fieldChar3 = new FieldChar() { FieldCharType = FieldCharValues.End };

                run6.Append(runProperties3);
                run6.Append(fieldChar3);

                paragraph2.Append(paragraphProperties2);
                paragraph2.Append(run2);
                paragraph2.Append(run3);
                paragraph2.Append(run4);
                paragraph2.Append(run5);
                paragraph2.Append(run6);

                textBoxContent1.Append(paragraph2);

                textBoxInfo21.Append(textBoxContent1);

                Wps.TextBodyProperties textBodyProperties3 = new Wps.TextBodyProperties() { Rotation = 0, Vertical = A.TextVerticalValues.Vertical270, Wrap = A.TextWrappingValues.Square, LeftInset = 0, TopInset = 0, RightInset = 0, BottomInset = 0, Anchor = A.TextAnchoringTypeValues.Top, AnchorCenter = false, UpRight = true };
                A.NoAutoFit noAutoFit2 = new A.NoAutoFit();

                textBodyProperties3.Append(noAutoFit2);

                wordprocessingShape3.Append(nonVisualDrawingProperties4);
                wordprocessingShape3.Append(nonVisualDrawingShapeProperties2);
                wordprocessingShape3.Append(shapeProperties3);
                wordprocessingShape3.Append(textBoxInfo21);
                wordprocessingShape3.Append(textBodyProperties3);

                wordprocessingGroup1.Append(nonVisualGroupDrawingShapeProperties1);
                wordprocessingGroup1.Append(groupShapeProperties1);
                wordprocessingGroup1.Append(groupShape1);
                wordprocessingGroup1.Append(wordprocessingShape3);

                graphicData1.Append(wordprocessingGroup1);

                graphic1.Append(graphicData1);

                Wp14.RelativeWidth relativeWidth1 = new Wp14.RelativeWidth() { ObjectId = Wp14.SizeRelativeHorizontallyValues.LeftMargin };
                Wp14.PercentageWidth percentageWidth1 = new Wp14.PercentageWidth();
                percentageWidth1.Text = "100000";

                relativeWidth1.Append(percentageWidth1);

                Wp14.RelativeHeight relativeHeight1 = new Wp14.RelativeHeight() { RelativeFrom = Wp14.SizeRelativeVerticallyValues.Page };
                Wp14.PercentageHeight percentageHeight1 = new Wp14.PercentageHeight();
                percentageHeight1.Text = "0";

                relativeHeight1.Append(percentageHeight1);

                anchor1.Append(simplePosition1);
                anchor1.Append(horizontalPosition1);
                anchor1.Append(verticalPosition1);
                anchor1.Append(extent1);
                anchor1.Append(effectExtent1);
                anchor1.Append(wrapNone1);
                anchor1.Append(docProperties1);
                anchor1.Append(nonVisualGraphicFrameDrawingProperties1);
                anchor1.Append(graphic1);
                anchor1.Append(relativeWidth1);
                anchor1.Append(relativeHeight1);

                drawing1.Append(anchor1);

                alternateContentChoice1.Append(drawing1);

                AlternateContentFallback alternateContentFallback1 = new AlternateContentFallback();

                Picture picture1 = new Picture();

                V.Group group1 = new V.Group() { Id = "Group 15", Style = "position:absolute;margin-left:0;margin-top:0;width:71.05pt;height:149.8pt;flip:x y;z-index:251659264;mso-width-percent:1000;mso-position-horizontal:left;mso-position-horizontal-relative:right-margin-area;mso-position-vertical:top;mso-position-vertical-relative:margin;mso-width-percent:1000;mso-width-relative:left-margin-area", CoordinateSize = "1425,2996", CoordinateOrigin = "13,11415", OptionalString = "_x0000_s1026", AllowInCell = false };
                group1.SetAttribute(new OpenXmlAttribute("w14", "anchorId", "http://schemas.microsoft.com/office/word/2010/wordml", "028B62BF"));
                group1.SetAttribute(new OpenXmlAttribute("o", "gfxdata", "urn:schemas-microsoft-com:office:office", "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\n"));

                V.Group group2 = new V.Group() { Id = "Group 7", Style = "position:absolute;left:13;top:14340;width:1410;height:71;flip:y", CoordinateSize = "1218,71", CoordinateOrigin = "-83,540", OptionalString = "_x0000_s1027" };
                group2.SetAttribute(new OpenXmlAttribute("o", "gfxdata", "urn:schemas-microsoft-com:office:office", "UEsDBBQABgAIAAAAIQDb4fbL7gAAAIUBAAATAAAAW0NvbnRlbnRfVHlwZXNdLnhtbHyQz07DMAyH\n70i8Q+QralM4IITa7kDhCAiNB7ASt43WOlEcyvb2pNu4IODoPz9/n1xv9vOkForiPDdwXVagiI23\njocG3rdPxR0oScgWJ8/UwIEENu3lRb09BBKV0ywNjCmFe63FjDSjlD4Q50nv44wpl3HQAc0OB9I3\nVXWrjedEnIq03oC27qjHjympx31un0wiTQLq4bS4shrAECZnMGVTvbD9QSnOhDInjzsyuiBXWQP0\nr4R18jfgnHvJr4nOknrFmJ5xzhraRtHWf3Kkpfz/yGo5S+H73hkquyhdjr3R8m2lj09svwAAAP//\nAwBQSwMEFAAGAAgAAAAhAFr0LFu/AAAAFQEAAAsAAABfcmVscy8ucmVsc2zPwWrDMAwG4Ptg72B0\nX5TuUMaI01uh19I+gLGVxCy2jGSy9e1nemrHjpL4P0nD4SetZiPRyNnCruvBUPYcYp4tXC/Htw8w\nWl0ObuVMFm6kcBhfX4Yzra62kC6xqGlKVgtLreUTUf1CyWnHhXKbTCzJ1VbKjMX5LzcTvvf9HuXR\ngPHJNKdgQU5hB+ZyK23zHztFL6w81c5zQp6m6P9TMfB3PtPWFCczVQtB9N4U2rp2HOA44NMz4y8A\nAAD//wMAUEsDBBQABgAIAAAAIQCvT5vfwQAAANsAAAAPAAAAZHJzL2Rvd25yZXYueG1sRE/JasMw\nEL0H+g9iAr0lcooJwbESQqAllF7qLPg4WBNbxBoZS7Xdv68Khd7m8dbJ95NtxUC9N44VrJYJCOLK\nacO1gsv5dbEB4QOyxtYxKfgmD/vd0yzHTLuRP2koQi1iCPsMFTQhdJmUvmrIol+6jjhyd9dbDBH2\ntdQ9jjHctvIlSdbSouHY0GBHx4aqR/FlFVwPJqX0Vr5/JBXRScvyrTCpUs/z6bAFEWgK/+I/90nH\n+Wv4/SUeIHc/AAAA//8DAFBLAQItABQABgAIAAAAIQDb4fbL7gAAAIUBAAATAAAAAAAAAAAAAAAA\nAAAAAABbQ29udGVudF9UeXBlc10ueG1sUEsBAi0AFAAGAAgAAAAhAFr0LFu/AAAAFQEAAAsAAAAA\nAAAAAAAAAAAAHwEAAF9yZWxzLy5yZWxzUEsBAi0AFAAGAAgAAAAhAK9Pm9/BAAAA2wAAAA8AAAAA\nAAAAAAAAAAAABwIAAGRycy9kb3ducmV2LnhtbFBLBQYAAAAAAwADALcAAAD1AgAAAAA=\n"));

                V.Rectangle rectangle1 = new V.Rectangle() { Id = "Rectangle 8", Style = "position:absolute;left:678;top:540;width:457;height:71;visibility:visible;mso-wrap-style:square;v-text-anchor:top", OptionalString = "_x0000_s1028", FillColor = "#5f497a", StrokeColor = "#5f497a" };
                rectangle1.SetAttribute(new OpenXmlAttribute("o", "gfxdata", "urn:schemas-microsoft-com:office:office", "UEsDBBQABgAIAAAAIQDb4fbL7gAAAIUBAAATAAAAW0NvbnRlbnRfVHlwZXNdLnhtbHyQz07DMAyH\n70i8Q+QralM4IITa7kDhCAiNB7ASt43WOlEcyvb2pNu4IODoPz9/n1xv9vOkForiPDdwXVagiI23\njocG3rdPxR0oScgWJ8/UwIEENu3lRb09BBKV0ywNjCmFe63FjDSjlD4Q50nv44wpl3HQAc0OB9I3\nVXWrjedEnIq03oC27qjHjympx31un0wiTQLq4bS4shrAECZnMGVTvbD9QSnOhDInjzsyuiBXWQP0\nr4R18jfgnHvJr4nOknrFmJ5xzhraRtHWf3Kkpfz/yGo5S+H73hkquyhdjr3R8m2lj09svwAAAP//\nAwBQSwMEFAAGAAgAAAAhAFr0LFu/AAAAFQEAAAsAAABfcmVscy8ucmVsc2zPwWrDMAwG4Ptg72B0\nX5TuUMaI01uh19I+gLGVxCy2jGSy9e1nemrHjpL4P0nD4SetZiPRyNnCruvBUPYcYp4tXC/Htw8w\nWl0ObuVMFm6kcBhfX4Yzra62kC6xqGlKVgtLreUTUf1CyWnHhXKbTCzJ1VbKjMX5LzcTvvf9HuXR\ngPHJNKdgQU5hB+ZyK23zHztFL6w81c5zQp6m6P9TMfB3PtPWFCczVQtB9N4U2rp2HOA44NMz4y8A\nAAD//wMAUEsDBBQABgAIAAAAIQApYTJ8wwAAANsAAAAPAAAAZHJzL2Rvd25yZXYueG1sRE9Na8JA\nEL0X/A/LCL3VTT3UNrqKiJWCUqlVvA7ZaRLNzqTZbYz/3i0UepvH+5zJrHOVaqnxpbCBx0ECijgT\nW3JuYP/5+vAMygdki5UwGbiSh9m0dzfB1MqFP6jdhVzFEPYpGihCqFOtfVaQQz+QmjhyX9I4DBE2\nubYNXmK4q/QwSZ60w5JjQ4E1LQrKzrsfZ+AkR2kP77LdbL4pWZ7mq+3LemXMfb+bj0EF6sK/+M/9\nZuP8Efz+Eg/Q0xsAAAD//wMAUEsBAi0AFAAGAAgAAAAhANvh9svuAAAAhQEAABMAAAAAAAAAAAAA\nAAAAAAAAAFtDb250ZW50X1R5cGVzXS54bWxQSwECLQAUAAYACAAAACEAWvQsW78AAAAVAQAACwAA\nAAAAAAAAAAAAAAAfAQAAX3JlbHMvLnJlbHNQSwECLQAUAAYACAAAACEAKWEyfMMAAADbAAAADwAA\nAAAAAAAAAAAAAAAHAgAAZHJzL2Rvd25yZXYueG1sUEsFBgAAAAADAAMAtwAAAPcCAAAAAA==\n"));

                V.Shapetype shapetype1 = new V.Shapetype() { Id = "_x0000_t32", CoordinateSize = "21600,21600", Oned = true, Filled = false, OptionalNumber = 32, EdgePath = "m,l21600,21600e" };
                V.Path path1 = new V.Path() { AllowFill = false, ShowArrowhead = true, ConnectionPointType = Ovml.ConnectValues.None };
                Ovml.Lock lock1 = new Ovml.Lock() { Extension = V.ExtensionHandlingBehaviorValues.Edit, ShapeType = true };

                shapetype1.Append(path1);
                shapetype1.Append(lock1);
                V.Shape shape1 = new V.Shape() { Id = "AutoShape 9", Style = "position:absolute;left:-83;top:540;width:761;height:0;flip:x;visibility:visible;mso-wrap-style:square", OptionalString = "_x0000_s1029", StrokeColor = "#5f497a", ConnectorType = Ovml.ConnectorValues.Straight, Type = "#_x0000_t32", EncodedPackage = "UEsDBBQABgAIAAAAIQDb4fbL7gAAAIUBAAATAAAAW0NvbnRlbnRfVHlwZXNdLnhtbHyQz07DMAyH\n70i8Q+QralM4IITa7kDhCAiNB7ASt43WOlEcyvb2pNu4IODoPz9/n1xv9vOkForiPDdwXVagiI23\njocG3rdPxR0oScgWJ8/UwIEENu3lRb09BBKV0ywNjCmFe63FjDSjlD4Q50nv44wpl3HQAc0OB9I3\nVXWrjedEnIq03oC27qjHjympx31un0wiTQLq4bS4shrAECZnMGVTvbD9QSnOhDInjzsyuiBXWQP0\nr4R18jfgnHvJr4nOknrFmJ5xzhraRtHWf3Kkpfz/yGo5S+H73hkquyhdjr3R8m2lj09svwAAAP//\nAwBQSwMEFAAGAAgAAAAhAFr0LFu/AAAAFQEAAAsAAABfcmVscy8ucmVsc2zPwWrDMAwG4Ptg72B0\nX5TuUMaI01uh19I+gLGVxCy2jGSy9e1nemrHjpL4P0nD4SetZiPRyNnCruvBUPYcYp4tXC/Htw8w\nWl0ObuVMFm6kcBhfX4Yzra62kC6xqGlKVgtLreUTUf1CyWnHhXKbTCzJ1VbKjMX5LzcTvvf9HuXR\ngPHJNKdgQU5hB+ZyK23zHztFL6w81c5zQp6m6P9TMfB3PtPWFCczVQtB9N4U2rp2HOA44NMz4y8A\nAAD//wMAUEsDBBQABgAIAAAAIQDjiYNYxAAAANsAAAAPAAAAZHJzL2Rvd25yZXYueG1sRI9Ba8JA\nEIXvQv/DMkIvUje1UCTNRkQo5NY29QeM2TGJZmfT7GpSf71zKPQ2w3vz3jfZZnKdutIQWs8GnpcJ\nKOLK25ZrA/vv96c1qBCRLXaeycAvBdjkD7MMU+tH/qJrGWslIRxSNNDE2Kdah6ohh2Hpe2LRjn5w\nGGUdam0HHCXcdXqVJK/aYcvS0GBPu4aqc3lxBvyi+NnpA58u061fvVTHz4+iHI15nE/bN1CRpvhv\n/rsurOALrPwiA+j8DgAA//8DAFBLAQItABQABgAIAAAAIQDb4fbL7gAAAIUBAAATAAAAAAAAAAAA\nAAAAAAAAAABbQ29udGVudF9UeXBlc10ueG1sUEsBAi0AFAAGAAgAAAAhAFr0LFu/AAAAFQEAAAsA\nAAAAAAAAAAAAAAAAHwEAAF9yZWxzLy5yZWxzUEsBAi0AFAAGAAgAAAAhAOOJg1jEAAAA2wAAAA8A\nAAAAAAAAAAAAAAAABwIAAGRycy9kb3ducmV2LnhtbFBLBQYAAAAAAwADALcAAAD4AgAAAAA=\n" };

                group2.Append(rectangle1);
                group2.Append(shapetype1);
                group2.Append(shape1);

                V.Rectangle rectangle2 = new V.Rectangle() { Id = "Rectangle 10", Style = "position:absolute;left:405;top:11415;width:1033;height:2805;visibility:visible;mso-wrap-style:square;v-text-anchor:top", OptionalString = "_x0000_s1030", Stroked = false };
                rectangle2.SetAttribute(new OpenXmlAttribute("o", "gfxdata", "urn:schemas-microsoft-com:office:office", "UEsDBBQABgAIAAAAIQDb4fbL7gAAAIUBAAATAAAAW0NvbnRlbnRfVHlwZXNdLnhtbHyQz07DMAyH\n70i8Q+QralM4IITa7kDhCAiNB7ASt43WOlEcyvb2pNu4IODoPz9/n1xv9vOkForiPDdwXVagiI23\njocG3rdPxR0oScgWJ8/UwIEENu3lRb09BBKV0ywNjCmFe63FjDSjlD4Q50nv44wpl3HQAc0OB9I3\nVXWrjedEnIq03oC27qjHjympx31un0wiTQLq4bS4shrAECZnMGVTvbD9QSnOhDInjzsyuiBXWQP0\nr4R18jfgnHvJr4nOknrFmJ5xzhraRtHWf3Kkpfz/yGo5S+H73hkquyhdjr3R8m2lj09svwAAAP//\nAwBQSwMEFAAGAAgAAAAhAFr0LFu/AAAAFQEAAAsAAABfcmVscy8ucmVsc2zPwWrDMAwG4Ptg72B0\nX5TuUMaI01uh19I+gLGVxCy2jGSy9e1nemrHjpL4P0nD4SetZiPRyNnCruvBUPYcYp4tXC/Htw8w\nWl0ObuVMFm6kcBhfX4Yzra62kC6xqGlKVgtLreUTUf1CyWnHhXKbTCzJ1VbKjMX5LzcTvvf9HuXR\ngPHJNKdgQU5hB+ZyK23zHztFL6w81c5zQp6m6P9TMfB3PtPWFCczVQtB9N4U2rp2HOA44NMz4y8A\nAAD//wMAUEsDBBQABgAIAAAAIQB2SxQ/wQAAANsAAAAPAAAAZHJzL2Rvd25yZXYueG1sRE9Li8Iw\nEL4L+x/CLHiRNXUPUqtRlgVlD4LPy96GZmyLzSQ0sbb/3giCt/n4nrNYdaYWLTW+sqxgMk5AEOdW\nV1woOJ/WXykIH5A11pZJQU8eVsuPwQIzbe98oPYYChFD2GeooAzBZVL6vCSDfmwdceQutjEYImwK\nqRu8x3BTy+8kmUqDFceGEh39lpRfjzejYHv+d/3IJX2125vLNm1Hqd+QUsPP7mcOIlAX3uKX+0/H\n+TN4/hIPkMsHAAAA//8DAFBLAQItABQABgAIAAAAIQDb4fbL7gAAAIUBAAATAAAAAAAAAAAAAAAA\nAAAAAABbQ29udGVudF9UeXBlc10ueG1sUEsBAi0AFAAGAAgAAAAhAFr0LFu/AAAAFQEAAAsAAAAA\nAAAAAAAAAAAAHwEAAF9yZWxzLy5yZWxzUEsBAi0AFAAGAAgAAAAhAHZLFD/BAAAA2wAAAA8AAAAA\nAAAAAAAAAAAABwIAAGRycy9kb3ducmV2LnhtbFBLBQYAAAAAAwADALcAAAD1AgAAAAA=\n"));

                V.TextBox textBox1 = new V.TextBox() { Style = "layout-flow:vertical;mso-layout-flow-alt:bottom-to-top", Inset = "0,0,0,0" };

                TextBoxContent textBoxContent2 = new TextBoxContent();

                Paragraph paragraph3 = new Paragraph() { RsidParagraphAddition = "002C0D80", RsidRunAdditionDefault = "002C0D80", ParagraphId = "3857EF7C", TextId = "77777777" };

                ParagraphProperties paragraphProperties3 = new ParagraphProperties();
                ParagraphStyleId paragraphStyleId3 = new ParagraphStyleId() { Val = "NoSpacing" };
                Justification justification2 = new Justification() { Val = JustificationValues.Right };

                paragraphProperties3.Append(paragraphStyleId3);
                paragraphProperties3.Append(justification2);

                Run run7 = new Run();
                FieldChar fieldChar4 = new FieldChar() { FieldCharType = FieldCharValues.Begin };

                run7.Append(fieldChar4);

                Run run8 = new Run();
                FieldCode fieldCode2 = new FieldCode() { Space = SpaceProcessingModeValues.Preserve };
                fieldCode2.Text = " PAGE    \\* MERGEFORMAT ";

                run8.Append(fieldCode2);

                Run run9 = new Run();
                FieldChar fieldChar5 = new FieldChar() { FieldCharType = FieldCharValues.Separate };

                run9.Append(fieldChar5);

                Run run10 = new Run();

                RunProperties runProperties4 = new RunProperties();
                Bold bold3 = new Bold();
                BoldComplexScript boldComplexScript3 = new BoldComplexScript();
                NoProof noProof4 = new NoProof();
                Color color3 = new Color() { Val = "BF8F00", ThemeColor = ThemeColorValues.Accent4, ThemeShade = "BF" };
                FontSize fontSize3 = new FontSize() { Val = "52" };
                FontSizeComplexScript fontSizeComplexScript3 = new FontSizeComplexScript() { Val = "52" };

                runProperties4.Append(bold3);
                runProperties4.Append(boldComplexScript3);
                runProperties4.Append(noProof4);
                runProperties4.Append(color3);
                runProperties4.Append(fontSize3);
                runProperties4.Append(fontSizeComplexScript3);
                Text text2 = new Text();
                text2.Text = "2";

                run10.Append(runProperties4);
                run10.Append(text2);

                Run run11 = new Run();

                RunProperties runProperties5 = new RunProperties();
                Bold bold4 = new Bold();
                BoldComplexScript boldComplexScript4 = new BoldComplexScript();
                NoProof noProof5 = new NoProof();
                Color color4 = new Color() { Val = "BF8F00", ThemeColor = ThemeColorValues.Accent4, ThemeShade = "BF" };
                FontSize fontSize4 = new FontSize() { Val = "52" };
                FontSizeComplexScript fontSizeComplexScript4 = new FontSizeComplexScript() { Val = "52" };

                runProperties5.Append(bold4);
                runProperties5.Append(boldComplexScript4);
                runProperties5.Append(noProof5);
                runProperties5.Append(color4);
                runProperties5.Append(fontSize4);
                runProperties5.Append(fontSizeComplexScript4);
                FieldChar fieldChar6 = new FieldChar() { FieldCharType = FieldCharValues.End };

                run11.Append(runProperties5);
                run11.Append(fieldChar6);

                paragraph3.Append(paragraphProperties3);
                paragraph3.Append(run7);
                paragraph3.Append(run8);
                paragraph3.Append(run9);
                paragraph3.Append(run10);
                paragraph3.Append(run11);

                textBoxContent2.Append(paragraph3);

                textBox1.Append(textBoxContent2);

                rectangle2.Append(textBox1);
                Wvml.TextWrap textWrap1 = new Wvml.TextWrap() { AnchorX = Wvml.HorizontalAnchorValues.Margin, AnchorY = Wvml.VerticalAnchorValues.Margin };

                group1.Append(group2);
                group1.Append(rectangle2);
                group1.Append(textWrap1);

                picture1.Append(group1);

                alternateContentFallback1.Append(picture1);

                alternateContent1.Append(alternateContentChoice1);
                alternateContent1.Append(alternateContentFallback1);

                run1.Append(runProperties1);
                run1.Append(alternateContent1);

                paragraph1.Append(paragraphProperties1);
                paragraph1.Append(run1);

                sdtContentBlock1.Append(paragraph1);

                sdtBlock1.Append(sdtProperties1);
                sdtBlock1.Append(sdtContentBlock1);
                return sdtBlock1;
            }
        }
    }
}
