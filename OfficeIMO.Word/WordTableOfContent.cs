using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public enum TableOfContentStyle {
        Template1,
        Template2
    }

    public class WordTableOfContent : WordElement {
        private readonly WordDocument _document;
        private readonly SdtBlock _sdtBlock;

        public string Text {
            get {
                if (_sdtBlock != null) {
                    var paragraphs = _sdtBlock.SdtContentBlock.ChildElements.OfType<Paragraph>();
                    foreach (var paragraph in paragraphs) {
                        var run = paragraph.OfType<Run>().FirstOrDefault();
                        if (run != null) {
                            Text text = run.OfType<Text>().FirstOrDefault();
                            if (text != null) {
                                return text.Text;
                            }
                        }
                    }
                }
                return "";
            }
            set {
                if (_sdtBlock != null) {
                    var paragraphs = _sdtBlock.SdtContentBlock.ChildElements.OfType<Paragraph>();
                    foreach (var paragraph in paragraphs) {
                        var run = paragraph.OfType<Run>().FirstOrDefault();
                        if (run != null) {
                            Text text = run.OfType<Text>().FirstOrDefault();
                            if (text != null) {
                                text.Text = value;
                            }
                        }
                    }
                }
            }
        }
        public string TextNoContent {
            get {
                if (_sdtBlock != null) {
                    var paragraphs = _sdtBlock.SdtContentBlock.ChildElements.OfType<Paragraph>();
                    foreach (var paragraph in paragraphs) {
                        var simpleField = paragraph.OfType<SimpleField>().FirstOrDefault();
                        if (simpleField != null) {
                            var run = simpleField.OfType<Run>().FirstOrDefault();
                            if (run != null) {
                                Text text = run.OfType<Text>().FirstOrDefault();
                                if (text != null) {
                                    return text.Text;
                                }
                            }
                        }
                    }
                }
                return "";
            }
            set {
                if (_sdtBlock != null) {
                    var paragraphs = _sdtBlock.SdtContentBlock.ChildElements.OfType<Paragraph>();
                    foreach (var paragraph in paragraphs) {
                        var simpleField = paragraph.OfType<SimpleField>().FirstOrDefault();
                        if (simpleField != null) {
                            var run = simpleField.OfType<Run>().FirstOrDefault();
                            if (run != null) {
                                Text text = run.OfType<Text>().FirstOrDefault();
                                if (text != null) {
                                    text.Text = value;
                                }
                            }
                        }
                    }
                }
            }
        }



        public WordTableOfContent(WordDocument wordDocument, TableOfContentStyle tableOfContentStyle) {
            this._document = wordDocument;
            this._sdtBlock = GetStyle(tableOfContentStyle);
            this._document._wordprocessingDocument.MainDocumentPart.Document.Body.Append(_sdtBlock);

            //var currentStdBlock = this._document._wordprocessingDocument.MainDocumentPart.Document.Body.OfType<SdtBlock>();
            //if (currentStdBlock.ToList().Count > 0) {
            //    this._document._wordprocessingDocument.MainDocumentPart.Document.Body.InsertAt(_sdtBlock, 1);
            //} else {
            //    this._document._wordprocessingDocument.MainDocumentPart.Document.Body.InsertAt(_sdtBlock, 0);
            //}
        }

        public WordTableOfContent(WordDocument wordDocument, SdtBlock sdtBlock) {
            this._document = wordDocument;
            this._sdtBlock = sdtBlock;
        }

        public void Update() {
            this._document.Settings.UpdateFieldsOnOpen = true;
        }

        private static SdtBlock GetStyle(TableOfContentStyle style) {
            switch (style) {
                case TableOfContentStyle.Template1: return Template1;
                case TableOfContentStyle.Template2: return Template2;
            }
            throw new ArgumentOutOfRangeException(nameof(style));
        }
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
