using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word;

/// <summary>
/// Represents a page-numbering building block.
/// </summary>
public partial class WordPageNumber {
    private static SdtBlock LargeItalics1 {
        get {
            SdtBlock sdtBlock1 = new SdtBlock();

            SdtProperties sdtProperties1 = new SdtProperties();

            RunProperties runProperties1 = new RunProperties();
            RunFonts runFonts1 = new RunFonts() { EastAsiaTheme = ThemeFontValues.MinorEastAsia };

            runProperties1.Append(runFonts1);
            SdtId sdtId1 = new SdtId();

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
}
