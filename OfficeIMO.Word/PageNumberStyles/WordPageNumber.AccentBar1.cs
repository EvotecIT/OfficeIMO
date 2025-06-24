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
}
