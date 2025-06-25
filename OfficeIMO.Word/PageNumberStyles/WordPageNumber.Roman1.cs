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

/// <summary>
/// Represents a page-numbering building block.
/// </summary>
public partial class WordPageNumber {
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
}
