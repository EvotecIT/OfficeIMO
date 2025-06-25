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
}
