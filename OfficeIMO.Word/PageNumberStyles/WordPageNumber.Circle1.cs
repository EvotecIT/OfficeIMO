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
}
