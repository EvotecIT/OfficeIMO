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
}
