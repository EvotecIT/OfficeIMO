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
            group1.SetAttribute(new OpenXmlAttribute("o", "gfxdata", "urn:schemas-microsoft-com:office:office", "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\nAAAA\n"));

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
}
