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
}
