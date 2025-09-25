using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using A14 = DocumentFormat.OpenXml.Office2010.Drawing;
using Ovml = DocumentFormat.OpenXml.Vml.Office;
using V = DocumentFormat.OpenXml.Vml;
using Wp = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using Wp14 = DocumentFormat.OpenXml.Office2010.Word.Drawing;
using Wpg = DocumentFormat.OpenXml.Office2010.Word.DrawingGroup;
using Wps = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;
using Wvml = DocumentFormat.OpenXml.Vml.Wordprocessing;

namespace OfficeIMO.Word;

/// <summary>
/// Represents a page-numbering building block.
/// </summary>
public partial class WordPageNumber {
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
}
