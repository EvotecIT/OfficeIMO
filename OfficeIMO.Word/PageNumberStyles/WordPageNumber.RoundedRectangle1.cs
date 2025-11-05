using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using A14 = DocumentFormat.OpenXml.Office2010.Drawing;
using Ovml = DocumentFormat.OpenXml.Vml.Office;
using V = DocumentFormat.OpenXml.Vml;
using Wp = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using Wpg = DocumentFormat.OpenXml.Office2010.Word.DrawingGroup;
using Wps = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;
using Wvml = DocumentFormat.OpenXml.Vml.Wordprocessing;

namespace OfficeIMO.Word;

/// <summary>
/// Represents a page-numbering building block.
/// </summary>
public partial class WordPageNumber {
    private static SdtBlock RoundedRectangle1 {
        get {
            SdtBlock sdtBlock1 = new SdtBlock();

            SdtProperties sdtProperties1 = new SdtProperties();
            SdtId sdtId1 = new SdtId();

            SdtContentDocPartObject sdtContentDocPartObject1 = new SdtContentDocPartObject();
            DocPartGallery docPartGallery1 = new DocPartGallery() { Val = "Page Numbers (Top of Page)" };
            DocPartUnique docPartUnique1 = new DocPartUnique();

            sdtContentDocPartObject1.Append(docPartGallery1);
            sdtContentDocPartObject1.Append(docPartUnique1);

            sdtProperties1.Append(sdtId1);
            sdtProperties1.Append(sdtContentDocPartObject1);

            SdtContentBlock sdtContentBlock1 = new SdtContentBlock();

            Paragraph paragraph1 = new Paragraph() { RsidParagraphAddition = "009D383E", RsidRunAdditionDefault = "009D383E", ParagraphId = "2427BF06", TextId = "68F62570" };

            ParagraphProperties paragraphProperties1 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId1 = new ParagraphStyleId() { Val = "Header" };
            Indentation indentation1 = new Indentation() { Left = "-864" };

            paragraphProperties1.Append(paragraphStyleId1);
            paragraphProperties1.Append(indentation1);

            Run run1 = new Run();

            RunProperties runProperties1 = new RunProperties();
            NoProof noProof1 = new NoProof();

            runProperties1.Append(noProof1);

            AlternateContent alternateContent1 = new AlternateContent();

            AlternateContentChoice alternateContentChoice1 = new AlternateContentChoice() { Requires = "wpg" };

            Drawing drawing1 = new Drawing();

            Wp.Inline inline1 = new Wp.Inline() { DistanceFromTop = (UInt32Value)0U, DistanceFromBottom = (UInt32Value)0U, DistanceFromLeft = (UInt32Value)0U, DistanceFromRight = (UInt32Value)0U, AnchorId = "553944E0", EditId = "4F1437F3" };
            Wp.Extent extent1 = new Wp.Extent() { Cx = 548640L, Cy = 237490L };
            Wp.EffectExtent effectExtent1 = new Wp.EffectExtent() { LeftEdge = 9525L, TopEdge = 9525L, RightEdge = 13335L, BottomEdge = 10160L };
            Wp.DocProperties docProperties1 = new Wp.DocProperties() { Id = (UInt32Value)22U, Name = "Group 22" };

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
            A.Extents extents1 = new A.Extents() { Cx = 548640L, Cy = 237490L };
            A.ChildOffset childOffset1 = new A.ChildOffset() { X = 614L, Y = 660L };
            A.ChildExtents childExtents1 = new A.ChildExtents() { Cx = 864L, Cy = 374L };

            transformGroup1.Append(offset1);
            transformGroup1.Append(extents1);
            transformGroup1.Append(childOffset1);
            transformGroup1.Append(childExtents1);

            groupShapeProperties1.Append(transformGroup1);

            Wps.WordprocessingShape wordprocessingShape1 = new Wps.WordprocessingShape();
            Wps.NonVisualDrawingProperties nonVisualDrawingProperties1 = new Wps.NonVisualDrawingProperties() { Id = (UInt32Value)23U, Name = "AutoShape 42" };

            Wps.NonVisualDrawingShapeProperties nonVisualDrawingShapeProperties1 = new Wps.NonVisualDrawingShapeProperties();
            A.ShapeLocks shapeLocks1 = new A.ShapeLocks() { NoChangeArrowheads = true };

            nonVisualDrawingShapeProperties1.Append(shapeLocks1);

            Wps.ShapeProperties shapeProperties1 = new Wps.ShapeProperties() { BlackWhiteMode = A.BlackWhiteModeValues.Auto };

            A.Transform2D transform2D1 = new A.Transform2D() { Rotation = -5400000 };
            A.Offset offset2 = new A.Offset() { X = 859L, Y = 415L };
            A.Extents extents2 = new A.Extents() { Cx = 374L, Cy = 864L };

            transform2D1.Append(offset2);
            transform2D1.Append(extents2);

            A.PresetGeometry presetGeometry1 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.RoundRectangle };

            A.AdjustValueList adjustValueList1 = new A.AdjustValueList();
            A.ShapeGuide shapeGuide1 = new A.ShapeGuide() { Name = "adj", Formula = "val 16667" };

            adjustValueList1.Append(shapeGuide1);

            presetGeometry1.Append(adjustValueList1);

            A.SolidFill solidFill1 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex1 = new A.RgbColorModelHex() { Val = "FFFFFF" };

            solidFill1.Append(rgbColorModelHex1);

            A.Outline outline1 = new A.Outline() { Width = 9525 };

            A.SolidFill solidFill2 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex2 = new A.RgbColorModelHex() { Val = "E4BE84" };

            solidFill2.Append(rgbColorModelHex2);
            A.Round round1 = new A.Round();
            A.HeadEnd headEnd1 = new A.HeadEnd();
            A.TailEnd tailEnd1 = new A.TailEnd();

            outline1.Append(solidFill2);
            outline1.Append(round1);
            outline1.Append(headEnd1);
            outline1.Append(tailEnd1);

            shapeProperties1.Append(transform2D1);
            shapeProperties1.Append(presetGeometry1);
            shapeProperties1.Append(solidFill1);
            shapeProperties1.Append(outline1);

            Wps.TextBodyProperties textBodyProperties1 = new Wps.TextBodyProperties() { Rotation = 0, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, LeftInset = 91440, TopInset = 45720, RightInset = 91440, BottomInset = 45720, Anchor = A.TextAnchoringTypeValues.Top, AnchorCenter = false, UpRight = true };
            A.NoAutoFit noAutoFit1 = new A.NoAutoFit();

            textBodyProperties1.Append(noAutoFit1);

            wordprocessingShape1.Append(nonVisualDrawingProperties1);
            wordprocessingShape1.Append(nonVisualDrawingShapeProperties1);
            wordprocessingShape1.Append(shapeProperties1);
            wordprocessingShape1.Append(textBodyProperties1);

            Wps.WordprocessingShape wordprocessingShape2 = new Wps.WordprocessingShape();
            Wps.NonVisualDrawingProperties nonVisualDrawingProperties2 = new Wps.NonVisualDrawingProperties() { Id = (UInt32Value)24U, Name = "AutoShape 43" };

            Wps.NonVisualDrawingShapeProperties nonVisualDrawingShapeProperties2 = new Wps.NonVisualDrawingShapeProperties();
            A.ShapeLocks shapeLocks2 = new A.ShapeLocks() { NoChangeArrowheads = true };

            nonVisualDrawingShapeProperties2.Append(shapeLocks2);

            Wps.ShapeProperties shapeProperties2 = new Wps.ShapeProperties() { BlackWhiteMode = A.BlackWhiteModeValues.Auto };

            A.Transform2D transform2D2 = new A.Transform2D() { Rotation = -5400000 };
            A.Offset offset3 = new A.Offset() { X = 898L, Y = 451L };
            A.Extents extents3 = new A.Extents() { Cx = 296L, Cy = 792L };

            transform2D2.Append(offset3);
            transform2D2.Append(extents3);

            A.PresetGeometry presetGeometry2 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.RoundRectangle };

            A.AdjustValueList adjustValueList2 = new A.AdjustValueList();
            A.ShapeGuide shapeGuide2 = new A.ShapeGuide() { Name = "adj", Formula = "val 16667" };

            adjustValueList2.Append(shapeGuide2);

            presetGeometry2.Append(adjustValueList2);

            A.SolidFill solidFill3 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex3 = new A.RgbColorModelHex() { Val = "E4BE84" };

            solidFill3.Append(rgbColorModelHex3);

            A.Outline outline2 = new A.Outline() { Width = 9525 };

            A.SolidFill solidFill4 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex4 = new A.RgbColorModelHex() { Val = "E4BE84" };

            solidFill4.Append(rgbColorModelHex4);
            A.Round round2 = new A.Round();
            A.HeadEnd headEnd2 = new A.HeadEnd();
            A.TailEnd tailEnd2 = new A.TailEnd();

            outline2.Append(solidFill4);
            outline2.Append(round2);
            outline2.Append(headEnd2);
            outline2.Append(tailEnd2);

            shapeProperties2.Append(transform2D2);
            shapeProperties2.Append(presetGeometry2);
            shapeProperties2.Append(solidFill3);
            shapeProperties2.Append(outline2);

            Wps.TextBodyProperties textBodyProperties2 = new Wps.TextBodyProperties() { Rotation = 0, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, LeftInset = 91440, TopInset = 45720, RightInset = 91440, BottomInset = 45720, Anchor = A.TextAnchoringTypeValues.Top, AnchorCenter = false, UpRight = true };
            A.NoAutoFit noAutoFit2 = new A.NoAutoFit();

            textBodyProperties2.Append(noAutoFit2);

            wordprocessingShape2.Append(nonVisualDrawingProperties2);
            wordprocessingShape2.Append(nonVisualDrawingShapeProperties2);
            wordprocessingShape2.Append(shapeProperties2);
            wordprocessingShape2.Append(textBodyProperties2);

            Wps.WordprocessingShape wordprocessingShape3 = new Wps.WordprocessingShape();
            Wps.NonVisualDrawingProperties nonVisualDrawingProperties3 = new Wps.NonVisualDrawingProperties() { Id = (UInt32Value)25U, Name = "Text Box 44" };

            Wps.NonVisualDrawingShapeProperties nonVisualDrawingShapeProperties3 = new Wps.NonVisualDrawingShapeProperties() { TextBox = true };
            A.ShapeLocks shapeLocks3 = new A.ShapeLocks() { NoChangeArrowheads = true };

            nonVisualDrawingShapeProperties3.Append(shapeLocks3);

            Wps.ShapeProperties shapeProperties3 = new Wps.ShapeProperties() { BlackWhiteMode = A.BlackWhiteModeValues.Auto };

            A.Transform2D transform2D3 = new A.Transform2D();
            A.Offset offset4 = new A.Offset() { X = 732L, Y = 716L };
            A.Extents extents4 = new A.Extents() { Cx = 659L, Cy = 288L };

            transform2D3.Append(offset4);
            transform2D3.Append(extents4);

            A.PresetGeometry presetGeometry3 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList3 = new A.AdjustValueList();

            presetGeometry3.Append(adjustValueList3);
            A.NoFill noFill1 = new A.NoFill();

            A.Outline outline3 = new A.Outline();
            A.NoFill noFill2 = new A.NoFill();

            outline3.Append(noFill2);

            A.ShapePropertiesExtensionList shapePropertiesExtensionList1 = new A.ShapePropertiesExtensionList();

            A.ShapePropertiesExtension shapePropertiesExtension1 = new A.ShapePropertiesExtension() { Uri = "{909E8E84-426E-40DD-AFC4-6F175D3DCCD1}" };

            A14.HiddenFillProperties hiddenFillProperties1 = new A14.HiddenFillProperties();
            hiddenFillProperties1.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");

            A.SolidFill solidFill5 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex5 = new A.RgbColorModelHex() { Val = "FFFFFF" };

            solidFill5.Append(rgbColorModelHex5);

            hiddenFillProperties1.Append(solidFill5);

            shapePropertiesExtension1.Append(hiddenFillProperties1);

            A.ShapePropertiesExtension shapePropertiesExtension2 = new A.ShapePropertiesExtension() { Uri = "{91240B29-F687-4F45-9708-019B960494DF}" };

            A14.HiddenLineProperties hiddenLineProperties1 = new A14.HiddenLineProperties() { Width = 9525 };
            hiddenLineProperties1.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");

            A.SolidFill solidFill6 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex6 = new A.RgbColorModelHex() { Val = "000000" };

            solidFill6.Append(rgbColorModelHex6);
            A.Miter miter1 = new A.Miter() { Limit = 800000 };
            A.HeadEnd headEnd3 = new A.HeadEnd();
            A.TailEnd tailEnd3 = new A.TailEnd();

            hiddenLineProperties1.Append(solidFill6);
            hiddenLineProperties1.Append(miter1);
            hiddenLineProperties1.Append(headEnd3);
            hiddenLineProperties1.Append(tailEnd3);

            shapePropertiesExtension2.Append(hiddenLineProperties1);

            shapePropertiesExtensionList1.Append(shapePropertiesExtension1);
            shapePropertiesExtensionList1.Append(shapePropertiesExtension2);

            shapeProperties3.Append(transform2D3);
            shapeProperties3.Append(presetGeometry3);
            shapeProperties3.Append(noFill1);
            shapeProperties3.Append(outline3);
            shapeProperties3.Append(shapePropertiesExtensionList1);

            Wps.TextBoxInfo2 textBoxInfo21 = new Wps.TextBoxInfo2();

            TextBoxContent textBoxContent1 = new TextBoxContent();

            Paragraph paragraph2 = new Paragraph() { RsidParagraphAddition = "009D383E", RsidRunAdditionDefault = "009D383E", ParagraphId = "4B129D42", TextId = "77777777" };

            ParagraphProperties paragraphProperties2 = new ParagraphProperties();
            Justification justification1 = new Justification() { Val = JustificationValues.Right };

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
            Bold bold1 = new Bold();
            BoldComplexScript boldComplexScript1 = new BoldComplexScript();
            NoProof noProof2 = new NoProof();
            Color color1 = new Color() { Val = "FFFFFF", ThemeColor = ThemeColorValues.Background1 };

            runProperties2.Append(bold1);
            runProperties2.Append(boldComplexScript1);
            runProperties2.Append(noProof2);
            runProperties2.Append(color1);
            Text text1 = new Text();
            text1.Text = "2";

            run5.Append(runProperties2);
            run5.Append(text1);

            Run run6 = new Run();

            RunProperties runProperties3 = new RunProperties();
            Bold bold2 = new Bold();
            BoldComplexScript boldComplexScript2 = new BoldComplexScript();
            NoProof noProof3 = new NoProof();
            Color color2 = new Color() { Val = "FFFFFF", ThemeColor = ThemeColorValues.Background1 };

            runProperties3.Append(bold2);
            runProperties3.Append(boldComplexScript2);
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

            Wps.TextBodyProperties textBodyProperties3 = new Wps.TextBodyProperties() { Rotation = 0, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, LeftInset = 0, TopInset = 0, RightInset = 0, BottomInset = 0, Anchor = A.TextAnchoringTypeValues.Top, AnchorCenter = false, UpRight = true };
            A.NoAutoFit noAutoFit3 = new A.NoAutoFit();

            textBodyProperties3.Append(noAutoFit3);

            wordprocessingShape3.Append(nonVisualDrawingProperties3);
            wordprocessingShape3.Append(nonVisualDrawingShapeProperties3);
            wordprocessingShape3.Append(shapeProperties3);
            wordprocessingShape3.Append(textBoxInfo21);
            wordprocessingShape3.Append(textBodyProperties3);

            wordprocessingGroup1.Append(nonVisualGroupDrawingShapeProperties1);
            wordprocessingGroup1.Append(groupShapeProperties1);
            wordprocessingGroup1.Append(wordprocessingShape1);
            wordprocessingGroup1.Append(wordprocessingShape2);
            wordprocessingGroup1.Append(wordprocessingShape3);

            graphicData1.Append(wordprocessingGroup1);

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

            V.Group group1 = new V.Group() { Id = "Group 22", Style = "width:43.2pt;height:18.7pt;mso-position-horizontal-relative:char;mso-position-vertical-relative:line", CoordinateSize = "864,374", CoordinateOrigin = "614,660", OptionalString = "_x0000_s1026" };
            group1.SetAttribute(new OpenXmlAttribute("o", "gfxdata", "urn:schemas-microsoft-com:office:office", "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\n"));

            V.RoundRectangle roundRectangle1 = new V.RoundRectangle() { Id = "AutoShape 42", Style = "position:absolute;left:859;top:415;width:374;height:864;rotation:-90;visibility:visible;mso-wrap-style:square;v-text-anchor:top", OptionalString = "_x0000_s1027", StrokeColor = "#e4be84", ArcSize = "10923f" };
            roundRectangle1.SetAttribute(new OpenXmlAttribute("o", "gfxdata", "urn:schemas-microsoft-com:office:office", "UEsDBBQABgAIAAAAIQDb4fbL7gAAAIUBAAATAAAAW0NvbnRlbnRfVHlwZXNdLnhtbHyQz07DMAyH\n70i8Q+QralM4IITa7kDhCAiNB7ASt43WOlEcyvb2pNu4IODoPz9/n1xv9vOkForiPDdwXVagiI23\njocG3rdPxR0oScgWJ8/UwIEENu3lRb09BBKV0ywNjCmFe63FjDSjlD4Q50nv44wpl3HQAc0OB9I3\nVXWrjedEnIq03oC27qjHjympx31un0wiTQLq4bS4shrAECZnMGVTvbD9QSnOhDInjzsyuiBXWQP0\nr4R18jfgnHvJr4nOknrFmJ5xzhraRtHWf3Kkpfz/yGo5S+H73hkquyhdjr3R8m2lj09svwAAAP//\nAwBQSwMEFAAGAAgAAAAhAFr0LFu/AAAAFQEAAAsAAABfcmVscy8ucmVsc2zPwWrDMAwG4Ptg72B0\nX5TuUMaI01uh19I+gLGVxCy2jGSy9e1nemrHjpL4P0nD4SetZiPRyNnCruvBUPYcYp4tXC/Htw8w\nWl0ObuVMFm6kcBhfX4Yzra62kC6xqGlKVgtLreUTUf1CyWnHhXKbTCzJ1VbKjMX5LzcTvvf9HuXR\ngPHJNKdgQU5hB+ZyK23zHztFL6w81c5zQp6m6P9TMfB3PtPWFCczVQtB9N4U2rp2HOA44NMz4y8A\nAAD//wMAUEsDBBQABgAIAAAAIQCr/nuhxAAAANsAAAAPAAAAZHJzL2Rvd25yZXYueG1sRI9Ba4NA\nFITvhfyH5QVykbjWQgnGTQgBoYeA1PbQ48N9UYn7VtyNmv76bqHQ4zAz3zD5cTG9mGh0nWUFz3EC\ngri2uuNGwedHsd2BcB5ZY2+ZFDzIwfGwesox03bmd5oq34gAYZehgtb7IZPS1S0ZdLEdiIN3taNB\nH+TYSD3iHOCml2mSvEqDHYeFFgc6t1TfqrtRoNPHTkZl0X9HRTndv3x1mYtKqc16Oe1BeFr8f/iv\n/aYVpC/w+yX8AHn4AQAA//8DAFBLAQItABQABgAIAAAAIQDb4fbL7gAAAIUBAAATAAAAAAAAAAAA\nAAAAAAAAAABbQ29udGVudF9UeXBlc10ueG1sUEsBAi0AFAAGAAgAAAAhAFr0LFu/AAAAFQEAAAsA\nAAAAAAAAAAAAAAAAHwEAAF9yZWxzLy5yZWxzUEsBAi0AFAAGAAgAAAAhAKv+e6HEAAAA2wAAAA8A\nAAAAAAAAAAAAAAAABwIAAGRycy9kb3ducmV2LnhtbFBLBQYAAAAAAwADALcAAAD4AgAAAAA=\n"));

            V.RoundRectangle roundRectangle2 = new V.RoundRectangle() { Id = "AutoShape 43", Style = "position:absolute;left:898;top:451;width:296;height:792;rotation:-90;visibility:visible;mso-wrap-style:square;v-text-anchor:top", OptionalString = "_x0000_s1028", FillColor = "#e4be84", StrokeColor = "#e4be84", ArcSize = "10923f" };
            roundRectangle2.SetAttribute(new OpenXmlAttribute("o", "gfxdata", "urn:schemas-microsoft-com:office:office", "UEsDBBQABgAIAAAAIQDb4fbL7gAAAIUBAAATAAAAW0NvbnRlbnRfVHlwZXNdLnhtbHyQz07DMAyH\n70i8Q+QralM4IITa7kDhCAiNB7ASt43WOlEcyvb2pNu4IODoPz9/n1xv9vOkForiPDdwXVagiI23\njocG3rdPxR0oScgWJ8/UwIEENu3lRb09BBKV0ywNjCmFe63FjDSjlD4Q50nv44wpl3HQAc0OB9I3\nVXWrjedEnIq03oC27qjHjympx31un0wiTQLq4bS4shrAECZnMGVTvbD9QSnOhDInjzsyuiBXWQP0\nr4R18jfgnHvJr4nOknrFmJ5xzhraRtHWf3Kkpfz/yGo5S+H73hkquyhdjr3R8m2lj09svwAAAP//\nAwBQSwMEFAAGAAgAAAAhAFr0LFu/AAAAFQEAAAsAAABfcmVscy8ucmVsc2zPwWrDMAwG4Ptg72B0\nX5TuUMaI01uh19I+gLGVxCy2jGSy9e1nemrHjpL4P0nD4SetZiPRyNnCruvBUPYcYp4tXC/Htw8w\nWl0ObuVMFm6kcBhfX4Yzra62kC6xqGlKVgtLreUTUf1CyWnHhXKbTCzJ1VbKjMX5LzcTvvf9HuXR\ngPHJNKdgQU5hB+ZyK23zHztFL6w81c5zQp6m6P9TMfB3PtPWFCczVQtB9N4U2rp2HOA44NMz4y8A\nAAD//wMAUEsDBBQABgAIAAAAIQATn8OnxQAAANsAAAAPAAAAZHJzL2Rvd25yZXYueG1sRI9BawIx\nFITvhf6H8ApepJtV2lq2RhFB8FbUInt83Tw3q5uXJYm69dc3BaHHYWa+Yabz3rbiQj40jhWMshwE\nceV0w7WCr93q+R1EiMgaW8ek4IcCzGePD1MstLvyhi7bWIsE4VCgAhNjV0gZKkMWQ+Y64uQdnLcY\nk/S11B6vCW5bOc7zN2mx4bRgsKOloeq0PVsFn6Usl6/l92SzyP3tMNrfaGiOSg2e+sUHiEh9/A/f\n22utYPwCf1/SD5CzXwAAAP//AwBQSwECLQAUAAYACAAAACEA2+H2y+4AAACFAQAAEwAAAAAAAAAA\nAAAAAAAAAAAAW0NvbnRlbnRfVHlwZXNdLnhtbFBLAQItABQABgAIAAAAIQBa9CxbvwAAABUBAAAL\nAAAAAAAAAAAAAAAAAB8BAABfcmVscy8ucmVsc1BLAQItABQABgAIAAAAIQATn8OnxQAAANsAAAAP\nAAAAAAAAAAAAAAAAAAcCAABkcnMvZG93bnJldi54bWxQSwUGAAAAAAMAAwC3AAAA+QIAAAAA\n"));

            V.Shapetype shapetype1 = new V.Shapetype() { Id = "_x0000_t202", CoordinateSize = "21600,21600", OptionalNumber = 202, EdgePath = "m,l,21600r21600,l21600,xe" };
            V.Stroke stroke1 = new V.Stroke() { JoinStyle = V.StrokeJoinStyleValues.Miter };
            V.Path path1 = new V.Path() { AllowGradientShape = true, ConnectionPointType = Ovml.ConnectValues.Rectangle };

            shapetype1.Append(stroke1);
            shapetype1.Append(path1);

            V.Shape shape1 = new V.Shape() { Id = "Text Box 44", Style = "position:absolute;left:732;top:716;width:659;height:288;visibility:visible;mso-wrap-style:square;v-text-anchor:top", OptionalString = "_x0000_s1029", Filled = false, Stroked = false, Type = "#_x0000_t202", EncodedPackage = "UEsDBBQABgAIAAAAIQDb4fbL7gAAAIUBAAATAAAAW0NvbnRlbnRfVHlwZXNdLnhtbHyQz07DMAyH\n70i8Q+QralM4IITa7kDhCAiNB7ASt43WOlEcyvb2pNu4IODoPz9/n1xv9vOkForiPDdwXVagiI23\njocG3rdPxR0oScgWJ8/UwIEENu3lRb09BBKV0ywNjCmFe63FjDSjlD4Q50nv44wpl3HQAc0OB9I3\nVXWrjedEnIq03oC27qjHjympx31un0wiTQLq4bS4shrAECZnMGVTvbD9QSnOhDInjzsyuiBXWQP0\nr4R18jfgnHvJr4nOknrFmJ5xzhraRtHWf3Kkpfz/yGo5S+H73hkquyhdjr3R8m2lj09svwAAAP//\nAwBQSwMEFAAGAAgAAAAhAFr0LFu/AAAAFQEAAAsAAABfcmVscy8ucmVsc2zPwWrDMAwG4Ptg72B0\nX5TuUMaI01uh19I+gLGVxCy2jGSy9e1nemrHjpL4P0nD4SetZiPRyNnCruvBUPYcYp4tXC/Htw8w\nWl0ObuVMFm6kcBhfX4Yzra62kC6xqGlKVgtLreUTUf1CyWnHhXKbTCzJ1VbKjMX5LzcTvvf9HuXR\ngPHJNKdgQU5hB+ZyK23zHztFL6w81c5zQp6m6P9TMfB3PtPWFCczVQtB9N4U2rp2HOA44NMz4y8A\nAAD//wMAUEsDBBQABgAIAAAAIQDLIsuSwwAAANsAAAAPAAAAZHJzL2Rvd25yZXYueG1sRI9Ba8JA\nFITvQv/D8gredKOg2OgqUhQKQjGmB4/P7DNZzL5Ns1uN/74rCB6HmfmGWaw6W4srtd44VjAaJiCI\nC6cNlwp+8u1gBsIHZI21Y1JwJw+r5Vtvgal2N87oegiliBD2KSqoQmhSKX1RkUU/dA1x9M6utRii\nbEupW7xFuK3lOEmm0qLhuFBhQ58VFZfDn1WwPnK2Mb/fp312zkyefyS8m16U6r936zmIQF14hZ/t\nL61gPIHHl/gD5PIfAAD//wMAUEsBAi0AFAAGAAgAAAAhANvh9svuAAAAhQEAABMAAAAAAAAAAAAA\nAAAAAAAAAFtDb250ZW50X1R5cGVzXS54bWxQSwECLQAUAAYACAAAACEAWvQsW78AAAAVAQAACwAA\nAAAAAAAAAAAAAAAfAQAAX3JlbHMvLnJlbHNQSwECLQAUAAYACAAAACEAyyLLksMAAADbAAAADwAA\nAAAAAAAAAAAAAAAHAgAAZHJzL2Rvd25yZXYueG1sUEsFBgAAAAADAAMAtwAAAPcCAAAAAA==\n" };

            V.TextBox textBox1 = new V.TextBox() { Inset = "0,0,0,0" };

            TextBoxContent textBoxContent2 = new TextBoxContent();

            Paragraph paragraph3 = new Paragraph() { RsidParagraphAddition = "009D383E", RsidRunAdditionDefault = "009D383E", ParagraphId = "4B129D42", TextId = "77777777" };

            ParagraphProperties paragraphProperties3 = new ParagraphProperties();
            Justification justification2 = new Justification() { Val = JustificationValues.Right };

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
            Bold bold3 = new Bold();
            BoldComplexScript boldComplexScript3 = new BoldComplexScript();
            NoProof noProof4 = new NoProof();
            Color color3 = new Color() { Val = "FFFFFF", ThemeColor = ThemeColorValues.Background1 };

            runProperties4.Append(bold3);
            runProperties4.Append(boldComplexScript3);
            runProperties4.Append(noProof4);
            runProperties4.Append(color3);
            Text text2 = new Text();
            text2.Text = "2";

            run10.Append(runProperties4);
            run10.Append(text2);

            Run run11 = new Run();

            RunProperties runProperties5 = new RunProperties();
            Bold bold4 = new Bold();
            BoldComplexScript boldComplexScript4 = new BoldComplexScript();
            NoProof noProof5 = new NoProof();
            Color color4 = new Color() { Val = "FFFFFF", ThemeColor = ThemeColorValues.Background1 };

            runProperties5.Append(bold4);
            runProperties5.Append(boldComplexScript4);
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
            Wvml.AnchorLock anchorLock1 = new Wvml.AnchorLock();

            group1.Append(roundRectangle1);
            group1.Append(roundRectangle2);
            group1.Append(shapetype1);
            group1.Append(shape1);
            group1.Append(anchorLock1);

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
