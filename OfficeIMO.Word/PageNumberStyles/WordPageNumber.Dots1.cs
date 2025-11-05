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
    private static SdtBlock Dots1 {
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

            Paragraph paragraph1 = new Paragraph() { RsidParagraphAddition = "004918B9", RsidRunAdditionDefault = "004918B9", ParagraphId = "00BD54B3", TextId = "2A9D606F" };

            ParagraphProperties paragraphProperties1 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId1 = new ParagraphStyleId() { Val = "Header" };
            Justification justification1 = new Justification() { Val = JustificationValues.Center };

            paragraphProperties1.Append(paragraphStyleId1);
            paragraphProperties1.Append(justification1);

            Run run1 = new Run();

            RunProperties runProperties1 = new RunProperties();
            NoProof noProof1 = new NoProof();

            runProperties1.Append(noProof1);

            AlternateContent alternateContent1 = new AlternateContent();

            AlternateContentChoice alternateContentChoice1 = new AlternateContentChoice() { Requires = "wpg" };

            Drawing drawing1 = new Drawing();

            Wp.Inline inline1 = new Wp.Inline() { DistanceFromTop = (UInt32Value)0U, DistanceFromBottom = (UInt32Value)0U, DistanceFromLeft = (UInt32Value)0U, DistanceFromRight = (UInt32Value)0U, AnchorId = "34BA9F6C", EditId = "62CF8404" };
            Wp.Extent extent1 = new Wp.Extent() { Cx = 418465L, Cy = 221615L };
            Wp.EffectExtent effectExtent1 = new Wp.EffectExtent() { LeftEdge = 0L, TopEdge = 0L, RightEdge = 635L, BottomEdge = 0L };
            Wp.DocProperties docProperties1 = new Wp.DocProperties() { Id = (UInt32Value)4U, Name = "Group 4" };

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
            A.Extents extents1 = new A.Extents() { Cx = 418465L, Cy = 221615L };
            A.ChildOffset childOffset1 = new A.ChildOffset() { X = 5351L, Y = 739L };
            A.ChildExtents childExtents1 = new A.ChildExtents() { Cx = 659L, Cy = 349L };

            transformGroup1.Append(offset1);
            transformGroup1.Append(extents1);
            transformGroup1.Append(childOffset1);
            transformGroup1.Append(childExtents1);

            groupShapeProperties1.Append(transformGroup1);

            Wps.WordprocessingShape wordprocessingShape1 = new Wps.WordprocessingShape();
            Wps.NonVisualDrawingProperties nonVisualDrawingProperties1 = new Wps.NonVisualDrawingProperties() { Id = (UInt32Value)5U, Name = "Text Box 56" };

            Wps.NonVisualDrawingShapeProperties nonVisualDrawingShapeProperties1 = new Wps.NonVisualDrawingShapeProperties() { TextBox = true };
            A.ShapeLocks shapeLocks1 = new A.ShapeLocks() { NoChangeArrowheads = true };

            nonVisualDrawingShapeProperties1.Append(shapeLocks1);

            Wps.ShapeProperties shapeProperties1 = new Wps.ShapeProperties() { BlackWhiteMode = A.BlackWhiteModeValues.Auto };

            A.Transform2D transform2D1 = new A.Transform2D();
            A.Offset offset2 = new A.Offset() { X = 5351L, Y = 800L };
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

            Paragraph paragraph2 = new Paragraph() { RsidParagraphAddition = "004918B9", RsidRunAdditionDefault = "004918B9", ParagraphId = "12772B1B", TextId = "77777777" };

            ParagraphProperties paragraphProperties2 = new ParagraphProperties();
            Justification justification2 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties1 = new ParagraphMarkRunProperties();
            FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "18" };

            paragraphMarkRunProperties1.Append(fontSizeComplexScript1);

            paragraphProperties2.Append(justification2);
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
            Italic italic1 = new Italic();
            ItalicComplexScript italicComplexScript1 = new ItalicComplexScript();
            NoProof noProof2 = new NoProof();
            FontSize fontSize1 = new FontSize() { Val = "18" };
            FontSizeComplexScript fontSizeComplexScript2 = new FontSizeComplexScript() { Val = "18" };

            runProperties2.Append(italic1);
            runProperties2.Append(italicComplexScript1);
            runProperties2.Append(noProof2);
            runProperties2.Append(fontSize1);
            runProperties2.Append(fontSizeComplexScript2);
            Text text1 = new Text();
            text1.Text = "2";

            run5.Append(runProperties2);
            run5.Append(text1);

            Run run6 = new Run();

            RunProperties runProperties3 = new RunProperties();
            Italic italic2 = new Italic();
            ItalicComplexScript italicComplexScript2 = new ItalicComplexScript();
            NoProof noProof3 = new NoProof();
            FontSize fontSize2 = new FontSize() { Val = "18" };
            FontSizeComplexScript fontSizeComplexScript3 = new FontSizeComplexScript() { Val = "18" };

            runProperties3.Append(italic2);
            runProperties3.Append(italicComplexScript2);
            runProperties3.Append(noProof3);
            runProperties3.Append(fontSize2);
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

            Wps.TextBodyProperties textBodyProperties1 = new Wps.TextBodyProperties() { Rotation = 0, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, LeftInset = 0, TopInset = 0, RightInset = 0, BottomInset = 0, Anchor = A.TextAnchoringTypeValues.Top, AnchorCenter = false, UpRight = true };
            A.NoAutoFit noAutoFit1 = new A.NoAutoFit();

            textBodyProperties1.Append(noAutoFit1);

            wordprocessingShape1.Append(nonVisualDrawingProperties1);
            wordprocessingShape1.Append(nonVisualDrawingShapeProperties1);
            wordprocessingShape1.Append(shapeProperties1);
            wordprocessingShape1.Append(textBoxInfo21);
            wordprocessingShape1.Append(textBodyProperties1);

            Wpg.GroupShape groupShape1 = new Wpg.GroupShape();
            Wpg.NonVisualDrawingProperties nonVisualDrawingProperties2 = new Wpg.NonVisualDrawingProperties() { Id = (UInt32Value)6U, Name = "Group 57" };

            Wpg.NonVisualGroupDrawingShapeProperties nonVisualGroupDrawingShapeProperties2 = new Wpg.NonVisualGroupDrawingShapeProperties();
            A.GroupShapeLocks groupShapeLocks2 = new A.GroupShapeLocks();

            nonVisualGroupDrawingShapeProperties2.Append(groupShapeLocks2);

            Wpg.GroupShapeProperties groupShapeProperties2 = new Wpg.GroupShapeProperties() { BlackWhiteMode = A.BlackWhiteModeValues.Auto };

            A.TransformGroup transformGroup2 = new A.TransformGroup();
            A.Offset offset3 = new A.Offset() { X = 5494L, Y = 739L };
            A.Extents extents3 = new A.Extents() { Cx = 372L, Cy = 72L };
            A.ChildOffset childOffset2 = new A.ChildOffset() { X = 5486L, Y = 739L };
            A.ChildExtents childExtents2 = new A.ChildExtents() { Cx = 372L, Cy = 72L };

            transformGroup2.Append(offset3);
            transformGroup2.Append(extents3);
            transformGroup2.Append(childOffset2);
            transformGroup2.Append(childExtents2);

            groupShapeProperties2.Append(transformGroup2);

            Wps.WordprocessingShape wordprocessingShape2 = new Wps.WordprocessingShape();
            Wps.NonVisualDrawingProperties nonVisualDrawingProperties3 = new Wps.NonVisualDrawingProperties() { Id = (UInt32Value)7U, Name = "Oval 58" };

            Wps.NonVisualDrawingShapeProperties nonVisualDrawingShapeProperties2 = new Wps.NonVisualDrawingShapeProperties();
            A.ShapeLocks shapeLocks2 = new A.ShapeLocks() { NoChangeArrowheads = true };

            nonVisualDrawingShapeProperties2.Append(shapeLocks2);

            Wps.ShapeProperties shapeProperties2 = new Wps.ShapeProperties() { BlackWhiteMode = A.BlackWhiteModeValues.Auto };

            A.Transform2D transform2D2 = new A.Transform2D();
            A.Offset offset4 = new A.Offset() { X = 5486L, Y = 739L };
            A.Extents extents4 = new A.Extents() { Cx = 72L, Cy = 72L };

            transform2D2.Append(offset4);
            transform2D2.Append(extents4);

            A.PresetGeometry presetGeometry2 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Ellipse };
            A.AdjustValueList adjustValueList2 = new A.AdjustValueList();

            presetGeometry2.Append(adjustValueList2);

            A.SolidFill solidFill3 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex3 = new A.RgbColorModelHex() { Val = "84A2C6" };

            solidFill3.Append(rgbColorModelHex3);

            A.Outline outline2 = new A.Outline();
            A.NoFill noFill3 = new A.NoFill();

            outline2.Append(noFill3);

            A.ShapePropertiesExtensionList shapePropertiesExtensionList2 = new A.ShapePropertiesExtensionList();

            A.ShapePropertiesExtension shapePropertiesExtension3 = new A.ShapePropertiesExtension() { Uri = "{91240B29-F687-4F45-9708-019B960494DF}" };

            A14.HiddenLineProperties hiddenLineProperties2 = new A14.HiddenLineProperties() { Width = 9525 };
            hiddenLineProperties2.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");

            A.SolidFill solidFill4 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex4 = new A.RgbColorModelHex() { Val = "000000" };

            solidFill4.Append(rgbColorModelHex4);
            A.Round round1 = new A.Round();
            A.HeadEnd headEnd2 = new A.HeadEnd();
            A.TailEnd tailEnd2 = new A.TailEnd();

            hiddenLineProperties2.Append(solidFill4);
            hiddenLineProperties2.Append(round1);
            hiddenLineProperties2.Append(headEnd2);
            hiddenLineProperties2.Append(tailEnd2);

            shapePropertiesExtension3.Append(hiddenLineProperties2);

            shapePropertiesExtensionList2.Append(shapePropertiesExtension3);

            shapeProperties2.Append(transform2D2);
            shapeProperties2.Append(presetGeometry2);
            shapeProperties2.Append(solidFill3);
            shapeProperties2.Append(outline2);
            shapeProperties2.Append(shapePropertiesExtensionList2);

            Wps.TextBodyProperties textBodyProperties2 = new Wps.TextBodyProperties() { Rotation = 0, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, LeftInset = 91440, TopInset = 45720, RightInset = 91440, BottomInset = 45720, Anchor = A.TextAnchoringTypeValues.Top, AnchorCenter = false, UpRight = true };
            A.NoAutoFit noAutoFit2 = new A.NoAutoFit();

            textBodyProperties2.Append(noAutoFit2);

            wordprocessingShape2.Append(nonVisualDrawingProperties3);
            wordprocessingShape2.Append(nonVisualDrawingShapeProperties2);
            wordprocessingShape2.Append(shapeProperties2);
            wordprocessingShape2.Append(textBodyProperties2);

            Wps.WordprocessingShape wordprocessingShape3 = new Wps.WordprocessingShape();
            Wps.NonVisualDrawingProperties nonVisualDrawingProperties4 = new Wps.NonVisualDrawingProperties() { Id = (UInt32Value)8U, Name = "Oval 59" };

            Wps.NonVisualDrawingShapeProperties nonVisualDrawingShapeProperties3 = new Wps.NonVisualDrawingShapeProperties();
            A.ShapeLocks shapeLocks3 = new A.ShapeLocks() { NoChangeArrowheads = true };

            nonVisualDrawingShapeProperties3.Append(shapeLocks3);

            Wps.ShapeProperties shapeProperties3 = new Wps.ShapeProperties() { BlackWhiteMode = A.BlackWhiteModeValues.Auto };

            A.Transform2D transform2D3 = new A.Transform2D();
            A.Offset offset5 = new A.Offset() { X = 5636L, Y = 739L };
            A.Extents extents5 = new A.Extents() { Cx = 72L, Cy = 72L };

            transform2D3.Append(offset5);
            transform2D3.Append(extents5);

            A.PresetGeometry presetGeometry3 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Ellipse };
            A.AdjustValueList adjustValueList3 = new A.AdjustValueList();

            presetGeometry3.Append(adjustValueList3);

            A.SolidFill solidFill5 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex5 = new A.RgbColorModelHex() { Val = "84A2C6" };

            solidFill5.Append(rgbColorModelHex5);

            A.Outline outline3 = new A.Outline();
            A.NoFill noFill4 = new A.NoFill();

            outline3.Append(noFill4);

            A.ShapePropertiesExtensionList shapePropertiesExtensionList3 = new A.ShapePropertiesExtensionList();

            A.ShapePropertiesExtension shapePropertiesExtension4 = new A.ShapePropertiesExtension() { Uri = "{91240B29-F687-4F45-9708-019B960494DF}" };

            A14.HiddenLineProperties hiddenLineProperties3 = new A14.HiddenLineProperties() { Width = 9525 };
            hiddenLineProperties3.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");

            A.SolidFill solidFill6 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex6 = new A.RgbColorModelHex() { Val = "000000" };

            solidFill6.Append(rgbColorModelHex6);
            A.Round round2 = new A.Round();
            A.HeadEnd headEnd3 = new A.HeadEnd();
            A.TailEnd tailEnd3 = new A.TailEnd();

            hiddenLineProperties3.Append(solidFill6);
            hiddenLineProperties3.Append(round2);
            hiddenLineProperties3.Append(headEnd3);
            hiddenLineProperties3.Append(tailEnd3);

            shapePropertiesExtension4.Append(hiddenLineProperties3);

            shapePropertiesExtensionList3.Append(shapePropertiesExtension4);

            shapeProperties3.Append(transform2D3);
            shapeProperties3.Append(presetGeometry3);
            shapeProperties3.Append(solidFill5);
            shapeProperties3.Append(outline3);
            shapeProperties3.Append(shapePropertiesExtensionList3);

            Wps.TextBodyProperties textBodyProperties3 = new Wps.TextBodyProperties() { Rotation = 0, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, LeftInset = 91440, TopInset = 45720, RightInset = 91440, BottomInset = 45720, Anchor = A.TextAnchoringTypeValues.Top, AnchorCenter = false, UpRight = true };
            A.NoAutoFit noAutoFit3 = new A.NoAutoFit();

            textBodyProperties3.Append(noAutoFit3);

            wordprocessingShape3.Append(nonVisualDrawingProperties4);
            wordprocessingShape3.Append(nonVisualDrawingShapeProperties3);
            wordprocessingShape3.Append(shapeProperties3);
            wordprocessingShape3.Append(textBodyProperties3);

            Wps.WordprocessingShape wordprocessingShape4 = new Wps.WordprocessingShape();
            Wps.NonVisualDrawingProperties nonVisualDrawingProperties5 = new Wps.NonVisualDrawingProperties() { Id = (UInt32Value)9U, Name = "Oval 60" };

            Wps.NonVisualDrawingShapeProperties nonVisualDrawingShapeProperties4 = new Wps.NonVisualDrawingShapeProperties();
            A.ShapeLocks shapeLocks4 = new A.ShapeLocks() { NoChangeArrowheads = true };

            nonVisualDrawingShapeProperties4.Append(shapeLocks4);

            Wps.ShapeProperties shapeProperties4 = new Wps.ShapeProperties() { BlackWhiteMode = A.BlackWhiteModeValues.Auto };

            A.Transform2D transform2D4 = new A.Transform2D();
            A.Offset offset6 = new A.Offset() { X = 5786L, Y = 739L };
            A.Extents extents6 = new A.Extents() { Cx = 72L, Cy = 72L };

            transform2D4.Append(offset6);
            transform2D4.Append(extents6);

            A.PresetGeometry presetGeometry4 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Ellipse };
            A.AdjustValueList adjustValueList4 = new A.AdjustValueList();

            presetGeometry4.Append(adjustValueList4);

            A.SolidFill solidFill7 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex7 = new A.RgbColorModelHex() { Val = "84A2C6" };

            solidFill7.Append(rgbColorModelHex7);

            A.Outline outline4 = new A.Outline();
            A.NoFill noFill5 = new A.NoFill();

            outline4.Append(noFill5);

            A.ShapePropertiesExtensionList shapePropertiesExtensionList4 = new A.ShapePropertiesExtensionList();

            A.ShapePropertiesExtension shapePropertiesExtension5 = new A.ShapePropertiesExtension() { Uri = "{91240B29-F687-4F45-9708-019B960494DF}" };

            A14.HiddenLineProperties hiddenLineProperties4 = new A14.HiddenLineProperties() { Width = 9525 };
            hiddenLineProperties4.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");

            A.SolidFill solidFill8 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex8 = new A.RgbColorModelHex() { Val = "000000" };

            solidFill8.Append(rgbColorModelHex8);
            A.Round round3 = new A.Round();
            A.HeadEnd headEnd4 = new A.HeadEnd();
            A.TailEnd tailEnd4 = new A.TailEnd();

            hiddenLineProperties4.Append(solidFill8);
            hiddenLineProperties4.Append(round3);
            hiddenLineProperties4.Append(headEnd4);
            hiddenLineProperties4.Append(tailEnd4);

            shapePropertiesExtension5.Append(hiddenLineProperties4);

            shapePropertiesExtensionList4.Append(shapePropertiesExtension5);

            shapeProperties4.Append(transform2D4);
            shapeProperties4.Append(presetGeometry4);
            shapeProperties4.Append(solidFill7);
            shapeProperties4.Append(outline4);
            shapeProperties4.Append(shapePropertiesExtensionList4);

            Wps.TextBodyProperties textBodyProperties4 = new Wps.TextBodyProperties() { Rotation = 0, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, LeftInset = 91440, TopInset = 45720, RightInset = 91440, BottomInset = 45720, Anchor = A.TextAnchoringTypeValues.Top, AnchorCenter = false, UpRight = true };
            A.NoAutoFit noAutoFit4 = new A.NoAutoFit();

            textBodyProperties4.Append(noAutoFit4);

            wordprocessingShape4.Append(nonVisualDrawingProperties5);
            wordprocessingShape4.Append(nonVisualDrawingShapeProperties4);
            wordprocessingShape4.Append(shapeProperties4);
            wordprocessingShape4.Append(textBodyProperties4);

            groupShape1.Append(nonVisualDrawingProperties2);
            groupShape1.Append(nonVisualGroupDrawingShapeProperties2);
            groupShape1.Append(groupShapeProperties2);
            groupShape1.Append(wordprocessingShape2);
            groupShape1.Append(wordprocessingShape3);
            groupShape1.Append(wordprocessingShape4);

            wordprocessingGroup1.Append(nonVisualGroupDrawingShapeProperties1);
            wordprocessingGroup1.Append(groupShapeProperties1);
            wordprocessingGroup1.Append(wordprocessingShape1);
            wordprocessingGroup1.Append(groupShape1);

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

            V.Group group1 = new V.Group() { Id = "Group 4", Style = "width:32.95pt;height:17.45pt;mso-position-horizontal-relative:char;mso-position-vertical-relative:line", CoordinateSize = "659,349", CoordinateOrigin = "5351,739", OptionalString = "_x0000_s1026" };
            group1.SetAttribute(new OpenXmlAttribute("o", "gfxdata", "urn:schemas-microsoft-com:office:office", "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\n"));

            V.Shapetype shapetype1 = new V.Shapetype() { Id = "_x0000_t202", CoordinateSize = "21600,21600", OptionalNumber = 202, EdgePath = "m,l,21600r21600,l21600,xe" };
            V.Stroke stroke1 = new V.Stroke() { JoinStyle = V.StrokeJoinStyleValues.Miter };
            V.Path path1 = new V.Path() { AllowGradientShape = true, ConnectionPointType = Ovml.ConnectValues.Rectangle };

            shapetype1.Append(stroke1);
            shapetype1.Append(path1);

            V.Shape shape1 = new V.Shape() { Id = "Text Box 56", Style = "position:absolute;left:5351;top:800;width:659;height:288;visibility:visible;mso-wrap-style:square;v-text-anchor:top", OptionalString = "_x0000_s1027", Filled = false, Stroked = false, Type = "#_x0000_t202", EncodedPackage = "UEsDBBQABgAIAAAAIQDb4fbL7gAAAIUBAAATAAAAW0NvbnRlbnRfVHlwZXNdLnhtbHyQz07DMAyH\n70i8Q+QralM4IITa7kDhCAiNB7ASt43WOlEcyvb2pNu4IODoPz9/n1xv9vOkForiPDdwXVagiI23\njocG3rdPxR0oScgWJ8/UwIEENu3lRb09BBKV0ywNjCmFe63FjDSjlD4Q50nv44wpl3HQAc0OB9I3\nVXWrjedEnIq03oC27qjHjympx31un0wiTQLq4bS4shrAECZnMGVTvbD9QSnOhDInjzsyuiBXWQP0\nr4R18jfgnHvJr4nOknrFmJ5xzhraRtHWf3Kkpfz/yGo5S+H73hkquyhdjr3R8m2lj09svwAAAP//\nAwBQSwMEFAAGAAgAAAAhAFr0LFu/AAAAFQEAAAsAAABfcmVscy8ucmVsc2zPwWrDMAwG4Ptg72B0\nX5TuUMaI01uh19I+gLGVxCy2jGSy9e1nemrHjpL4P0nD4SetZiPRyNnCruvBUPYcYp4tXC/Htw8w\nWl0ObuVMFm6kcBhfX4Yzra62kC6xqGlKVgtLreUTUf1CyWnHhXKbTCzJ1VbKjMX5LzcTvvf9HuXR\ngPHJNKdgQU5hB+ZyK23zHztFL6w81c5zQp6m6P9TMfB3PtPWFCczVQtB9N4U2rp2HOA44NMz4y8A\nAAD//wMAUEsDBBQABgAIAAAAIQBXZw8vwwAAANoAAAAPAAAAZHJzL2Rvd25yZXYueG1sRI9Ba8JA\nFITvBf/D8gRvdWNBqdFVRCwUhGKMB4/P7DNZzL6N2a3Gf98VCh6HmfmGmS87W4sbtd44VjAaJiCI\nC6cNlwoO+df7JwgfkDXWjknBgzwsF723Oaba3Tmj2z6UIkLYp6igCqFJpfRFRRb90DXE0Tu71mKI\nsi2lbvEe4baWH0kykRYNx4UKG1pXVFz2v1bB6sjZxlx/TrvsnJk8nya8nVyUGvS71QxEoC68wv/t\nb61gDM8r8QbIxR8AAAD//wMAUEsBAi0AFAAGAAgAAAAhANvh9svuAAAAhQEAABMAAAAAAAAAAAAA\nAAAAAAAAAFtDb250ZW50X1R5cGVzXS54bWxQSwECLQAUAAYACAAAACEAWvQsW78AAAAVAQAACwAA\nAAAAAAAAAAAAAAAfAQAAX3JlbHMvLnJlbHNQSwECLQAUAAYACAAAACEAV2cPL8MAAADaAAAADwAA\nAAAAAAAAAAAAAAAHAgAAZHJzL2Rvd25yZXYueG1sUEsFBgAAAAADAAMAtwAAAPcCAAAAAA==\n" };

            V.TextBox textBox1 = new V.TextBox() { Inset = "0,0,0,0" };

            TextBoxContent textBoxContent2 = new TextBoxContent();

            Paragraph paragraph3 = new Paragraph() { RsidParagraphAddition = "004918B9", RsidRunAdditionDefault = "004918B9", ParagraphId = "12772B1B", TextId = "77777777" };

            ParagraphProperties paragraphProperties3 = new ParagraphProperties();
            Justification justification3 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties2 = new ParagraphMarkRunProperties();
            FontSizeComplexScript fontSizeComplexScript4 = new FontSizeComplexScript() { Val = "18" };

            paragraphMarkRunProperties2.Append(fontSizeComplexScript4);

            paragraphProperties3.Append(justification3);
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
            Italic italic3 = new Italic();
            ItalicComplexScript italicComplexScript3 = new ItalicComplexScript();
            NoProof noProof4 = new NoProof();
            FontSize fontSize3 = new FontSize() { Val = "18" };
            FontSizeComplexScript fontSizeComplexScript5 = new FontSizeComplexScript() { Val = "18" };

            runProperties4.Append(italic3);
            runProperties4.Append(italicComplexScript3);
            runProperties4.Append(noProof4);
            runProperties4.Append(fontSize3);
            runProperties4.Append(fontSizeComplexScript5);
            Text text2 = new Text();
            text2.Text = "2";

            run10.Append(runProperties4);
            run10.Append(text2);

            Run run11 = new Run();

            RunProperties runProperties5 = new RunProperties();
            Italic italic4 = new Italic();
            ItalicComplexScript italicComplexScript4 = new ItalicComplexScript();
            NoProof noProof5 = new NoProof();
            FontSize fontSize4 = new FontSize() { Val = "18" };
            FontSizeComplexScript fontSizeComplexScript6 = new FontSizeComplexScript() { Val = "18" };

            runProperties5.Append(italic4);
            runProperties5.Append(italicComplexScript4);
            runProperties5.Append(noProof5);
            runProperties5.Append(fontSize4);
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

            shape1.Append(textBox1);

            V.Group group2 = new V.Group() { Id = "Group 57", Style = "position:absolute;left:5494;top:739;width:372;height:72", CoordinateSize = "372,72", CoordinateOrigin = "5486,739", OptionalString = "_x0000_s1028" };
            group2.SetAttribute(new OpenXmlAttribute("o", "gfxdata", "urn:schemas-microsoft-com:office:office", "UEsDBBQABgAIAAAAIQDb4fbL7gAAAIUBAAATAAAAW0NvbnRlbnRfVHlwZXNdLnhtbHyQz07DMAyH\n70i8Q+QralM4IITa7kDhCAiNB7ASt43WOlEcyvb2pNu4IODoPz9/n1xv9vOkForiPDdwXVagiI23\njocG3rdPxR0oScgWJ8/UwIEENu3lRb09BBKV0ywNjCmFe63FjDSjlD4Q50nv44wpl3HQAc0OB9I3\nVXWrjedEnIq03oC27qjHjympx31un0wiTQLq4bS4shrAECZnMGVTvbD9QSnOhDInjzsyuiBXWQP0\nr4R18jfgnHvJr4nOknrFmJ5xzhraRtHWf3Kkpfz/yGo5S+H73hkquyhdjr3R8m2lj09svwAAAP//\nAwBQSwMEFAAGAAgAAAAhAFr0LFu/AAAAFQEAAAsAAABfcmVscy8ucmVsc2zPwWrDMAwG4Ptg72B0\nX5TuUMaI01uh19I+gLGVxCy2jGSy9e1nemrHjpL4P0nD4SetZiPRyNnCruvBUPYcYp4tXC/Htw8w\nWl0ObuVMFm6kcBhfX4Yzra62kC6xqGlKVgtLreUTUf1CyWnHhXKbTCzJ1VbKjMX5LzcTvvf9HuXR\ngPHJNKdgQU5hB+ZyK23zHztFL6w81c5zQp6m6P9TMfB3PtPWFCczVQtB9N4U2rp2HOA44NMz4y8A\nAAD//wMAUEsDBBQABgAIAAAAIQDNCx1fwgAAANoAAAAPAAAAZHJzL2Rvd25yZXYueG1sRI9Bi8Iw\nFITvC/6H8ARva1plRapRRFQ8iLAqiLdH82yLzUtpYlv/vVkQ9jjMzDfMfNmZUjRUu8KygngYgSBO\nrS44U3A5b7+nIJxH1lhaJgUvcrBc9L7mmGjb8i81J5+JAGGXoILc+yqR0qU5GXRDWxEH725rgz7I\nOpO6xjbATSlHUTSRBgsOCzlWtM4pfZyeRsGuxXY1jjfN4XFfv27nn+P1EJNSg363moHw1Pn/8Ke9\n1wom8Hcl3AC5eAMAAP//AwBQSwECLQAUAAYACAAAACEA2+H2y+4AAACFAQAAEwAAAAAAAAAAAAAA\nAAAAAAAAW0NvbnRlbnRfVHlwZXNdLnhtbFBLAQItABQABgAIAAAAIQBa9CxbvwAAABUBAAALAAAA\nAAAAAAAAAAAAAB8BAABfcmVscy8ucmVsc1BLAQItABQABgAIAAAAIQDNCx1fwgAAANoAAAAPAAAA\nAAAAAAAAAAAAAAcCAABkcnMvZG93bnJldi54bWxQSwUGAAAAAAMAAwC3AAAA9gIAAAAA\n"));

            V.Oval oval1 = new V.Oval() { Id = "Oval 58", Style = "position:absolute;left:5486;top:739;width:72;height:72;visibility:visible;mso-wrap-style:square;v-text-anchor:top", OptionalString = "_x0000_s1029", FillColor = "#84a2c6", Stroked = false };
            oval1.SetAttribute(new OpenXmlAttribute("o", "gfxdata", "urn:schemas-microsoft-com:office:office", "UEsDBBQABgAIAAAAIQDb4fbL7gAAAIUBAAATAAAAW0NvbnRlbnRfVHlwZXNdLnhtbHyQz07DMAyH\n70i8Q+QralM4IITa7kDhCAiNB7ASt43WOlEcyvb2pNu4IODoPz9/n1xv9vOkForiPDdwXVagiI23\njocG3rdPxR0oScgWJ8/UwIEENu3lRb09BBKV0ywNjCmFe63FjDSjlD4Q50nv44wpl3HQAc0OB9I3\nVXWrjedEnIq03oC27qjHjympx31un0wiTQLq4bS4shrAECZnMGVTvbD9QSnOhDInjzsyuiBXWQP0\nr4R18jfgnHvJr4nOknrFmJ5xzhraRtHWf3Kkpfz/yGo5S+H73hkquyhdjr3R8m2lj09svwAAAP//\nAwBQSwMEFAAGAAgAAAAhAFr0LFu/AAAAFQEAAAsAAABfcmVscy8ucmVsc2zPwWrDMAwG4Ptg72B0\nX5TuUMaI01uh19I+gLGVxCy2jGSy9e1nemrHjpL4P0nD4SetZiPRyNnCruvBUPYcYp4tXC/Htw8w\nWl0ObuVMFm6kcBhfX4Yzra62kC6xqGlKVgtLreUTUf1CyWnHhXKbTCzJ1VbKjMX5LzcTvvf9HuXR\ngPHJNKdgQU5hB+ZyK23zHztFL6w81c5zQp6m6P9TMfB3PtPWFCczVQtB9N4U2rp2HOA44NMz4y8A\nAAD//wMAUEsDBBQABgAIAAAAIQCVU+0kvgAAANoAAAAPAAAAZHJzL2Rvd25yZXYueG1sRI9BD8FA\nFITvEv9h8yRubDkgZQkS4qo4uD3dp2103zbdVfXvrUTiOJmZbzKLVWtK0VDtCssKRsMIBHFqdcGZ\ngvNpN5iBcB5ZY2mZFLzJwWrZ7Sww1vbFR2oSn4kAYRejgtz7KpbSpTkZdENbEQfvbmuDPsg6k7rG\nV4CbUo6jaCINFhwWcqxom1P6SJ5GQbG3o8tukxzdtZls5bq8bezlplS/167nIDy1/h/+tQ9awRS+\nV8INkMsPAAAA//8DAFBLAQItABQABgAIAAAAIQDb4fbL7gAAAIUBAAATAAAAAAAAAAAAAAAAAAAA\nAABbQ29udGVudF9UeXBlc10ueG1sUEsBAi0AFAAGAAgAAAAhAFr0LFu/AAAAFQEAAAsAAAAAAAAA\nAAAAAAAAHwEAAF9yZWxzLy5yZWxzUEsBAi0AFAAGAAgAAAAhAJVT7SS+AAAA2gAAAA8AAAAAAAAA\nAAAAAAAABwIAAGRycy9kb3ducmV2LnhtbFBLBQYAAAAAAwADALcAAADyAgAAAAA=\n"));

            V.Oval oval2 = new V.Oval() { Id = "Oval 59", Style = "position:absolute;left:5636;top:739;width:72;height:72;visibility:visible;mso-wrap-style:square;v-text-anchor:top", OptionalString = "_x0000_s1030", FillColor = "#84a2c6", Stroked = false };
            oval2.SetAttribute(new OpenXmlAttribute("o", "gfxdata", "urn:schemas-microsoft-com:office:office", "UEsDBBQABgAIAAAAIQDb4fbL7gAAAIUBAAATAAAAW0NvbnRlbnRfVHlwZXNdLnhtbHyQz07DMAyH\n70i8Q+QralM4IITa7kDhCAiNB7ASt43WOlEcyvb2pNu4IODoPz9/n1xv9vOkForiPDdwXVagiI23\njocG3rdPxR0oScgWJ8/UwIEENu3lRb09BBKV0ywNjCmFe63FjDSjlD4Q50nv44wpl3HQAc0OB9I3\nVXWrjedEnIq03oC27qjHjympx31un0wiTQLq4bS4shrAECZnMGVTvbD9QSnOhDInjzsyuiBXWQP0\nr4R18jfgnHvJr4nOknrFmJ5xzhraRtHWf3Kkpfz/yGo5S+H73hkquyhdjr3R8m2lj09svwAAAP//\nAwBQSwMEFAAGAAgAAAAhAFr0LFu/AAAAFQEAAAsAAABfcmVscy8ucmVsc2zPwWrDMAwG4Ptg72B0\nX5TuUMaI01uh19I+gLGVxCy2jGSy9e1nemrHjpL4P0nD4SetZiPRyNnCruvBUPYcYp4tXC/Htw8w\nWl0ObuVMFm6kcBhfX4Yzra62kC6xqGlKVgtLreUTUf1CyWnHhXKbTCzJ1VbKjMX5LzcTvvf9HuXR\ngPHJNKdgQU5hB+ZyK23zHztFL6w81c5zQp6m6P9TMfB3PtPWFCczVQtB9N4U2rp2HOA44NMz4y8A\nAAD//wMAUEsDBBQABgAIAAAAIQDkzHlWuwAAANoAAAAPAAAAZHJzL2Rvd25yZXYueG1sRE+9CsIw\nEN4F3yGc4KapDiLVWKqguFp1cDubsy02l9LEWt/eDILjx/e/TnpTi45aV1lWMJtGIIhzqysuFFzO\n+8kShPPIGmvLpOBDDpLNcLDGWNs3n6jLfCFCCLsYFZTeN7GULi/JoJvahjhwD9sa9AG2hdQtvkO4\nqeU8ihbSYMWhocSGdiXlz+xlFFQHO7vut9nJ3brFTqb1fWuvd6XGoz5dgfDU+7/45z5qBWFruBJu\ngNx8AQAA//8DAFBLAQItABQABgAIAAAAIQDb4fbL7gAAAIUBAAATAAAAAAAAAAAAAAAAAAAAAABb\nQ29udGVudF9UeXBlc10ueG1sUEsBAi0AFAAGAAgAAAAhAFr0LFu/AAAAFQEAAAsAAAAAAAAAAAAA\nAAAAHwEAAF9yZWxzLy5yZWxzUEsBAi0AFAAGAAgAAAAhAOTMeVa7AAAA2gAAAA8AAAAAAAAAAAAA\nAAAABwIAAGRycy9kb3ducmV2LnhtbFBLBQYAAAAAAwADALcAAADvAgAAAAA=\n"));

            V.Oval oval3 = new V.Oval() { Id = "Oval 60", Style = "position:absolute;left:5786;top:739;width:72;height:72;visibility:visible;mso-wrap-style:square;v-text-anchor:top", OptionalString = "_x0000_s1031", FillColor = "#84a2c6", Stroked = false };
            oval3.SetAttribute(new OpenXmlAttribute("o", "gfxdata", "urn:schemas-microsoft-com:office:office", "UEsDBBQABgAIAAAAIQDb4fbL7gAAAIUBAAATAAAAW0NvbnRlbnRfVHlwZXNdLnhtbHyQz07DMAyH\n70i8Q+QralM4IITa7kDhCAiNB7ASt43WOlEcyvb2pNu4IODoPz9/n1xv9vOkForiPDdwXVagiI23\njocG3rdPxR0oScgWJ8/UwIEENu3lRb09BBKV0ywNjCmFe63FjDSjlD4Q50nv44wpl3HQAc0OB9I3\nVXWrjedEnIq03oC27qjHjympx31un0wiTQLq4bS4shrAECZnMGVTvbD9QSnOhDInjzsyuiBXWQP0\nr4R18jfgnHvJr4nOknrFmJ5xzhraRtHWf3Kkpfz/yGo5S+H73hkquyhdjr3R8m2lj09svwAAAP//\nAwBQSwMEFAAGAAgAAAAhAFr0LFu/AAAAFQEAAAsAAABfcmVscy8ucmVsc2zPwWrDMAwG4Ptg72B0\nX5TuUMaI01uh19I+gLGVxCy2jGSy9e1nemrHjpL4P0nD4SetZiPRyNnCruvBUPYcYp4tXC/Htw8w\nWl0ObuVMFm6kcBhfX4Yzra62kC6xqGlKVgtLreUTUf1CyWnHhXKbTCzJ1VbKjMX5LzcTvvf9HuXR\ngPHJNKdgQU5hB+ZyK23zHztFL6w81c5zQp6m6P9TMfB3PtPWFCczVQtB9N4U2rp2HOA44NMz4y8A\nAAD//wMAUEsDBBQABgAIAAAAIQCLgNzNvgAAANoAAAAPAAAAZHJzL2Rvd25yZXYueG1sRI9BD8FA\nFITvEv9h8yRubDkIZQkS4qo4uD3dp2103zbdVfXvrUTiOJmZbzKLVWtK0VDtCssKRsMIBHFqdcGZ\ngvNpN5iCcB5ZY2mZFLzJwWrZ7Sww1vbFR2oSn4kAYRejgtz7KpbSpTkZdENbEQfvbmuDPsg6k7rG\nV4CbUo6jaCINFhwWcqxom1P6SJ5GQbG3o8tukxzdtZls5bq8bezlplS/167nIDy1/h/+tQ9awQy+\nV8INkMsPAAAA//8DAFBLAQItABQABgAIAAAAIQDb4fbL7gAAAIUBAAATAAAAAAAAAAAAAAAAAAAA\nAABbQ29udGVudF9UeXBlc10ueG1sUEsBAi0AFAAGAAgAAAAhAFr0LFu/AAAAFQEAAAsAAAAAAAAA\nAAAAAAAAHwEAAF9yZWxzLy5yZWxzUEsBAi0AFAAGAAgAAAAhAIuA3M2+AAAA2gAAAA8AAAAAAAAA\nAAAAAAAABwIAAGRycy9kb3ducmV2LnhtbFBLBQYAAAAAAwADALcAAADyAgAAAAA=\n"));

            group2.Append(oval1);
            group2.Append(oval2);
            group2.Append(oval3);
            Wvml.AnchorLock anchorLock1 = new Wvml.AnchorLock();

            group1.Append(shapetype1);
            group1.Append(shape1);
            group1.Append(group2);
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
