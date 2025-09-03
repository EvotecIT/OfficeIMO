using DocumentFormat.OpenXml.Packaging;
using Ap = DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using Wp = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using Dgm = DocumentFormat.OpenXml.Drawing.Diagrams;
using Dsp = DocumentFormat.OpenXml.Office.Drawing;
using M = DocumentFormat.OpenXml.Math;
using Ovml = DocumentFormat.OpenXml.Vml.Office;
using V = DocumentFormat.OpenXml.Vml;
using W14 = DocumentFormat.OpenXml.Office2010.Word;
using W15 = DocumentFormat.OpenXml.Office2013.Word;
using Thm15 = DocumentFormat.OpenXml.Office2013.Theme;

namespace GeneratedCode
{
    public class GeneratedClass
    {
        // Creates a WordprocessingDocument.
        public void CreatePackage(string filePath)
        {
            using(WordprocessingDocument package = WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document))
            {
                CreateParts(package);
            }
        }

        // Adds child parts and generates content of the specified part.
        private void CreateParts(WordprocessingDocument document)
        {
            ExtendedFilePropertiesPart extendedFilePropertiesPart1 = document.AddNewPart<ExtendedFilePropertiesPart>("rId3");
            GenerateExtendedFilePropertiesPart1Content(extendedFilePropertiesPart1);

            MainDocumentPart mainDocumentPart1 = document.AddMainDocumentPart();
            GenerateMainDocumentPart1Content(mainDocumentPart1);

            DiagramPersistLayoutPart diagramPersistLayoutPart1 = mainDocumentPart1.AddNewPart<DiagramPersistLayoutPart>("rId8");
            GenerateDiagramPersistLayoutPart1Content(diagramPersistLayoutPart1);

            WebSettingsPart webSettingsPart1 = mainDocumentPart1.AddNewPart<WebSettingsPart>("rId3");
            GenerateWebSettingsPart1Content(webSettingsPart1);

            DiagramColorsPart diagramColorsPart1 = mainDocumentPart1.AddNewPart<DiagramColorsPart>("rId7");
            GenerateDiagramColorsPart1Content(diagramColorsPart1);

            DocumentSettingsPart documentSettingsPart1 = mainDocumentPart1.AddNewPart<DocumentSettingsPart>("rId2");
            GenerateDocumentSettingsPart1Content(documentSettingsPart1);

            StyleDefinitionsPart styleDefinitionsPart1 = mainDocumentPart1.AddNewPart<StyleDefinitionsPart>("rId1");
            GenerateStyleDefinitionsPart1Content(styleDefinitionsPart1);

            DiagramStylePart diagramStylePart1 = mainDocumentPart1.AddNewPart<DiagramStylePart>("rId6");
            GenerateDiagramStylePart1Content(diagramStylePart1);

            DiagramLayoutDefinitionPart diagramLayoutDefinitionPart1 = mainDocumentPart1.AddNewPart<DiagramLayoutDefinitionPart>("rId5");
            GenerateDiagramLayoutDefinitionPart1Content(diagramLayoutDefinitionPart1);

            ThemePart themePart1 = mainDocumentPart1.AddNewPart<ThemePart>("rId10");
            GenerateThemePart1Content(themePart1);

            DiagramDataPart diagramDataPart1 = mainDocumentPart1.AddNewPart<DiagramDataPart>("rId4");
            GenerateDiagramDataPart1Content(diagramDataPart1);

            FontTablePart fontTablePart1 = mainDocumentPart1.AddNewPart<FontTablePart>("rId9");
            GenerateFontTablePart1Content(fontTablePart1);

            SetPackageProperties(document);
        }

        // Generates content of extendedFilePropertiesPart1.
        private void GenerateExtendedFilePropertiesPart1Content(ExtendedFilePropertiesPart extendedFilePropertiesPart1)
        {
            Ap.Properties properties1 = new Ap.Properties();
            properties1.AddNamespaceDeclaration("vt", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes");
            Ap.Template template1 = new Ap.Template();
            template1.Text = "Normal.dotm";
            Ap.TotalTime totalTime1 = new Ap.TotalTime();
            totalTime1.Text = "1";
            Ap.Pages pages1 = new Ap.Pages();
            pages1.Text = "1";
            Ap.Words words1 = new Ap.Words();
            words1.Text = "0";
            Ap.Characters characters1 = new Ap.Characters();
            characters1.Text = "1";
            Ap.Application application1 = new Ap.Application();
            application1.Text = "Microsoft Office Word";
            Ap.DocumentSecurity documentSecurity1 = new Ap.DocumentSecurity();
            documentSecurity1.Text = "0";
            Ap.Lines lines1 = new Ap.Lines();
            lines1.Text = "1";
            Ap.Paragraphs paragraphs1 = new Ap.Paragraphs();
            paragraphs1.Text = "1";
            Ap.ScaleCrop scaleCrop1 = new Ap.ScaleCrop();
            scaleCrop1.Text = "false";
            Ap.Company company1 = new Ap.Company();
            company1.Text = "";
            Ap.LinksUpToDate linksUpToDate1 = new Ap.LinksUpToDate();
            linksUpToDate1.Text = "false";
            Ap.CharactersWithSpaces charactersWithSpaces1 = new Ap.CharactersWithSpaces();
            charactersWithSpaces1.Text = "1";
            Ap.SharedDocument sharedDocument1 = new Ap.SharedDocument();
            sharedDocument1.Text = "false";
            Ap.HyperlinksChanged hyperlinksChanged1 = new Ap.HyperlinksChanged();
            hyperlinksChanged1.Text = "false";
            Ap.ApplicationVersion applicationVersion1 = new Ap.ApplicationVersion();
            applicationVersion1.Text = "16.0000";

            properties1.Append(template1);
            properties1.Append(totalTime1);
            properties1.Append(pages1);
            properties1.Append(words1);
            properties1.Append(characters1);
            properties1.Append(application1);
            properties1.Append(documentSecurity1);
            properties1.Append(lines1);
            properties1.Append(paragraphs1);
            properties1.Append(scaleCrop1);
            properties1.Append(company1);
            properties1.Append(linksUpToDate1);
            properties1.Append(charactersWithSpaces1);
            properties1.Append(sharedDocument1);
            properties1.Append(hyperlinksChanged1);
            properties1.Append(applicationVersion1);

            extendedFilePropertiesPart1.Properties = properties1;
        }

        // Generates content of mainDocumentPart1.
        private void GenerateMainDocumentPart1Content(MainDocumentPart mainDocumentPart1)
        {
            Document document1 = new Document(){ MCAttributes = new MarkupCompatibilityAttributes(){ Ignorable = "w14 w15 w16se w16cid w16 w16cex w16sdtdh w16sdtfl w16du wp14" }  };
            document1.AddNamespaceDeclaration("wpc", "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas");
            document1.AddNamespaceDeclaration("cx", "http://schemas.microsoft.com/office/drawing/2014/chartex");
            document1.AddNamespaceDeclaration("cx1", "http://schemas.microsoft.com/office/drawing/2015/9/8/chartex");
            document1.AddNamespaceDeclaration("cx2", "http://schemas.microsoft.com/office/drawing/2015/10/21/chartex");
            document1.AddNamespaceDeclaration("cx3", "http://schemas.microsoft.com/office/drawing/2016/5/9/chartex");
            document1.AddNamespaceDeclaration("cx4", "http://schemas.microsoft.com/office/drawing/2016/5/10/chartex");
            document1.AddNamespaceDeclaration("cx5", "http://schemas.microsoft.com/office/drawing/2016/5/11/chartex");
            document1.AddNamespaceDeclaration("cx6", "http://schemas.microsoft.com/office/drawing/2016/5/12/chartex");
            document1.AddNamespaceDeclaration("cx7", "http://schemas.microsoft.com/office/drawing/2016/5/13/chartex");
            document1.AddNamespaceDeclaration("cx8", "http://schemas.microsoft.com/office/drawing/2016/5/14/chartex");
            document1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            document1.AddNamespaceDeclaration("aink", "http://schemas.microsoft.com/office/drawing/2016/ink");
            document1.AddNamespaceDeclaration("am3d", "http://schemas.microsoft.com/office/drawing/2017/model3d");
            document1.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            document1.AddNamespaceDeclaration("oel", "http://schemas.microsoft.com/office/2019/extlst");
            document1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            document1.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            document1.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            document1.AddNamespaceDeclaration("wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing");
            document1.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            document1.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
            document1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            document1.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            document1.AddNamespaceDeclaration("w15", "http://schemas.microsoft.com/office/word/2012/wordml");
            document1.AddNamespaceDeclaration("w16cex", "http://schemas.microsoft.com/office/word/2018/wordml/cex");
            document1.AddNamespaceDeclaration("w16cid", "http://schemas.microsoft.com/office/word/2016/wordml/cid");
            document1.AddNamespaceDeclaration("w16", "http://schemas.microsoft.com/office/word/2018/wordml");
            document1.AddNamespaceDeclaration("w16du", "http://schemas.microsoft.com/office/word/2023/wordml/word16du");
            document1.AddNamespaceDeclaration("w16sdtdh", "http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash");
            document1.AddNamespaceDeclaration("w16sdtfl", "http://schemas.microsoft.com/office/word/2024/wordml/sdtformatlock");
            document1.AddNamespaceDeclaration("w16se", "http://schemas.microsoft.com/office/word/2015/wordml/symex");
            document1.AddNamespaceDeclaration("wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup");
            document1.AddNamespaceDeclaration("wpi", "http://schemas.microsoft.com/office/word/2010/wordprocessingInk");
            document1.AddNamespaceDeclaration("wne", "http://schemas.microsoft.com/office/word/2006/wordml");
            document1.AddNamespaceDeclaration("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");

            Body body1 = new Body();

            Paragraph paragraph1 = new Paragraph(){ RsidParagraphAddition = "0018596D", RsidRunAdditionDefault = "00440271", ParagraphId = "0B052151", TextId = "6BF634FF" };

            Run run1 = new Run();

            RunProperties runProperties1 = new RunProperties();
            NoProof noProof1 = new NoProof();

            runProperties1.Append(noProof1);

            Drawing drawing1 = new Drawing();

            Wp.Inline inline1 = new Wp.Inline(){ DistanceFromTop = (UInt32Value)0U, DistanceFromBottom = (UInt32Value)0U, DistanceFromLeft = (UInt32Value)0U, DistanceFromRight = (UInt32Value)0U, AnchorId = "5F5BFB11", EditId = "49C2D8E9" };
            Wp.Extent extent1 = new Wp.Extent(){ Cx = 5486400L, Cy = 3200400L };
            Wp.EffectExtent effectExtent1 = new Wp.EffectExtent(){ LeftEdge = 38100L, TopEdge = 0L, RightEdge = 57150L, BottomEdge = 0L };
            Wp.DocProperties docProperties1 = new Wp.DocProperties(){ Id = (UInt32Value)1175469687U, Name = "Diagram 1" };
            Wp.NonVisualGraphicFrameDrawingProperties nonVisualGraphicFrameDrawingProperties1 = new Wp.NonVisualGraphicFrameDrawingProperties();

            A.Graphic graphic1 = new A.Graphic();
            graphic1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            A.GraphicData graphicData1 = new A.GraphicData(){ Uri = "http://schemas.openxmlformats.org/drawingml/2006/diagram" };

            Dgm.RelationshipIds relationshipIds1 = new Dgm.RelationshipIds(){ DataPart = "rId4", LayoutPart = "rId5", StylePart = "rId6", ColorPart = "rId7" };
            relationshipIds1.AddNamespaceDeclaration("dgm", "http://schemas.openxmlformats.org/drawingml/2006/diagram");
            relationshipIds1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

            graphicData1.Append(relationshipIds1);

            graphic1.Append(graphicData1);

            inline1.Append(extent1);
            inline1.Append(effectExtent1);
            inline1.Append(docProperties1);
            inline1.Append(nonVisualGraphicFrameDrawingProperties1);
            inline1.Append(graphic1);

            drawing1.Append(inline1);

            run1.Append(runProperties1);
            run1.Append(drawing1);

            paragraph1.Append(run1);

            SectionProperties sectionProperties1 = new SectionProperties(){ RsidR = "0018596D" };
            PageSize pageSize1 = new PageSize(){ Width = (UInt32Value)11906U, Height = (UInt32Value)16838U };
            PageMargin pageMargin1 = new PageMargin(){ Top = 1440, Right = (UInt32Value)1440U, Bottom = 1440, Left = (UInt32Value)1440U, Header = (UInt32Value)708U, Footer = (UInt32Value)708U, Gutter = (UInt32Value)0U };
            Columns columns1 = new Columns(){ Space = "708" };
            DocGrid docGrid1 = new DocGrid(){ LinePitch = 360 };

            sectionProperties1.Append(pageSize1);
            sectionProperties1.Append(pageMargin1);
            sectionProperties1.Append(columns1);
            sectionProperties1.Append(docGrid1);

            body1.Append(paragraph1);
            body1.Append(sectionProperties1);

            document1.Append(body1);

            mainDocumentPart1.Document = document1;
        }

        // Generates content of diagramPersistLayoutPart1.
        private void GenerateDiagramPersistLayoutPart1Content(DiagramPersistLayoutPart diagramPersistLayoutPart1)
        {
            Dsp.Drawing drawing2 = new Dsp.Drawing();
            drawing2.AddNamespaceDeclaration("dgm", "http://schemas.openxmlformats.org/drawingml/2006/diagram");
            drawing2.AddNamespaceDeclaration("dsp", "http://schemas.microsoft.com/office/drawing/2008/diagram");
            drawing2.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            Dsp.ShapeTree shapeTree1 = new Dsp.ShapeTree();

            Dsp.GroupShapeNonVisualProperties groupShapeNonVisualProperties1 = new Dsp.GroupShapeNonVisualProperties();
            Dsp.NonVisualDrawingProperties nonVisualDrawingProperties1 = new Dsp.NonVisualDrawingProperties(){ Id = (UInt32Value)0U, Name = "" };
            Dsp.NonVisualGroupDrawingShapeProperties nonVisualGroupDrawingShapeProperties1 = new Dsp.NonVisualGroupDrawingShapeProperties();

            groupShapeNonVisualProperties1.Append(nonVisualDrawingProperties1);
            groupShapeNonVisualProperties1.Append(nonVisualGroupDrawingShapeProperties1);
            Dsp.GroupShapeProperties groupShapeProperties1 = new Dsp.GroupShapeProperties();

            Dsp.Shape shape1 = new Dsp.Shape(){ ModelId = "{068D8084-B4F1-4349-BF7B-3A540F7ACE9A}" };

            Dsp.ShapeNonVisualProperties shapeNonVisualProperties1 = new Dsp.ShapeNonVisualProperties();
            Dsp.NonVisualDrawingProperties nonVisualDrawingProperties2 = new Dsp.NonVisualDrawingProperties(){ Id = (UInt32Value)0U, Name = "" };
            Dsp.NonVisualDrawingShapeProperties nonVisualDrawingShapeProperties1 = new Dsp.NonVisualDrawingShapeProperties();

            shapeNonVisualProperties1.Append(nonVisualDrawingProperties2);
            shapeNonVisualProperties1.Append(nonVisualDrawingShapeProperties1);

            Dsp.ShapeProperties shapeProperties1 = new Dsp.ShapeProperties();

            A.Transform2D transform2D1 = new A.Transform2D();
            A.Offset offset1 = new A.Offset(){ X = 0L, Y = 485774L };
            A.Extents extents1 = new A.Extents(){ Cx = 1714499L, Cy = 1028700L };

            transform2D1.Append(offset1);
            transform2D1.Append(extents1);

            A.PresetGeometry presetGeometry1 = new A.PresetGeometry(){ Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList1 = new A.AdjustValueList();

            presetGeometry1.Append(adjustValueList1);

            A.SolidFill solidFill1 = new A.SolidFill();

            A.SchemeColor schemeColor1 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent1 };
            A.HueOffset hueOffset1 = new A.HueOffset(){ Val = 0 };
            A.SaturationOffset saturationOffset1 = new A.SaturationOffset(){ Val = 0 };
            A.LuminanceOffset luminanceOffset1 = new A.LuminanceOffset(){ Val = 0 };
            A.AlphaOffset alphaOffset1 = new A.AlphaOffset(){ Val = 0 };

            schemeColor1.Append(hueOffset1);
            schemeColor1.Append(saturationOffset1);
            schemeColor1.Append(luminanceOffset1);
            schemeColor1.Append(alphaOffset1);

            solidFill1.Append(schemeColor1);

            A.Outline outline1 = new A.Outline(){ Width = 19050, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill2 = new A.SolidFill();

            A.SchemeColor schemeColor2 = new A.SchemeColor(){ Val = A.SchemeColorValues.Light1 };
            A.HueOffset hueOffset2 = new A.HueOffset(){ Val = 0 };
            A.SaturationOffset saturationOffset2 = new A.SaturationOffset(){ Val = 0 };
            A.LuminanceOffset luminanceOffset2 = new A.LuminanceOffset(){ Val = 0 };
            A.AlphaOffset alphaOffset2 = new A.AlphaOffset(){ Val = 0 };

            schemeColor2.Append(hueOffset2);
            schemeColor2.Append(saturationOffset2);
            schemeColor2.Append(luminanceOffset2);
            schemeColor2.Append(alphaOffset2);

            solidFill2.Append(schemeColor2);
            A.PresetDash presetDash1 = new A.PresetDash(){ Val = A.PresetLineDashValues.Solid };
            A.Miter miter1 = new A.Miter(){ Limit = 800000 };

            outline1.Append(solidFill2);
            outline1.Append(presetDash1);
            outline1.Append(miter1);
            A.EffectList effectList1 = new A.EffectList();

            shapeProperties1.Append(transform2D1);
            shapeProperties1.Append(presetGeometry1);
            shapeProperties1.Append(solidFill1);
            shapeProperties1.Append(outline1);
            shapeProperties1.Append(effectList1);

            Dsp.ShapeStyle shapeStyle1 = new Dsp.ShapeStyle();

            A.LineReference lineReference1 = new A.LineReference(){ Index = (UInt32Value)2U };
            A.RgbColorModelPercentage rgbColorModelPercentage1 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            lineReference1.Append(rgbColorModelPercentage1);

            A.FillReference fillReference1 = new A.FillReference(){ Index = (UInt32Value)1U };
            A.RgbColorModelPercentage rgbColorModelPercentage2 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            fillReference1.Append(rgbColorModelPercentage2);

            A.EffectReference effectReference1 = new A.EffectReference(){ Index = (UInt32Value)0U };
            A.RgbColorModelPercentage rgbColorModelPercentage3 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            effectReference1.Append(rgbColorModelPercentage3);

            A.FontReference fontReference1 = new A.FontReference(){ Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor3 = new A.SchemeColor(){ Val = A.SchemeColorValues.Light1 };

            fontReference1.Append(schemeColor3);

            shapeStyle1.Append(lineReference1);
            shapeStyle1.Append(fillReference1);
            shapeStyle1.Append(effectReference1);
            shapeStyle1.Append(fontReference1);

            Dsp.TextBody textBody1 = new Dsp.TextBody();

            A.BodyProperties bodyProperties1 = new A.BodyProperties(){ UseParagraphSpacing = false, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, LeftInset = 179070, TopInset = 179070, RightInset = 179070, BottomInset = 179070, ColumnCount = 1, ColumnSpacing = 1270, Anchor = A.TextAnchoringTypeValues.Center, AnchorCenter = false };
            A.NoAutoFit noAutoFit1 = new A.NoAutoFit();

            bodyProperties1.Append(noAutoFit1);
            A.ListStyle listStyle1 = new A.ListStyle();

            A.Paragraph paragraph2 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties1 = new A.ParagraphProperties(){ LeftMargin = 0, Level = 0, Indent = 0, Alignment = A.TextAlignmentTypeValues.Center, DefaultTabSize = 2089150 };

            A.LineSpacing lineSpacing1 = new A.LineSpacing();
            A.SpacingPercent spacingPercent1 = new A.SpacingPercent(){ Val = 90000 };

            lineSpacing1.Append(spacingPercent1);

            A.SpaceBefore spaceBefore1 = new A.SpaceBefore();
            A.SpacingPercent spacingPercent2 = new A.SpacingPercent(){ Val = 0 };

            spaceBefore1.Append(spacingPercent2);

            A.SpaceAfter spaceAfter1 = new A.SpaceAfter();
            A.SpacingPercent spacingPercent3 = new A.SpacingPercent(){ Val = 35000 };

            spaceAfter1.Append(spacingPercent3);
            A.NoBullet noBullet1 = new A.NoBullet();

            paragraphProperties1.Append(lineSpacing1);
            paragraphProperties1.Append(spaceBefore1);
            paragraphProperties1.Append(spaceAfter1);
            paragraphProperties1.Append(noBullet1);
            A.EndParagraphRunProperties endParagraphRunProperties1 = new A.EndParagraphRunProperties(){ Language = "pl-PL", FontSize = 4700, Kerning = 1200 };

            paragraph2.Append(paragraphProperties1);
            paragraph2.Append(endParagraphRunProperties1);

            textBody1.Append(bodyProperties1);
            textBody1.Append(listStyle1);
            textBody1.Append(paragraph2);

            Dsp.Transform2D transform2D2 = new Dsp.Transform2D();
            A.Offset offset2 = new A.Offset(){ X = 0L, Y = 485774L };
            A.Extents extents2 = new A.Extents(){ Cx = 1714499L, Cy = 1028700L };

            transform2D2.Append(offset2);
            transform2D2.Append(extents2);

            shape1.Append(shapeNonVisualProperties1);
            shape1.Append(shapeProperties1);
            shape1.Append(shapeStyle1);
            shape1.Append(textBody1);
            shape1.Append(transform2D2);

            Dsp.Shape shape2 = new Dsp.Shape(){ ModelId = "{DF06976E-7188-463E-AE39-BDA19617EFC4}" };

            Dsp.ShapeNonVisualProperties shapeNonVisualProperties2 = new Dsp.ShapeNonVisualProperties();
            Dsp.NonVisualDrawingProperties nonVisualDrawingProperties3 = new Dsp.NonVisualDrawingProperties(){ Id = (UInt32Value)0U, Name = "" };
            Dsp.NonVisualDrawingShapeProperties nonVisualDrawingShapeProperties2 = new Dsp.NonVisualDrawingShapeProperties();

            shapeNonVisualProperties2.Append(nonVisualDrawingProperties3);
            shapeNonVisualProperties2.Append(nonVisualDrawingShapeProperties2);

            Dsp.ShapeProperties shapeProperties2 = new Dsp.ShapeProperties();

            A.Transform2D transform2D3 = new A.Transform2D();
            A.Offset offset3 = new A.Offset(){ X = 1885950L, Y = 485774L };
            A.Extents extents3 = new A.Extents(){ Cx = 1714499L, Cy = 1028700L };

            transform2D3.Append(offset3);
            transform2D3.Append(extents3);

            A.PresetGeometry presetGeometry2 = new A.PresetGeometry(){ Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList2 = new A.AdjustValueList();

            presetGeometry2.Append(adjustValueList2);

            A.SolidFill solidFill3 = new A.SolidFill();

            A.SchemeColor schemeColor4 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent1 };
            A.HueOffset hueOffset3 = new A.HueOffset(){ Val = 0 };
            A.SaturationOffset saturationOffset3 = new A.SaturationOffset(){ Val = 0 };
            A.LuminanceOffset luminanceOffset3 = new A.LuminanceOffset(){ Val = 0 };
            A.AlphaOffset alphaOffset3 = new A.AlphaOffset(){ Val = 0 };

            schemeColor4.Append(hueOffset3);
            schemeColor4.Append(saturationOffset3);
            schemeColor4.Append(luminanceOffset3);
            schemeColor4.Append(alphaOffset3);

            solidFill3.Append(schemeColor4);

            A.Outline outline2 = new A.Outline(){ Width = 19050, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill4 = new A.SolidFill();

            A.SchemeColor schemeColor5 = new A.SchemeColor(){ Val = A.SchemeColorValues.Light1 };
            A.HueOffset hueOffset4 = new A.HueOffset(){ Val = 0 };
            A.SaturationOffset saturationOffset4 = new A.SaturationOffset(){ Val = 0 };
            A.LuminanceOffset luminanceOffset4 = new A.LuminanceOffset(){ Val = 0 };
            A.AlphaOffset alphaOffset4 = new A.AlphaOffset(){ Val = 0 };

            schemeColor5.Append(hueOffset4);
            schemeColor5.Append(saturationOffset4);
            schemeColor5.Append(luminanceOffset4);
            schemeColor5.Append(alphaOffset4);

            solidFill4.Append(schemeColor5);
            A.PresetDash presetDash2 = new A.PresetDash(){ Val = A.PresetLineDashValues.Solid };
            A.Miter miter2 = new A.Miter(){ Limit = 800000 };

            outline2.Append(solidFill4);
            outline2.Append(presetDash2);
            outline2.Append(miter2);
            A.EffectList effectList2 = new A.EffectList();

            shapeProperties2.Append(transform2D3);
            shapeProperties2.Append(presetGeometry2);
            shapeProperties2.Append(solidFill3);
            shapeProperties2.Append(outline2);
            shapeProperties2.Append(effectList2);

            Dsp.ShapeStyle shapeStyle2 = new Dsp.ShapeStyle();

            A.LineReference lineReference2 = new A.LineReference(){ Index = (UInt32Value)2U };
            A.RgbColorModelPercentage rgbColorModelPercentage4 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            lineReference2.Append(rgbColorModelPercentage4);

            A.FillReference fillReference2 = new A.FillReference(){ Index = (UInt32Value)1U };
            A.RgbColorModelPercentage rgbColorModelPercentage5 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            fillReference2.Append(rgbColorModelPercentage5);

            A.EffectReference effectReference2 = new A.EffectReference(){ Index = (UInt32Value)0U };
            A.RgbColorModelPercentage rgbColorModelPercentage6 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            effectReference2.Append(rgbColorModelPercentage6);

            A.FontReference fontReference2 = new A.FontReference(){ Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor6 = new A.SchemeColor(){ Val = A.SchemeColorValues.Light1 };

            fontReference2.Append(schemeColor6);

            shapeStyle2.Append(lineReference2);
            shapeStyle2.Append(fillReference2);
            shapeStyle2.Append(effectReference2);
            shapeStyle2.Append(fontReference2);

            Dsp.TextBody textBody2 = new Dsp.TextBody();

            A.BodyProperties bodyProperties2 = new A.BodyProperties(){ UseParagraphSpacing = false, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, LeftInset = 179070, TopInset = 179070, RightInset = 179070, BottomInset = 179070, ColumnCount = 1, ColumnSpacing = 1270, Anchor = A.TextAnchoringTypeValues.Center, AnchorCenter = false };
            A.NoAutoFit noAutoFit2 = new A.NoAutoFit();

            bodyProperties2.Append(noAutoFit2);
            A.ListStyle listStyle2 = new A.ListStyle();

            A.Paragraph paragraph3 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties2 = new A.ParagraphProperties(){ LeftMargin = 0, Level = 0, Indent = 0, Alignment = A.TextAlignmentTypeValues.Center, DefaultTabSize = 2089150 };

            A.LineSpacing lineSpacing2 = new A.LineSpacing();
            A.SpacingPercent spacingPercent4 = new A.SpacingPercent(){ Val = 90000 };

            lineSpacing2.Append(spacingPercent4);

            A.SpaceBefore spaceBefore2 = new A.SpaceBefore();
            A.SpacingPercent spacingPercent5 = new A.SpacingPercent(){ Val = 0 };

            spaceBefore2.Append(spacingPercent5);

            A.SpaceAfter spaceAfter2 = new A.SpaceAfter();
            A.SpacingPercent spacingPercent6 = new A.SpacingPercent(){ Val = 35000 };

            spaceAfter2.Append(spacingPercent6);
            A.NoBullet noBullet2 = new A.NoBullet();

            paragraphProperties2.Append(lineSpacing2);
            paragraphProperties2.Append(spaceBefore2);
            paragraphProperties2.Append(spaceAfter2);
            paragraphProperties2.Append(noBullet2);
            A.EndParagraphRunProperties endParagraphRunProperties2 = new A.EndParagraphRunProperties(){ Language = "pl-PL", FontSize = 4700, Kerning = 1200 };

            paragraph3.Append(paragraphProperties2);
            paragraph3.Append(endParagraphRunProperties2);

            textBody2.Append(bodyProperties2);
            textBody2.Append(listStyle2);
            textBody2.Append(paragraph3);

            Dsp.Transform2D transform2D4 = new Dsp.Transform2D();
            A.Offset offset4 = new A.Offset(){ X = 1885950L, Y = 485774L };
            A.Extents extents4 = new A.Extents(){ Cx = 1714499L, Cy = 1028700L };

            transform2D4.Append(offset4);
            transform2D4.Append(extents4);

            shape2.Append(shapeNonVisualProperties2);
            shape2.Append(shapeProperties2);
            shape2.Append(shapeStyle2);
            shape2.Append(textBody2);
            shape2.Append(transform2D4);

            Dsp.Shape shape3 = new Dsp.Shape(){ ModelId = "{73C7BCEA-927D-4615-95E9-BB89F1A66540}" };

            Dsp.ShapeNonVisualProperties shapeNonVisualProperties3 = new Dsp.ShapeNonVisualProperties();
            Dsp.NonVisualDrawingProperties nonVisualDrawingProperties4 = new Dsp.NonVisualDrawingProperties(){ Id = (UInt32Value)0U, Name = "" };
            Dsp.NonVisualDrawingShapeProperties nonVisualDrawingShapeProperties3 = new Dsp.NonVisualDrawingShapeProperties();

            shapeNonVisualProperties3.Append(nonVisualDrawingProperties4);
            shapeNonVisualProperties3.Append(nonVisualDrawingShapeProperties3);

            Dsp.ShapeProperties shapeProperties3 = new Dsp.ShapeProperties();

            A.Transform2D transform2D5 = new A.Transform2D();
            A.Offset offset5 = new A.Offset(){ X = 3771900L, Y = 485774L };
            A.Extents extents5 = new A.Extents(){ Cx = 1714499L, Cy = 1028700L };

            transform2D5.Append(offset5);
            transform2D5.Append(extents5);

            A.PresetGeometry presetGeometry3 = new A.PresetGeometry(){ Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList3 = new A.AdjustValueList();

            presetGeometry3.Append(adjustValueList3);

            A.SolidFill solidFill5 = new A.SolidFill();

            A.SchemeColor schemeColor7 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent1 };
            A.HueOffset hueOffset5 = new A.HueOffset(){ Val = 0 };
            A.SaturationOffset saturationOffset5 = new A.SaturationOffset(){ Val = 0 };
            A.LuminanceOffset luminanceOffset5 = new A.LuminanceOffset(){ Val = 0 };
            A.AlphaOffset alphaOffset5 = new A.AlphaOffset(){ Val = 0 };

            schemeColor7.Append(hueOffset5);
            schemeColor7.Append(saturationOffset5);
            schemeColor7.Append(luminanceOffset5);
            schemeColor7.Append(alphaOffset5);

            solidFill5.Append(schemeColor7);

            A.Outline outline3 = new A.Outline(){ Width = 19050, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill6 = new A.SolidFill();

            A.SchemeColor schemeColor8 = new A.SchemeColor(){ Val = A.SchemeColorValues.Light1 };
            A.HueOffset hueOffset6 = new A.HueOffset(){ Val = 0 };
            A.SaturationOffset saturationOffset6 = new A.SaturationOffset(){ Val = 0 };
            A.LuminanceOffset luminanceOffset6 = new A.LuminanceOffset(){ Val = 0 };
            A.AlphaOffset alphaOffset6 = new A.AlphaOffset(){ Val = 0 };

            schemeColor8.Append(hueOffset6);
            schemeColor8.Append(saturationOffset6);
            schemeColor8.Append(luminanceOffset6);
            schemeColor8.Append(alphaOffset6);

            solidFill6.Append(schemeColor8);
            A.PresetDash presetDash3 = new A.PresetDash(){ Val = A.PresetLineDashValues.Solid };
            A.Miter miter3 = new A.Miter(){ Limit = 800000 };

            outline3.Append(solidFill6);
            outline3.Append(presetDash3);
            outline3.Append(miter3);
            A.EffectList effectList3 = new A.EffectList();

            shapeProperties3.Append(transform2D5);
            shapeProperties3.Append(presetGeometry3);
            shapeProperties3.Append(solidFill5);
            shapeProperties3.Append(outline3);
            shapeProperties3.Append(effectList3);

            Dsp.ShapeStyle shapeStyle3 = new Dsp.ShapeStyle();

            A.LineReference lineReference3 = new A.LineReference(){ Index = (UInt32Value)2U };
            A.RgbColorModelPercentage rgbColorModelPercentage7 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            lineReference3.Append(rgbColorModelPercentage7);

            A.FillReference fillReference3 = new A.FillReference(){ Index = (UInt32Value)1U };
            A.RgbColorModelPercentage rgbColorModelPercentage8 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            fillReference3.Append(rgbColorModelPercentage8);

            A.EffectReference effectReference3 = new A.EffectReference(){ Index = (UInt32Value)0U };
            A.RgbColorModelPercentage rgbColorModelPercentage9 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            effectReference3.Append(rgbColorModelPercentage9);

            A.FontReference fontReference3 = new A.FontReference(){ Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor9 = new A.SchemeColor(){ Val = A.SchemeColorValues.Light1 };

            fontReference3.Append(schemeColor9);

            shapeStyle3.Append(lineReference3);
            shapeStyle3.Append(fillReference3);
            shapeStyle3.Append(effectReference3);
            shapeStyle3.Append(fontReference3);

            Dsp.TextBody textBody3 = new Dsp.TextBody();

            A.BodyProperties bodyProperties3 = new A.BodyProperties(){ UseParagraphSpacing = false, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, LeftInset = 179070, TopInset = 179070, RightInset = 179070, BottomInset = 179070, ColumnCount = 1, ColumnSpacing = 1270, Anchor = A.TextAnchoringTypeValues.Center, AnchorCenter = false };
            A.NoAutoFit noAutoFit3 = new A.NoAutoFit();

            bodyProperties3.Append(noAutoFit3);
            A.ListStyle listStyle3 = new A.ListStyle();

            A.Paragraph paragraph4 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties3 = new A.ParagraphProperties(){ LeftMargin = 0, Level = 0, Indent = 0, Alignment = A.TextAlignmentTypeValues.Center, DefaultTabSize = 2089150 };

            A.LineSpacing lineSpacing3 = new A.LineSpacing();
            A.SpacingPercent spacingPercent7 = new A.SpacingPercent(){ Val = 90000 };

            lineSpacing3.Append(spacingPercent7);

            A.SpaceBefore spaceBefore3 = new A.SpaceBefore();
            A.SpacingPercent spacingPercent8 = new A.SpacingPercent(){ Val = 0 };

            spaceBefore3.Append(spacingPercent8);

            A.SpaceAfter spaceAfter3 = new A.SpaceAfter();
            A.SpacingPercent spacingPercent9 = new A.SpacingPercent(){ Val = 35000 };

            spaceAfter3.Append(spacingPercent9);
            A.NoBullet noBullet3 = new A.NoBullet();

            paragraphProperties3.Append(lineSpacing3);
            paragraphProperties3.Append(spaceBefore3);
            paragraphProperties3.Append(spaceAfter3);
            paragraphProperties3.Append(noBullet3);
            A.EndParagraphRunProperties endParagraphRunProperties3 = new A.EndParagraphRunProperties(){ Language = "pl-PL", FontSize = 4700, Kerning = 1200 };

            paragraph4.Append(paragraphProperties3);
            paragraph4.Append(endParagraphRunProperties3);

            textBody3.Append(bodyProperties3);
            textBody3.Append(listStyle3);
            textBody3.Append(paragraph4);

            Dsp.Transform2D transform2D6 = new Dsp.Transform2D();
            A.Offset offset6 = new A.Offset(){ X = 3771900L, Y = 485774L };
            A.Extents extents6 = new A.Extents(){ Cx = 1714499L, Cy = 1028700L };

            transform2D6.Append(offset6);
            transform2D6.Append(extents6);

            shape3.Append(shapeNonVisualProperties3);
            shape3.Append(shapeProperties3);
            shape3.Append(shapeStyle3);
            shape3.Append(textBody3);
            shape3.Append(transform2D6);

            Dsp.Shape shape4 = new Dsp.Shape(){ ModelId = "{72A7A719-1E8E-46AB-8256-BB41F6817818}" };

            Dsp.ShapeNonVisualProperties shapeNonVisualProperties4 = new Dsp.ShapeNonVisualProperties();
            Dsp.NonVisualDrawingProperties nonVisualDrawingProperties5 = new Dsp.NonVisualDrawingProperties(){ Id = (UInt32Value)0U, Name = "" };
            Dsp.NonVisualDrawingShapeProperties nonVisualDrawingShapeProperties4 = new Dsp.NonVisualDrawingShapeProperties();

            shapeNonVisualProperties4.Append(nonVisualDrawingProperties5);
            shapeNonVisualProperties4.Append(nonVisualDrawingShapeProperties4);

            Dsp.ShapeProperties shapeProperties4 = new Dsp.ShapeProperties();

            A.Transform2D transform2D7 = new A.Transform2D();
            A.Offset offset7 = new A.Offset(){ X = 942975L, Y = 1685925L };
            A.Extents extents7 = new A.Extents(){ Cx = 1714499L, Cy = 1028700L };

            transform2D7.Append(offset7);
            transform2D7.Append(extents7);

            A.PresetGeometry presetGeometry4 = new A.PresetGeometry(){ Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList4 = new A.AdjustValueList();

            presetGeometry4.Append(adjustValueList4);

            A.SolidFill solidFill7 = new A.SolidFill();

            A.SchemeColor schemeColor10 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent1 };
            A.HueOffset hueOffset7 = new A.HueOffset(){ Val = 0 };
            A.SaturationOffset saturationOffset7 = new A.SaturationOffset(){ Val = 0 };
            A.LuminanceOffset luminanceOffset7 = new A.LuminanceOffset(){ Val = 0 };
            A.AlphaOffset alphaOffset7 = new A.AlphaOffset(){ Val = 0 };

            schemeColor10.Append(hueOffset7);
            schemeColor10.Append(saturationOffset7);
            schemeColor10.Append(luminanceOffset7);
            schemeColor10.Append(alphaOffset7);

            solidFill7.Append(schemeColor10);

            A.Outline outline4 = new A.Outline(){ Width = 19050, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill8 = new A.SolidFill();

            A.SchemeColor schemeColor11 = new A.SchemeColor(){ Val = A.SchemeColorValues.Light1 };
            A.HueOffset hueOffset8 = new A.HueOffset(){ Val = 0 };
            A.SaturationOffset saturationOffset8 = new A.SaturationOffset(){ Val = 0 };
            A.LuminanceOffset luminanceOffset8 = new A.LuminanceOffset(){ Val = 0 };
            A.AlphaOffset alphaOffset8 = new A.AlphaOffset(){ Val = 0 };

            schemeColor11.Append(hueOffset8);
            schemeColor11.Append(saturationOffset8);
            schemeColor11.Append(luminanceOffset8);
            schemeColor11.Append(alphaOffset8);

            solidFill8.Append(schemeColor11);
            A.PresetDash presetDash4 = new A.PresetDash(){ Val = A.PresetLineDashValues.Solid };
            A.Miter miter4 = new A.Miter(){ Limit = 800000 };

            outline4.Append(solidFill8);
            outline4.Append(presetDash4);
            outline4.Append(miter4);
            A.EffectList effectList4 = new A.EffectList();

            shapeProperties4.Append(transform2D7);
            shapeProperties4.Append(presetGeometry4);
            shapeProperties4.Append(solidFill7);
            shapeProperties4.Append(outline4);
            shapeProperties4.Append(effectList4);

            Dsp.ShapeStyle shapeStyle4 = new Dsp.ShapeStyle();

            A.LineReference lineReference4 = new A.LineReference(){ Index = (UInt32Value)2U };
            A.RgbColorModelPercentage rgbColorModelPercentage10 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            lineReference4.Append(rgbColorModelPercentage10);

            A.FillReference fillReference4 = new A.FillReference(){ Index = (UInt32Value)1U };
            A.RgbColorModelPercentage rgbColorModelPercentage11 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            fillReference4.Append(rgbColorModelPercentage11);

            A.EffectReference effectReference4 = new A.EffectReference(){ Index = (UInt32Value)0U };
            A.RgbColorModelPercentage rgbColorModelPercentage12 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            effectReference4.Append(rgbColorModelPercentage12);

            A.FontReference fontReference4 = new A.FontReference(){ Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor12 = new A.SchemeColor(){ Val = A.SchemeColorValues.Light1 };

            fontReference4.Append(schemeColor12);

            shapeStyle4.Append(lineReference4);
            shapeStyle4.Append(fillReference4);
            shapeStyle4.Append(effectReference4);
            shapeStyle4.Append(fontReference4);

            Dsp.TextBody textBody4 = new Dsp.TextBody();

            A.BodyProperties bodyProperties4 = new A.BodyProperties(){ UseParagraphSpacing = false, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, LeftInset = 179070, TopInset = 179070, RightInset = 179070, BottomInset = 179070, ColumnCount = 1, ColumnSpacing = 1270, Anchor = A.TextAnchoringTypeValues.Center, AnchorCenter = false };
            A.NoAutoFit noAutoFit4 = new A.NoAutoFit();

            bodyProperties4.Append(noAutoFit4);
            A.ListStyle listStyle4 = new A.ListStyle();

            A.Paragraph paragraph5 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties4 = new A.ParagraphProperties(){ LeftMargin = 0, Level = 0, Indent = 0, Alignment = A.TextAlignmentTypeValues.Center, DefaultTabSize = 2089150 };

            A.LineSpacing lineSpacing4 = new A.LineSpacing();
            A.SpacingPercent spacingPercent10 = new A.SpacingPercent(){ Val = 90000 };

            lineSpacing4.Append(spacingPercent10);

            A.SpaceBefore spaceBefore4 = new A.SpaceBefore();
            A.SpacingPercent spacingPercent11 = new A.SpacingPercent(){ Val = 0 };

            spaceBefore4.Append(spacingPercent11);

            A.SpaceAfter spaceAfter4 = new A.SpaceAfter();
            A.SpacingPercent spacingPercent12 = new A.SpacingPercent(){ Val = 35000 };

            spaceAfter4.Append(spacingPercent12);
            A.NoBullet noBullet4 = new A.NoBullet();

            paragraphProperties4.Append(lineSpacing4);
            paragraphProperties4.Append(spaceBefore4);
            paragraphProperties4.Append(spaceAfter4);
            paragraphProperties4.Append(noBullet4);
            A.EndParagraphRunProperties endParagraphRunProperties4 = new A.EndParagraphRunProperties(){ Language = "pl-PL", FontSize = 4700, Kerning = 1200 };

            paragraph5.Append(paragraphProperties4);
            paragraph5.Append(endParagraphRunProperties4);

            textBody4.Append(bodyProperties4);
            textBody4.Append(listStyle4);
            textBody4.Append(paragraph5);

            Dsp.Transform2D transform2D8 = new Dsp.Transform2D();
            A.Offset offset8 = new A.Offset(){ X = 942975L, Y = 1685925L };
            A.Extents extents8 = new A.Extents(){ Cx = 1714499L, Cy = 1028700L };

            transform2D8.Append(offset8);
            transform2D8.Append(extents8);

            shape4.Append(shapeNonVisualProperties4);
            shape4.Append(shapeProperties4);
            shape4.Append(shapeStyle4);
            shape4.Append(textBody4);
            shape4.Append(transform2D8);

            Dsp.Shape shape5 = new Dsp.Shape(){ ModelId = "{2B5DA877-174D-4060-8E30-014EB5090235}" };

            Dsp.ShapeNonVisualProperties shapeNonVisualProperties5 = new Dsp.ShapeNonVisualProperties();
            Dsp.NonVisualDrawingProperties nonVisualDrawingProperties6 = new Dsp.NonVisualDrawingProperties(){ Id = (UInt32Value)0U, Name = "" };
            Dsp.NonVisualDrawingShapeProperties nonVisualDrawingShapeProperties5 = new Dsp.NonVisualDrawingShapeProperties();

            shapeNonVisualProperties5.Append(nonVisualDrawingProperties6);
            shapeNonVisualProperties5.Append(nonVisualDrawingShapeProperties5);

            Dsp.ShapeProperties shapeProperties5 = new Dsp.ShapeProperties();

            A.Transform2D transform2D9 = new A.Transform2D();
            A.Offset offset9 = new A.Offset(){ X = 2828925L, Y = 1685925L };
            A.Extents extents9 = new A.Extents(){ Cx = 1714499L, Cy = 1028700L };

            transform2D9.Append(offset9);
            transform2D9.Append(extents9);

            A.PresetGeometry presetGeometry5 = new A.PresetGeometry(){ Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList5 = new A.AdjustValueList();

            presetGeometry5.Append(adjustValueList5);

            A.SolidFill solidFill9 = new A.SolidFill();

            A.SchemeColor schemeColor13 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent1 };
            A.HueOffset hueOffset9 = new A.HueOffset(){ Val = 0 };
            A.SaturationOffset saturationOffset9 = new A.SaturationOffset(){ Val = 0 };
            A.LuminanceOffset luminanceOffset9 = new A.LuminanceOffset(){ Val = 0 };
            A.AlphaOffset alphaOffset9 = new A.AlphaOffset(){ Val = 0 };

            schemeColor13.Append(hueOffset9);
            schemeColor13.Append(saturationOffset9);
            schemeColor13.Append(luminanceOffset9);
            schemeColor13.Append(alphaOffset9);

            solidFill9.Append(schemeColor13);

            A.Outline outline5 = new A.Outline(){ Width = 19050, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill10 = new A.SolidFill();

            A.SchemeColor schemeColor14 = new A.SchemeColor(){ Val = A.SchemeColorValues.Light1 };
            A.HueOffset hueOffset10 = new A.HueOffset(){ Val = 0 };
            A.SaturationOffset saturationOffset10 = new A.SaturationOffset(){ Val = 0 };
            A.LuminanceOffset luminanceOffset10 = new A.LuminanceOffset(){ Val = 0 };
            A.AlphaOffset alphaOffset10 = new A.AlphaOffset(){ Val = 0 };

            schemeColor14.Append(hueOffset10);
            schemeColor14.Append(saturationOffset10);
            schemeColor14.Append(luminanceOffset10);
            schemeColor14.Append(alphaOffset10);

            solidFill10.Append(schemeColor14);
            A.PresetDash presetDash5 = new A.PresetDash(){ Val = A.PresetLineDashValues.Solid };
            A.Miter miter5 = new A.Miter(){ Limit = 800000 };

            outline5.Append(solidFill10);
            outline5.Append(presetDash5);
            outline5.Append(miter5);
            A.EffectList effectList5 = new A.EffectList();

            shapeProperties5.Append(transform2D9);
            shapeProperties5.Append(presetGeometry5);
            shapeProperties5.Append(solidFill9);
            shapeProperties5.Append(outline5);
            shapeProperties5.Append(effectList5);

            Dsp.ShapeStyle shapeStyle5 = new Dsp.ShapeStyle();

            A.LineReference lineReference5 = new A.LineReference(){ Index = (UInt32Value)2U };
            A.RgbColorModelPercentage rgbColorModelPercentage13 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            lineReference5.Append(rgbColorModelPercentage13);

            A.FillReference fillReference5 = new A.FillReference(){ Index = (UInt32Value)1U };
            A.RgbColorModelPercentage rgbColorModelPercentage14 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            fillReference5.Append(rgbColorModelPercentage14);

            A.EffectReference effectReference5 = new A.EffectReference(){ Index = (UInt32Value)0U };
            A.RgbColorModelPercentage rgbColorModelPercentage15 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            effectReference5.Append(rgbColorModelPercentage15);

            A.FontReference fontReference5 = new A.FontReference(){ Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor15 = new A.SchemeColor(){ Val = A.SchemeColorValues.Light1 };

            fontReference5.Append(schemeColor15);

            shapeStyle5.Append(lineReference5);
            shapeStyle5.Append(fillReference5);
            shapeStyle5.Append(effectReference5);
            shapeStyle5.Append(fontReference5);

            Dsp.TextBody textBody5 = new Dsp.TextBody();

            A.BodyProperties bodyProperties5 = new A.BodyProperties(){ UseParagraphSpacing = false, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, LeftInset = 179070, TopInset = 179070, RightInset = 179070, BottomInset = 179070, ColumnCount = 1, ColumnSpacing = 1270, Anchor = A.TextAnchoringTypeValues.Center, AnchorCenter = false };
            A.NoAutoFit noAutoFit5 = new A.NoAutoFit();

            bodyProperties5.Append(noAutoFit5);
            A.ListStyle listStyle5 = new A.ListStyle();

            A.Paragraph paragraph6 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties5 = new A.ParagraphProperties(){ LeftMargin = 0, Level = 0, Indent = 0, Alignment = A.TextAlignmentTypeValues.Center, DefaultTabSize = 2089150 };

            A.LineSpacing lineSpacing5 = new A.LineSpacing();
            A.SpacingPercent spacingPercent13 = new A.SpacingPercent(){ Val = 90000 };

            lineSpacing5.Append(spacingPercent13);

            A.SpaceBefore spaceBefore5 = new A.SpaceBefore();
            A.SpacingPercent spacingPercent14 = new A.SpacingPercent(){ Val = 0 };

            spaceBefore5.Append(spacingPercent14);

            A.SpaceAfter spaceAfter5 = new A.SpaceAfter();
            A.SpacingPercent spacingPercent15 = new A.SpacingPercent(){ Val = 35000 };

            spaceAfter5.Append(spacingPercent15);
            A.NoBullet noBullet5 = new A.NoBullet();

            paragraphProperties5.Append(lineSpacing5);
            paragraphProperties5.Append(spaceBefore5);
            paragraphProperties5.Append(spaceAfter5);
            paragraphProperties5.Append(noBullet5);
            A.EndParagraphRunProperties endParagraphRunProperties5 = new A.EndParagraphRunProperties(){ Language = "pl-PL", FontSize = 4700, Kerning = 1200 };

            paragraph6.Append(paragraphProperties5);
            paragraph6.Append(endParagraphRunProperties5);

            textBody5.Append(bodyProperties5);
            textBody5.Append(listStyle5);
            textBody5.Append(paragraph6);

            Dsp.Transform2D transform2D10 = new Dsp.Transform2D();
            A.Offset offset10 = new A.Offset(){ X = 2828925L, Y = 1685925L };
            A.Extents extents10 = new A.Extents(){ Cx = 1714499L, Cy = 1028700L };

            transform2D10.Append(offset10);
            transform2D10.Append(extents10);

            shape5.Append(shapeNonVisualProperties5);
            shape5.Append(shapeProperties5);
            shape5.Append(shapeStyle5);
            shape5.Append(textBody5);
            shape5.Append(transform2D10);

            shapeTree1.Append(groupShapeNonVisualProperties1);
            shapeTree1.Append(groupShapeProperties1);
            shapeTree1.Append(shape1);
            shapeTree1.Append(shape2);
            shapeTree1.Append(shape3);
            shapeTree1.Append(shape4);
            shapeTree1.Append(shape5);

            drawing2.Append(shapeTree1);

            diagramPersistLayoutPart1.Drawing = drawing2;
        }

        // Generates content of webSettingsPart1.
        private void GenerateWebSettingsPart1Content(WebSettingsPart webSettingsPart1)
        {
            WebSettings webSettings1 = new WebSettings(){ MCAttributes = new MarkupCompatibilityAttributes(){ Ignorable = "w14 w15 w16se w16cid w16 w16cex w16sdtdh w16sdtfl w16du" }  };
            webSettings1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            webSettings1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            webSettings1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            webSettings1.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            webSettings1.AddNamespaceDeclaration("w15", "http://schemas.microsoft.com/office/word/2012/wordml");
            webSettings1.AddNamespaceDeclaration("w16cex", "http://schemas.microsoft.com/office/word/2018/wordml/cex");
            webSettings1.AddNamespaceDeclaration("w16cid", "http://schemas.microsoft.com/office/word/2016/wordml/cid");
            webSettings1.AddNamespaceDeclaration("w16", "http://schemas.microsoft.com/office/word/2018/wordml");
            webSettings1.AddNamespaceDeclaration("w16du", "http://schemas.microsoft.com/office/word/2023/wordml/word16du");
            webSettings1.AddNamespaceDeclaration("w16sdtdh", "http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash");
            webSettings1.AddNamespaceDeclaration("w16sdtfl", "http://schemas.microsoft.com/office/word/2024/wordml/sdtformatlock");
            webSettings1.AddNamespaceDeclaration("w16se", "http://schemas.microsoft.com/office/word/2015/wordml/symex");
            OptimizeForBrowser optimizeForBrowser1 = new OptimizeForBrowser();
            AllowPNG allowPNG1 = new AllowPNG();

            webSettings1.Append(optimizeForBrowser1);
            webSettings1.Append(allowPNG1);

            webSettingsPart1.WebSettings = webSettings1;
        }

        // Generates content of diagramColorsPart1.
        private void GenerateDiagramColorsPart1Content(DiagramColorsPart diagramColorsPart1)
        {
            Dgm.ColorsDefinition colorsDefinition1 = new Dgm.ColorsDefinition(){ UniqueId = "urn:microsoft.com/office/officeart/2005/8/colors/accent1_2" };
            colorsDefinition1.AddNamespaceDeclaration("dgm", "http://schemas.openxmlformats.org/drawingml/2006/diagram");
            colorsDefinition1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            Dgm.ColorDefinitionTitle colorDefinitionTitle1 = new Dgm.ColorDefinitionTitle(){ Val = "" };
            Dgm.ColorTransformDescription colorTransformDescription1 = new Dgm.ColorTransformDescription(){ Val = "" };

            Dgm.ColorTransformCategories colorTransformCategories1 = new Dgm.ColorTransformCategories();
            Dgm.ColorTransformCategory colorTransformCategory1 = new Dgm.ColorTransformCategory(){ Type = "accent1", Priority = (UInt32Value)11200U };

            colorTransformCategories1.Append(colorTransformCategory1);

            Dgm.ColorTransformStyleLabel colorTransformStyleLabel1 = new Dgm.ColorTransformStyleLabel(){ Name = "node0" };

            Dgm.FillColorList fillColorList1 = new Dgm.FillColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };
            A.SchemeColor schemeColor16 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent1 };

            fillColorList1.Append(schemeColor16);

            Dgm.LineColorList lineColorList1 = new Dgm.LineColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };
            A.SchemeColor schemeColor17 = new A.SchemeColor(){ Val = A.SchemeColorValues.Light1 };

            lineColorList1.Append(schemeColor17);
            Dgm.EffectColorList effectColorList1 = new Dgm.EffectColorList();
            Dgm.TextLineColorList textLineColorList1 = new Dgm.TextLineColorList();
            Dgm.TextFillColorList textFillColorList1 = new Dgm.TextFillColorList();
            Dgm.TextEffectColorList textEffectColorList1 = new Dgm.TextEffectColorList();

            colorTransformStyleLabel1.Append(fillColorList1);
            colorTransformStyleLabel1.Append(lineColorList1);
            colorTransformStyleLabel1.Append(effectColorList1);
            colorTransformStyleLabel1.Append(textLineColorList1);
            colorTransformStyleLabel1.Append(textFillColorList1);
            colorTransformStyleLabel1.Append(textEffectColorList1);

            Dgm.ColorTransformStyleLabel colorTransformStyleLabel2 = new Dgm.ColorTransformStyleLabel(){ Name = "alignNode1" };

            Dgm.FillColorList fillColorList2 = new Dgm.FillColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };
            A.SchemeColor schemeColor18 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent1 };

            fillColorList2.Append(schemeColor18);

            Dgm.LineColorList lineColorList2 = new Dgm.LineColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };
            A.SchemeColor schemeColor19 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent1 };

            lineColorList2.Append(schemeColor19);
            Dgm.EffectColorList effectColorList2 = new Dgm.EffectColorList();
            Dgm.TextLineColorList textLineColorList2 = new Dgm.TextLineColorList();
            Dgm.TextFillColorList textFillColorList2 = new Dgm.TextFillColorList();
            Dgm.TextEffectColorList textEffectColorList2 = new Dgm.TextEffectColorList();

            colorTransformStyleLabel2.Append(fillColorList2);
            colorTransformStyleLabel2.Append(lineColorList2);
            colorTransformStyleLabel2.Append(effectColorList2);
            colorTransformStyleLabel2.Append(textLineColorList2);
            colorTransformStyleLabel2.Append(textFillColorList2);
            colorTransformStyleLabel2.Append(textEffectColorList2);

            Dgm.ColorTransformStyleLabel colorTransformStyleLabel3 = new Dgm.ColorTransformStyleLabel(){ Name = "node1" };

            Dgm.FillColorList fillColorList3 = new Dgm.FillColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };
            A.SchemeColor schemeColor20 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent1 };

            fillColorList3.Append(schemeColor20);

            Dgm.LineColorList lineColorList3 = new Dgm.LineColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };
            A.SchemeColor schemeColor21 = new A.SchemeColor(){ Val = A.SchemeColorValues.Light1 };

            lineColorList3.Append(schemeColor21);
            Dgm.EffectColorList effectColorList3 = new Dgm.EffectColorList();
            Dgm.TextLineColorList textLineColorList3 = new Dgm.TextLineColorList();
            Dgm.TextFillColorList textFillColorList3 = new Dgm.TextFillColorList();
            Dgm.TextEffectColorList textEffectColorList3 = new Dgm.TextEffectColorList();

            colorTransformStyleLabel3.Append(fillColorList3);
            colorTransformStyleLabel3.Append(lineColorList3);
            colorTransformStyleLabel3.Append(effectColorList3);
            colorTransformStyleLabel3.Append(textLineColorList3);
            colorTransformStyleLabel3.Append(textFillColorList3);
            colorTransformStyleLabel3.Append(textEffectColorList3);

            Dgm.ColorTransformStyleLabel colorTransformStyleLabel4 = new Dgm.ColorTransformStyleLabel(){ Name = "lnNode1" };

            Dgm.FillColorList fillColorList4 = new Dgm.FillColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };
            A.SchemeColor schemeColor22 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent1 };

            fillColorList4.Append(schemeColor22);

            Dgm.LineColorList lineColorList4 = new Dgm.LineColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };
            A.SchemeColor schemeColor23 = new A.SchemeColor(){ Val = A.SchemeColorValues.Light1 };

            lineColorList4.Append(schemeColor23);
            Dgm.EffectColorList effectColorList4 = new Dgm.EffectColorList();
            Dgm.TextLineColorList textLineColorList4 = new Dgm.TextLineColorList();
            Dgm.TextFillColorList textFillColorList4 = new Dgm.TextFillColorList();
            Dgm.TextEffectColorList textEffectColorList4 = new Dgm.TextEffectColorList();

            colorTransformStyleLabel4.Append(fillColorList4);
            colorTransformStyleLabel4.Append(lineColorList4);
            colorTransformStyleLabel4.Append(effectColorList4);
            colorTransformStyleLabel4.Append(textLineColorList4);
            colorTransformStyleLabel4.Append(textFillColorList4);
            colorTransformStyleLabel4.Append(textEffectColorList4);

            Dgm.ColorTransformStyleLabel colorTransformStyleLabel5 = new Dgm.ColorTransformStyleLabel(){ Name = "vennNode1" };

            Dgm.FillColorList fillColorList5 = new Dgm.FillColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };

            A.SchemeColor schemeColor24 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent1 };
            A.Alpha alpha1 = new A.Alpha(){ Val = 50000 };

            schemeColor24.Append(alpha1);

            fillColorList5.Append(schemeColor24);

            Dgm.LineColorList lineColorList5 = new Dgm.LineColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };
            A.SchemeColor schemeColor25 = new A.SchemeColor(){ Val = A.SchemeColorValues.Light1 };

            lineColorList5.Append(schemeColor25);
            Dgm.EffectColorList effectColorList5 = new Dgm.EffectColorList();
            Dgm.TextLineColorList textLineColorList5 = new Dgm.TextLineColorList();
            Dgm.TextFillColorList textFillColorList5 = new Dgm.TextFillColorList();
            Dgm.TextEffectColorList textEffectColorList5 = new Dgm.TextEffectColorList();

            colorTransformStyleLabel5.Append(fillColorList5);
            colorTransformStyleLabel5.Append(lineColorList5);
            colorTransformStyleLabel5.Append(effectColorList5);
            colorTransformStyleLabel5.Append(textLineColorList5);
            colorTransformStyleLabel5.Append(textFillColorList5);
            colorTransformStyleLabel5.Append(textEffectColorList5);

            Dgm.ColorTransformStyleLabel colorTransformStyleLabel6 = new Dgm.ColorTransformStyleLabel(){ Name = "node2" };

            Dgm.FillColorList fillColorList6 = new Dgm.FillColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };
            A.SchemeColor schemeColor26 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent1 };

            fillColorList6.Append(schemeColor26);

            Dgm.LineColorList lineColorList6 = new Dgm.LineColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };
            A.SchemeColor schemeColor27 = new A.SchemeColor(){ Val = A.SchemeColorValues.Light1 };

            lineColorList6.Append(schemeColor27);
            Dgm.EffectColorList effectColorList6 = new Dgm.EffectColorList();
            Dgm.TextLineColorList textLineColorList6 = new Dgm.TextLineColorList();
            Dgm.TextFillColorList textFillColorList6 = new Dgm.TextFillColorList();
            Dgm.TextEffectColorList textEffectColorList6 = new Dgm.TextEffectColorList();

            colorTransformStyleLabel6.Append(fillColorList6);
            colorTransformStyleLabel6.Append(lineColorList6);
            colorTransformStyleLabel6.Append(effectColorList6);
            colorTransformStyleLabel6.Append(textLineColorList6);
            colorTransformStyleLabel6.Append(textFillColorList6);
            colorTransformStyleLabel6.Append(textEffectColorList6);

            Dgm.ColorTransformStyleLabel colorTransformStyleLabel7 = new Dgm.ColorTransformStyleLabel(){ Name = "node3" };

            Dgm.FillColorList fillColorList7 = new Dgm.FillColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };
            A.SchemeColor schemeColor28 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent1 };

            fillColorList7.Append(schemeColor28);

            Dgm.LineColorList lineColorList7 = new Dgm.LineColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };
            A.SchemeColor schemeColor29 = new A.SchemeColor(){ Val = A.SchemeColorValues.Light1 };

            lineColorList7.Append(schemeColor29);
            Dgm.EffectColorList effectColorList7 = new Dgm.EffectColorList();
            Dgm.TextLineColorList textLineColorList7 = new Dgm.TextLineColorList();
            Dgm.TextFillColorList textFillColorList7 = new Dgm.TextFillColorList();
            Dgm.TextEffectColorList textEffectColorList7 = new Dgm.TextEffectColorList();

            colorTransformStyleLabel7.Append(fillColorList7);
            colorTransformStyleLabel7.Append(lineColorList7);
            colorTransformStyleLabel7.Append(effectColorList7);
            colorTransformStyleLabel7.Append(textLineColorList7);
            colorTransformStyleLabel7.Append(textFillColorList7);
            colorTransformStyleLabel7.Append(textEffectColorList7);

            Dgm.ColorTransformStyleLabel colorTransformStyleLabel8 = new Dgm.ColorTransformStyleLabel(){ Name = "node4" };

            Dgm.FillColorList fillColorList8 = new Dgm.FillColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };
            A.SchemeColor schemeColor30 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent1 };

            fillColorList8.Append(schemeColor30);

            Dgm.LineColorList lineColorList8 = new Dgm.LineColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };
            A.SchemeColor schemeColor31 = new A.SchemeColor(){ Val = A.SchemeColorValues.Light1 };

            lineColorList8.Append(schemeColor31);
            Dgm.EffectColorList effectColorList8 = new Dgm.EffectColorList();
            Dgm.TextLineColorList textLineColorList8 = new Dgm.TextLineColorList();
            Dgm.TextFillColorList textFillColorList8 = new Dgm.TextFillColorList();
            Dgm.TextEffectColorList textEffectColorList8 = new Dgm.TextEffectColorList();

            colorTransformStyleLabel8.Append(fillColorList8);
            colorTransformStyleLabel8.Append(lineColorList8);
            colorTransformStyleLabel8.Append(effectColorList8);
            colorTransformStyleLabel8.Append(textLineColorList8);
            colorTransformStyleLabel8.Append(textFillColorList8);
            colorTransformStyleLabel8.Append(textEffectColorList8);

            Dgm.ColorTransformStyleLabel colorTransformStyleLabel9 = new Dgm.ColorTransformStyleLabel(){ Name = "fgImgPlace1" };

            Dgm.FillColorList fillColorList9 = new Dgm.FillColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };

            A.SchemeColor schemeColor32 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent1 };
            A.Tint tint1 = new A.Tint(){ Val = 50000 };

            schemeColor32.Append(tint1);

            fillColorList9.Append(schemeColor32);

            Dgm.LineColorList lineColorList9 = new Dgm.LineColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };
            A.SchemeColor schemeColor33 = new A.SchemeColor(){ Val = A.SchemeColorValues.Light1 };

            lineColorList9.Append(schemeColor33);
            Dgm.EffectColorList effectColorList9 = new Dgm.EffectColorList();
            Dgm.TextLineColorList textLineColorList9 = new Dgm.TextLineColorList();

            Dgm.TextFillColorList textFillColorList9 = new Dgm.TextFillColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };
            A.SchemeColor schemeColor34 = new A.SchemeColor(){ Val = A.SchemeColorValues.Light1 };

            textFillColorList9.Append(schemeColor34);
            Dgm.TextEffectColorList textEffectColorList9 = new Dgm.TextEffectColorList();

            colorTransformStyleLabel9.Append(fillColorList9);
            colorTransformStyleLabel9.Append(lineColorList9);
            colorTransformStyleLabel9.Append(effectColorList9);
            colorTransformStyleLabel9.Append(textLineColorList9);
            colorTransformStyleLabel9.Append(textFillColorList9);
            colorTransformStyleLabel9.Append(textEffectColorList9);

            Dgm.ColorTransformStyleLabel colorTransformStyleLabel10 = new Dgm.ColorTransformStyleLabel(){ Name = "alignImgPlace1" };

            Dgm.FillColorList fillColorList10 = new Dgm.FillColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };

            A.SchemeColor schemeColor35 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent1 };
            A.Tint tint2 = new A.Tint(){ Val = 50000 };

            schemeColor35.Append(tint2);

            fillColorList10.Append(schemeColor35);

            Dgm.LineColorList lineColorList10 = new Dgm.LineColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };
            A.SchemeColor schemeColor36 = new A.SchemeColor(){ Val = A.SchemeColorValues.Light1 };

            lineColorList10.Append(schemeColor36);
            Dgm.EffectColorList effectColorList10 = new Dgm.EffectColorList();
            Dgm.TextLineColorList textLineColorList10 = new Dgm.TextLineColorList();

            Dgm.TextFillColorList textFillColorList10 = new Dgm.TextFillColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };
            A.SchemeColor schemeColor37 = new A.SchemeColor(){ Val = A.SchemeColorValues.Light1 };

            textFillColorList10.Append(schemeColor37);
            Dgm.TextEffectColorList textEffectColorList10 = new Dgm.TextEffectColorList();

            colorTransformStyleLabel10.Append(fillColorList10);
            colorTransformStyleLabel10.Append(lineColorList10);
            colorTransformStyleLabel10.Append(effectColorList10);
            colorTransformStyleLabel10.Append(textLineColorList10);
            colorTransformStyleLabel10.Append(textFillColorList10);
            colorTransformStyleLabel10.Append(textEffectColorList10);

            Dgm.ColorTransformStyleLabel colorTransformStyleLabel11 = new Dgm.ColorTransformStyleLabel(){ Name = "bgImgPlace1" };

            Dgm.FillColorList fillColorList11 = new Dgm.FillColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };

            A.SchemeColor schemeColor38 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent1 };
            A.Tint tint3 = new A.Tint(){ Val = 50000 };

            schemeColor38.Append(tint3);

            fillColorList11.Append(schemeColor38);

            Dgm.LineColorList lineColorList11 = new Dgm.LineColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };
            A.SchemeColor schemeColor39 = new A.SchemeColor(){ Val = A.SchemeColorValues.Light1 };

            lineColorList11.Append(schemeColor39);
            Dgm.EffectColorList effectColorList11 = new Dgm.EffectColorList();
            Dgm.TextLineColorList textLineColorList11 = new Dgm.TextLineColorList();

            Dgm.TextFillColorList textFillColorList11 = new Dgm.TextFillColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };
            A.SchemeColor schemeColor40 = new A.SchemeColor(){ Val = A.SchemeColorValues.Light1 };

            textFillColorList11.Append(schemeColor40);
            Dgm.TextEffectColorList textEffectColorList11 = new Dgm.TextEffectColorList();

            colorTransformStyleLabel11.Append(fillColorList11);
            colorTransformStyleLabel11.Append(lineColorList11);
            colorTransformStyleLabel11.Append(effectColorList11);
            colorTransformStyleLabel11.Append(textLineColorList11);
            colorTransformStyleLabel11.Append(textFillColorList11);
            colorTransformStyleLabel11.Append(textEffectColorList11);

            Dgm.ColorTransformStyleLabel colorTransformStyleLabel12 = new Dgm.ColorTransformStyleLabel(){ Name = "sibTrans2D1" };

            Dgm.FillColorList fillColorList12 = new Dgm.FillColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };

            A.SchemeColor schemeColor41 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent1 };
            A.Tint tint4 = new A.Tint(){ Val = 60000 };

            schemeColor41.Append(tint4);

            fillColorList12.Append(schemeColor41);

            Dgm.LineColorList lineColorList12 = new Dgm.LineColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };

            A.SchemeColor schemeColor42 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent1 };
            A.Tint tint5 = new A.Tint(){ Val = 60000 };

            schemeColor42.Append(tint5);

            lineColorList12.Append(schemeColor42);
            Dgm.EffectColorList effectColorList12 = new Dgm.EffectColorList();
            Dgm.TextLineColorList textLineColorList12 = new Dgm.TextLineColorList();
            Dgm.TextFillColorList textFillColorList12 = new Dgm.TextFillColorList();
            Dgm.TextEffectColorList textEffectColorList12 = new Dgm.TextEffectColorList();

            colorTransformStyleLabel12.Append(fillColorList12);
            colorTransformStyleLabel12.Append(lineColorList12);
            colorTransformStyleLabel12.Append(effectColorList12);
            colorTransformStyleLabel12.Append(textLineColorList12);
            colorTransformStyleLabel12.Append(textFillColorList12);
            colorTransformStyleLabel12.Append(textEffectColorList12);

            Dgm.ColorTransformStyleLabel colorTransformStyleLabel13 = new Dgm.ColorTransformStyleLabel(){ Name = "fgSibTrans2D1" };

            Dgm.FillColorList fillColorList13 = new Dgm.FillColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };

            A.SchemeColor schemeColor43 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent1 };
            A.Tint tint6 = new A.Tint(){ Val = 60000 };

            schemeColor43.Append(tint6);

            fillColorList13.Append(schemeColor43);

            Dgm.LineColorList lineColorList13 = new Dgm.LineColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };

            A.SchemeColor schemeColor44 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent1 };
            A.Tint tint7 = new A.Tint(){ Val = 60000 };

            schemeColor44.Append(tint7);

            lineColorList13.Append(schemeColor44);
            Dgm.EffectColorList effectColorList13 = new Dgm.EffectColorList();
            Dgm.TextLineColorList textLineColorList13 = new Dgm.TextLineColorList();
            Dgm.TextFillColorList textFillColorList13 = new Dgm.TextFillColorList();
            Dgm.TextEffectColorList textEffectColorList13 = new Dgm.TextEffectColorList();

            colorTransformStyleLabel13.Append(fillColorList13);
            colorTransformStyleLabel13.Append(lineColorList13);
            colorTransformStyleLabel13.Append(effectColorList13);
            colorTransformStyleLabel13.Append(textLineColorList13);
            colorTransformStyleLabel13.Append(textFillColorList13);
            colorTransformStyleLabel13.Append(textEffectColorList13);

            Dgm.ColorTransformStyleLabel colorTransformStyleLabel14 = new Dgm.ColorTransformStyleLabel(){ Name = "bgSibTrans2D1" };

            Dgm.FillColorList fillColorList14 = new Dgm.FillColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };

            A.SchemeColor schemeColor45 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent1 };
            A.Tint tint8 = new A.Tint(){ Val = 60000 };

            schemeColor45.Append(tint8);

            fillColorList14.Append(schemeColor45);

            Dgm.LineColorList lineColorList14 = new Dgm.LineColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };

            A.SchemeColor schemeColor46 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent1 };
            A.Tint tint9 = new A.Tint(){ Val = 60000 };

            schemeColor46.Append(tint9);

            lineColorList14.Append(schemeColor46);
            Dgm.EffectColorList effectColorList14 = new Dgm.EffectColorList();
            Dgm.TextLineColorList textLineColorList14 = new Dgm.TextLineColorList();
            Dgm.TextFillColorList textFillColorList14 = new Dgm.TextFillColorList();
            Dgm.TextEffectColorList textEffectColorList14 = new Dgm.TextEffectColorList();

            colorTransformStyleLabel14.Append(fillColorList14);
            colorTransformStyleLabel14.Append(lineColorList14);
            colorTransformStyleLabel14.Append(effectColorList14);
            colorTransformStyleLabel14.Append(textLineColorList14);
            colorTransformStyleLabel14.Append(textFillColorList14);
            colorTransformStyleLabel14.Append(textEffectColorList14);

            Dgm.ColorTransformStyleLabel colorTransformStyleLabel15 = new Dgm.ColorTransformStyleLabel(){ Name = "sibTrans1D1" };

            Dgm.FillColorList fillColorList15 = new Dgm.FillColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };
            A.SchemeColor schemeColor47 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent1 };

            fillColorList15.Append(schemeColor47);

            Dgm.LineColorList lineColorList15 = new Dgm.LineColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };
            A.SchemeColor schemeColor48 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent1 };

            lineColorList15.Append(schemeColor48);
            Dgm.EffectColorList effectColorList15 = new Dgm.EffectColorList();
            Dgm.TextLineColorList textLineColorList15 = new Dgm.TextLineColorList();

            Dgm.TextFillColorList textFillColorList15 = new Dgm.TextFillColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };
            A.SchemeColor schemeColor49 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            textFillColorList15.Append(schemeColor49);
            Dgm.TextEffectColorList textEffectColorList15 = new Dgm.TextEffectColorList();

            colorTransformStyleLabel15.Append(fillColorList15);
            colorTransformStyleLabel15.Append(lineColorList15);
            colorTransformStyleLabel15.Append(effectColorList15);
            colorTransformStyleLabel15.Append(textLineColorList15);
            colorTransformStyleLabel15.Append(textFillColorList15);
            colorTransformStyleLabel15.Append(textEffectColorList15);

            Dgm.ColorTransformStyleLabel colorTransformStyleLabel16 = new Dgm.ColorTransformStyleLabel(){ Name = "callout" };

            Dgm.FillColorList fillColorList16 = new Dgm.FillColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };
            A.SchemeColor schemeColor50 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent1 };

            fillColorList16.Append(schemeColor50);

            Dgm.LineColorList lineColorList16 = new Dgm.LineColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };

            A.SchemeColor schemeColor51 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent1 };
            A.Tint tint10 = new A.Tint(){ Val = 50000 };

            schemeColor51.Append(tint10);

            lineColorList16.Append(schemeColor51);
            Dgm.EffectColorList effectColorList16 = new Dgm.EffectColorList();
            Dgm.TextLineColorList textLineColorList16 = new Dgm.TextLineColorList();

            Dgm.TextFillColorList textFillColorList16 = new Dgm.TextFillColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };
            A.SchemeColor schemeColor52 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            textFillColorList16.Append(schemeColor52);
            Dgm.TextEffectColorList textEffectColorList16 = new Dgm.TextEffectColorList();

            colorTransformStyleLabel16.Append(fillColorList16);
            colorTransformStyleLabel16.Append(lineColorList16);
            colorTransformStyleLabel16.Append(effectColorList16);
            colorTransformStyleLabel16.Append(textLineColorList16);
            colorTransformStyleLabel16.Append(textFillColorList16);
            colorTransformStyleLabel16.Append(textEffectColorList16);

            Dgm.ColorTransformStyleLabel colorTransformStyleLabel17 = new Dgm.ColorTransformStyleLabel(){ Name = "asst0" };

            Dgm.FillColorList fillColorList17 = new Dgm.FillColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };
            A.SchemeColor schemeColor53 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent1 };

            fillColorList17.Append(schemeColor53);

            Dgm.LineColorList lineColorList17 = new Dgm.LineColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };
            A.SchemeColor schemeColor54 = new A.SchemeColor(){ Val = A.SchemeColorValues.Light1 };

            lineColorList17.Append(schemeColor54);
            Dgm.EffectColorList effectColorList17 = new Dgm.EffectColorList();
            Dgm.TextLineColorList textLineColorList17 = new Dgm.TextLineColorList();
            Dgm.TextFillColorList textFillColorList17 = new Dgm.TextFillColorList();
            Dgm.TextEffectColorList textEffectColorList17 = new Dgm.TextEffectColorList();

            colorTransformStyleLabel17.Append(fillColorList17);
            colorTransformStyleLabel17.Append(lineColorList17);
            colorTransformStyleLabel17.Append(effectColorList17);
            colorTransformStyleLabel17.Append(textLineColorList17);
            colorTransformStyleLabel17.Append(textFillColorList17);
            colorTransformStyleLabel17.Append(textEffectColorList17);

            Dgm.ColorTransformStyleLabel colorTransformStyleLabel18 = new Dgm.ColorTransformStyleLabel(){ Name = "asst1" };

            Dgm.FillColorList fillColorList18 = new Dgm.FillColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };
            A.SchemeColor schemeColor55 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent1 };

            fillColorList18.Append(schemeColor55);

            Dgm.LineColorList lineColorList18 = new Dgm.LineColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };
            A.SchemeColor schemeColor56 = new A.SchemeColor(){ Val = A.SchemeColorValues.Light1 };

            lineColorList18.Append(schemeColor56);
            Dgm.EffectColorList effectColorList18 = new Dgm.EffectColorList();
            Dgm.TextLineColorList textLineColorList18 = new Dgm.TextLineColorList();
            Dgm.TextFillColorList textFillColorList18 = new Dgm.TextFillColorList();
            Dgm.TextEffectColorList textEffectColorList18 = new Dgm.TextEffectColorList();

            colorTransformStyleLabel18.Append(fillColorList18);
            colorTransformStyleLabel18.Append(lineColorList18);
            colorTransformStyleLabel18.Append(effectColorList18);
            colorTransformStyleLabel18.Append(textLineColorList18);
            colorTransformStyleLabel18.Append(textFillColorList18);
            colorTransformStyleLabel18.Append(textEffectColorList18);

            Dgm.ColorTransformStyleLabel colorTransformStyleLabel19 = new Dgm.ColorTransformStyleLabel(){ Name = "asst2" };

            Dgm.FillColorList fillColorList19 = new Dgm.FillColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };
            A.SchemeColor schemeColor57 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent1 };

            fillColorList19.Append(schemeColor57);

            Dgm.LineColorList lineColorList19 = new Dgm.LineColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };
            A.SchemeColor schemeColor58 = new A.SchemeColor(){ Val = A.SchemeColorValues.Light1 };

            lineColorList19.Append(schemeColor58);
            Dgm.EffectColorList effectColorList19 = new Dgm.EffectColorList();
            Dgm.TextLineColorList textLineColorList19 = new Dgm.TextLineColorList();
            Dgm.TextFillColorList textFillColorList19 = new Dgm.TextFillColorList();
            Dgm.TextEffectColorList textEffectColorList19 = new Dgm.TextEffectColorList();

            colorTransformStyleLabel19.Append(fillColorList19);
            colorTransformStyleLabel19.Append(lineColorList19);
            colorTransformStyleLabel19.Append(effectColorList19);
            colorTransformStyleLabel19.Append(textLineColorList19);
            colorTransformStyleLabel19.Append(textFillColorList19);
            colorTransformStyleLabel19.Append(textEffectColorList19);

            Dgm.ColorTransformStyleLabel colorTransformStyleLabel20 = new Dgm.ColorTransformStyleLabel(){ Name = "asst3" };

            Dgm.FillColorList fillColorList20 = new Dgm.FillColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };
            A.SchemeColor schemeColor59 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent1 };

            fillColorList20.Append(schemeColor59);

            Dgm.LineColorList lineColorList20 = new Dgm.LineColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };
            A.SchemeColor schemeColor60 = new A.SchemeColor(){ Val = A.SchemeColorValues.Light1 };

            lineColorList20.Append(schemeColor60);
            Dgm.EffectColorList effectColorList20 = new Dgm.EffectColorList();
            Dgm.TextLineColorList textLineColorList20 = new Dgm.TextLineColorList();
            Dgm.TextFillColorList textFillColorList20 = new Dgm.TextFillColorList();
            Dgm.TextEffectColorList textEffectColorList20 = new Dgm.TextEffectColorList();

            colorTransformStyleLabel20.Append(fillColorList20);
            colorTransformStyleLabel20.Append(lineColorList20);
            colorTransformStyleLabel20.Append(effectColorList20);
            colorTransformStyleLabel20.Append(textLineColorList20);
            colorTransformStyleLabel20.Append(textFillColorList20);
            colorTransformStyleLabel20.Append(textEffectColorList20);

            Dgm.ColorTransformStyleLabel colorTransformStyleLabel21 = new Dgm.ColorTransformStyleLabel(){ Name = "asst4" };

            Dgm.FillColorList fillColorList21 = new Dgm.FillColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };
            A.SchemeColor schemeColor61 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent1 };

            fillColorList21.Append(schemeColor61);

            Dgm.LineColorList lineColorList21 = new Dgm.LineColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };
            A.SchemeColor schemeColor62 = new A.SchemeColor(){ Val = A.SchemeColorValues.Light1 };

            lineColorList21.Append(schemeColor62);
            Dgm.EffectColorList effectColorList21 = new Dgm.EffectColorList();
            Dgm.TextLineColorList textLineColorList21 = new Dgm.TextLineColorList();
            Dgm.TextFillColorList textFillColorList21 = new Dgm.TextFillColorList();
            Dgm.TextEffectColorList textEffectColorList21 = new Dgm.TextEffectColorList();

            colorTransformStyleLabel21.Append(fillColorList21);
            colorTransformStyleLabel21.Append(lineColorList21);
            colorTransformStyleLabel21.Append(effectColorList21);
            colorTransformStyleLabel21.Append(textLineColorList21);
            colorTransformStyleLabel21.Append(textFillColorList21);
            colorTransformStyleLabel21.Append(textEffectColorList21);

            Dgm.ColorTransformStyleLabel colorTransformStyleLabel22 = new Dgm.ColorTransformStyleLabel(){ Name = "parChTrans2D1" };

            Dgm.FillColorList fillColorList22 = new Dgm.FillColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };

            A.SchemeColor schemeColor63 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent1 };
            A.Tint tint11 = new A.Tint(){ Val = 60000 };

            schemeColor63.Append(tint11);

            fillColorList22.Append(schemeColor63);

            Dgm.LineColorList lineColorList22 = new Dgm.LineColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };

            A.SchemeColor schemeColor64 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent1 };
            A.Tint tint12 = new A.Tint(){ Val = 60000 };

            schemeColor64.Append(tint12);

            lineColorList22.Append(schemeColor64);
            Dgm.EffectColorList effectColorList22 = new Dgm.EffectColorList();
            Dgm.TextLineColorList textLineColorList22 = new Dgm.TextLineColorList();

            Dgm.TextFillColorList textFillColorList22 = new Dgm.TextFillColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };
            A.SchemeColor schemeColor65 = new A.SchemeColor(){ Val = A.SchemeColorValues.Light1 };

            textFillColorList22.Append(schemeColor65);
            Dgm.TextEffectColorList textEffectColorList22 = new Dgm.TextEffectColorList();

            colorTransformStyleLabel22.Append(fillColorList22);
            colorTransformStyleLabel22.Append(lineColorList22);
            colorTransformStyleLabel22.Append(effectColorList22);
            colorTransformStyleLabel22.Append(textLineColorList22);
            colorTransformStyleLabel22.Append(textFillColorList22);
            colorTransformStyleLabel22.Append(textEffectColorList22);

            Dgm.ColorTransformStyleLabel colorTransformStyleLabel23 = new Dgm.ColorTransformStyleLabel(){ Name = "parChTrans2D2" };

            Dgm.FillColorList fillColorList23 = new Dgm.FillColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };
            A.SchemeColor schemeColor66 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent1 };

            fillColorList23.Append(schemeColor66);

            Dgm.LineColorList lineColorList23 = new Dgm.LineColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };
            A.SchemeColor schemeColor67 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent1 };

            lineColorList23.Append(schemeColor67);
            Dgm.EffectColorList effectColorList23 = new Dgm.EffectColorList();
            Dgm.TextLineColorList textLineColorList23 = new Dgm.TextLineColorList();

            Dgm.TextFillColorList textFillColorList23 = new Dgm.TextFillColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };
            A.SchemeColor schemeColor68 = new A.SchemeColor(){ Val = A.SchemeColorValues.Light1 };

            textFillColorList23.Append(schemeColor68);
            Dgm.TextEffectColorList textEffectColorList23 = new Dgm.TextEffectColorList();

            colorTransformStyleLabel23.Append(fillColorList23);
            colorTransformStyleLabel23.Append(lineColorList23);
            colorTransformStyleLabel23.Append(effectColorList23);
            colorTransformStyleLabel23.Append(textLineColorList23);
            colorTransformStyleLabel23.Append(textFillColorList23);
            colorTransformStyleLabel23.Append(textEffectColorList23);

            Dgm.ColorTransformStyleLabel colorTransformStyleLabel24 = new Dgm.ColorTransformStyleLabel(){ Name = "parChTrans2D3" };

            Dgm.FillColorList fillColorList24 = new Dgm.FillColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };
            A.SchemeColor schemeColor69 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent1 };

            fillColorList24.Append(schemeColor69);

            Dgm.LineColorList lineColorList24 = new Dgm.LineColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };
            A.SchemeColor schemeColor70 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent1 };

            lineColorList24.Append(schemeColor70);
            Dgm.EffectColorList effectColorList24 = new Dgm.EffectColorList();
            Dgm.TextLineColorList textLineColorList24 = new Dgm.TextLineColorList();

            Dgm.TextFillColorList textFillColorList24 = new Dgm.TextFillColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };
            A.SchemeColor schemeColor71 = new A.SchemeColor(){ Val = A.SchemeColorValues.Light1 };

            textFillColorList24.Append(schemeColor71);
            Dgm.TextEffectColorList textEffectColorList24 = new Dgm.TextEffectColorList();

            colorTransformStyleLabel24.Append(fillColorList24);
            colorTransformStyleLabel24.Append(lineColorList24);
            colorTransformStyleLabel24.Append(effectColorList24);
            colorTransformStyleLabel24.Append(textLineColorList24);
            colorTransformStyleLabel24.Append(textFillColorList24);
            colorTransformStyleLabel24.Append(textEffectColorList24);

            Dgm.ColorTransformStyleLabel colorTransformStyleLabel25 = new Dgm.ColorTransformStyleLabel(){ Name = "parChTrans2D4" };

            Dgm.FillColorList fillColorList25 = new Dgm.FillColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };
            A.SchemeColor schemeColor72 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent1 };

            fillColorList25.Append(schemeColor72);

            Dgm.LineColorList lineColorList25 = new Dgm.LineColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };
            A.SchemeColor schemeColor73 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent1 };

            lineColorList25.Append(schemeColor73);
            Dgm.EffectColorList effectColorList25 = new Dgm.EffectColorList();
            Dgm.TextLineColorList textLineColorList25 = new Dgm.TextLineColorList();

            Dgm.TextFillColorList textFillColorList25 = new Dgm.TextFillColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };
            A.SchemeColor schemeColor74 = new A.SchemeColor(){ Val = A.SchemeColorValues.Light1 };

            textFillColorList25.Append(schemeColor74);
            Dgm.TextEffectColorList textEffectColorList25 = new Dgm.TextEffectColorList();

            colorTransformStyleLabel25.Append(fillColorList25);
            colorTransformStyleLabel25.Append(lineColorList25);
            colorTransformStyleLabel25.Append(effectColorList25);
            colorTransformStyleLabel25.Append(textLineColorList25);
            colorTransformStyleLabel25.Append(textFillColorList25);
            colorTransformStyleLabel25.Append(textEffectColorList25);

            Dgm.ColorTransformStyleLabel colorTransformStyleLabel26 = new Dgm.ColorTransformStyleLabel(){ Name = "parChTrans1D1" };

            Dgm.FillColorList fillColorList26 = new Dgm.FillColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };
            A.SchemeColor schemeColor75 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent1 };

            fillColorList26.Append(schemeColor75);

            Dgm.LineColorList lineColorList26 = new Dgm.LineColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };

            A.SchemeColor schemeColor76 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent1 };
            A.Shade shade1 = new A.Shade(){ Val = 60000 };

            schemeColor76.Append(shade1);

            lineColorList26.Append(schemeColor76);
            Dgm.EffectColorList effectColorList26 = new Dgm.EffectColorList();
            Dgm.TextLineColorList textLineColorList26 = new Dgm.TextLineColorList();

            Dgm.TextFillColorList textFillColorList26 = new Dgm.TextFillColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };
            A.SchemeColor schemeColor77 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            textFillColorList26.Append(schemeColor77);
            Dgm.TextEffectColorList textEffectColorList26 = new Dgm.TextEffectColorList();

            colorTransformStyleLabel26.Append(fillColorList26);
            colorTransformStyleLabel26.Append(lineColorList26);
            colorTransformStyleLabel26.Append(effectColorList26);
            colorTransformStyleLabel26.Append(textLineColorList26);
            colorTransformStyleLabel26.Append(textFillColorList26);
            colorTransformStyleLabel26.Append(textEffectColorList26);

            Dgm.ColorTransformStyleLabel colorTransformStyleLabel27 = new Dgm.ColorTransformStyleLabel(){ Name = "parChTrans1D2" };

            Dgm.FillColorList fillColorList27 = new Dgm.FillColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };
            A.SchemeColor schemeColor78 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent1 };

            fillColorList27.Append(schemeColor78);

            Dgm.LineColorList lineColorList27 = new Dgm.LineColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };

            A.SchemeColor schemeColor79 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent1 };
            A.Shade shade2 = new A.Shade(){ Val = 60000 };

            schemeColor79.Append(shade2);

            lineColorList27.Append(schemeColor79);
            Dgm.EffectColorList effectColorList27 = new Dgm.EffectColorList();
            Dgm.TextLineColorList textLineColorList27 = new Dgm.TextLineColorList();

            Dgm.TextFillColorList textFillColorList27 = new Dgm.TextFillColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };
            A.SchemeColor schemeColor80 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            textFillColorList27.Append(schemeColor80);
            Dgm.TextEffectColorList textEffectColorList27 = new Dgm.TextEffectColorList();

            colorTransformStyleLabel27.Append(fillColorList27);
            colorTransformStyleLabel27.Append(lineColorList27);
            colorTransformStyleLabel27.Append(effectColorList27);
            colorTransformStyleLabel27.Append(textLineColorList27);
            colorTransformStyleLabel27.Append(textFillColorList27);
            colorTransformStyleLabel27.Append(textEffectColorList27);

            Dgm.ColorTransformStyleLabel colorTransformStyleLabel28 = new Dgm.ColorTransformStyleLabel(){ Name = "parChTrans1D3" };

            Dgm.FillColorList fillColorList28 = new Dgm.FillColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };
            A.SchemeColor schemeColor81 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent1 };

            fillColorList28.Append(schemeColor81);

            Dgm.LineColorList lineColorList28 = new Dgm.LineColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };

            A.SchemeColor schemeColor82 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent1 };
            A.Shade shade3 = new A.Shade(){ Val = 80000 };

            schemeColor82.Append(shade3);

            lineColorList28.Append(schemeColor82);
            Dgm.EffectColorList effectColorList28 = new Dgm.EffectColorList();
            Dgm.TextLineColorList textLineColorList28 = new Dgm.TextLineColorList();

            Dgm.TextFillColorList textFillColorList28 = new Dgm.TextFillColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };
            A.SchemeColor schemeColor83 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            textFillColorList28.Append(schemeColor83);
            Dgm.TextEffectColorList textEffectColorList28 = new Dgm.TextEffectColorList();

            colorTransformStyleLabel28.Append(fillColorList28);
            colorTransformStyleLabel28.Append(lineColorList28);
            colorTransformStyleLabel28.Append(effectColorList28);
            colorTransformStyleLabel28.Append(textLineColorList28);
            colorTransformStyleLabel28.Append(textFillColorList28);
            colorTransformStyleLabel28.Append(textEffectColorList28);

            Dgm.ColorTransformStyleLabel colorTransformStyleLabel29 = new Dgm.ColorTransformStyleLabel(){ Name = "parChTrans1D4" };

            Dgm.FillColorList fillColorList29 = new Dgm.FillColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };
            A.SchemeColor schemeColor84 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent1 };

            fillColorList29.Append(schemeColor84);

            Dgm.LineColorList lineColorList29 = new Dgm.LineColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };

            A.SchemeColor schemeColor85 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent1 };
            A.Shade shade4 = new A.Shade(){ Val = 80000 };

            schemeColor85.Append(shade4);

            lineColorList29.Append(schemeColor85);
            Dgm.EffectColorList effectColorList29 = new Dgm.EffectColorList();
            Dgm.TextLineColorList textLineColorList29 = new Dgm.TextLineColorList();

            Dgm.TextFillColorList textFillColorList29 = new Dgm.TextFillColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };
            A.SchemeColor schemeColor86 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            textFillColorList29.Append(schemeColor86);
            Dgm.TextEffectColorList textEffectColorList29 = new Dgm.TextEffectColorList();

            colorTransformStyleLabel29.Append(fillColorList29);
            colorTransformStyleLabel29.Append(lineColorList29);
            colorTransformStyleLabel29.Append(effectColorList29);
            colorTransformStyleLabel29.Append(textLineColorList29);
            colorTransformStyleLabel29.Append(textFillColorList29);
            colorTransformStyleLabel29.Append(textEffectColorList29);

            Dgm.ColorTransformStyleLabel colorTransformStyleLabel30 = new Dgm.ColorTransformStyleLabel(){ Name = "fgAcc1" };

            Dgm.FillColorList fillColorList30 = new Dgm.FillColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };

            A.SchemeColor schemeColor87 = new A.SchemeColor(){ Val = A.SchemeColorValues.Light1 };
            A.Alpha alpha2 = new A.Alpha(){ Val = 90000 };

            schemeColor87.Append(alpha2);

            fillColorList30.Append(schemeColor87);

            Dgm.LineColorList lineColorList30 = new Dgm.LineColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };
            A.SchemeColor schemeColor88 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent1 };

            lineColorList30.Append(schemeColor88);
            Dgm.EffectColorList effectColorList30 = new Dgm.EffectColorList();
            Dgm.TextLineColorList textLineColorList30 = new Dgm.TextLineColorList();

            Dgm.TextFillColorList textFillColorList30 = new Dgm.TextFillColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };
            A.SchemeColor schemeColor89 = new A.SchemeColor(){ Val = A.SchemeColorValues.Dark1 };

            textFillColorList30.Append(schemeColor89);
            Dgm.TextEffectColorList textEffectColorList30 = new Dgm.TextEffectColorList();

            colorTransformStyleLabel30.Append(fillColorList30);
            colorTransformStyleLabel30.Append(lineColorList30);
            colorTransformStyleLabel30.Append(effectColorList30);
            colorTransformStyleLabel30.Append(textLineColorList30);
            colorTransformStyleLabel30.Append(textFillColorList30);
            colorTransformStyleLabel30.Append(textEffectColorList30);

            Dgm.ColorTransformStyleLabel colorTransformStyleLabel31 = new Dgm.ColorTransformStyleLabel(){ Name = "conFgAcc1" };

            Dgm.FillColorList fillColorList31 = new Dgm.FillColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };

            A.SchemeColor schemeColor90 = new A.SchemeColor(){ Val = A.SchemeColorValues.Light1 };
            A.Alpha alpha3 = new A.Alpha(){ Val = 90000 };

            schemeColor90.Append(alpha3);

            fillColorList31.Append(schemeColor90);

            Dgm.LineColorList lineColorList31 = new Dgm.LineColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };
            A.SchemeColor schemeColor91 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent1 };

            lineColorList31.Append(schemeColor91);
            Dgm.EffectColorList effectColorList31 = new Dgm.EffectColorList();
            Dgm.TextLineColorList textLineColorList31 = new Dgm.TextLineColorList();

            Dgm.TextFillColorList textFillColorList31 = new Dgm.TextFillColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };
            A.SchemeColor schemeColor92 = new A.SchemeColor(){ Val = A.SchemeColorValues.Dark1 };

            textFillColorList31.Append(schemeColor92);
            Dgm.TextEffectColorList textEffectColorList31 = new Dgm.TextEffectColorList();

            colorTransformStyleLabel31.Append(fillColorList31);
            colorTransformStyleLabel31.Append(lineColorList31);
            colorTransformStyleLabel31.Append(effectColorList31);
            colorTransformStyleLabel31.Append(textLineColorList31);
            colorTransformStyleLabel31.Append(textFillColorList31);
            colorTransformStyleLabel31.Append(textEffectColorList31);

            Dgm.ColorTransformStyleLabel colorTransformStyleLabel32 = new Dgm.ColorTransformStyleLabel(){ Name = "alignAcc1" };

            Dgm.FillColorList fillColorList32 = new Dgm.FillColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };

            A.SchemeColor schemeColor93 = new A.SchemeColor(){ Val = A.SchemeColorValues.Light1 };
            A.Alpha alpha4 = new A.Alpha(){ Val = 90000 };

            schemeColor93.Append(alpha4);

            fillColorList32.Append(schemeColor93);

            Dgm.LineColorList lineColorList32 = new Dgm.LineColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };
            A.SchemeColor schemeColor94 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent1 };

            lineColorList32.Append(schemeColor94);
            Dgm.EffectColorList effectColorList32 = new Dgm.EffectColorList();
            Dgm.TextLineColorList textLineColorList32 = new Dgm.TextLineColorList();

            Dgm.TextFillColorList textFillColorList32 = new Dgm.TextFillColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };
            A.SchemeColor schemeColor95 = new A.SchemeColor(){ Val = A.SchemeColorValues.Dark1 };

            textFillColorList32.Append(schemeColor95);
            Dgm.TextEffectColorList textEffectColorList32 = new Dgm.TextEffectColorList();

            colorTransformStyleLabel32.Append(fillColorList32);
            colorTransformStyleLabel32.Append(lineColorList32);
            colorTransformStyleLabel32.Append(effectColorList32);
            colorTransformStyleLabel32.Append(textLineColorList32);
            colorTransformStyleLabel32.Append(textFillColorList32);
            colorTransformStyleLabel32.Append(textEffectColorList32);

            Dgm.ColorTransformStyleLabel colorTransformStyleLabel33 = new Dgm.ColorTransformStyleLabel(){ Name = "trAlignAcc1" };

            Dgm.FillColorList fillColorList33 = new Dgm.FillColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };

            A.SchemeColor schemeColor96 = new A.SchemeColor(){ Val = A.SchemeColorValues.Light1 };
            A.Alpha alpha5 = new A.Alpha(){ Val = 40000 };

            schemeColor96.Append(alpha5);

            fillColorList33.Append(schemeColor96);

            Dgm.LineColorList lineColorList33 = new Dgm.LineColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };
            A.SchemeColor schemeColor97 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent1 };

            lineColorList33.Append(schemeColor97);
            Dgm.EffectColorList effectColorList33 = new Dgm.EffectColorList();
            Dgm.TextLineColorList textLineColorList33 = new Dgm.TextLineColorList();

            Dgm.TextFillColorList textFillColorList33 = new Dgm.TextFillColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };
            A.SchemeColor schemeColor98 = new A.SchemeColor(){ Val = A.SchemeColorValues.Dark1 };

            textFillColorList33.Append(schemeColor98);
            Dgm.TextEffectColorList textEffectColorList33 = new Dgm.TextEffectColorList();

            colorTransformStyleLabel33.Append(fillColorList33);
            colorTransformStyleLabel33.Append(lineColorList33);
            colorTransformStyleLabel33.Append(effectColorList33);
            colorTransformStyleLabel33.Append(textLineColorList33);
            colorTransformStyleLabel33.Append(textFillColorList33);
            colorTransformStyleLabel33.Append(textEffectColorList33);

            Dgm.ColorTransformStyleLabel colorTransformStyleLabel34 = new Dgm.ColorTransformStyleLabel(){ Name = "bgAcc1" };

            Dgm.FillColorList fillColorList34 = new Dgm.FillColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };

            A.SchemeColor schemeColor99 = new A.SchemeColor(){ Val = A.SchemeColorValues.Light1 };
            A.Alpha alpha6 = new A.Alpha(){ Val = 90000 };

            schemeColor99.Append(alpha6);

            fillColorList34.Append(schemeColor99);

            Dgm.LineColorList lineColorList34 = new Dgm.LineColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };
            A.SchemeColor schemeColor100 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent1 };

            lineColorList34.Append(schemeColor100);
            Dgm.EffectColorList effectColorList34 = new Dgm.EffectColorList();
            Dgm.TextLineColorList textLineColorList34 = new Dgm.TextLineColorList();

            Dgm.TextFillColorList textFillColorList34 = new Dgm.TextFillColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };
            A.SchemeColor schemeColor101 = new A.SchemeColor(){ Val = A.SchemeColorValues.Dark1 };

            textFillColorList34.Append(schemeColor101);
            Dgm.TextEffectColorList textEffectColorList34 = new Dgm.TextEffectColorList();

            colorTransformStyleLabel34.Append(fillColorList34);
            colorTransformStyleLabel34.Append(lineColorList34);
            colorTransformStyleLabel34.Append(effectColorList34);
            colorTransformStyleLabel34.Append(textLineColorList34);
            colorTransformStyleLabel34.Append(textFillColorList34);
            colorTransformStyleLabel34.Append(textEffectColorList34);

            Dgm.ColorTransformStyleLabel colorTransformStyleLabel35 = new Dgm.ColorTransformStyleLabel(){ Name = "solidFgAcc1" };

            Dgm.FillColorList fillColorList35 = new Dgm.FillColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };
            A.SchemeColor schemeColor102 = new A.SchemeColor(){ Val = A.SchemeColorValues.Light1 };

            fillColorList35.Append(schemeColor102);

            Dgm.LineColorList lineColorList35 = new Dgm.LineColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };
            A.SchemeColor schemeColor103 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent1 };

            lineColorList35.Append(schemeColor103);
            Dgm.EffectColorList effectColorList35 = new Dgm.EffectColorList();
            Dgm.TextLineColorList textLineColorList35 = new Dgm.TextLineColorList();

            Dgm.TextFillColorList textFillColorList35 = new Dgm.TextFillColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };
            A.SchemeColor schemeColor104 = new A.SchemeColor(){ Val = A.SchemeColorValues.Dark1 };

            textFillColorList35.Append(schemeColor104);
            Dgm.TextEffectColorList textEffectColorList35 = new Dgm.TextEffectColorList();

            colorTransformStyleLabel35.Append(fillColorList35);
            colorTransformStyleLabel35.Append(lineColorList35);
            colorTransformStyleLabel35.Append(effectColorList35);
            colorTransformStyleLabel35.Append(textLineColorList35);
            colorTransformStyleLabel35.Append(textFillColorList35);
            colorTransformStyleLabel35.Append(textEffectColorList35);

            Dgm.ColorTransformStyleLabel colorTransformStyleLabel36 = new Dgm.ColorTransformStyleLabel(){ Name = "solidAlignAcc1" };

            Dgm.FillColorList fillColorList36 = new Dgm.FillColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };
            A.SchemeColor schemeColor105 = new A.SchemeColor(){ Val = A.SchemeColorValues.Light1 };

            fillColorList36.Append(schemeColor105);

            Dgm.LineColorList lineColorList36 = new Dgm.LineColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };
            A.SchemeColor schemeColor106 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent1 };

            lineColorList36.Append(schemeColor106);
            Dgm.EffectColorList effectColorList36 = new Dgm.EffectColorList();
            Dgm.TextLineColorList textLineColorList36 = new Dgm.TextLineColorList();

            Dgm.TextFillColorList textFillColorList36 = new Dgm.TextFillColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };
            A.SchemeColor schemeColor107 = new A.SchemeColor(){ Val = A.SchemeColorValues.Dark1 };

            textFillColorList36.Append(schemeColor107);
            Dgm.TextEffectColorList textEffectColorList36 = new Dgm.TextEffectColorList();

            colorTransformStyleLabel36.Append(fillColorList36);
            colorTransformStyleLabel36.Append(lineColorList36);
            colorTransformStyleLabel36.Append(effectColorList36);
            colorTransformStyleLabel36.Append(textLineColorList36);
            colorTransformStyleLabel36.Append(textFillColorList36);
            colorTransformStyleLabel36.Append(textEffectColorList36);

            Dgm.ColorTransformStyleLabel colorTransformStyleLabel37 = new Dgm.ColorTransformStyleLabel(){ Name = "solidBgAcc1" };

            Dgm.FillColorList fillColorList37 = new Dgm.FillColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };
            A.SchemeColor schemeColor108 = new A.SchemeColor(){ Val = A.SchemeColorValues.Light1 };

            fillColorList37.Append(schemeColor108);

            Dgm.LineColorList lineColorList37 = new Dgm.LineColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };
            A.SchemeColor schemeColor109 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent1 };

            lineColorList37.Append(schemeColor109);
            Dgm.EffectColorList effectColorList37 = new Dgm.EffectColorList();
            Dgm.TextLineColorList textLineColorList37 = new Dgm.TextLineColorList();

            Dgm.TextFillColorList textFillColorList37 = new Dgm.TextFillColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };
            A.SchemeColor schemeColor110 = new A.SchemeColor(){ Val = A.SchemeColorValues.Dark1 };

            textFillColorList37.Append(schemeColor110);
            Dgm.TextEffectColorList textEffectColorList37 = new Dgm.TextEffectColorList();

            colorTransformStyleLabel37.Append(fillColorList37);
            colorTransformStyleLabel37.Append(lineColorList37);
            colorTransformStyleLabel37.Append(effectColorList37);
            colorTransformStyleLabel37.Append(textLineColorList37);
            colorTransformStyleLabel37.Append(textFillColorList37);
            colorTransformStyleLabel37.Append(textEffectColorList37);

            Dgm.ColorTransformStyleLabel colorTransformStyleLabel38 = new Dgm.ColorTransformStyleLabel(){ Name = "fgAccFollowNode1" };

            Dgm.FillColorList fillColorList38 = new Dgm.FillColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };

            A.SchemeColor schemeColor111 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent1 };
            A.Alpha alpha7 = new A.Alpha(){ Val = 90000 };
            A.Tint tint13 = new A.Tint(){ Val = 40000 };

            schemeColor111.Append(alpha7);
            schemeColor111.Append(tint13);

            fillColorList38.Append(schemeColor111);

            Dgm.LineColorList lineColorList38 = new Dgm.LineColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };

            A.SchemeColor schemeColor112 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent1 };
            A.Alpha alpha8 = new A.Alpha(){ Val = 90000 };
            A.Tint tint14 = new A.Tint(){ Val = 40000 };

            schemeColor112.Append(alpha8);
            schemeColor112.Append(tint14);

            lineColorList38.Append(schemeColor112);
            Dgm.EffectColorList effectColorList38 = new Dgm.EffectColorList();
            Dgm.TextLineColorList textLineColorList38 = new Dgm.TextLineColorList();

            Dgm.TextFillColorList textFillColorList38 = new Dgm.TextFillColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };
            A.SchemeColor schemeColor113 = new A.SchemeColor(){ Val = A.SchemeColorValues.Dark1 };

            textFillColorList38.Append(schemeColor113);
            Dgm.TextEffectColorList textEffectColorList38 = new Dgm.TextEffectColorList();

            colorTransformStyleLabel38.Append(fillColorList38);
            colorTransformStyleLabel38.Append(lineColorList38);
            colorTransformStyleLabel38.Append(effectColorList38);
            colorTransformStyleLabel38.Append(textLineColorList38);
            colorTransformStyleLabel38.Append(textFillColorList38);
            colorTransformStyleLabel38.Append(textEffectColorList38);

            Dgm.ColorTransformStyleLabel colorTransformStyleLabel39 = new Dgm.ColorTransformStyleLabel(){ Name = "alignAccFollowNode1" };

            Dgm.FillColorList fillColorList39 = new Dgm.FillColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };

            A.SchemeColor schemeColor114 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent1 };
            A.Alpha alpha9 = new A.Alpha(){ Val = 90000 };
            A.Tint tint15 = new A.Tint(){ Val = 40000 };

            schemeColor114.Append(alpha9);
            schemeColor114.Append(tint15);

            fillColorList39.Append(schemeColor114);

            Dgm.LineColorList lineColorList39 = new Dgm.LineColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };

            A.SchemeColor schemeColor115 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent1 };
            A.Alpha alpha10 = new A.Alpha(){ Val = 90000 };
            A.Tint tint16 = new A.Tint(){ Val = 40000 };

            schemeColor115.Append(alpha10);
            schemeColor115.Append(tint16);

            lineColorList39.Append(schemeColor115);
            Dgm.EffectColorList effectColorList39 = new Dgm.EffectColorList();
            Dgm.TextLineColorList textLineColorList39 = new Dgm.TextLineColorList();

            Dgm.TextFillColorList textFillColorList39 = new Dgm.TextFillColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };
            A.SchemeColor schemeColor116 = new A.SchemeColor(){ Val = A.SchemeColorValues.Dark1 };

            textFillColorList39.Append(schemeColor116);
            Dgm.TextEffectColorList textEffectColorList39 = new Dgm.TextEffectColorList();

            colorTransformStyleLabel39.Append(fillColorList39);
            colorTransformStyleLabel39.Append(lineColorList39);
            colorTransformStyleLabel39.Append(effectColorList39);
            colorTransformStyleLabel39.Append(textLineColorList39);
            colorTransformStyleLabel39.Append(textFillColorList39);
            colorTransformStyleLabel39.Append(textEffectColorList39);

            Dgm.ColorTransformStyleLabel colorTransformStyleLabel40 = new Dgm.ColorTransformStyleLabel(){ Name = "bgAccFollowNode1" };

            Dgm.FillColorList fillColorList40 = new Dgm.FillColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };

            A.SchemeColor schemeColor117 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent1 };
            A.Alpha alpha11 = new A.Alpha(){ Val = 90000 };
            A.Tint tint17 = new A.Tint(){ Val = 40000 };

            schemeColor117.Append(alpha11);
            schemeColor117.Append(tint17);

            fillColorList40.Append(schemeColor117);

            Dgm.LineColorList lineColorList40 = new Dgm.LineColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };

            A.SchemeColor schemeColor118 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent1 };
            A.Alpha alpha12 = new A.Alpha(){ Val = 90000 };
            A.Tint tint18 = new A.Tint(){ Val = 40000 };

            schemeColor118.Append(alpha12);
            schemeColor118.Append(tint18);

            lineColorList40.Append(schemeColor118);
            Dgm.EffectColorList effectColorList40 = new Dgm.EffectColorList();
            Dgm.TextLineColorList textLineColorList40 = new Dgm.TextLineColorList();

            Dgm.TextFillColorList textFillColorList40 = new Dgm.TextFillColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };
            A.SchemeColor schemeColor119 = new A.SchemeColor(){ Val = A.SchemeColorValues.Dark1 };

            textFillColorList40.Append(schemeColor119);
            Dgm.TextEffectColorList textEffectColorList40 = new Dgm.TextEffectColorList();

            colorTransformStyleLabel40.Append(fillColorList40);
            colorTransformStyleLabel40.Append(lineColorList40);
            colorTransformStyleLabel40.Append(effectColorList40);
            colorTransformStyleLabel40.Append(textLineColorList40);
            colorTransformStyleLabel40.Append(textFillColorList40);
            colorTransformStyleLabel40.Append(textEffectColorList40);

            Dgm.ColorTransformStyleLabel colorTransformStyleLabel41 = new Dgm.ColorTransformStyleLabel(){ Name = "fgAcc0" };

            Dgm.FillColorList fillColorList41 = new Dgm.FillColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };

            A.SchemeColor schemeColor120 = new A.SchemeColor(){ Val = A.SchemeColorValues.Light1 };
            A.Alpha alpha13 = new A.Alpha(){ Val = 90000 };

            schemeColor120.Append(alpha13);

            fillColorList41.Append(schemeColor120);

            Dgm.LineColorList lineColorList41 = new Dgm.LineColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };
            A.SchemeColor schemeColor121 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent1 };

            lineColorList41.Append(schemeColor121);
            Dgm.EffectColorList effectColorList41 = new Dgm.EffectColorList();
            Dgm.TextLineColorList textLineColorList41 = new Dgm.TextLineColorList();

            Dgm.TextFillColorList textFillColorList41 = new Dgm.TextFillColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };
            A.SchemeColor schemeColor122 = new A.SchemeColor(){ Val = A.SchemeColorValues.Dark1 };

            textFillColorList41.Append(schemeColor122);
            Dgm.TextEffectColorList textEffectColorList41 = new Dgm.TextEffectColorList();

            colorTransformStyleLabel41.Append(fillColorList41);
            colorTransformStyleLabel41.Append(lineColorList41);
            colorTransformStyleLabel41.Append(effectColorList41);
            colorTransformStyleLabel41.Append(textLineColorList41);
            colorTransformStyleLabel41.Append(textFillColorList41);
            colorTransformStyleLabel41.Append(textEffectColorList41);

            Dgm.ColorTransformStyleLabel colorTransformStyleLabel42 = new Dgm.ColorTransformStyleLabel(){ Name = "fgAcc2" };

            Dgm.FillColorList fillColorList42 = new Dgm.FillColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };

            A.SchemeColor schemeColor123 = new A.SchemeColor(){ Val = A.SchemeColorValues.Light1 };
            A.Alpha alpha14 = new A.Alpha(){ Val = 90000 };

            schemeColor123.Append(alpha14);

            fillColorList42.Append(schemeColor123);

            Dgm.LineColorList lineColorList42 = new Dgm.LineColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };
            A.SchemeColor schemeColor124 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent1 };

            lineColorList42.Append(schemeColor124);
            Dgm.EffectColorList effectColorList42 = new Dgm.EffectColorList();
            Dgm.TextLineColorList textLineColorList42 = new Dgm.TextLineColorList();

            Dgm.TextFillColorList textFillColorList42 = new Dgm.TextFillColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };
            A.SchemeColor schemeColor125 = new A.SchemeColor(){ Val = A.SchemeColorValues.Dark1 };

            textFillColorList42.Append(schemeColor125);
            Dgm.TextEffectColorList textEffectColorList42 = new Dgm.TextEffectColorList();

            colorTransformStyleLabel42.Append(fillColorList42);
            colorTransformStyleLabel42.Append(lineColorList42);
            colorTransformStyleLabel42.Append(effectColorList42);
            colorTransformStyleLabel42.Append(textLineColorList42);
            colorTransformStyleLabel42.Append(textFillColorList42);
            colorTransformStyleLabel42.Append(textEffectColorList42);

            Dgm.ColorTransformStyleLabel colorTransformStyleLabel43 = new Dgm.ColorTransformStyleLabel(){ Name = "fgAcc3" };

            Dgm.FillColorList fillColorList43 = new Dgm.FillColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };

            A.SchemeColor schemeColor126 = new A.SchemeColor(){ Val = A.SchemeColorValues.Light1 };
            A.Alpha alpha15 = new A.Alpha(){ Val = 90000 };

            schemeColor126.Append(alpha15);

            fillColorList43.Append(schemeColor126);

            Dgm.LineColorList lineColorList43 = new Dgm.LineColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };
            A.SchemeColor schemeColor127 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent1 };

            lineColorList43.Append(schemeColor127);
            Dgm.EffectColorList effectColorList43 = new Dgm.EffectColorList();
            Dgm.TextLineColorList textLineColorList43 = new Dgm.TextLineColorList();

            Dgm.TextFillColorList textFillColorList43 = new Dgm.TextFillColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };
            A.SchemeColor schemeColor128 = new A.SchemeColor(){ Val = A.SchemeColorValues.Dark1 };

            textFillColorList43.Append(schemeColor128);
            Dgm.TextEffectColorList textEffectColorList43 = new Dgm.TextEffectColorList();

            colorTransformStyleLabel43.Append(fillColorList43);
            colorTransformStyleLabel43.Append(lineColorList43);
            colorTransformStyleLabel43.Append(effectColorList43);
            colorTransformStyleLabel43.Append(textLineColorList43);
            colorTransformStyleLabel43.Append(textFillColorList43);
            colorTransformStyleLabel43.Append(textEffectColorList43);

            Dgm.ColorTransformStyleLabel colorTransformStyleLabel44 = new Dgm.ColorTransformStyleLabel(){ Name = "fgAcc4" };

            Dgm.FillColorList fillColorList44 = new Dgm.FillColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };

            A.SchemeColor schemeColor129 = new A.SchemeColor(){ Val = A.SchemeColorValues.Light1 };
            A.Alpha alpha16 = new A.Alpha(){ Val = 90000 };

            schemeColor129.Append(alpha16);

            fillColorList44.Append(schemeColor129);

            Dgm.LineColorList lineColorList44 = new Dgm.LineColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };
            A.SchemeColor schemeColor130 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent1 };

            lineColorList44.Append(schemeColor130);
            Dgm.EffectColorList effectColorList44 = new Dgm.EffectColorList();
            Dgm.TextLineColorList textLineColorList44 = new Dgm.TextLineColorList();

            Dgm.TextFillColorList textFillColorList44 = new Dgm.TextFillColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };
            A.SchemeColor schemeColor131 = new A.SchemeColor(){ Val = A.SchemeColorValues.Dark1 };

            textFillColorList44.Append(schemeColor131);
            Dgm.TextEffectColorList textEffectColorList44 = new Dgm.TextEffectColorList();

            colorTransformStyleLabel44.Append(fillColorList44);
            colorTransformStyleLabel44.Append(lineColorList44);
            colorTransformStyleLabel44.Append(effectColorList44);
            colorTransformStyleLabel44.Append(textLineColorList44);
            colorTransformStyleLabel44.Append(textFillColorList44);
            colorTransformStyleLabel44.Append(textEffectColorList44);

            Dgm.ColorTransformStyleLabel colorTransformStyleLabel45 = new Dgm.ColorTransformStyleLabel(){ Name = "bgShp" };

            Dgm.FillColorList fillColorList45 = new Dgm.FillColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };

            A.SchemeColor schemeColor132 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent1 };
            A.Tint tint19 = new A.Tint(){ Val = 40000 };

            schemeColor132.Append(tint19);

            fillColorList45.Append(schemeColor132);

            Dgm.LineColorList lineColorList45 = new Dgm.LineColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };
            A.SchemeColor schemeColor133 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent1 };

            lineColorList45.Append(schemeColor133);
            Dgm.EffectColorList effectColorList45 = new Dgm.EffectColorList();
            Dgm.TextLineColorList textLineColorList45 = new Dgm.TextLineColorList();

            Dgm.TextFillColorList textFillColorList45 = new Dgm.TextFillColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };
            A.SchemeColor schemeColor134 = new A.SchemeColor(){ Val = A.SchemeColorValues.Dark1 };

            textFillColorList45.Append(schemeColor134);
            Dgm.TextEffectColorList textEffectColorList45 = new Dgm.TextEffectColorList();

            colorTransformStyleLabel45.Append(fillColorList45);
            colorTransformStyleLabel45.Append(lineColorList45);
            colorTransformStyleLabel45.Append(effectColorList45);
            colorTransformStyleLabel45.Append(textLineColorList45);
            colorTransformStyleLabel45.Append(textFillColorList45);
            colorTransformStyleLabel45.Append(textEffectColorList45);

            Dgm.ColorTransformStyleLabel colorTransformStyleLabel46 = new Dgm.ColorTransformStyleLabel(){ Name = "dkBgShp" };

            Dgm.FillColorList fillColorList46 = new Dgm.FillColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };

            A.SchemeColor schemeColor135 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent1 };
            A.Shade shade5 = new A.Shade(){ Val = 80000 };

            schemeColor135.Append(shade5);

            fillColorList46.Append(schemeColor135);

            Dgm.LineColorList lineColorList46 = new Dgm.LineColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };
            A.SchemeColor schemeColor136 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent1 };

            lineColorList46.Append(schemeColor136);
            Dgm.EffectColorList effectColorList46 = new Dgm.EffectColorList();
            Dgm.TextLineColorList textLineColorList46 = new Dgm.TextLineColorList();

            Dgm.TextFillColorList textFillColorList46 = new Dgm.TextFillColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };
            A.SchemeColor schemeColor137 = new A.SchemeColor(){ Val = A.SchemeColorValues.Light1 };

            textFillColorList46.Append(schemeColor137);
            Dgm.TextEffectColorList textEffectColorList46 = new Dgm.TextEffectColorList();

            colorTransformStyleLabel46.Append(fillColorList46);
            colorTransformStyleLabel46.Append(lineColorList46);
            colorTransformStyleLabel46.Append(effectColorList46);
            colorTransformStyleLabel46.Append(textLineColorList46);
            colorTransformStyleLabel46.Append(textFillColorList46);
            colorTransformStyleLabel46.Append(textEffectColorList46);

            Dgm.ColorTransformStyleLabel colorTransformStyleLabel47 = new Dgm.ColorTransformStyleLabel(){ Name = "trBgShp" };

            Dgm.FillColorList fillColorList47 = new Dgm.FillColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };

            A.SchemeColor schemeColor138 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent1 };
            A.Tint tint20 = new A.Tint(){ Val = 50000 };
            A.Alpha alpha17 = new A.Alpha(){ Val = 40000 };

            schemeColor138.Append(tint20);
            schemeColor138.Append(alpha17);

            fillColorList47.Append(schemeColor138);

            Dgm.LineColorList lineColorList47 = new Dgm.LineColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };
            A.SchemeColor schemeColor139 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent1 };

            lineColorList47.Append(schemeColor139);
            Dgm.EffectColorList effectColorList47 = new Dgm.EffectColorList();
            Dgm.TextLineColorList textLineColorList47 = new Dgm.TextLineColorList();

            Dgm.TextFillColorList textFillColorList47 = new Dgm.TextFillColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };
            A.SchemeColor schemeColor140 = new A.SchemeColor(){ Val = A.SchemeColorValues.Light1 };

            textFillColorList47.Append(schemeColor140);
            Dgm.TextEffectColorList textEffectColorList47 = new Dgm.TextEffectColorList();

            colorTransformStyleLabel47.Append(fillColorList47);
            colorTransformStyleLabel47.Append(lineColorList47);
            colorTransformStyleLabel47.Append(effectColorList47);
            colorTransformStyleLabel47.Append(textLineColorList47);
            colorTransformStyleLabel47.Append(textFillColorList47);
            colorTransformStyleLabel47.Append(textEffectColorList47);

            Dgm.ColorTransformStyleLabel colorTransformStyleLabel48 = new Dgm.ColorTransformStyleLabel(){ Name = "fgShp" };

            Dgm.FillColorList fillColorList48 = new Dgm.FillColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };

            A.SchemeColor schemeColor141 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent1 };
            A.Tint tint21 = new A.Tint(){ Val = 60000 };

            schemeColor141.Append(tint21);

            fillColorList48.Append(schemeColor141);

            Dgm.LineColorList lineColorList48 = new Dgm.LineColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };
            A.SchemeColor schemeColor142 = new A.SchemeColor(){ Val = A.SchemeColorValues.Light1 };

            lineColorList48.Append(schemeColor142);
            Dgm.EffectColorList effectColorList48 = new Dgm.EffectColorList();
            Dgm.TextLineColorList textLineColorList48 = new Dgm.TextLineColorList();

            Dgm.TextFillColorList textFillColorList48 = new Dgm.TextFillColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };
            A.SchemeColor schemeColor143 = new A.SchemeColor(){ Val = A.SchemeColorValues.Dark1 };

            textFillColorList48.Append(schemeColor143);
            Dgm.TextEffectColorList textEffectColorList48 = new Dgm.TextEffectColorList();

            colorTransformStyleLabel48.Append(fillColorList48);
            colorTransformStyleLabel48.Append(lineColorList48);
            colorTransformStyleLabel48.Append(effectColorList48);
            colorTransformStyleLabel48.Append(textLineColorList48);
            colorTransformStyleLabel48.Append(textFillColorList48);
            colorTransformStyleLabel48.Append(textEffectColorList48);

            Dgm.ColorTransformStyleLabel colorTransformStyleLabel49 = new Dgm.ColorTransformStyleLabel(){ Name = "revTx" };

            Dgm.FillColorList fillColorList49 = new Dgm.FillColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };

            A.SchemeColor schemeColor144 = new A.SchemeColor(){ Val = A.SchemeColorValues.Light1 };
            A.Alpha alpha18 = new A.Alpha(){ Val = 0 };

            schemeColor144.Append(alpha18);

            fillColorList49.Append(schemeColor144);

            Dgm.LineColorList lineColorList49 = new Dgm.LineColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };

            A.SchemeColor schemeColor145 = new A.SchemeColor(){ Val = A.SchemeColorValues.Dark1 };
            A.Alpha alpha19 = new A.Alpha(){ Val = 0 };

            schemeColor145.Append(alpha19);

            lineColorList49.Append(schemeColor145);
            Dgm.EffectColorList effectColorList49 = new Dgm.EffectColorList();
            Dgm.TextLineColorList textLineColorList49 = new Dgm.TextLineColorList();

            Dgm.TextFillColorList textFillColorList49 = new Dgm.TextFillColorList(){ Method = Dgm.ColorApplicationMethodValues.Repeat };
            A.SchemeColor schemeColor146 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            textFillColorList49.Append(schemeColor146);
            Dgm.TextEffectColorList textEffectColorList49 = new Dgm.TextEffectColorList();

            colorTransformStyleLabel49.Append(fillColorList49);
            colorTransformStyleLabel49.Append(lineColorList49);
            colorTransformStyleLabel49.Append(effectColorList49);
            colorTransformStyleLabel49.Append(textLineColorList49);
            colorTransformStyleLabel49.Append(textFillColorList49);
            colorTransformStyleLabel49.Append(textEffectColorList49);

            colorsDefinition1.Append(colorDefinitionTitle1);
            colorsDefinition1.Append(colorTransformDescription1);
            colorsDefinition1.Append(colorTransformCategories1);
            colorsDefinition1.Append(colorTransformStyleLabel1);
            colorsDefinition1.Append(colorTransformStyleLabel2);
            colorsDefinition1.Append(colorTransformStyleLabel3);
            colorsDefinition1.Append(colorTransformStyleLabel4);
            colorsDefinition1.Append(colorTransformStyleLabel5);
            colorsDefinition1.Append(colorTransformStyleLabel6);
            colorsDefinition1.Append(colorTransformStyleLabel7);
            colorsDefinition1.Append(colorTransformStyleLabel8);
            colorsDefinition1.Append(colorTransformStyleLabel9);
            colorsDefinition1.Append(colorTransformStyleLabel10);
            colorsDefinition1.Append(colorTransformStyleLabel11);
            colorsDefinition1.Append(colorTransformStyleLabel12);
            colorsDefinition1.Append(colorTransformStyleLabel13);
            colorsDefinition1.Append(colorTransformStyleLabel14);
            colorsDefinition1.Append(colorTransformStyleLabel15);
            colorsDefinition1.Append(colorTransformStyleLabel16);
            colorsDefinition1.Append(colorTransformStyleLabel17);
            colorsDefinition1.Append(colorTransformStyleLabel18);
            colorsDefinition1.Append(colorTransformStyleLabel19);
            colorsDefinition1.Append(colorTransformStyleLabel20);
            colorsDefinition1.Append(colorTransformStyleLabel21);
            colorsDefinition1.Append(colorTransformStyleLabel22);
            colorsDefinition1.Append(colorTransformStyleLabel23);
            colorsDefinition1.Append(colorTransformStyleLabel24);
            colorsDefinition1.Append(colorTransformStyleLabel25);
            colorsDefinition1.Append(colorTransformStyleLabel26);
            colorsDefinition1.Append(colorTransformStyleLabel27);
            colorsDefinition1.Append(colorTransformStyleLabel28);
            colorsDefinition1.Append(colorTransformStyleLabel29);
            colorsDefinition1.Append(colorTransformStyleLabel30);
            colorsDefinition1.Append(colorTransformStyleLabel31);
            colorsDefinition1.Append(colorTransformStyleLabel32);
            colorsDefinition1.Append(colorTransformStyleLabel33);
            colorsDefinition1.Append(colorTransformStyleLabel34);
            colorsDefinition1.Append(colorTransformStyleLabel35);
            colorsDefinition1.Append(colorTransformStyleLabel36);
            colorsDefinition1.Append(colorTransformStyleLabel37);
            colorsDefinition1.Append(colorTransformStyleLabel38);
            colorsDefinition1.Append(colorTransformStyleLabel39);
            colorsDefinition1.Append(colorTransformStyleLabel40);
            colorsDefinition1.Append(colorTransformStyleLabel41);
            colorsDefinition1.Append(colorTransformStyleLabel42);
            colorsDefinition1.Append(colorTransformStyleLabel43);
            colorsDefinition1.Append(colorTransformStyleLabel44);
            colorsDefinition1.Append(colorTransformStyleLabel45);
            colorsDefinition1.Append(colorTransformStyleLabel46);
            colorsDefinition1.Append(colorTransformStyleLabel47);
            colorsDefinition1.Append(colorTransformStyleLabel48);
            colorsDefinition1.Append(colorTransformStyleLabel49);

            diagramColorsPart1.ColorsDefinition = colorsDefinition1;
        }

        // Generates content of documentSettingsPart1.
        private void GenerateDocumentSettingsPart1Content(DocumentSettingsPart documentSettingsPart1)
        {
            Settings settings1 = new Settings(){ MCAttributes = new MarkupCompatibilityAttributes(){ Ignorable = "w14 w15 w16se w16cid w16 w16cex w16sdtdh w16sdtfl w16du" }  };
            settings1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            settings1.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            settings1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            settings1.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            settings1.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            settings1.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
            settings1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            settings1.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            settings1.AddNamespaceDeclaration("w15", "http://schemas.microsoft.com/office/word/2012/wordml");
            settings1.AddNamespaceDeclaration("w16cex", "http://schemas.microsoft.com/office/word/2018/wordml/cex");
            settings1.AddNamespaceDeclaration("w16cid", "http://schemas.microsoft.com/office/word/2016/wordml/cid");
            settings1.AddNamespaceDeclaration("w16", "http://schemas.microsoft.com/office/word/2018/wordml");
            settings1.AddNamespaceDeclaration("w16du", "http://schemas.microsoft.com/office/word/2023/wordml/word16du");
            settings1.AddNamespaceDeclaration("w16sdtdh", "http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash");
            settings1.AddNamespaceDeclaration("w16sdtfl", "http://schemas.microsoft.com/office/word/2024/wordml/sdtformatlock");
            settings1.AddNamespaceDeclaration("w16se", "http://schemas.microsoft.com/office/word/2015/wordml/symex");
            settings1.AddNamespaceDeclaration("sl", "http://schemas.openxmlformats.org/schemaLibrary/2006/main");
            Zoom zoom1 = new Zoom(){ Percent = "100" };
            ProofState proofState1 = new ProofState(){ Grammar = ProofingStateValues.Clean };
            DefaultTabStop defaultTabStop1 = new DefaultTabStop(){ Val = 708 };
            HyphenationZone hyphenationZone1 = new HyphenationZone(){ Val = "425" };
            CharacterSpacingControl characterSpacingControl1 = new CharacterSpacingControl(){ Val = CharacterSpacingValues.DoNotCompress };

            Compatibility compatibility1 = new Compatibility();
            CompatibilitySetting compatibilitySetting1 = new CompatibilitySetting(){ Name = CompatSettingNameValues.CompatibilityMode, Uri = "http://schemas.microsoft.com/office/word", Val = "15" };
            CompatibilitySetting compatibilitySetting2 = new CompatibilitySetting(){ Name = CompatSettingNameValues.OverrideTableStyleFontSizeAndJustification, Uri = "http://schemas.microsoft.com/office/word", Val = "1" };
            CompatibilitySetting compatibilitySetting3 = new CompatibilitySetting(){ Name = CompatSettingNameValues.EnableOpenTypeFeatures, Uri = "http://schemas.microsoft.com/office/word", Val = "1" };
            CompatibilitySetting compatibilitySetting4 = new CompatibilitySetting(){ Name = CompatSettingNameValues.DoNotFlipMirrorIndents, Uri = "http://schemas.microsoft.com/office/word", Val = "1" };
            CompatibilitySetting compatibilitySetting5 = new CompatibilitySetting(){ Name = CompatSettingNameValues.DifferentiateMultirowTableHeaders, Uri = "http://schemas.microsoft.com/office/word", Val = "1" };
            CompatibilitySetting compatibilitySetting6 = new CompatibilitySetting(){ Name = new EnumValue<CompatSettingNameValues>() { InnerText = "useWord2013TrackBottomHyphenation"}, Uri = "http://schemas.microsoft.com/office/word", Val = "0" };

            compatibility1.Append(compatibilitySetting1);
            compatibility1.Append(compatibilitySetting2);
            compatibility1.Append(compatibilitySetting3);
            compatibility1.Append(compatibilitySetting4);
            compatibility1.Append(compatibilitySetting5);
            compatibility1.Append(compatibilitySetting6);

            Rsids rsids1 = new Rsids();
            RsidRoot rsidRoot1 = new RsidRoot(){ Val = "00440271" };
            Rsid rsid1 = new Rsid(){ Val = "0018596D" };
            Rsid rsid2 = new Rsid(){ Val = "00440271" };
            Rsid rsid3 = new Rsid(){ Val = "009674C8" };
            Rsid rsid4 = new Rsid(){ Val = "00C33987" };

            rsids1.Append(rsidRoot1);
            rsids1.Append(rsid1);
            rsids1.Append(rsid2);
            rsids1.Append(rsid3);
            rsids1.Append(rsid4);

            M.MathProperties mathProperties1 = new M.MathProperties();
            M.MathFont mathFont1 = new M.MathFont(){ Val = "Cambria Math" };
            M.BreakBinary breakBinary1 = new M.BreakBinary(){ Val = M.BreakBinaryOperatorValues.Before };
            M.BreakBinarySubtraction breakBinarySubtraction1 = new M.BreakBinarySubtraction(){ Val = M.BreakBinarySubtractionValues.MinusMinus };
            M.SmallFraction smallFraction1 = new M.SmallFraction(){ Val = M.BooleanValues.Zero };
            M.DisplayDefaults displayDefaults1 = new M.DisplayDefaults();
            M.LeftMargin leftMargin1 = new M.LeftMargin(){ Val = (UInt32Value)0U };
            M.RightMargin rightMargin1 = new M.RightMargin(){ Val = (UInt32Value)0U };
            M.DefaultJustification defaultJustification1 = new M.DefaultJustification(){ Val = M.JustificationValues.CenterGroup };
            M.WrapIndent wrapIndent1 = new M.WrapIndent(){ Val = (UInt32Value)1440U };
            M.IntegralLimitLocation integralLimitLocation1 = new M.IntegralLimitLocation(){ Val = M.LimitLocationValues.SubscriptSuperscript };
            M.NaryLimitLocation naryLimitLocation1 = new M.NaryLimitLocation(){ Val = M.LimitLocationValues.UnderOver };

            mathProperties1.Append(mathFont1);
            mathProperties1.Append(breakBinary1);
            mathProperties1.Append(breakBinarySubtraction1);
            mathProperties1.Append(smallFraction1);
            mathProperties1.Append(displayDefaults1);
            mathProperties1.Append(leftMargin1);
            mathProperties1.Append(rightMargin1);
            mathProperties1.Append(defaultJustification1);
            mathProperties1.Append(wrapIndent1);
            mathProperties1.Append(integralLimitLocation1);
            mathProperties1.Append(naryLimitLocation1);
            ThemeFontLanguages themeFontLanguages1 = new ThemeFontLanguages(){ Val = "pl-PL" };
            ColorSchemeMapping colorSchemeMapping1 = new ColorSchemeMapping(){ Background1 = ColorSchemeIndexValues.Light1, Text1 = ColorSchemeIndexValues.Dark1, Background2 = ColorSchemeIndexValues.Light2, Text2 = ColorSchemeIndexValues.Dark2, Accent1 = ColorSchemeIndexValues.Accent1, Accent2 = ColorSchemeIndexValues.Accent2, Accent3 = ColorSchemeIndexValues.Accent3, Accent4 = ColorSchemeIndexValues.Accent4, Accent5 = ColorSchemeIndexValues.Accent5, Accent6 = ColorSchemeIndexValues.Accent6, Hyperlink = ColorSchemeIndexValues.Hyperlink, FollowedHyperlink = ColorSchemeIndexValues.FollowedHyperlink };

            ShapeDefaults shapeDefaults1 = new ShapeDefaults();
            Ovml.ShapeDefaults shapeDefaults2 = new Ovml.ShapeDefaults(){ Extension = V.ExtensionHandlingBehaviorValues.Edit, MaxShapeId = 1026 };

            Ovml.ShapeLayout shapeLayout1 = new Ovml.ShapeLayout(){ Extension = V.ExtensionHandlingBehaviorValues.Edit };
            Ovml.ShapeIdMap shapeIdMap1 = new Ovml.ShapeIdMap(){ Extension = V.ExtensionHandlingBehaviorValues.Edit, Data = "1" };

            shapeLayout1.Append(shapeIdMap1);

            shapeDefaults1.Append(shapeDefaults2);
            shapeDefaults1.Append(shapeLayout1);
            DecimalSymbol decimalSymbol1 = new DecimalSymbol(){ Val = "," };
            ListSeparator listSeparator1 = new ListSeparator(){ Val = ";" };
            W14.DocumentId documentId1 = new W14.DocumentId(){ Val = "393193E3" };
            W15.ChartTrackingRefBased chartTrackingRefBased1 = new W15.ChartTrackingRefBased();
            W15.PersistentDocumentId persistentDocumentId1 = new W15.PersistentDocumentId(){ Val = "{4CBB8C6D-3693-4734-B9BA-AA046105008B}" };

            settings1.Append(zoom1);
            settings1.Append(proofState1);
            settings1.Append(defaultTabStop1);
            settings1.Append(hyphenationZone1);
            settings1.Append(characterSpacingControl1);
            settings1.Append(compatibility1);
            settings1.Append(rsids1);
            settings1.Append(mathProperties1);
            settings1.Append(themeFontLanguages1);
            settings1.Append(colorSchemeMapping1);
            settings1.Append(shapeDefaults1);
            settings1.Append(decimalSymbol1);
            settings1.Append(listSeparator1);
            settings1.Append(documentId1);
            settings1.Append(chartTrackingRefBased1);
            settings1.Append(persistentDocumentId1);

            documentSettingsPart1.Settings = settings1;
        }

        // Generates content of styleDefinitionsPart1.
        private void GenerateStyleDefinitionsPart1Content(StyleDefinitionsPart styleDefinitionsPart1)
        {
            Styles styles1 = new Styles(){ MCAttributes = new MarkupCompatibilityAttributes(){ Ignorable = "w14 w15 w16se w16cid w16 w16cex w16sdtdh w16sdtfl w16du" }  };
            styles1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            styles1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            styles1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            styles1.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            styles1.AddNamespaceDeclaration("w15", "http://schemas.microsoft.com/office/word/2012/wordml");
            styles1.AddNamespaceDeclaration("w16cex", "http://schemas.microsoft.com/office/word/2018/wordml/cex");
            styles1.AddNamespaceDeclaration("w16cid", "http://schemas.microsoft.com/office/word/2016/wordml/cid");
            styles1.AddNamespaceDeclaration("w16", "http://schemas.microsoft.com/office/word/2018/wordml");
            styles1.AddNamespaceDeclaration("w16du", "http://schemas.microsoft.com/office/word/2023/wordml/word16du");
            styles1.AddNamespaceDeclaration("w16sdtdh", "http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash");
            styles1.AddNamespaceDeclaration("w16sdtfl", "http://schemas.microsoft.com/office/word/2024/wordml/sdtformatlock");
            styles1.AddNamespaceDeclaration("w16se", "http://schemas.microsoft.com/office/word/2015/wordml/symex");

            DocDefaults docDefaults1 = new DocDefaults();

            RunPropertiesDefault runPropertiesDefault1 = new RunPropertiesDefault();

            RunPropertiesBaseStyle runPropertiesBaseStyle1 = new RunPropertiesBaseStyle();
            RunFonts runFonts1 = new RunFonts(){ AsciiTheme = ThemeFontValues.MinorHighAnsi, HighAnsiTheme = ThemeFontValues.MinorHighAnsi, EastAsiaTheme = ThemeFontValues.MinorHighAnsi, ComplexScriptTheme = ThemeFontValues.MinorBidi };
            Kern kern1 = new Kern(){ Val = (UInt32Value)2U };
            FontSize fontSize1 = new FontSize(){ Val = "24" };
            FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript(){ Val = "24" };
            Languages languages1 = new Languages(){ Val = "pl-PL", EastAsia = "en-US", Bidi = "ar-SA" };

            OpenXmlUnknownElement openXmlUnknownElement1 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<w14:ligatures w14:val=\"standardContextual\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" />");

            runPropertiesBaseStyle1.Append(runFonts1);
            runPropertiesBaseStyle1.Append(kern1);
            runPropertiesBaseStyle1.Append(fontSize1);
            runPropertiesBaseStyle1.Append(fontSizeComplexScript1);
            runPropertiesBaseStyle1.Append(languages1);
            runPropertiesBaseStyle1.Append(openXmlUnknownElement1);

            runPropertiesDefault1.Append(runPropertiesBaseStyle1);

            ParagraphPropertiesDefault paragraphPropertiesDefault1 = new ParagraphPropertiesDefault();

            ParagraphPropertiesBaseStyle paragraphPropertiesBaseStyle1 = new ParagraphPropertiesBaseStyle();
            SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines(){ After = "160", Line = "278", LineRule = LineSpacingRuleValues.Auto };

            paragraphPropertiesBaseStyle1.Append(spacingBetweenLines1);

            paragraphPropertiesDefault1.Append(paragraphPropertiesBaseStyle1);

            docDefaults1.Append(runPropertiesDefault1);
            docDefaults1.Append(paragraphPropertiesDefault1);

            LatentStyles latentStyles1 = new LatentStyles(){ DefaultLockedState = false, DefaultUiPriority = 99, DefaultSemiHidden = false, DefaultUnhideWhenUsed = false, DefaultPrimaryStyle = false, Count = 376 };
            LatentStyleExceptionInfo latentStyleExceptionInfo1 = new LatentStyleExceptionInfo(){ Name = "Normal", UiPriority = 0, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo2 = new LatentStyleExceptionInfo(){ Name = "heading 1", UiPriority = 9, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo3 = new LatentStyleExceptionInfo(){ Name = "heading 2", UiPriority = 9, SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo4 = new LatentStyleExceptionInfo(){ Name = "heading 3", UiPriority = 9, SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo5 = new LatentStyleExceptionInfo(){ Name = "heading 4", UiPriority = 9, SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo6 = new LatentStyleExceptionInfo(){ Name = "heading 5", UiPriority = 9, SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo7 = new LatentStyleExceptionInfo(){ Name = "heading 6", UiPriority = 9, SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo8 = new LatentStyleExceptionInfo(){ Name = "heading 7", UiPriority = 9, SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo9 = new LatentStyleExceptionInfo(){ Name = "heading 8", UiPriority = 9, SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo10 = new LatentStyleExceptionInfo(){ Name = "heading 9", UiPriority = 9, SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo11 = new LatentStyleExceptionInfo(){ Name = "index 1", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo12 = new LatentStyleExceptionInfo(){ Name = "index 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo13 = new LatentStyleExceptionInfo(){ Name = "index 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo14 = new LatentStyleExceptionInfo(){ Name = "index 4", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo15 = new LatentStyleExceptionInfo(){ Name = "index 5", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo16 = new LatentStyleExceptionInfo(){ Name = "index 6", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo17 = new LatentStyleExceptionInfo(){ Name = "index 7", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo18 = new LatentStyleExceptionInfo(){ Name = "index 8", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo19 = new LatentStyleExceptionInfo(){ Name = "index 9", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo20 = new LatentStyleExceptionInfo(){ Name = "toc 1", UiPriority = 39, SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo21 = new LatentStyleExceptionInfo(){ Name = "toc 2", UiPriority = 39, SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo22 = new LatentStyleExceptionInfo(){ Name = "toc 3", UiPriority = 39, SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo23 = new LatentStyleExceptionInfo(){ Name = "toc 4", UiPriority = 39, SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo24 = new LatentStyleExceptionInfo(){ Name = "toc 5", UiPriority = 39, SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo25 = new LatentStyleExceptionInfo(){ Name = "toc 6", UiPriority = 39, SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo26 = new LatentStyleExceptionInfo(){ Name = "toc 7", UiPriority = 39, SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo27 = new LatentStyleExceptionInfo(){ Name = "toc 8", UiPriority = 39, SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo28 = new LatentStyleExceptionInfo(){ Name = "toc 9", UiPriority = 39, SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo29 = new LatentStyleExceptionInfo(){ Name = "Normal Indent", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo30 = new LatentStyleExceptionInfo(){ Name = "footnote text", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo31 = new LatentStyleExceptionInfo(){ Name = "annotation text", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo32 = new LatentStyleExceptionInfo(){ Name = "header", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo33 = new LatentStyleExceptionInfo(){ Name = "footer", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo34 = new LatentStyleExceptionInfo(){ Name = "index heading", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo35 = new LatentStyleExceptionInfo(){ Name = "caption", UiPriority = 35, SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo36 = new LatentStyleExceptionInfo(){ Name = "table of figures", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo37 = new LatentStyleExceptionInfo(){ Name = "envelope address", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo38 = new LatentStyleExceptionInfo(){ Name = "envelope return", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo39 = new LatentStyleExceptionInfo(){ Name = "footnote reference", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo40 = new LatentStyleExceptionInfo(){ Name = "annotation reference", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo41 = new LatentStyleExceptionInfo(){ Name = "line number", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo42 = new LatentStyleExceptionInfo(){ Name = "page number", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo43 = new LatentStyleExceptionInfo(){ Name = "endnote reference", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo44 = new LatentStyleExceptionInfo(){ Name = "endnote text", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo45 = new LatentStyleExceptionInfo(){ Name = "table of authorities", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo46 = new LatentStyleExceptionInfo(){ Name = "macro", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo47 = new LatentStyleExceptionInfo(){ Name = "toa heading", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo48 = new LatentStyleExceptionInfo(){ Name = "List", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo49 = new LatentStyleExceptionInfo(){ Name = "List Bullet", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo50 = new LatentStyleExceptionInfo(){ Name = "List Number", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo51 = new LatentStyleExceptionInfo(){ Name = "List 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo52 = new LatentStyleExceptionInfo(){ Name = "List 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo53 = new LatentStyleExceptionInfo(){ Name = "List 4", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo54 = new LatentStyleExceptionInfo(){ Name = "List 5", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo55 = new LatentStyleExceptionInfo(){ Name = "List Bullet 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo56 = new LatentStyleExceptionInfo(){ Name = "List Bullet 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo57 = new LatentStyleExceptionInfo(){ Name = "List Bullet 4", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo58 = new LatentStyleExceptionInfo(){ Name = "List Bullet 5", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo59 = new LatentStyleExceptionInfo(){ Name = "List Number 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo60 = new LatentStyleExceptionInfo(){ Name = "List Number 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo61 = new LatentStyleExceptionInfo(){ Name = "List Number 4", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo62 = new LatentStyleExceptionInfo(){ Name = "List Number 5", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo63 = new LatentStyleExceptionInfo(){ Name = "Title", UiPriority = 10, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo64 = new LatentStyleExceptionInfo(){ Name = "Closing", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo65 = new LatentStyleExceptionInfo(){ Name = "Signature", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo66 = new LatentStyleExceptionInfo(){ Name = "Default Paragraph Font", UiPriority = 1, SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo67 = new LatentStyleExceptionInfo(){ Name = "Body Text", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo68 = new LatentStyleExceptionInfo(){ Name = "Body Text Indent", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo69 = new LatentStyleExceptionInfo(){ Name = "List Continue", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo70 = new LatentStyleExceptionInfo(){ Name = "List Continue 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo71 = new LatentStyleExceptionInfo(){ Name = "List Continue 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo72 = new LatentStyleExceptionInfo(){ Name = "List Continue 4", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo73 = new LatentStyleExceptionInfo(){ Name = "List Continue 5", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo74 = new LatentStyleExceptionInfo(){ Name = "Message Header", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo75 = new LatentStyleExceptionInfo(){ Name = "Subtitle", UiPriority = 11, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo76 = new LatentStyleExceptionInfo(){ Name = "Salutation", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo77 = new LatentStyleExceptionInfo(){ Name = "Date", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo78 = new LatentStyleExceptionInfo(){ Name = "Body Text First Indent", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo79 = new LatentStyleExceptionInfo(){ Name = "Body Text First Indent 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo80 = new LatentStyleExceptionInfo(){ Name = "Note Heading", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo81 = new LatentStyleExceptionInfo(){ Name = "Body Text 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo82 = new LatentStyleExceptionInfo(){ Name = "Body Text 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo83 = new LatentStyleExceptionInfo(){ Name = "Body Text Indent 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo84 = new LatentStyleExceptionInfo(){ Name = "Body Text Indent 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo85 = new LatentStyleExceptionInfo(){ Name = "Block Text", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo86 = new LatentStyleExceptionInfo(){ Name = "Hyperlink", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo87 = new LatentStyleExceptionInfo(){ Name = "FollowedHyperlink", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo88 = new LatentStyleExceptionInfo(){ Name = "Strong", UiPriority = 22, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo89 = new LatentStyleExceptionInfo(){ Name = "Emphasis", UiPriority = 20, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo90 = new LatentStyleExceptionInfo(){ Name = "Document Map", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo91 = new LatentStyleExceptionInfo(){ Name = "Plain Text", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo92 = new LatentStyleExceptionInfo(){ Name = "E-mail Signature", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo93 = new LatentStyleExceptionInfo(){ Name = "HTML Top of Form", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo94 = new LatentStyleExceptionInfo(){ Name = "HTML Bottom of Form", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo95 = new LatentStyleExceptionInfo(){ Name = "Normal (Web)", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo96 = new LatentStyleExceptionInfo(){ Name = "HTML Acronym", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo97 = new LatentStyleExceptionInfo(){ Name = "HTML Address", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo98 = new LatentStyleExceptionInfo(){ Name = "HTML Cite", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo99 = new LatentStyleExceptionInfo(){ Name = "HTML Code", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo100 = new LatentStyleExceptionInfo(){ Name = "HTML Definition", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo101 = new LatentStyleExceptionInfo(){ Name = "HTML Keyboard", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo102 = new LatentStyleExceptionInfo(){ Name = "HTML Preformatted", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo103 = new LatentStyleExceptionInfo(){ Name = "HTML Sample", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo104 = new LatentStyleExceptionInfo(){ Name = "HTML Typewriter", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo105 = new LatentStyleExceptionInfo(){ Name = "HTML Variable", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo106 = new LatentStyleExceptionInfo(){ Name = "Normal Table", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo107 = new LatentStyleExceptionInfo(){ Name = "annotation subject", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo108 = new LatentStyleExceptionInfo(){ Name = "No List", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo109 = new LatentStyleExceptionInfo(){ Name = "Outline List 1", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo110 = new LatentStyleExceptionInfo(){ Name = "Outline List 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo111 = new LatentStyleExceptionInfo(){ Name = "Outline List 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo112 = new LatentStyleExceptionInfo(){ Name = "Table Simple 1", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo113 = new LatentStyleExceptionInfo(){ Name = "Table Simple 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo114 = new LatentStyleExceptionInfo(){ Name = "Table Simple 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo115 = new LatentStyleExceptionInfo(){ Name = "Table Classic 1", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo116 = new LatentStyleExceptionInfo(){ Name = "Table Classic 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo117 = new LatentStyleExceptionInfo(){ Name = "Table Classic 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo118 = new LatentStyleExceptionInfo(){ Name = "Table Classic 4", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo119 = new LatentStyleExceptionInfo(){ Name = "Table Colorful 1", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo120 = new LatentStyleExceptionInfo(){ Name = "Table Colorful 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo121 = new LatentStyleExceptionInfo(){ Name = "Table Colorful 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo122 = new LatentStyleExceptionInfo(){ Name = "Table Columns 1", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo123 = new LatentStyleExceptionInfo(){ Name = "Table Columns 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo124 = new LatentStyleExceptionInfo(){ Name = "Table Columns 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo125 = new LatentStyleExceptionInfo(){ Name = "Table Columns 4", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo126 = new LatentStyleExceptionInfo(){ Name = "Table Columns 5", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo127 = new LatentStyleExceptionInfo(){ Name = "Table Grid 1", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo128 = new LatentStyleExceptionInfo(){ Name = "Table Grid 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo129 = new LatentStyleExceptionInfo(){ Name = "Table Grid 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo130 = new LatentStyleExceptionInfo(){ Name = "Table Grid 4", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo131 = new LatentStyleExceptionInfo(){ Name = "Table Grid 5", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo132 = new LatentStyleExceptionInfo(){ Name = "Table Grid 6", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo133 = new LatentStyleExceptionInfo(){ Name = "Table Grid 7", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo134 = new LatentStyleExceptionInfo(){ Name = "Table Grid 8", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo135 = new LatentStyleExceptionInfo(){ Name = "Table List 1", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo136 = new LatentStyleExceptionInfo(){ Name = "Table List 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo137 = new LatentStyleExceptionInfo(){ Name = "Table List 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo138 = new LatentStyleExceptionInfo(){ Name = "Table List 4", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo139 = new LatentStyleExceptionInfo(){ Name = "Table List 5", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo140 = new LatentStyleExceptionInfo(){ Name = "Table List 6", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo141 = new LatentStyleExceptionInfo(){ Name = "Table List 7", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo142 = new LatentStyleExceptionInfo(){ Name = "Table List 8", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo143 = new LatentStyleExceptionInfo(){ Name = "Table 3D effects 1", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo144 = new LatentStyleExceptionInfo(){ Name = "Table 3D effects 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo145 = new LatentStyleExceptionInfo(){ Name = "Table 3D effects 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo146 = new LatentStyleExceptionInfo(){ Name = "Table Contemporary", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo147 = new LatentStyleExceptionInfo(){ Name = "Table Elegant", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo148 = new LatentStyleExceptionInfo(){ Name = "Table Professional", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo149 = new LatentStyleExceptionInfo(){ Name = "Table Subtle 1", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo150 = new LatentStyleExceptionInfo(){ Name = "Table Subtle 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo151 = new LatentStyleExceptionInfo(){ Name = "Table Web 1", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo152 = new LatentStyleExceptionInfo(){ Name = "Table Web 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo153 = new LatentStyleExceptionInfo(){ Name = "Table Web 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo154 = new LatentStyleExceptionInfo(){ Name = "Balloon Text", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo155 = new LatentStyleExceptionInfo(){ Name = "Table Grid", UiPriority = 39 };
            LatentStyleExceptionInfo latentStyleExceptionInfo156 = new LatentStyleExceptionInfo(){ Name = "Table Theme", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo157 = new LatentStyleExceptionInfo(){ Name = "Placeholder Text", SemiHidden = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo158 = new LatentStyleExceptionInfo(){ Name = "No Spacing", UiPriority = 1, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo159 = new LatentStyleExceptionInfo(){ Name = "Light Shading", UiPriority = 60 };
            LatentStyleExceptionInfo latentStyleExceptionInfo160 = new LatentStyleExceptionInfo(){ Name = "Light List", UiPriority = 61 };
            LatentStyleExceptionInfo latentStyleExceptionInfo161 = new LatentStyleExceptionInfo(){ Name = "Light Grid", UiPriority = 62 };
            LatentStyleExceptionInfo latentStyleExceptionInfo162 = new LatentStyleExceptionInfo(){ Name = "Medium Shading 1", UiPriority = 63 };
            LatentStyleExceptionInfo latentStyleExceptionInfo163 = new LatentStyleExceptionInfo(){ Name = "Medium Shading 2", UiPriority = 64 };
            LatentStyleExceptionInfo latentStyleExceptionInfo164 = new LatentStyleExceptionInfo(){ Name = "Medium List 1", UiPriority = 65 };
            LatentStyleExceptionInfo latentStyleExceptionInfo165 = new LatentStyleExceptionInfo(){ Name = "Medium List 2", UiPriority = 66 };
            LatentStyleExceptionInfo latentStyleExceptionInfo166 = new LatentStyleExceptionInfo(){ Name = "Medium Grid 1", UiPriority = 67 };
            LatentStyleExceptionInfo latentStyleExceptionInfo167 = new LatentStyleExceptionInfo(){ Name = "Medium Grid 2", UiPriority = 68 };
            LatentStyleExceptionInfo latentStyleExceptionInfo168 = new LatentStyleExceptionInfo(){ Name = "Medium Grid 3", UiPriority = 69 };
            LatentStyleExceptionInfo latentStyleExceptionInfo169 = new LatentStyleExceptionInfo(){ Name = "Dark List", UiPriority = 70 };
            LatentStyleExceptionInfo latentStyleExceptionInfo170 = new LatentStyleExceptionInfo(){ Name = "Colorful Shading", UiPriority = 71 };
            LatentStyleExceptionInfo latentStyleExceptionInfo171 = new LatentStyleExceptionInfo(){ Name = "Colorful List", UiPriority = 72 };
            LatentStyleExceptionInfo latentStyleExceptionInfo172 = new LatentStyleExceptionInfo(){ Name = "Colorful Grid", UiPriority = 73 };
            LatentStyleExceptionInfo latentStyleExceptionInfo173 = new LatentStyleExceptionInfo(){ Name = "Light Shading Accent 1", UiPriority = 60 };
            LatentStyleExceptionInfo latentStyleExceptionInfo174 = new LatentStyleExceptionInfo(){ Name = "Light List Accent 1", UiPriority = 61 };
            LatentStyleExceptionInfo latentStyleExceptionInfo175 = new LatentStyleExceptionInfo(){ Name = "Light Grid Accent 1", UiPriority = 62 };
            LatentStyleExceptionInfo latentStyleExceptionInfo176 = new LatentStyleExceptionInfo(){ Name = "Medium Shading 1 Accent 1", UiPriority = 63 };
            LatentStyleExceptionInfo latentStyleExceptionInfo177 = new LatentStyleExceptionInfo(){ Name = "Medium Shading 2 Accent 1", UiPriority = 64 };
            LatentStyleExceptionInfo latentStyleExceptionInfo178 = new LatentStyleExceptionInfo(){ Name = "Medium List 1 Accent 1", UiPriority = 65 };
            LatentStyleExceptionInfo latentStyleExceptionInfo179 = new LatentStyleExceptionInfo(){ Name = "Revision", SemiHidden = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo180 = new LatentStyleExceptionInfo(){ Name = "List Paragraph", UiPriority = 34, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo181 = new LatentStyleExceptionInfo(){ Name = "Quote", UiPriority = 29, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo182 = new LatentStyleExceptionInfo(){ Name = "Intense Quote", UiPriority = 30, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo183 = new LatentStyleExceptionInfo(){ Name = "Medium List 2 Accent 1", UiPriority = 66 };
            LatentStyleExceptionInfo latentStyleExceptionInfo184 = new LatentStyleExceptionInfo(){ Name = "Medium Grid 1 Accent 1", UiPriority = 67 };
            LatentStyleExceptionInfo latentStyleExceptionInfo185 = new LatentStyleExceptionInfo(){ Name = "Medium Grid 2 Accent 1", UiPriority = 68 };
            LatentStyleExceptionInfo latentStyleExceptionInfo186 = new LatentStyleExceptionInfo(){ Name = "Medium Grid 3 Accent 1", UiPriority = 69 };
            LatentStyleExceptionInfo latentStyleExceptionInfo187 = new LatentStyleExceptionInfo(){ Name = "Dark List Accent 1", UiPriority = 70 };
            LatentStyleExceptionInfo latentStyleExceptionInfo188 = new LatentStyleExceptionInfo(){ Name = "Colorful Shading Accent 1", UiPriority = 71 };
            LatentStyleExceptionInfo latentStyleExceptionInfo189 = new LatentStyleExceptionInfo(){ Name = "Colorful List Accent 1", UiPriority = 72 };
            LatentStyleExceptionInfo latentStyleExceptionInfo190 = new LatentStyleExceptionInfo(){ Name = "Colorful Grid Accent 1", UiPriority = 73 };
            LatentStyleExceptionInfo latentStyleExceptionInfo191 = new LatentStyleExceptionInfo(){ Name = "Light Shading Accent 2", UiPriority = 60 };
            LatentStyleExceptionInfo latentStyleExceptionInfo192 = new LatentStyleExceptionInfo(){ Name = "Light List Accent 2", UiPriority = 61 };
            LatentStyleExceptionInfo latentStyleExceptionInfo193 = new LatentStyleExceptionInfo(){ Name = "Light Grid Accent 2", UiPriority = 62 };
            LatentStyleExceptionInfo latentStyleExceptionInfo194 = new LatentStyleExceptionInfo(){ Name = "Medium Shading 1 Accent 2", UiPriority = 63 };
            LatentStyleExceptionInfo latentStyleExceptionInfo195 = new LatentStyleExceptionInfo(){ Name = "Medium Shading 2 Accent 2", UiPriority = 64 };
            LatentStyleExceptionInfo latentStyleExceptionInfo196 = new LatentStyleExceptionInfo(){ Name = "Medium List 1 Accent 2", UiPriority = 65 };
            LatentStyleExceptionInfo latentStyleExceptionInfo197 = new LatentStyleExceptionInfo(){ Name = "Medium List 2 Accent 2", UiPriority = 66 };
            LatentStyleExceptionInfo latentStyleExceptionInfo198 = new LatentStyleExceptionInfo(){ Name = "Medium Grid 1 Accent 2", UiPriority = 67 };
            LatentStyleExceptionInfo latentStyleExceptionInfo199 = new LatentStyleExceptionInfo(){ Name = "Medium Grid 2 Accent 2", UiPriority = 68 };
            LatentStyleExceptionInfo latentStyleExceptionInfo200 = new LatentStyleExceptionInfo(){ Name = "Medium Grid 3 Accent 2", UiPriority = 69 };
            LatentStyleExceptionInfo latentStyleExceptionInfo201 = new LatentStyleExceptionInfo(){ Name = "Dark List Accent 2", UiPriority = 70 };
            LatentStyleExceptionInfo latentStyleExceptionInfo202 = new LatentStyleExceptionInfo(){ Name = "Colorful Shading Accent 2", UiPriority = 71 };
            LatentStyleExceptionInfo latentStyleExceptionInfo203 = new LatentStyleExceptionInfo(){ Name = "Colorful List Accent 2", UiPriority = 72 };
            LatentStyleExceptionInfo latentStyleExceptionInfo204 = new LatentStyleExceptionInfo(){ Name = "Colorful Grid Accent 2", UiPriority = 73 };
            LatentStyleExceptionInfo latentStyleExceptionInfo205 = new LatentStyleExceptionInfo(){ Name = "Light Shading Accent 3", UiPriority = 60 };
            LatentStyleExceptionInfo latentStyleExceptionInfo206 = new LatentStyleExceptionInfo(){ Name = "Light List Accent 3", UiPriority = 61 };
            LatentStyleExceptionInfo latentStyleExceptionInfo207 = new LatentStyleExceptionInfo(){ Name = "Light Grid Accent 3", UiPriority = 62 };
            LatentStyleExceptionInfo latentStyleExceptionInfo208 = new LatentStyleExceptionInfo(){ Name = "Medium Shading 1 Accent 3", UiPriority = 63 };
            LatentStyleExceptionInfo latentStyleExceptionInfo209 = new LatentStyleExceptionInfo(){ Name = "Medium Shading 2 Accent 3", UiPriority = 64 };
            LatentStyleExceptionInfo latentStyleExceptionInfo210 = new LatentStyleExceptionInfo(){ Name = "Medium List 1 Accent 3", UiPriority = 65 };
            LatentStyleExceptionInfo latentStyleExceptionInfo211 = new LatentStyleExceptionInfo(){ Name = "Medium List 2 Accent 3", UiPriority = 66 };
            LatentStyleExceptionInfo latentStyleExceptionInfo212 = new LatentStyleExceptionInfo(){ Name = "Medium Grid 1 Accent 3", UiPriority = 67 };
            LatentStyleExceptionInfo latentStyleExceptionInfo213 = new LatentStyleExceptionInfo(){ Name = "Medium Grid 2 Accent 3", UiPriority = 68 };
            LatentStyleExceptionInfo latentStyleExceptionInfo214 = new LatentStyleExceptionInfo(){ Name = "Medium Grid 3 Accent 3", UiPriority = 69 };
            LatentStyleExceptionInfo latentStyleExceptionInfo215 = new LatentStyleExceptionInfo(){ Name = "Dark List Accent 3", UiPriority = 70 };
            LatentStyleExceptionInfo latentStyleExceptionInfo216 = new LatentStyleExceptionInfo(){ Name = "Colorful Shading Accent 3", UiPriority = 71 };
            LatentStyleExceptionInfo latentStyleExceptionInfo217 = new LatentStyleExceptionInfo(){ Name = "Colorful List Accent 3", UiPriority = 72 };
            LatentStyleExceptionInfo latentStyleExceptionInfo218 = new LatentStyleExceptionInfo(){ Name = "Colorful Grid Accent 3", UiPriority = 73 };
            LatentStyleExceptionInfo latentStyleExceptionInfo219 = new LatentStyleExceptionInfo(){ Name = "Light Shading Accent 4", UiPriority = 60 };
            LatentStyleExceptionInfo latentStyleExceptionInfo220 = new LatentStyleExceptionInfo(){ Name = "Light List Accent 4", UiPriority = 61 };
            LatentStyleExceptionInfo latentStyleExceptionInfo221 = new LatentStyleExceptionInfo(){ Name = "Light Grid Accent 4", UiPriority = 62 };
            LatentStyleExceptionInfo latentStyleExceptionInfo222 = new LatentStyleExceptionInfo(){ Name = "Medium Shading 1 Accent 4", UiPriority = 63 };
            LatentStyleExceptionInfo latentStyleExceptionInfo223 = new LatentStyleExceptionInfo(){ Name = "Medium Shading 2 Accent 4", UiPriority = 64 };
            LatentStyleExceptionInfo latentStyleExceptionInfo224 = new LatentStyleExceptionInfo(){ Name = "Medium List 1 Accent 4", UiPriority = 65 };
            LatentStyleExceptionInfo latentStyleExceptionInfo225 = new LatentStyleExceptionInfo(){ Name = "Medium List 2 Accent 4", UiPriority = 66 };
            LatentStyleExceptionInfo latentStyleExceptionInfo226 = new LatentStyleExceptionInfo(){ Name = "Medium Grid 1 Accent 4", UiPriority = 67 };
            LatentStyleExceptionInfo latentStyleExceptionInfo227 = new LatentStyleExceptionInfo(){ Name = "Medium Grid 2 Accent 4", UiPriority = 68 };
            LatentStyleExceptionInfo latentStyleExceptionInfo228 = new LatentStyleExceptionInfo(){ Name = "Medium Grid 3 Accent 4", UiPriority = 69 };
            LatentStyleExceptionInfo latentStyleExceptionInfo229 = new LatentStyleExceptionInfo(){ Name = "Dark List Accent 4", UiPriority = 70 };
            LatentStyleExceptionInfo latentStyleExceptionInfo230 = new LatentStyleExceptionInfo(){ Name = "Colorful Shading Accent 4", UiPriority = 71 };
            LatentStyleExceptionInfo latentStyleExceptionInfo231 = new LatentStyleExceptionInfo(){ Name = "Colorful List Accent 4", UiPriority = 72 };
            LatentStyleExceptionInfo latentStyleExceptionInfo232 = new LatentStyleExceptionInfo(){ Name = "Colorful Grid Accent 4", UiPriority = 73 };
            LatentStyleExceptionInfo latentStyleExceptionInfo233 = new LatentStyleExceptionInfo(){ Name = "Light Shading Accent 5", UiPriority = 60 };
            LatentStyleExceptionInfo latentStyleExceptionInfo234 = new LatentStyleExceptionInfo(){ Name = "Light List Accent 5", UiPriority = 61 };
            LatentStyleExceptionInfo latentStyleExceptionInfo235 = new LatentStyleExceptionInfo(){ Name = "Light Grid Accent 5", UiPriority = 62 };
            LatentStyleExceptionInfo latentStyleExceptionInfo236 = new LatentStyleExceptionInfo(){ Name = "Medium Shading 1 Accent 5", UiPriority = 63 };
            LatentStyleExceptionInfo latentStyleExceptionInfo237 = new LatentStyleExceptionInfo(){ Name = "Medium Shading 2 Accent 5", UiPriority = 64 };
            LatentStyleExceptionInfo latentStyleExceptionInfo238 = new LatentStyleExceptionInfo(){ Name = "Medium List 1 Accent 5", UiPriority = 65 };
            LatentStyleExceptionInfo latentStyleExceptionInfo239 = new LatentStyleExceptionInfo(){ Name = "Medium List 2 Accent 5", UiPriority = 66 };
            LatentStyleExceptionInfo latentStyleExceptionInfo240 = new LatentStyleExceptionInfo(){ Name = "Medium Grid 1 Accent 5", UiPriority = 67 };
            LatentStyleExceptionInfo latentStyleExceptionInfo241 = new LatentStyleExceptionInfo(){ Name = "Medium Grid 2 Accent 5", UiPriority = 68 };
            LatentStyleExceptionInfo latentStyleExceptionInfo242 = new LatentStyleExceptionInfo(){ Name = "Medium Grid 3 Accent 5", UiPriority = 69 };
            LatentStyleExceptionInfo latentStyleExceptionInfo243 = new LatentStyleExceptionInfo(){ Name = "Dark List Accent 5", UiPriority = 70 };
            LatentStyleExceptionInfo latentStyleExceptionInfo244 = new LatentStyleExceptionInfo(){ Name = "Colorful Shading Accent 5", UiPriority = 71 };
            LatentStyleExceptionInfo latentStyleExceptionInfo245 = new LatentStyleExceptionInfo(){ Name = "Colorful List Accent 5", UiPriority = 72 };
            LatentStyleExceptionInfo latentStyleExceptionInfo246 = new LatentStyleExceptionInfo(){ Name = "Colorful Grid Accent 5", UiPriority = 73 };
            LatentStyleExceptionInfo latentStyleExceptionInfo247 = new LatentStyleExceptionInfo(){ Name = "Light Shading Accent 6", UiPriority = 60 };
            LatentStyleExceptionInfo latentStyleExceptionInfo248 = new LatentStyleExceptionInfo(){ Name = "Light List Accent 6", UiPriority = 61 };
            LatentStyleExceptionInfo latentStyleExceptionInfo249 = new LatentStyleExceptionInfo(){ Name = "Light Grid Accent 6", UiPriority = 62 };
            LatentStyleExceptionInfo latentStyleExceptionInfo250 = new LatentStyleExceptionInfo(){ Name = "Medium Shading 1 Accent 6", UiPriority = 63 };
            LatentStyleExceptionInfo latentStyleExceptionInfo251 = new LatentStyleExceptionInfo(){ Name = "Medium Shading 2 Accent 6", UiPriority = 64 };
            LatentStyleExceptionInfo latentStyleExceptionInfo252 = new LatentStyleExceptionInfo(){ Name = "Medium List 1 Accent 6", UiPriority = 65 };
            LatentStyleExceptionInfo latentStyleExceptionInfo253 = new LatentStyleExceptionInfo(){ Name = "Medium List 2 Accent 6", UiPriority = 66 };
            LatentStyleExceptionInfo latentStyleExceptionInfo254 = new LatentStyleExceptionInfo(){ Name = "Medium Grid 1 Accent 6", UiPriority = 67 };
            LatentStyleExceptionInfo latentStyleExceptionInfo255 = new LatentStyleExceptionInfo(){ Name = "Medium Grid 2 Accent 6", UiPriority = 68 };
            LatentStyleExceptionInfo latentStyleExceptionInfo256 = new LatentStyleExceptionInfo(){ Name = "Medium Grid 3 Accent 6", UiPriority = 69 };
            LatentStyleExceptionInfo latentStyleExceptionInfo257 = new LatentStyleExceptionInfo(){ Name = "Dark List Accent 6", UiPriority = 70 };
            LatentStyleExceptionInfo latentStyleExceptionInfo258 = new LatentStyleExceptionInfo(){ Name = "Colorful Shading Accent 6", UiPriority = 71 };
            LatentStyleExceptionInfo latentStyleExceptionInfo259 = new LatentStyleExceptionInfo(){ Name = "Colorful List Accent 6", UiPriority = 72 };
            LatentStyleExceptionInfo latentStyleExceptionInfo260 = new LatentStyleExceptionInfo(){ Name = "Colorful Grid Accent 6", UiPriority = 73 };
            LatentStyleExceptionInfo latentStyleExceptionInfo261 = new LatentStyleExceptionInfo(){ Name = "Subtle Emphasis", UiPriority = 19, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo262 = new LatentStyleExceptionInfo(){ Name = "Intense Emphasis", UiPriority = 21, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo263 = new LatentStyleExceptionInfo(){ Name = "Subtle Reference", UiPriority = 31, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo264 = new LatentStyleExceptionInfo(){ Name = "Intense Reference", UiPriority = 32, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo265 = new LatentStyleExceptionInfo(){ Name = "Book Title", UiPriority = 33, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo266 = new LatentStyleExceptionInfo(){ Name = "Bibliography", UiPriority = 37, SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo267 = new LatentStyleExceptionInfo(){ Name = "TOC Heading", UiPriority = 39, SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo268 = new LatentStyleExceptionInfo(){ Name = "Plain Table 1", UiPriority = 41 };
            LatentStyleExceptionInfo latentStyleExceptionInfo269 = new LatentStyleExceptionInfo(){ Name = "Plain Table 2", UiPriority = 42 };
            LatentStyleExceptionInfo latentStyleExceptionInfo270 = new LatentStyleExceptionInfo(){ Name = "Plain Table 3", UiPriority = 43 };
            LatentStyleExceptionInfo latentStyleExceptionInfo271 = new LatentStyleExceptionInfo(){ Name = "Plain Table 4", UiPriority = 44 };
            LatentStyleExceptionInfo latentStyleExceptionInfo272 = new LatentStyleExceptionInfo(){ Name = "Plain Table 5", UiPriority = 45 };
            LatentStyleExceptionInfo latentStyleExceptionInfo273 = new LatentStyleExceptionInfo(){ Name = "Grid Table Light", UiPriority = 40 };
            LatentStyleExceptionInfo latentStyleExceptionInfo274 = new LatentStyleExceptionInfo(){ Name = "Grid Table 1 Light", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo275 = new LatentStyleExceptionInfo(){ Name = "Grid Table 2", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo276 = new LatentStyleExceptionInfo(){ Name = "Grid Table 3", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo277 = new LatentStyleExceptionInfo(){ Name = "Grid Table 4", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo278 = new LatentStyleExceptionInfo(){ Name = "Grid Table 5 Dark", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo279 = new LatentStyleExceptionInfo(){ Name = "Grid Table 6 Colorful", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo280 = new LatentStyleExceptionInfo(){ Name = "Grid Table 7 Colorful", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo281 = new LatentStyleExceptionInfo(){ Name = "Grid Table 1 Light Accent 1", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo282 = new LatentStyleExceptionInfo(){ Name = "Grid Table 2 Accent 1", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo283 = new LatentStyleExceptionInfo(){ Name = "Grid Table 3 Accent 1", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo284 = new LatentStyleExceptionInfo(){ Name = "Grid Table 4 Accent 1", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo285 = new LatentStyleExceptionInfo(){ Name = "Grid Table 5 Dark Accent 1", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo286 = new LatentStyleExceptionInfo(){ Name = "Grid Table 6 Colorful Accent 1", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo287 = new LatentStyleExceptionInfo(){ Name = "Grid Table 7 Colorful Accent 1", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo288 = new LatentStyleExceptionInfo(){ Name = "Grid Table 1 Light Accent 2", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo289 = new LatentStyleExceptionInfo(){ Name = "Grid Table 2 Accent 2", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo290 = new LatentStyleExceptionInfo(){ Name = "Grid Table 3 Accent 2", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo291 = new LatentStyleExceptionInfo(){ Name = "Grid Table 4 Accent 2", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo292 = new LatentStyleExceptionInfo(){ Name = "Grid Table 5 Dark Accent 2", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo293 = new LatentStyleExceptionInfo(){ Name = "Grid Table 6 Colorful Accent 2", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo294 = new LatentStyleExceptionInfo(){ Name = "Grid Table 7 Colorful Accent 2", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo295 = new LatentStyleExceptionInfo(){ Name = "Grid Table 1 Light Accent 3", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo296 = new LatentStyleExceptionInfo(){ Name = "Grid Table 2 Accent 3", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo297 = new LatentStyleExceptionInfo(){ Name = "Grid Table 3 Accent 3", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo298 = new LatentStyleExceptionInfo(){ Name = "Grid Table 4 Accent 3", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo299 = new LatentStyleExceptionInfo(){ Name = "Grid Table 5 Dark Accent 3", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo300 = new LatentStyleExceptionInfo(){ Name = "Grid Table 6 Colorful Accent 3", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo301 = new LatentStyleExceptionInfo(){ Name = "Grid Table 7 Colorful Accent 3", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo302 = new LatentStyleExceptionInfo(){ Name = "Grid Table 1 Light Accent 4", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo303 = new LatentStyleExceptionInfo(){ Name = "Grid Table 2 Accent 4", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo304 = new LatentStyleExceptionInfo(){ Name = "Grid Table 3 Accent 4", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo305 = new LatentStyleExceptionInfo(){ Name = "Grid Table 4 Accent 4", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo306 = new LatentStyleExceptionInfo(){ Name = "Grid Table 5 Dark Accent 4", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo307 = new LatentStyleExceptionInfo(){ Name = "Grid Table 6 Colorful Accent 4", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo308 = new LatentStyleExceptionInfo(){ Name = "Grid Table 7 Colorful Accent 4", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo309 = new LatentStyleExceptionInfo(){ Name = "Grid Table 1 Light Accent 5", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo310 = new LatentStyleExceptionInfo(){ Name = "Grid Table 2 Accent 5", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo311 = new LatentStyleExceptionInfo(){ Name = "Grid Table 3 Accent 5", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo312 = new LatentStyleExceptionInfo(){ Name = "Grid Table 4 Accent 5", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo313 = new LatentStyleExceptionInfo(){ Name = "Grid Table 5 Dark Accent 5", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo314 = new LatentStyleExceptionInfo(){ Name = "Grid Table 6 Colorful Accent 5", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo315 = new LatentStyleExceptionInfo(){ Name = "Grid Table 7 Colorful Accent 5", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo316 = new LatentStyleExceptionInfo(){ Name = "Grid Table 1 Light Accent 6", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo317 = new LatentStyleExceptionInfo(){ Name = "Grid Table 2 Accent 6", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo318 = new LatentStyleExceptionInfo(){ Name = "Grid Table 3 Accent 6", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo319 = new LatentStyleExceptionInfo(){ Name = "Grid Table 4 Accent 6", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo320 = new LatentStyleExceptionInfo(){ Name = "Grid Table 5 Dark Accent 6", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo321 = new LatentStyleExceptionInfo(){ Name = "Grid Table 6 Colorful Accent 6", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo322 = new LatentStyleExceptionInfo(){ Name = "Grid Table 7 Colorful Accent 6", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo323 = new LatentStyleExceptionInfo(){ Name = "List Table 1 Light", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo324 = new LatentStyleExceptionInfo(){ Name = "List Table 2", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo325 = new LatentStyleExceptionInfo(){ Name = "List Table 3", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo326 = new LatentStyleExceptionInfo(){ Name = "List Table 4", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo327 = new LatentStyleExceptionInfo(){ Name = "List Table 5 Dark", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo328 = new LatentStyleExceptionInfo(){ Name = "List Table 6 Colorful", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo329 = new LatentStyleExceptionInfo(){ Name = "List Table 7 Colorful", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo330 = new LatentStyleExceptionInfo(){ Name = "List Table 1 Light Accent 1", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo331 = new LatentStyleExceptionInfo(){ Name = "List Table 2 Accent 1", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo332 = new LatentStyleExceptionInfo(){ Name = "List Table 3 Accent 1", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo333 = new LatentStyleExceptionInfo(){ Name = "List Table 4 Accent 1", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo334 = new LatentStyleExceptionInfo(){ Name = "List Table 5 Dark Accent 1", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo335 = new LatentStyleExceptionInfo(){ Name = "List Table 6 Colorful Accent 1", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo336 = new LatentStyleExceptionInfo(){ Name = "List Table 7 Colorful Accent 1", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo337 = new LatentStyleExceptionInfo(){ Name = "List Table 1 Light Accent 2", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo338 = new LatentStyleExceptionInfo(){ Name = "List Table 2 Accent 2", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo339 = new LatentStyleExceptionInfo(){ Name = "List Table 3 Accent 2", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo340 = new LatentStyleExceptionInfo(){ Name = "List Table 4 Accent 2", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo341 = new LatentStyleExceptionInfo(){ Name = "List Table 5 Dark Accent 2", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo342 = new LatentStyleExceptionInfo(){ Name = "List Table 6 Colorful Accent 2", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo343 = new LatentStyleExceptionInfo(){ Name = "List Table 7 Colorful Accent 2", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo344 = new LatentStyleExceptionInfo(){ Name = "List Table 1 Light Accent 3", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo345 = new LatentStyleExceptionInfo(){ Name = "List Table 2 Accent 3", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo346 = new LatentStyleExceptionInfo(){ Name = "List Table 3 Accent 3", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo347 = new LatentStyleExceptionInfo(){ Name = "List Table 4 Accent 3", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo348 = new LatentStyleExceptionInfo(){ Name = "List Table 5 Dark Accent 3", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo349 = new LatentStyleExceptionInfo(){ Name = "List Table 6 Colorful Accent 3", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo350 = new LatentStyleExceptionInfo(){ Name = "List Table 7 Colorful Accent 3", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo351 = new LatentStyleExceptionInfo(){ Name = "List Table 1 Light Accent 4", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo352 = new LatentStyleExceptionInfo(){ Name = "List Table 2 Accent 4", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo353 = new LatentStyleExceptionInfo(){ Name = "List Table 3 Accent 4", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo354 = new LatentStyleExceptionInfo(){ Name = "List Table 4 Accent 4", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo355 = new LatentStyleExceptionInfo(){ Name = "List Table 5 Dark Accent 4", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo356 = new LatentStyleExceptionInfo(){ Name = "List Table 6 Colorful Accent 4", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo357 = new LatentStyleExceptionInfo(){ Name = "List Table 7 Colorful Accent 4", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo358 = new LatentStyleExceptionInfo(){ Name = "List Table 1 Light Accent 5", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo359 = new LatentStyleExceptionInfo(){ Name = "List Table 2 Accent 5", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo360 = new LatentStyleExceptionInfo(){ Name = "List Table 3 Accent 5", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo361 = new LatentStyleExceptionInfo(){ Name = "List Table 4 Accent 5", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo362 = new LatentStyleExceptionInfo(){ Name = "List Table 5 Dark Accent 5", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo363 = new LatentStyleExceptionInfo(){ Name = "List Table 6 Colorful Accent 5", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo364 = new LatentStyleExceptionInfo(){ Name = "List Table 7 Colorful Accent 5", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo365 = new LatentStyleExceptionInfo(){ Name = "List Table 1 Light Accent 6", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo366 = new LatentStyleExceptionInfo(){ Name = "List Table 2 Accent 6", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo367 = new LatentStyleExceptionInfo(){ Name = "List Table 3 Accent 6", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo368 = new LatentStyleExceptionInfo(){ Name = "List Table 4 Accent 6", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo369 = new LatentStyleExceptionInfo(){ Name = "List Table 5 Dark Accent 6", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo370 = new LatentStyleExceptionInfo(){ Name = "List Table 6 Colorful Accent 6", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo371 = new LatentStyleExceptionInfo(){ Name = "List Table 7 Colorful Accent 6", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo372 = new LatentStyleExceptionInfo(){ Name = "Mention", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo373 = new LatentStyleExceptionInfo(){ Name = "Smart Hyperlink", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo374 = new LatentStyleExceptionInfo(){ Name = "Hashtag", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo375 = new LatentStyleExceptionInfo(){ Name = "Unresolved Mention", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo376 = new LatentStyleExceptionInfo(){ Name = "Smart Link", SemiHidden = true, UnhideWhenUsed = true };

            latentStyles1.Append(latentStyleExceptionInfo1);
            latentStyles1.Append(latentStyleExceptionInfo2);
            latentStyles1.Append(latentStyleExceptionInfo3);
            latentStyles1.Append(latentStyleExceptionInfo4);
            latentStyles1.Append(latentStyleExceptionInfo5);
            latentStyles1.Append(latentStyleExceptionInfo6);
            latentStyles1.Append(latentStyleExceptionInfo7);
            latentStyles1.Append(latentStyleExceptionInfo8);
            latentStyles1.Append(latentStyleExceptionInfo9);
            latentStyles1.Append(latentStyleExceptionInfo10);
            latentStyles1.Append(latentStyleExceptionInfo11);
            latentStyles1.Append(latentStyleExceptionInfo12);
            latentStyles1.Append(latentStyleExceptionInfo13);
            latentStyles1.Append(latentStyleExceptionInfo14);
            latentStyles1.Append(latentStyleExceptionInfo15);
            latentStyles1.Append(latentStyleExceptionInfo16);
            latentStyles1.Append(latentStyleExceptionInfo17);
            latentStyles1.Append(latentStyleExceptionInfo18);
            latentStyles1.Append(latentStyleExceptionInfo19);
            latentStyles1.Append(latentStyleExceptionInfo20);
            latentStyles1.Append(latentStyleExceptionInfo21);
            latentStyles1.Append(latentStyleExceptionInfo22);
            latentStyles1.Append(latentStyleExceptionInfo23);
            latentStyles1.Append(latentStyleExceptionInfo24);
            latentStyles1.Append(latentStyleExceptionInfo25);
            latentStyles1.Append(latentStyleExceptionInfo26);
            latentStyles1.Append(latentStyleExceptionInfo27);
            latentStyles1.Append(latentStyleExceptionInfo28);
            latentStyles1.Append(latentStyleExceptionInfo29);
            latentStyles1.Append(latentStyleExceptionInfo30);
            latentStyles1.Append(latentStyleExceptionInfo31);
            latentStyles1.Append(latentStyleExceptionInfo32);
            latentStyles1.Append(latentStyleExceptionInfo33);
            latentStyles1.Append(latentStyleExceptionInfo34);
            latentStyles1.Append(latentStyleExceptionInfo35);
            latentStyles1.Append(latentStyleExceptionInfo36);
            latentStyles1.Append(latentStyleExceptionInfo37);
            latentStyles1.Append(latentStyleExceptionInfo38);
            latentStyles1.Append(latentStyleExceptionInfo39);
            latentStyles1.Append(latentStyleExceptionInfo40);
            latentStyles1.Append(latentStyleExceptionInfo41);
            latentStyles1.Append(latentStyleExceptionInfo42);
            latentStyles1.Append(latentStyleExceptionInfo43);
            latentStyles1.Append(latentStyleExceptionInfo44);
            latentStyles1.Append(latentStyleExceptionInfo45);
            latentStyles1.Append(latentStyleExceptionInfo46);
            latentStyles1.Append(latentStyleExceptionInfo47);
            latentStyles1.Append(latentStyleExceptionInfo48);
            latentStyles1.Append(latentStyleExceptionInfo49);
            latentStyles1.Append(latentStyleExceptionInfo50);
            latentStyles1.Append(latentStyleExceptionInfo51);
            latentStyles1.Append(latentStyleExceptionInfo52);
            latentStyles1.Append(latentStyleExceptionInfo53);
            latentStyles1.Append(latentStyleExceptionInfo54);
            latentStyles1.Append(latentStyleExceptionInfo55);
            latentStyles1.Append(latentStyleExceptionInfo56);
            latentStyles1.Append(latentStyleExceptionInfo57);
            latentStyles1.Append(latentStyleExceptionInfo58);
            latentStyles1.Append(latentStyleExceptionInfo59);
            latentStyles1.Append(latentStyleExceptionInfo60);
            latentStyles1.Append(latentStyleExceptionInfo61);
            latentStyles1.Append(latentStyleExceptionInfo62);
            latentStyles1.Append(latentStyleExceptionInfo63);
            latentStyles1.Append(latentStyleExceptionInfo64);
            latentStyles1.Append(latentStyleExceptionInfo65);
            latentStyles1.Append(latentStyleExceptionInfo66);
            latentStyles1.Append(latentStyleExceptionInfo67);
            latentStyles1.Append(latentStyleExceptionInfo68);
            latentStyles1.Append(latentStyleExceptionInfo69);
            latentStyles1.Append(latentStyleExceptionInfo70);
            latentStyles1.Append(latentStyleExceptionInfo71);
            latentStyles1.Append(latentStyleExceptionInfo72);
            latentStyles1.Append(latentStyleExceptionInfo73);
            latentStyles1.Append(latentStyleExceptionInfo74);
            latentStyles1.Append(latentStyleExceptionInfo75);
            latentStyles1.Append(latentStyleExceptionInfo76);
            latentStyles1.Append(latentStyleExceptionInfo77);
            latentStyles1.Append(latentStyleExceptionInfo78);
            latentStyles1.Append(latentStyleExceptionInfo79);
            latentStyles1.Append(latentStyleExceptionInfo80);
            latentStyles1.Append(latentStyleExceptionInfo81);
            latentStyles1.Append(latentStyleExceptionInfo82);
            latentStyles1.Append(latentStyleExceptionInfo83);
            latentStyles1.Append(latentStyleExceptionInfo84);
            latentStyles1.Append(latentStyleExceptionInfo85);
            latentStyles1.Append(latentStyleExceptionInfo86);
            latentStyles1.Append(latentStyleExceptionInfo87);
            latentStyles1.Append(latentStyleExceptionInfo88);
            latentStyles1.Append(latentStyleExceptionInfo89);
            latentStyles1.Append(latentStyleExceptionInfo90);
            latentStyles1.Append(latentStyleExceptionInfo91);
            latentStyles1.Append(latentStyleExceptionInfo92);
            latentStyles1.Append(latentStyleExceptionInfo93);
            latentStyles1.Append(latentStyleExceptionInfo94);
            latentStyles1.Append(latentStyleExceptionInfo95);
            latentStyles1.Append(latentStyleExceptionInfo96);
            latentStyles1.Append(latentStyleExceptionInfo97);
            latentStyles1.Append(latentStyleExceptionInfo98);
            latentStyles1.Append(latentStyleExceptionInfo99);
            latentStyles1.Append(latentStyleExceptionInfo100);
            latentStyles1.Append(latentStyleExceptionInfo101);
            latentStyles1.Append(latentStyleExceptionInfo102);
            latentStyles1.Append(latentStyleExceptionInfo103);
            latentStyles1.Append(latentStyleExceptionInfo104);
            latentStyles1.Append(latentStyleExceptionInfo105);
            latentStyles1.Append(latentStyleExceptionInfo106);
            latentStyles1.Append(latentStyleExceptionInfo107);
            latentStyles1.Append(latentStyleExceptionInfo108);
            latentStyles1.Append(latentStyleExceptionInfo109);
            latentStyles1.Append(latentStyleExceptionInfo110);
            latentStyles1.Append(latentStyleExceptionInfo111);
            latentStyles1.Append(latentStyleExceptionInfo112);
            latentStyles1.Append(latentStyleExceptionInfo113);
            latentStyles1.Append(latentStyleExceptionInfo114);
            latentStyles1.Append(latentStyleExceptionInfo115);
            latentStyles1.Append(latentStyleExceptionInfo116);
            latentStyles1.Append(latentStyleExceptionInfo117);
            latentStyles1.Append(latentStyleExceptionInfo118);
            latentStyles1.Append(latentStyleExceptionInfo119);
            latentStyles1.Append(latentStyleExceptionInfo120);
            latentStyles1.Append(latentStyleExceptionInfo121);
            latentStyles1.Append(latentStyleExceptionInfo122);
            latentStyles1.Append(latentStyleExceptionInfo123);
            latentStyles1.Append(latentStyleExceptionInfo124);
            latentStyles1.Append(latentStyleExceptionInfo125);
            latentStyles1.Append(latentStyleExceptionInfo126);
            latentStyles1.Append(latentStyleExceptionInfo127);
            latentStyles1.Append(latentStyleExceptionInfo128);
            latentStyles1.Append(latentStyleExceptionInfo129);
            latentStyles1.Append(latentStyleExceptionInfo130);
            latentStyles1.Append(latentStyleExceptionInfo131);
            latentStyles1.Append(latentStyleExceptionInfo132);
            latentStyles1.Append(latentStyleExceptionInfo133);
            latentStyles1.Append(latentStyleExceptionInfo134);
            latentStyles1.Append(latentStyleExceptionInfo135);
            latentStyles1.Append(latentStyleExceptionInfo136);
            latentStyles1.Append(latentStyleExceptionInfo137);
            latentStyles1.Append(latentStyleExceptionInfo138);
            latentStyles1.Append(latentStyleExceptionInfo139);
            latentStyles1.Append(latentStyleExceptionInfo140);
            latentStyles1.Append(latentStyleExceptionInfo141);
            latentStyles1.Append(latentStyleExceptionInfo142);
            latentStyles1.Append(latentStyleExceptionInfo143);
            latentStyles1.Append(latentStyleExceptionInfo144);
            latentStyles1.Append(latentStyleExceptionInfo145);
            latentStyles1.Append(latentStyleExceptionInfo146);
            latentStyles1.Append(latentStyleExceptionInfo147);
            latentStyles1.Append(latentStyleExceptionInfo148);
            latentStyles1.Append(latentStyleExceptionInfo149);
            latentStyles1.Append(latentStyleExceptionInfo150);
            latentStyles1.Append(latentStyleExceptionInfo151);
            latentStyles1.Append(latentStyleExceptionInfo152);
            latentStyles1.Append(latentStyleExceptionInfo153);
            latentStyles1.Append(latentStyleExceptionInfo154);
            latentStyles1.Append(latentStyleExceptionInfo155);
            latentStyles1.Append(latentStyleExceptionInfo156);
            latentStyles1.Append(latentStyleExceptionInfo157);
            latentStyles1.Append(latentStyleExceptionInfo158);
            latentStyles1.Append(latentStyleExceptionInfo159);
            latentStyles1.Append(latentStyleExceptionInfo160);
            latentStyles1.Append(latentStyleExceptionInfo161);
            latentStyles1.Append(latentStyleExceptionInfo162);
            latentStyles1.Append(latentStyleExceptionInfo163);
            latentStyles1.Append(latentStyleExceptionInfo164);
            latentStyles1.Append(latentStyleExceptionInfo165);
            latentStyles1.Append(latentStyleExceptionInfo166);
            latentStyles1.Append(latentStyleExceptionInfo167);
            latentStyles1.Append(latentStyleExceptionInfo168);
            latentStyles1.Append(latentStyleExceptionInfo169);
            latentStyles1.Append(latentStyleExceptionInfo170);
            latentStyles1.Append(latentStyleExceptionInfo171);
            latentStyles1.Append(latentStyleExceptionInfo172);
            latentStyles1.Append(latentStyleExceptionInfo173);
            latentStyles1.Append(latentStyleExceptionInfo174);
            latentStyles1.Append(latentStyleExceptionInfo175);
            latentStyles1.Append(latentStyleExceptionInfo176);
            latentStyles1.Append(latentStyleExceptionInfo177);
            latentStyles1.Append(latentStyleExceptionInfo178);
            latentStyles1.Append(latentStyleExceptionInfo179);
            latentStyles1.Append(latentStyleExceptionInfo180);
            latentStyles1.Append(latentStyleExceptionInfo181);
            latentStyles1.Append(latentStyleExceptionInfo182);
            latentStyles1.Append(latentStyleExceptionInfo183);
            latentStyles1.Append(latentStyleExceptionInfo184);
            latentStyles1.Append(latentStyleExceptionInfo185);
            latentStyles1.Append(latentStyleExceptionInfo186);
            latentStyles1.Append(latentStyleExceptionInfo187);
            latentStyles1.Append(latentStyleExceptionInfo188);
            latentStyles1.Append(latentStyleExceptionInfo189);
            latentStyles1.Append(latentStyleExceptionInfo190);
            latentStyles1.Append(latentStyleExceptionInfo191);
            latentStyles1.Append(latentStyleExceptionInfo192);
            latentStyles1.Append(latentStyleExceptionInfo193);
            latentStyles1.Append(latentStyleExceptionInfo194);
            latentStyles1.Append(latentStyleExceptionInfo195);
            latentStyles1.Append(latentStyleExceptionInfo196);
            latentStyles1.Append(latentStyleExceptionInfo197);
            latentStyles1.Append(latentStyleExceptionInfo198);
            latentStyles1.Append(latentStyleExceptionInfo199);
            latentStyles1.Append(latentStyleExceptionInfo200);
            latentStyles1.Append(latentStyleExceptionInfo201);
            latentStyles1.Append(latentStyleExceptionInfo202);
            latentStyles1.Append(latentStyleExceptionInfo203);
            latentStyles1.Append(latentStyleExceptionInfo204);
            latentStyles1.Append(latentStyleExceptionInfo205);
            latentStyles1.Append(latentStyleExceptionInfo206);
            latentStyles1.Append(latentStyleExceptionInfo207);
            latentStyles1.Append(latentStyleExceptionInfo208);
            latentStyles1.Append(latentStyleExceptionInfo209);
            latentStyles1.Append(latentStyleExceptionInfo210);
            latentStyles1.Append(latentStyleExceptionInfo211);
            latentStyles1.Append(latentStyleExceptionInfo212);
            latentStyles1.Append(latentStyleExceptionInfo213);
            latentStyles1.Append(latentStyleExceptionInfo214);
            latentStyles1.Append(latentStyleExceptionInfo215);
            latentStyles1.Append(latentStyleExceptionInfo216);
            latentStyles1.Append(latentStyleExceptionInfo217);
            latentStyles1.Append(latentStyleExceptionInfo218);
            latentStyles1.Append(latentStyleExceptionInfo219);
            latentStyles1.Append(latentStyleExceptionInfo220);
            latentStyles1.Append(latentStyleExceptionInfo221);
            latentStyles1.Append(latentStyleExceptionInfo222);
            latentStyles1.Append(latentStyleExceptionInfo223);
            latentStyles1.Append(latentStyleExceptionInfo224);
            latentStyles1.Append(latentStyleExceptionInfo225);
            latentStyles1.Append(latentStyleExceptionInfo226);
            latentStyles1.Append(latentStyleExceptionInfo227);
            latentStyles1.Append(latentStyleExceptionInfo228);
            latentStyles1.Append(latentStyleExceptionInfo229);
            latentStyles1.Append(latentStyleExceptionInfo230);
            latentStyles1.Append(latentStyleExceptionInfo231);
            latentStyles1.Append(latentStyleExceptionInfo232);
            latentStyles1.Append(latentStyleExceptionInfo233);
            latentStyles1.Append(latentStyleExceptionInfo234);
            latentStyles1.Append(latentStyleExceptionInfo235);
            latentStyles1.Append(latentStyleExceptionInfo236);
            latentStyles1.Append(latentStyleExceptionInfo237);
            latentStyles1.Append(latentStyleExceptionInfo238);
            latentStyles1.Append(latentStyleExceptionInfo239);
            latentStyles1.Append(latentStyleExceptionInfo240);
            latentStyles1.Append(latentStyleExceptionInfo241);
            latentStyles1.Append(latentStyleExceptionInfo242);
            latentStyles1.Append(latentStyleExceptionInfo243);
            latentStyles1.Append(latentStyleExceptionInfo244);
            latentStyles1.Append(latentStyleExceptionInfo245);
            latentStyles1.Append(latentStyleExceptionInfo246);
            latentStyles1.Append(latentStyleExceptionInfo247);
            latentStyles1.Append(latentStyleExceptionInfo248);
            latentStyles1.Append(latentStyleExceptionInfo249);
            latentStyles1.Append(latentStyleExceptionInfo250);
            latentStyles1.Append(latentStyleExceptionInfo251);
            latentStyles1.Append(latentStyleExceptionInfo252);
            latentStyles1.Append(latentStyleExceptionInfo253);
            latentStyles1.Append(latentStyleExceptionInfo254);
            latentStyles1.Append(latentStyleExceptionInfo255);
            latentStyles1.Append(latentStyleExceptionInfo256);
            latentStyles1.Append(latentStyleExceptionInfo257);
            latentStyles1.Append(latentStyleExceptionInfo258);
            latentStyles1.Append(latentStyleExceptionInfo259);
            latentStyles1.Append(latentStyleExceptionInfo260);
            latentStyles1.Append(latentStyleExceptionInfo261);
            latentStyles1.Append(latentStyleExceptionInfo262);
            latentStyles1.Append(latentStyleExceptionInfo263);
            latentStyles1.Append(latentStyleExceptionInfo264);
            latentStyles1.Append(latentStyleExceptionInfo265);
            latentStyles1.Append(latentStyleExceptionInfo266);
            latentStyles1.Append(latentStyleExceptionInfo267);
            latentStyles1.Append(latentStyleExceptionInfo268);
            latentStyles1.Append(latentStyleExceptionInfo269);
            latentStyles1.Append(latentStyleExceptionInfo270);
            latentStyles1.Append(latentStyleExceptionInfo271);
            latentStyles1.Append(latentStyleExceptionInfo272);
            latentStyles1.Append(latentStyleExceptionInfo273);
            latentStyles1.Append(latentStyleExceptionInfo274);
            latentStyles1.Append(latentStyleExceptionInfo275);
            latentStyles1.Append(latentStyleExceptionInfo276);
            latentStyles1.Append(latentStyleExceptionInfo277);
            latentStyles1.Append(latentStyleExceptionInfo278);
            latentStyles1.Append(latentStyleExceptionInfo279);
            latentStyles1.Append(latentStyleExceptionInfo280);
            latentStyles1.Append(latentStyleExceptionInfo281);
            latentStyles1.Append(latentStyleExceptionInfo282);
            latentStyles1.Append(latentStyleExceptionInfo283);
            latentStyles1.Append(latentStyleExceptionInfo284);
            latentStyles1.Append(latentStyleExceptionInfo285);
            latentStyles1.Append(latentStyleExceptionInfo286);
            latentStyles1.Append(latentStyleExceptionInfo287);
            latentStyles1.Append(latentStyleExceptionInfo288);
            latentStyles1.Append(latentStyleExceptionInfo289);
            latentStyles1.Append(latentStyleExceptionInfo290);
            latentStyles1.Append(latentStyleExceptionInfo291);
            latentStyles1.Append(latentStyleExceptionInfo292);
            latentStyles1.Append(latentStyleExceptionInfo293);
            latentStyles1.Append(latentStyleExceptionInfo294);
            latentStyles1.Append(latentStyleExceptionInfo295);
            latentStyles1.Append(latentStyleExceptionInfo296);
            latentStyles1.Append(latentStyleExceptionInfo297);
            latentStyles1.Append(latentStyleExceptionInfo298);
            latentStyles1.Append(latentStyleExceptionInfo299);
            latentStyles1.Append(latentStyleExceptionInfo300);
            latentStyles1.Append(latentStyleExceptionInfo301);
            latentStyles1.Append(latentStyleExceptionInfo302);
            latentStyles1.Append(latentStyleExceptionInfo303);
            latentStyles1.Append(latentStyleExceptionInfo304);
            latentStyles1.Append(latentStyleExceptionInfo305);
            latentStyles1.Append(latentStyleExceptionInfo306);
            latentStyles1.Append(latentStyleExceptionInfo307);
            latentStyles1.Append(latentStyleExceptionInfo308);
            latentStyles1.Append(latentStyleExceptionInfo309);
            latentStyles1.Append(latentStyleExceptionInfo310);
            latentStyles1.Append(latentStyleExceptionInfo311);
            latentStyles1.Append(latentStyleExceptionInfo312);
            latentStyles1.Append(latentStyleExceptionInfo313);
            latentStyles1.Append(latentStyleExceptionInfo314);
            latentStyles1.Append(latentStyleExceptionInfo315);
            latentStyles1.Append(latentStyleExceptionInfo316);
            latentStyles1.Append(latentStyleExceptionInfo317);
            latentStyles1.Append(latentStyleExceptionInfo318);
            latentStyles1.Append(latentStyleExceptionInfo319);
            latentStyles1.Append(latentStyleExceptionInfo320);
            latentStyles1.Append(latentStyleExceptionInfo321);
            latentStyles1.Append(latentStyleExceptionInfo322);
            latentStyles1.Append(latentStyleExceptionInfo323);
            latentStyles1.Append(latentStyleExceptionInfo324);
            latentStyles1.Append(latentStyleExceptionInfo325);
            latentStyles1.Append(latentStyleExceptionInfo326);
            latentStyles1.Append(latentStyleExceptionInfo327);
            latentStyles1.Append(latentStyleExceptionInfo328);
            latentStyles1.Append(latentStyleExceptionInfo329);
            latentStyles1.Append(latentStyleExceptionInfo330);
            latentStyles1.Append(latentStyleExceptionInfo331);
            latentStyles1.Append(latentStyleExceptionInfo332);
            latentStyles1.Append(latentStyleExceptionInfo333);
            latentStyles1.Append(latentStyleExceptionInfo334);
            latentStyles1.Append(latentStyleExceptionInfo335);
            latentStyles1.Append(latentStyleExceptionInfo336);
            latentStyles1.Append(latentStyleExceptionInfo337);
            latentStyles1.Append(latentStyleExceptionInfo338);
            latentStyles1.Append(latentStyleExceptionInfo339);
            latentStyles1.Append(latentStyleExceptionInfo340);
            latentStyles1.Append(latentStyleExceptionInfo341);
            latentStyles1.Append(latentStyleExceptionInfo342);
            latentStyles1.Append(latentStyleExceptionInfo343);
            latentStyles1.Append(latentStyleExceptionInfo344);
            latentStyles1.Append(latentStyleExceptionInfo345);
            latentStyles1.Append(latentStyleExceptionInfo346);
            latentStyles1.Append(latentStyleExceptionInfo347);
            latentStyles1.Append(latentStyleExceptionInfo348);
            latentStyles1.Append(latentStyleExceptionInfo349);
            latentStyles1.Append(latentStyleExceptionInfo350);
            latentStyles1.Append(latentStyleExceptionInfo351);
            latentStyles1.Append(latentStyleExceptionInfo352);
            latentStyles1.Append(latentStyleExceptionInfo353);
            latentStyles1.Append(latentStyleExceptionInfo354);
            latentStyles1.Append(latentStyleExceptionInfo355);
            latentStyles1.Append(latentStyleExceptionInfo356);
            latentStyles1.Append(latentStyleExceptionInfo357);
            latentStyles1.Append(latentStyleExceptionInfo358);
            latentStyles1.Append(latentStyleExceptionInfo359);
            latentStyles1.Append(latentStyleExceptionInfo360);
            latentStyles1.Append(latentStyleExceptionInfo361);
            latentStyles1.Append(latentStyleExceptionInfo362);
            latentStyles1.Append(latentStyleExceptionInfo363);
            latentStyles1.Append(latentStyleExceptionInfo364);
            latentStyles1.Append(latentStyleExceptionInfo365);
            latentStyles1.Append(latentStyleExceptionInfo366);
            latentStyles1.Append(latentStyleExceptionInfo367);
            latentStyles1.Append(latentStyleExceptionInfo368);
            latentStyles1.Append(latentStyleExceptionInfo369);
            latentStyles1.Append(latentStyleExceptionInfo370);
            latentStyles1.Append(latentStyleExceptionInfo371);
            latentStyles1.Append(latentStyleExceptionInfo372);
            latentStyles1.Append(latentStyleExceptionInfo373);
            latentStyles1.Append(latentStyleExceptionInfo374);
            latentStyles1.Append(latentStyleExceptionInfo375);
            latentStyles1.Append(latentStyleExceptionInfo376);

            Style style1 = new Style(){ Type = StyleValues.Paragraph, StyleId = "Normal", Default = true };
            StyleName styleName1 = new StyleName(){ Val = "Normal" };
            PrimaryStyle primaryStyle1 = new PrimaryStyle();

            style1.Append(styleName1);
            style1.Append(primaryStyle1);

            Style style2 = new Style(){ Type = StyleValues.Paragraph, StyleId = "Heading1" };
            StyleName styleName2 = new StyleName(){ Val = "heading 1" };
            BasedOn basedOn1 = new BasedOn(){ Val = "Normal" };
            NextParagraphStyle nextParagraphStyle1 = new NextParagraphStyle(){ Val = "Normal" };
            LinkedStyle linkedStyle1 = new LinkedStyle(){ Val = "Heading1Char" };
            UIPriority uIPriority1 = new UIPriority(){ Val = 9 };
            PrimaryStyle primaryStyle2 = new PrimaryStyle();
            Rsid rsid5 = new Rsid(){ Val = "00440271" };

            StyleParagraphProperties styleParagraphProperties1 = new StyleParagraphProperties();
            KeepNext keepNext1 = new KeepNext();
            KeepLines keepLines1 = new KeepLines();
            SpacingBetweenLines spacingBetweenLines2 = new SpacingBetweenLines(){ Before = "360", After = "80" };
            OutlineLevel outlineLevel1 = new OutlineLevel(){ Val = 0 };

            styleParagraphProperties1.Append(keepNext1);
            styleParagraphProperties1.Append(keepLines1);
            styleParagraphProperties1.Append(spacingBetweenLines2);
            styleParagraphProperties1.Append(outlineLevel1);

            StyleRunProperties styleRunProperties1 = new StyleRunProperties();
            RunFonts runFonts2 = new RunFonts(){ AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
            Color color1 = new Color(){ Val = "0F4761", ThemeColor = ThemeColorValues.Accent1, ThemeShade = "BF" };
            FontSize fontSize2 = new FontSize(){ Val = "40" };
            FontSizeComplexScript fontSizeComplexScript2 = new FontSizeComplexScript(){ Val = "40" };

            styleRunProperties1.Append(runFonts2);
            styleRunProperties1.Append(color1);
            styleRunProperties1.Append(fontSize2);
            styleRunProperties1.Append(fontSizeComplexScript2);

            style2.Append(styleName2);
            style2.Append(basedOn1);
            style2.Append(nextParagraphStyle1);
            style2.Append(linkedStyle1);
            style2.Append(uIPriority1);
            style2.Append(primaryStyle2);
            style2.Append(rsid5);
            style2.Append(styleParagraphProperties1);
            style2.Append(styleRunProperties1);

            Style style3 = new Style(){ Type = StyleValues.Paragraph, StyleId = "Heading2" };
            StyleName styleName3 = new StyleName(){ Val = "heading 2" };
            BasedOn basedOn2 = new BasedOn(){ Val = "Normal" };
            NextParagraphStyle nextParagraphStyle2 = new NextParagraphStyle(){ Val = "Normal" };
            LinkedStyle linkedStyle2 = new LinkedStyle(){ Val = "Heading2Char" };
            UIPriority uIPriority2 = new UIPriority(){ Val = 9 };
            SemiHidden semiHidden1 = new SemiHidden();
            UnhideWhenUsed unhideWhenUsed1 = new UnhideWhenUsed();
            PrimaryStyle primaryStyle3 = new PrimaryStyle();
            Rsid rsid6 = new Rsid(){ Val = "00440271" };

            StyleParagraphProperties styleParagraphProperties2 = new StyleParagraphProperties();
            KeepNext keepNext2 = new KeepNext();
            KeepLines keepLines2 = new KeepLines();
            SpacingBetweenLines spacingBetweenLines3 = new SpacingBetweenLines(){ Before = "160", After = "80" };
            OutlineLevel outlineLevel2 = new OutlineLevel(){ Val = 1 };

            styleParagraphProperties2.Append(keepNext2);
            styleParagraphProperties2.Append(keepLines2);
            styleParagraphProperties2.Append(spacingBetweenLines3);
            styleParagraphProperties2.Append(outlineLevel2);

            StyleRunProperties styleRunProperties2 = new StyleRunProperties();
            RunFonts runFonts3 = new RunFonts(){ AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
            Color color2 = new Color(){ Val = "0F4761", ThemeColor = ThemeColorValues.Accent1, ThemeShade = "BF" };
            FontSize fontSize3 = new FontSize(){ Val = "32" };
            FontSizeComplexScript fontSizeComplexScript3 = new FontSizeComplexScript(){ Val = "32" };

            styleRunProperties2.Append(runFonts3);
            styleRunProperties2.Append(color2);
            styleRunProperties2.Append(fontSize3);
            styleRunProperties2.Append(fontSizeComplexScript3);

            style3.Append(styleName3);
            style3.Append(basedOn2);
            style3.Append(nextParagraphStyle2);
            style3.Append(linkedStyle2);
            style3.Append(uIPriority2);
            style3.Append(semiHidden1);
            style3.Append(unhideWhenUsed1);
            style3.Append(primaryStyle3);
            style3.Append(rsid6);
            style3.Append(styleParagraphProperties2);
            style3.Append(styleRunProperties2);

            Style style4 = new Style(){ Type = StyleValues.Paragraph, StyleId = "Heading3" };
            StyleName styleName4 = new StyleName(){ Val = "heading 3" };
            BasedOn basedOn3 = new BasedOn(){ Val = "Normal" };
            NextParagraphStyle nextParagraphStyle3 = new NextParagraphStyle(){ Val = "Normal" };
            LinkedStyle linkedStyle3 = new LinkedStyle(){ Val = "Heading3Char" };
            UIPriority uIPriority3 = new UIPriority(){ Val = 9 };
            SemiHidden semiHidden2 = new SemiHidden();
            UnhideWhenUsed unhideWhenUsed2 = new UnhideWhenUsed();
            PrimaryStyle primaryStyle4 = new PrimaryStyle();
            Rsid rsid7 = new Rsid(){ Val = "00440271" };

            StyleParagraphProperties styleParagraphProperties3 = new StyleParagraphProperties();
            KeepNext keepNext3 = new KeepNext();
            KeepLines keepLines3 = new KeepLines();
            SpacingBetweenLines spacingBetweenLines4 = new SpacingBetweenLines(){ Before = "160", After = "80" };
            OutlineLevel outlineLevel3 = new OutlineLevel(){ Val = 2 };

            styleParagraphProperties3.Append(keepNext3);
            styleParagraphProperties3.Append(keepLines3);
            styleParagraphProperties3.Append(spacingBetweenLines4);
            styleParagraphProperties3.Append(outlineLevel3);

            StyleRunProperties styleRunProperties3 = new StyleRunProperties();
            RunFonts runFonts4 = new RunFonts(){ EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
            Color color3 = new Color(){ Val = "0F4761", ThemeColor = ThemeColorValues.Accent1, ThemeShade = "BF" };
            FontSize fontSize4 = new FontSize(){ Val = "28" };
            FontSizeComplexScript fontSizeComplexScript4 = new FontSizeComplexScript(){ Val = "28" };

            styleRunProperties3.Append(runFonts4);
            styleRunProperties3.Append(color3);
            styleRunProperties3.Append(fontSize4);
            styleRunProperties3.Append(fontSizeComplexScript4);

            style4.Append(styleName4);
            style4.Append(basedOn3);
            style4.Append(nextParagraphStyle3);
            style4.Append(linkedStyle3);
            style4.Append(uIPriority3);
            style4.Append(semiHidden2);
            style4.Append(unhideWhenUsed2);
            style4.Append(primaryStyle4);
            style4.Append(rsid7);
            style4.Append(styleParagraphProperties3);
            style4.Append(styleRunProperties3);

            Style style5 = new Style(){ Type = StyleValues.Paragraph, StyleId = "Heading4" };
            StyleName styleName5 = new StyleName(){ Val = "heading 4" };
            BasedOn basedOn4 = new BasedOn(){ Val = "Normal" };
            NextParagraphStyle nextParagraphStyle4 = new NextParagraphStyle(){ Val = "Normal" };
            LinkedStyle linkedStyle4 = new LinkedStyle(){ Val = "Heading4Char" };
            UIPriority uIPriority4 = new UIPriority(){ Val = 9 };
            SemiHidden semiHidden3 = new SemiHidden();
            UnhideWhenUsed unhideWhenUsed3 = new UnhideWhenUsed();
            PrimaryStyle primaryStyle5 = new PrimaryStyle();
            Rsid rsid8 = new Rsid(){ Val = "00440271" };

            StyleParagraphProperties styleParagraphProperties4 = new StyleParagraphProperties();
            KeepNext keepNext4 = new KeepNext();
            KeepLines keepLines4 = new KeepLines();
            SpacingBetweenLines spacingBetweenLines5 = new SpacingBetweenLines(){ Before = "80", After = "40" };
            OutlineLevel outlineLevel4 = new OutlineLevel(){ Val = 3 };

            styleParagraphProperties4.Append(keepNext4);
            styleParagraphProperties4.Append(keepLines4);
            styleParagraphProperties4.Append(spacingBetweenLines5);
            styleParagraphProperties4.Append(outlineLevel4);

            StyleRunProperties styleRunProperties4 = new StyleRunProperties();
            RunFonts runFonts5 = new RunFonts(){ EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
            Italic italic1 = new Italic();
            ItalicComplexScript italicComplexScript1 = new ItalicComplexScript();
            Color color4 = new Color(){ Val = "0F4761", ThemeColor = ThemeColorValues.Accent1, ThemeShade = "BF" };

            styleRunProperties4.Append(runFonts5);
            styleRunProperties4.Append(italic1);
            styleRunProperties4.Append(italicComplexScript1);
            styleRunProperties4.Append(color4);

            style5.Append(styleName5);
            style5.Append(basedOn4);
            style5.Append(nextParagraphStyle4);
            style5.Append(linkedStyle4);
            style5.Append(uIPriority4);
            style5.Append(semiHidden3);
            style5.Append(unhideWhenUsed3);
            style5.Append(primaryStyle5);
            style5.Append(rsid8);
            style5.Append(styleParagraphProperties4);
            style5.Append(styleRunProperties4);

            Style style6 = new Style(){ Type = StyleValues.Paragraph, StyleId = "Heading5" };
            StyleName styleName6 = new StyleName(){ Val = "heading 5" };
            BasedOn basedOn5 = new BasedOn(){ Val = "Normal" };
            NextParagraphStyle nextParagraphStyle5 = new NextParagraphStyle(){ Val = "Normal" };
            LinkedStyle linkedStyle5 = new LinkedStyle(){ Val = "Heading5Char" };
            UIPriority uIPriority5 = new UIPriority(){ Val = 9 };
            SemiHidden semiHidden4 = new SemiHidden();
            UnhideWhenUsed unhideWhenUsed4 = new UnhideWhenUsed();
            PrimaryStyle primaryStyle6 = new PrimaryStyle();
            Rsid rsid9 = new Rsid(){ Val = "00440271" };

            StyleParagraphProperties styleParagraphProperties5 = new StyleParagraphProperties();
            KeepNext keepNext5 = new KeepNext();
            KeepLines keepLines5 = new KeepLines();
            SpacingBetweenLines spacingBetweenLines6 = new SpacingBetweenLines(){ Before = "80", After = "40" };
            OutlineLevel outlineLevel5 = new OutlineLevel(){ Val = 4 };

            styleParagraphProperties5.Append(keepNext5);
            styleParagraphProperties5.Append(keepLines5);
            styleParagraphProperties5.Append(spacingBetweenLines6);
            styleParagraphProperties5.Append(outlineLevel5);

            StyleRunProperties styleRunProperties5 = new StyleRunProperties();
            RunFonts runFonts6 = new RunFonts(){ EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
            Color color5 = new Color(){ Val = "0F4761", ThemeColor = ThemeColorValues.Accent1, ThemeShade = "BF" };

            styleRunProperties5.Append(runFonts6);
            styleRunProperties5.Append(color5);

            style6.Append(styleName6);
            style6.Append(basedOn5);
            style6.Append(nextParagraphStyle5);
            style6.Append(linkedStyle5);
            style6.Append(uIPriority5);
            style6.Append(semiHidden4);
            style6.Append(unhideWhenUsed4);
            style6.Append(primaryStyle6);
            style6.Append(rsid9);
            style6.Append(styleParagraphProperties5);
            style6.Append(styleRunProperties5);

            Style style7 = new Style(){ Type = StyleValues.Paragraph, StyleId = "Heading6" };
            StyleName styleName7 = new StyleName(){ Val = "heading 6" };
            BasedOn basedOn6 = new BasedOn(){ Val = "Normal" };
            NextParagraphStyle nextParagraphStyle6 = new NextParagraphStyle(){ Val = "Normal" };
            LinkedStyle linkedStyle6 = new LinkedStyle(){ Val = "Heading6Char" };
            UIPriority uIPriority6 = new UIPriority(){ Val = 9 };
            SemiHidden semiHidden5 = new SemiHidden();
            UnhideWhenUsed unhideWhenUsed5 = new UnhideWhenUsed();
            PrimaryStyle primaryStyle7 = new PrimaryStyle();
            Rsid rsid10 = new Rsid(){ Val = "00440271" };

            StyleParagraphProperties styleParagraphProperties6 = new StyleParagraphProperties();
            KeepNext keepNext6 = new KeepNext();
            KeepLines keepLines6 = new KeepLines();
            SpacingBetweenLines spacingBetweenLines7 = new SpacingBetweenLines(){ Before = "40", After = "0" };
            OutlineLevel outlineLevel6 = new OutlineLevel(){ Val = 5 };

            styleParagraphProperties6.Append(keepNext6);
            styleParagraphProperties6.Append(keepLines6);
            styleParagraphProperties6.Append(spacingBetweenLines7);
            styleParagraphProperties6.Append(outlineLevel6);

            StyleRunProperties styleRunProperties6 = new StyleRunProperties();
            RunFonts runFonts7 = new RunFonts(){ EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
            Italic italic2 = new Italic();
            ItalicComplexScript italicComplexScript2 = new ItalicComplexScript();
            Color color6 = new Color(){ Val = "595959", ThemeColor = ThemeColorValues.Text1, ThemeTint = "A6" };

            styleRunProperties6.Append(runFonts7);
            styleRunProperties6.Append(italic2);
            styleRunProperties6.Append(italicComplexScript2);
            styleRunProperties6.Append(color6);

            style7.Append(styleName7);
            style7.Append(basedOn6);
            style7.Append(nextParagraphStyle6);
            style7.Append(linkedStyle6);
            style7.Append(uIPriority6);
            style7.Append(semiHidden5);
            style7.Append(unhideWhenUsed5);
            style7.Append(primaryStyle7);
            style7.Append(rsid10);
            style7.Append(styleParagraphProperties6);
            style7.Append(styleRunProperties6);

            Style style8 = new Style(){ Type = StyleValues.Paragraph, StyleId = "Heading7" };
            StyleName styleName8 = new StyleName(){ Val = "heading 7" };
            BasedOn basedOn7 = new BasedOn(){ Val = "Normal" };
            NextParagraphStyle nextParagraphStyle7 = new NextParagraphStyle(){ Val = "Normal" };
            LinkedStyle linkedStyle7 = new LinkedStyle(){ Val = "Heading7Char" };
            UIPriority uIPriority7 = new UIPriority(){ Val = 9 };
            SemiHidden semiHidden6 = new SemiHidden();
            UnhideWhenUsed unhideWhenUsed6 = new UnhideWhenUsed();
            PrimaryStyle primaryStyle8 = new PrimaryStyle();
            Rsid rsid11 = new Rsid(){ Val = "00440271" };

            StyleParagraphProperties styleParagraphProperties7 = new StyleParagraphProperties();
            KeepNext keepNext7 = new KeepNext();
            KeepLines keepLines7 = new KeepLines();
            SpacingBetweenLines spacingBetweenLines8 = new SpacingBetweenLines(){ Before = "40", After = "0" };
            OutlineLevel outlineLevel7 = new OutlineLevel(){ Val = 6 };

            styleParagraphProperties7.Append(keepNext7);
            styleParagraphProperties7.Append(keepLines7);
            styleParagraphProperties7.Append(spacingBetweenLines8);
            styleParagraphProperties7.Append(outlineLevel7);

            StyleRunProperties styleRunProperties7 = new StyleRunProperties();
            RunFonts runFonts8 = new RunFonts(){ EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
            Color color7 = new Color(){ Val = "595959", ThemeColor = ThemeColorValues.Text1, ThemeTint = "A6" };

            styleRunProperties7.Append(runFonts8);
            styleRunProperties7.Append(color7);

            style8.Append(styleName8);
            style8.Append(basedOn7);
            style8.Append(nextParagraphStyle7);
            style8.Append(linkedStyle7);
            style8.Append(uIPriority7);
            style8.Append(semiHidden6);
            style8.Append(unhideWhenUsed6);
            style8.Append(primaryStyle8);
            style8.Append(rsid11);
            style8.Append(styleParagraphProperties7);
            style8.Append(styleRunProperties7);

            Style style9 = new Style(){ Type = StyleValues.Paragraph, StyleId = "Heading8" };
            StyleName styleName9 = new StyleName(){ Val = "heading 8" };
            BasedOn basedOn8 = new BasedOn(){ Val = "Normal" };
            NextParagraphStyle nextParagraphStyle8 = new NextParagraphStyle(){ Val = "Normal" };
            LinkedStyle linkedStyle8 = new LinkedStyle(){ Val = "Heading8Char" };
            UIPriority uIPriority8 = new UIPriority(){ Val = 9 };
            SemiHidden semiHidden7 = new SemiHidden();
            UnhideWhenUsed unhideWhenUsed7 = new UnhideWhenUsed();
            PrimaryStyle primaryStyle9 = new PrimaryStyle();
            Rsid rsid12 = new Rsid(){ Val = "00440271" };

            StyleParagraphProperties styleParagraphProperties8 = new StyleParagraphProperties();
            KeepNext keepNext8 = new KeepNext();
            KeepLines keepLines8 = new KeepLines();
            SpacingBetweenLines spacingBetweenLines9 = new SpacingBetweenLines(){ After = "0" };
            OutlineLevel outlineLevel8 = new OutlineLevel(){ Val = 7 };

            styleParagraphProperties8.Append(keepNext8);
            styleParagraphProperties8.Append(keepLines8);
            styleParagraphProperties8.Append(spacingBetweenLines9);
            styleParagraphProperties8.Append(outlineLevel8);

            StyleRunProperties styleRunProperties8 = new StyleRunProperties();
            RunFonts runFonts9 = new RunFonts(){ EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
            Italic italic3 = new Italic();
            ItalicComplexScript italicComplexScript3 = new ItalicComplexScript();
            Color color8 = new Color(){ Val = "272727", ThemeColor = ThemeColorValues.Text1, ThemeTint = "D8" };

            styleRunProperties8.Append(runFonts9);
            styleRunProperties8.Append(italic3);
            styleRunProperties8.Append(italicComplexScript3);
            styleRunProperties8.Append(color8);

            style9.Append(styleName9);
            style9.Append(basedOn8);
            style9.Append(nextParagraphStyle8);
            style9.Append(linkedStyle8);
            style9.Append(uIPriority8);
            style9.Append(semiHidden7);
            style9.Append(unhideWhenUsed7);
            style9.Append(primaryStyle9);
            style9.Append(rsid12);
            style9.Append(styleParagraphProperties8);
            style9.Append(styleRunProperties8);

            Style style10 = new Style(){ Type = StyleValues.Paragraph, StyleId = "Heading9" };
            StyleName styleName10 = new StyleName(){ Val = "heading 9" };
            BasedOn basedOn9 = new BasedOn(){ Val = "Normal" };
            NextParagraphStyle nextParagraphStyle9 = new NextParagraphStyle(){ Val = "Normal" };
            LinkedStyle linkedStyle9 = new LinkedStyle(){ Val = "Heading9Char" };
            UIPriority uIPriority9 = new UIPriority(){ Val = 9 };
            SemiHidden semiHidden8 = new SemiHidden();
            UnhideWhenUsed unhideWhenUsed8 = new UnhideWhenUsed();
            PrimaryStyle primaryStyle10 = new PrimaryStyle();
            Rsid rsid13 = new Rsid(){ Val = "00440271" };

            StyleParagraphProperties styleParagraphProperties9 = new StyleParagraphProperties();
            KeepNext keepNext9 = new KeepNext();
            KeepLines keepLines9 = new KeepLines();
            SpacingBetweenLines spacingBetweenLines10 = new SpacingBetweenLines(){ After = "0" };
            OutlineLevel outlineLevel9 = new OutlineLevel(){ Val = 8 };

            styleParagraphProperties9.Append(keepNext9);
            styleParagraphProperties9.Append(keepLines9);
            styleParagraphProperties9.Append(spacingBetweenLines10);
            styleParagraphProperties9.Append(outlineLevel9);

            StyleRunProperties styleRunProperties9 = new StyleRunProperties();
            RunFonts runFonts10 = new RunFonts(){ EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
            Color color9 = new Color(){ Val = "272727", ThemeColor = ThemeColorValues.Text1, ThemeTint = "D8" };

            styleRunProperties9.Append(runFonts10);
            styleRunProperties9.Append(color9);

            style10.Append(styleName10);
            style10.Append(basedOn9);
            style10.Append(nextParagraphStyle9);
            style10.Append(linkedStyle9);
            style10.Append(uIPriority9);
            style10.Append(semiHidden8);
            style10.Append(unhideWhenUsed8);
            style10.Append(primaryStyle10);
            style10.Append(rsid13);
            style10.Append(styleParagraphProperties9);
            style10.Append(styleRunProperties9);

            Style style11 = new Style(){ Type = StyleValues.Character, StyleId = "DefaultParagraphFont", Default = true };
            StyleName styleName11 = new StyleName(){ Val = "Default Paragraph Font" };
            UIPriority uIPriority10 = new UIPriority(){ Val = 1 };
            SemiHidden semiHidden9 = new SemiHidden();
            UnhideWhenUsed unhideWhenUsed9 = new UnhideWhenUsed();

            style11.Append(styleName11);
            style11.Append(uIPriority10);
            style11.Append(semiHidden9);
            style11.Append(unhideWhenUsed9);

            Style style12 = new Style(){ Type = StyleValues.Table, StyleId = "TableNormal", Default = true };
            StyleName styleName12 = new StyleName(){ Val = "Normal Table" };
            UIPriority uIPriority11 = new UIPriority(){ Val = 99 };
            SemiHidden semiHidden10 = new SemiHidden();
            UnhideWhenUsed unhideWhenUsed10 = new UnhideWhenUsed();

            StyleTableProperties styleTableProperties1 = new StyleTableProperties();
            TableIndentation tableIndentation1 = new TableIndentation(){ Width = 0, Type = TableWidthUnitValues.Dxa };

            TableCellMarginDefault tableCellMarginDefault1 = new TableCellMarginDefault();
            TopMargin topMargin1 = new TopMargin(){ Width = "0", Type = TableWidthUnitValues.Dxa };
            TableCellLeftMargin tableCellLeftMargin1 = new TableCellLeftMargin(){ Width = 108, Type = TableWidthValues.Dxa };
            BottomMargin bottomMargin1 = new BottomMargin(){ Width = "0", Type = TableWidthUnitValues.Dxa };
            TableCellRightMargin tableCellRightMargin1 = new TableCellRightMargin(){ Width = 108, Type = TableWidthValues.Dxa };

            tableCellMarginDefault1.Append(topMargin1);
            tableCellMarginDefault1.Append(tableCellLeftMargin1);
            tableCellMarginDefault1.Append(bottomMargin1);
            tableCellMarginDefault1.Append(tableCellRightMargin1);

            styleTableProperties1.Append(tableIndentation1);
            styleTableProperties1.Append(tableCellMarginDefault1);

            style12.Append(styleName12);
            style12.Append(uIPriority11);
            style12.Append(semiHidden10);
            style12.Append(unhideWhenUsed10);
            style12.Append(styleTableProperties1);

            Style style13 = new Style(){ Type = StyleValues.Numbering, StyleId = "NoList", Default = true };
            StyleName styleName13 = new StyleName(){ Val = "No List" };
            UIPriority uIPriority12 = new UIPriority(){ Val = 99 };
            SemiHidden semiHidden11 = new SemiHidden();
            UnhideWhenUsed unhideWhenUsed11 = new UnhideWhenUsed();

            style13.Append(styleName13);
            style13.Append(uIPriority12);
            style13.Append(semiHidden11);
            style13.Append(unhideWhenUsed11);

            Style style14 = new Style(){ Type = StyleValues.Character, StyleId = "Heading1Char", CustomStyle = true };
            StyleName styleName14 = new StyleName(){ Val = "Heading 1 Char" };
            BasedOn basedOn10 = new BasedOn(){ Val = "DefaultParagraphFont" };
            LinkedStyle linkedStyle10 = new LinkedStyle(){ Val = "Heading1" };
            UIPriority uIPriority13 = new UIPriority(){ Val = 9 };
            Rsid rsid14 = new Rsid(){ Val = "00440271" };

            StyleRunProperties styleRunProperties10 = new StyleRunProperties();
            RunFonts runFonts11 = new RunFonts(){ AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
            Color color10 = new Color(){ Val = "0F4761", ThemeColor = ThemeColorValues.Accent1, ThemeShade = "BF" };
            FontSize fontSize5 = new FontSize(){ Val = "40" };
            FontSizeComplexScript fontSizeComplexScript5 = new FontSizeComplexScript(){ Val = "40" };

            styleRunProperties10.Append(runFonts11);
            styleRunProperties10.Append(color10);
            styleRunProperties10.Append(fontSize5);
            styleRunProperties10.Append(fontSizeComplexScript5);

            style14.Append(styleName14);
            style14.Append(basedOn10);
            style14.Append(linkedStyle10);
            style14.Append(uIPriority13);
            style14.Append(rsid14);
            style14.Append(styleRunProperties10);

            Style style15 = new Style(){ Type = StyleValues.Character, StyleId = "Heading2Char", CustomStyle = true };
            StyleName styleName15 = new StyleName(){ Val = "Heading 2 Char" };
            BasedOn basedOn11 = new BasedOn(){ Val = "DefaultParagraphFont" };
            LinkedStyle linkedStyle11 = new LinkedStyle(){ Val = "Heading2" };
            UIPriority uIPriority14 = new UIPriority(){ Val = 9 };
            SemiHidden semiHidden12 = new SemiHidden();
            Rsid rsid15 = new Rsid(){ Val = "00440271" };

            StyleRunProperties styleRunProperties11 = new StyleRunProperties();
            RunFonts runFonts12 = new RunFonts(){ AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
            Color color11 = new Color(){ Val = "0F4761", ThemeColor = ThemeColorValues.Accent1, ThemeShade = "BF" };
            FontSize fontSize6 = new FontSize(){ Val = "32" };
            FontSizeComplexScript fontSizeComplexScript6 = new FontSizeComplexScript(){ Val = "32" };

            styleRunProperties11.Append(runFonts12);
            styleRunProperties11.Append(color11);
            styleRunProperties11.Append(fontSize6);
            styleRunProperties11.Append(fontSizeComplexScript6);

            style15.Append(styleName15);
            style15.Append(basedOn11);
            style15.Append(linkedStyle11);
            style15.Append(uIPriority14);
            style15.Append(semiHidden12);
            style15.Append(rsid15);
            style15.Append(styleRunProperties11);

            Style style16 = new Style(){ Type = StyleValues.Character, StyleId = "Heading3Char", CustomStyle = true };
            StyleName styleName16 = new StyleName(){ Val = "Heading 3 Char" };
            BasedOn basedOn12 = new BasedOn(){ Val = "DefaultParagraphFont" };
            LinkedStyle linkedStyle12 = new LinkedStyle(){ Val = "Heading3" };
            UIPriority uIPriority15 = new UIPriority(){ Val = 9 };
            SemiHidden semiHidden13 = new SemiHidden();
            Rsid rsid16 = new Rsid(){ Val = "00440271" };

            StyleRunProperties styleRunProperties12 = new StyleRunProperties();
            RunFonts runFonts13 = new RunFonts(){ EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
            Color color12 = new Color(){ Val = "0F4761", ThemeColor = ThemeColorValues.Accent1, ThemeShade = "BF" };
            FontSize fontSize7 = new FontSize(){ Val = "28" };
            FontSizeComplexScript fontSizeComplexScript7 = new FontSizeComplexScript(){ Val = "28" };

            styleRunProperties12.Append(runFonts13);
            styleRunProperties12.Append(color12);
            styleRunProperties12.Append(fontSize7);
            styleRunProperties12.Append(fontSizeComplexScript7);

            style16.Append(styleName16);
            style16.Append(basedOn12);
            style16.Append(linkedStyle12);
            style16.Append(uIPriority15);
            style16.Append(semiHidden13);
            style16.Append(rsid16);
            style16.Append(styleRunProperties12);

            Style style17 = new Style(){ Type = StyleValues.Character, StyleId = "Heading4Char", CustomStyle = true };
            StyleName styleName17 = new StyleName(){ Val = "Heading 4 Char" };
            BasedOn basedOn13 = new BasedOn(){ Val = "DefaultParagraphFont" };
            LinkedStyle linkedStyle13 = new LinkedStyle(){ Val = "Heading4" };
            UIPriority uIPriority16 = new UIPriority(){ Val = 9 };
            SemiHidden semiHidden14 = new SemiHidden();
            Rsid rsid17 = new Rsid(){ Val = "00440271" };

            StyleRunProperties styleRunProperties13 = new StyleRunProperties();
            RunFonts runFonts14 = new RunFonts(){ EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
            Italic italic4 = new Italic();
            ItalicComplexScript italicComplexScript4 = new ItalicComplexScript();
            Color color13 = new Color(){ Val = "0F4761", ThemeColor = ThemeColorValues.Accent1, ThemeShade = "BF" };

            styleRunProperties13.Append(runFonts14);
            styleRunProperties13.Append(italic4);
            styleRunProperties13.Append(italicComplexScript4);
            styleRunProperties13.Append(color13);

            style17.Append(styleName17);
            style17.Append(basedOn13);
            style17.Append(linkedStyle13);
            style17.Append(uIPriority16);
            style17.Append(semiHidden14);
            style17.Append(rsid17);
            style17.Append(styleRunProperties13);

            Style style18 = new Style(){ Type = StyleValues.Character, StyleId = "Heading5Char", CustomStyle = true };
            StyleName styleName18 = new StyleName(){ Val = "Heading 5 Char" };
            BasedOn basedOn14 = new BasedOn(){ Val = "DefaultParagraphFont" };
            LinkedStyle linkedStyle14 = new LinkedStyle(){ Val = "Heading5" };
            UIPriority uIPriority17 = new UIPriority(){ Val = 9 };
            SemiHidden semiHidden15 = new SemiHidden();
            Rsid rsid18 = new Rsid(){ Val = "00440271" };

            StyleRunProperties styleRunProperties14 = new StyleRunProperties();
            RunFonts runFonts15 = new RunFonts(){ EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
            Color color14 = new Color(){ Val = "0F4761", ThemeColor = ThemeColorValues.Accent1, ThemeShade = "BF" };

            styleRunProperties14.Append(runFonts15);
            styleRunProperties14.Append(color14);

            style18.Append(styleName18);
            style18.Append(basedOn14);
            style18.Append(linkedStyle14);
            style18.Append(uIPriority17);
            style18.Append(semiHidden15);
            style18.Append(rsid18);
            style18.Append(styleRunProperties14);

            Style style19 = new Style(){ Type = StyleValues.Character, StyleId = "Heading6Char", CustomStyle = true };
            StyleName styleName19 = new StyleName(){ Val = "Heading 6 Char" };
            BasedOn basedOn15 = new BasedOn(){ Val = "DefaultParagraphFont" };
            LinkedStyle linkedStyle15 = new LinkedStyle(){ Val = "Heading6" };
            UIPriority uIPriority18 = new UIPriority(){ Val = 9 };
            SemiHidden semiHidden16 = new SemiHidden();
            Rsid rsid19 = new Rsid(){ Val = "00440271" };

            StyleRunProperties styleRunProperties15 = new StyleRunProperties();
            RunFonts runFonts16 = new RunFonts(){ EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
            Italic italic5 = new Italic();
            ItalicComplexScript italicComplexScript5 = new ItalicComplexScript();
            Color color15 = new Color(){ Val = "595959", ThemeColor = ThemeColorValues.Text1, ThemeTint = "A6" };

            styleRunProperties15.Append(runFonts16);
            styleRunProperties15.Append(italic5);
            styleRunProperties15.Append(italicComplexScript5);
            styleRunProperties15.Append(color15);

            style19.Append(styleName19);
            style19.Append(basedOn15);
            style19.Append(linkedStyle15);
            style19.Append(uIPriority18);
            style19.Append(semiHidden16);
            style19.Append(rsid19);
            style19.Append(styleRunProperties15);

            Style style20 = new Style(){ Type = StyleValues.Character, StyleId = "Heading7Char", CustomStyle = true };
            StyleName styleName20 = new StyleName(){ Val = "Heading 7 Char" };
            BasedOn basedOn16 = new BasedOn(){ Val = "DefaultParagraphFont" };
            LinkedStyle linkedStyle16 = new LinkedStyle(){ Val = "Heading7" };
            UIPriority uIPriority19 = new UIPriority(){ Val = 9 };
            SemiHidden semiHidden17 = new SemiHidden();
            Rsid rsid20 = new Rsid(){ Val = "00440271" };

            StyleRunProperties styleRunProperties16 = new StyleRunProperties();
            RunFonts runFonts17 = new RunFonts(){ EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
            Color color16 = new Color(){ Val = "595959", ThemeColor = ThemeColorValues.Text1, ThemeTint = "A6" };

            styleRunProperties16.Append(runFonts17);
            styleRunProperties16.Append(color16);

            style20.Append(styleName20);
            style20.Append(basedOn16);
            style20.Append(linkedStyle16);
            style20.Append(uIPriority19);
            style20.Append(semiHidden17);
            style20.Append(rsid20);
            style20.Append(styleRunProperties16);

            Style style21 = new Style(){ Type = StyleValues.Character, StyleId = "Heading8Char", CustomStyle = true };
            StyleName styleName21 = new StyleName(){ Val = "Heading 8 Char" };
            BasedOn basedOn17 = new BasedOn(){ Val = "DefaultParagraphFont" };
            LinkedStyle linkedStyle17 = new LinkedStyle(){ Val = "Heading8" };
            UIPriority uIPriority20 = new UIPriority(){ Val = 9 };
            SemiHidden semiHidden18 = new SemiHidden();
            Rsid rsid21 = new Rsid(){ Val = "00440271" };

            StyleRunProperties styleRunProperties17 = new StyleRunProperties();
            RunFonts runFonts18 = new RunFonts(){ EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
            Italic italic6 = new Italic();
            ItalicComplexScript italicComplexScript6 = new ItalicComplexScript();
            Color color17 = new Color(){ Val = "272727", ThemeColor = ThemeColorValues.Text1, ThemeTint = "D8" };

            styleRunProperties17.Append(runFonts18);
            styleRunProperties17.Append(italic6);
            styleRunProperties17.Append(italicComplexScript6);
            styleRunProperties17.Append(color17);

            style21.Append(styleName21);
            style21.Append(basedOn17);
            style21.Append(linkedStyle17);
            style21.Append(uIPriority20);
            style21.Append(semiHidden18);
            style21.Append(rsid21);
            style21.Append(styleRunProperties17);

            Style style22 = new Style(){ Type = StyleValues.Character, StyleId = "Heading9Char", CustomStyle = true };
            StyleName styleName22 = new StyleName(){ Val = "Heading 9 Char" };
            BasedOn basedOn18 = new BasedOn(){ Val = "DefaultParagraphFont" };
            LinkedStyle linkedStyle18 = new LinkedStyle(){ Val = "Heading9" };
            UIPriority uIPriority21 = new UIPriority(){ Val = 9 };
            SemiHidden semiHidden19 = new SemiHidden();
            Rsid rsid22 = new Rsid(){ Val = "00440271" };

            StyleRunProperties styleRunProperties18 = new StyleRunProperties();
            RunFonts runFonts19 = new RunFonts(){ EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
            Color color18 = new Color(){ Val = "272727", ThemeColor = ThemeColorValues.Text1, ThemeTint = "D8" };

            styleRunProperties18.Append(runFonts19);
            styleRunProperties18.Append(color18);

            style22.Append(styleName22);
            style22.Append(basedOn18);
            style22.Append(linkedStyle18);
            style22.Append(uIPriority21);
            style22.Append(semiHidden19);
            style22.Append(rsid22);
            style22.Append(styleRunProperties18);

            Style style23 = new Style(){ Type = StyleValues.Paragraph, StyleId = "Title" };
            StyleName styleName23 = new StyleName(){ Val = "Title" };
            BasedOn basedOn19 = new BasedOn(){ Val = "Normal" };
            NextParagraphStyle nextParagraphStyle10 = new NextParagraphStyle(){ Val = "Normal" };
            LinkedStyle linkedStyle19 = new LinkedStyle(){ Val = "TitleChar" };
            UIPriority uIPriority22 = new UIPriority(){ Val = 10 };
            PrimaryStyle primaryStyle11 = new PrimaryStyle();
            Rsid rsid23 = new Rsid(){ Val = "00440271" };

            StyleParagraphProperties styleParagraphProperties10 = new StyleParagraphProperties();
            SpacingBetweenLines spacingBetweenLines11 = new SpacingBetweenLines(){ After = "80", Line = "240", LineRule = LineSpacingRuleValues.Auto };
            ContextualSpacing contextualSpacing1 = new ContextualSpacing();

            styleParagraphProperties10.Append(spacingBetweenLines11);
            styleParagraphProperties10.Append(contextualSpacing1);

            StyleRunProperties styleRunProperties19 = new StyleRunProperties();
            RunFonts runFonts20 = new RunFonts(){ AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
            Spacing spacing1 = new Spacing(){ Val = -10 };
            Kern kern2 = new Kern(){ Val = (UInt32Value)28U };
            FontSize fontSize8 = new FontSize(){ Val = "56" };
            FontSizeComplexScript fontSizeComplexScript8 = new FontSizeComplexScript(){ Val = "56" };

            styleRunProperties19.Append(runFonts20);
            styleRunProperties19.Append(spacing1);
            styleRunProperties19.Append(kern2);
            styleRunProperties19.Append(fontSize8);
            styleRunProperties19.Append(fontSizeComplexScript8);

            style23.Append(styleName23);
            style23.Append(basedOn19);
            style23.Append(nextParagraphStyle10);
            style23.Append(linkedStyle19);
            style23.Append(uIPriority22);
            style23.Append(primaryStyle11);
            style23.Append(rsid23);
            style23.Append(styleParagraphProperties10);
            style23.Append(styleRunProperties19);

            Style style24 = new Style(){ Type = StyleValues.Character, StyleId = "TitleChar", CustomStyle = true };
            StyleName styleName24 = new StyleName(){ Val = "Title Char" };
            BasedOn basedOn20 = new BasedOn(){ Val = "DefaultParagraphFont" };
            LinkedStyle linkedStyle20 = new LinkedStyle(){ Val = "Title" };
            UIPriority uIPriority23 = new UIPriority(){ Val = 10 };
            Rsid rsid24 = new Rsid(){ Val = "00440271" };

            StyleRunProperties styleRunProperties20 = new StyleRunProperties();
            RunFonts runFonts21 = new RunFonts(){ AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
            Spacing spacing2 = new Spacing(){ Val = -10 };
            Kern kern3 = new Kern(){ Val = (UInt32Value)28U };
            FontSize fontSize9 = new FontSize(){ Val = "56" };
            FontSizeComplexScript fontSizeComplexScript9 = new FontSizeComplexScript(){ Val = "56" };

            styleRunProperties20.Append(runFonts21);
            styleRunProperties20.Append(spacing2);
            styleRunProperties20.Append(kern3);
            styleRunProperties20.Append(fontSize9);
            styleRunProperties20.Append(fontSizeComplexScript9);

            style24.Append(styleName24);
            style24.Append(basedOn20);
            style24.Append(linkedStyle20);
            style24.Append(uIPriority23);
            style24.Append(rsid24);
            style24.Append(styleRunProperties20);

            Style style25 = new Style(){ Type = StyleValues.Paragraph, StyleId = "Subtitle" };
            StyleName styleName25 = new StyleName(){ Val = "Subtitle" };
            BasedOn basedOn21 = new BasedOn(){ Val = "Normal" };
            NextParagraphStyle nextParagraphStyle11 = new NextParagraphStyle(){ Val = "Normal" };
            LinkedStyle linkedStyle21 = new LinkedStyle(){ Val = "SubtitleChar" };
            UIPriority uIPriority24 = new UIPriority(){ Val = 11 };
            PrimaryStyle primaryStyle12 = new PrimaryStyle();
            Rsid rsid25 = new Rsid(){ Val = "00440271" };

            StyleParagraphProperties styleParagraphProperties11 = new StyleParagraphProperties();

            NumberingProperties numberingProperties1 = new NumberingProperties();
            NumberingLevelReference numberingLevelReference1 = new NumberingLevelReference(){ Val = 1 };

            numberingProperties1.Append(numberingLevelReference1);

            styleParagraphProperties11.Append(numberingProperties1);

            StyleRunProperties styleRunProperties21 = new StyleRunProperties();
            RunFonts runFonts22 = new RunFonts(){ EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
            Color color19 = new Color(){ Val = "595959", ThemeColor = ThemeColorValues.Text1, ThemeTint = "A6" };
            Spacing spacing3 = new Spacing(){ Val = 15 };
            FontSize fontSize10 = new FontSize(){ Val = "28" };
            FontSizeComplexScript fontSizeComplexScript10 = new FontSizeComplexScript(){ Val = "28" };

            styleRunProperties21.Append(runFonts22);
            styleRunProperties21.Append(color19);
            styleRunProperties21.Append(spacing3);
            styleRunProperties21.Append(fontSize10);
            styleRunProperties21.Append(fontSizeComplexScript10);

            style25.Append(styleName25);
            style25.Append(basedOn21);
            style25.Append(nextParagraphStyle11);
            style25.Append(linkedStyle21);
            style25.Append(uIPriority24);
            style25.Append(primaryStyle12);
            style25.Append(rsid25);
            style25.Append(styleParagraphProperties11);
            style25.Append(styleRunProperties21);

            Style style26 = new Style(){ Type = StyleValues.Character, StyleId = "SubtitleChar", CustomStyle = true };
            StyleName styleName26 = new StyleName(){ Val = "Subtitle Char" };
            BasedOn basedOn22 = new BasedOn(){ Val = "DefaultParagraphFont" };
            LinkedStyle linkedStyle22 = new LinkedStyle(){ Val = "Subtitle" };
            UIPriority uIPriority25 = new UIPriority(){ Val = 11 };
            Rsid rsid26 = new Rsid(){ Val = "00440271" };

            StyleRunProperties styleRunProperties22 = new StyleRunProperties();
            RunFonts runFonts23 = new RunFonts(){ EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
            Color color20 = new Color(){ Val = "595959", ThemeColor = ThemeColorValues.Text1, ThemeTint = "A6" };
            Spacing spacing4 = new Spacing(){ Val = 15 };
            FontSize fontSize11 = new FontSize(){ Val = "28" };
            FontSizeComplexScript fontSizeComplexScript11 = new FontSizeComplexScript(){ Val = "28" };

            styleRunProperties22.Append(runFonts23);
            styleRunProperties22.Append(color20);
            styleRunProperties22.Append(spacing4);
            styleRunProperties22.Append(fontSize11);
            styleRunProperties22.Append(fontSizeComplexScript11);

            style26.Append(styleName26);
            style26.Append(basedOn22);
            style26.Append(linkedStyle22);
            style26.Append(uIPriority25);
            style26.Append(rsid26);
            style26.Append(styleRunProperties22);

            Style style27 = new Style(){ Type = StyleValues.Paragraph, StyleId = "Quote" };
            StyleName styleName27 = new StyleName(){ Val = "Quote" };
            BasedOn basedOn23 = new BasedOn(){ Val = "Normal" };
            NextParagraphStyle nextParagraphStyle12 = new NextParagraphStyle(){ Val = "Normal" };
            LinkedStyle linkedStyle23 = new LinkedStyle(){ Val = "QuoteChar" };
            UIPriority uIPriority26 = new UIPriority(){ Val = 29 };
            PrimaryStyle primaryStyle13 = new PrimaryStyle();
            Rsid rsid27 = new Rsid(){ Val = "00440271" };

            StyleParagraphProperties styleParagraphProperties12 = new StyleParagraphProperties();
            SpacingBetweenLines spacingBetweenLines12 = new SpacingBetweenLines(){ Before = "160" };
            Justification justification1 = new Justification(){ Val = JustificationValues.Center };

            styleParagraphProperties12.Append(spacingBetweenLines12);
            styleParagraphProperties12.Append(justification1);

            StyleRunProperties styleRunProperties23 = new StyleRunProperties();
            Italic italic7 = new Italic();
            ItalicComplexScript italicComplexScript7 = new ItalicComplexScript();
            Color color21 = new Color(){ Val = "404040", ThemeColor = ThemeColorValues.Text1, ThemeTint = "BF" };

            styleRunProperties23.Append(italic7);
            styleRunProperties23.Append(italicComplexScript7);
            styleRunProperties23.Append(color21);

            style27.Append(styleName27);
            style27.Append(basedOn23);
            style27.Append(nextParagraphStyle12);
            style27.Append(linkedStyle23);
            style27.Append(uIPriority26);
            style27.Append(primaryStyle13);
            style27.Append(rsid27);
            style27.Append(styleParagraphProperties12);
            style27.Append(styleRunProperties23);

            Style style28 = new Style(){ Type = StyleValues.Character, StyleId = "QuoteChar", CustomStyle = true };
            StyleName styleName28 = new StyleName(){ Val = "Quote Char" };
            BasedOn basedOn24 = new BasedOn(){ Val = "DefaultParagraphFont" };
            LinkedStyle linkedStyle24 = new LinkedStyle(){ Val = "Quote" };
            UIPriority uIPriority27 = new UIPriority(){ Val = 29 };
            Rsid rsid28 = new Rsid(){ Val = "00440271" };

            StyleRunProperties styleRunProperties24 = new StyleRunProperties();
            Italic italic8 = new Italic();
            ItalicComplexScript italicComplexScript8 = new ItalicComplexScript();
            Color color22 = new Color(){ Val = "404040", ThemeColor = ThemeColorValues.Text1, ThemeTint = "BF" };

            styleRunProperties24.Append(italic8);
            styleRunProperties24.Append(italicComplexScript8);
            styleRunProperties24.Append(color22);

            style28.Append(styleName28);
            style28.Append(basedOn24);
            style28.Append(linkedStyle24);
            style28.Append(uIPriority27);
            style28.Append(rsid28);
            style28.Append(styleRunProperties24);

            Style style29 = new Style(){ Type = StyleValues.Paragraph, StyleId = "ListParagraph" };
            StyleName styleName29 = new StyleName(){ Val = "List Paragraph" };
            BasedOn basedOn25 = new BasedOn(){ Val = "Normal" };
            UIPriority uIPriority28 = new UIPriority(){ Val = 34 };
            PrimaryStyle primaryStyle14 = new PrimaryStyle();
            Rsid rsid29 = new Rsid(){ Val = "00440271" };

            StyleParagraphProperties styleParagraphProperties13 = new StyleParagraphProperties();
            Indentation indentation1 = new Indentation(){ Start = "720" };
            ContextualSpacing contextualSpacing2 = new ContextualSpacing();

            styleParagraphProperties13.Append(indentation1);
            styleParagraphProperties13.Append(contextualSpacing2);

            style29.Append(styleName29);
            style29.Append(basedOn25);
            style29.Append(uIPriority28);
            style29.Append(primaryStyle14);
            style29.Append(rsid29);
            style29.Append(styleParagraphProperties13);

            Style style30 = new Style(){ Type = StyleValues.Character, StyleId = "IntenseEmphasis" };
            StyleName styleName30 = new StyleName(){ Val = "Intense Emphasis" };
            BasedOn basedOn26 = new BasedOn(){ Val = "DefaultParagraphFont" };
            UIPriority uIPriority29 = new UIPriority(){ Val = 21 };
            PrimaryStyle primaryStyle15 = new PrimaryStyle();
            Rsid rsid30 = new Rsid(){ Val = "00440271" };

            StyleRunProperties styleRunProperties25 = new StyleRunProperties();
            Italic italic9 = new Italic();
            ItalicComplexScript italicComplexScript9 = new ItalicComplexScript();
            Color color23 = new Color(){ Val = "0F4761", ThemeColor = ThemeColorValues.Accent1, ThemeShade = "BF" };

            styleRunProperties25.Append(italic9);
            styleRunProperties25.Append(italicComplexScript9);
            styleRunProperties25.Append(color23);

            style30.Append(styleName30);
            style30.Append(basedOn26);
            style30.Append(uIPriority29);
            style30.Append(primaryStyle15);
            style30.Append(rsid30);
            style30.Append(styleRunProperties25);

            Style style31 = new Style(){ Type = StyleValues.Paragraph, StyleId = "IntenseQuote" };
            StyleName styleName31 = new StyleName(){ Val = "Intense Quote" };
            BasedOn basedOn27 = new BasedOn(){ Val = "Normal" };
            NextParagraphStyle nextParagraphStyle13 = new NextParagraphStyle(){ Val = "Normal" };
            LinkedStyle linkedStyle25 = new LinkedStyle(){ Val = "IntenseQuoteChar" };
            UIPriority uIPriority30 = new UIPriority(){ Val = 30 };
            PrimaryStyle primaryStyle16 = new PrimaryStyle();
            Rsid rsid31 = new Rsid(){ Val = "00440271" };

            StyleParagraphProperties styleParagraphProperties14 = new StyleParagraphProperties();

            ParagraphBorders paragraphBorders1 = new ParagraphBorders();
            TopBorder topBorder1 = new TopBorder(){ Val = BorderValues.Single, Color = "0F4761", ThemeColor = ThemeColorValues.Accent1, ThemeShade = "BF", Size = (UInt32Value)4U, Space = (UInt32Value)10U };
            BottomBorder bottomBorder1 = new BottomBorder(){ Val = BorderValues.Single, Color = "0F4761", ThemeColor = ThemeColorValues.Accent1, ThemeShade = "BF", Size = (UInt32Value)4U, Space = (UInt32Value)10U };

            paragraphBorders1.Append(topBorder1);
            paragraphBorders1.Append(bottomBorder1);
            SpacingBetweenLines spacingBetweenLines13 = new SpacingBetweenLines(){ Before = "360", After = "360" };
            Indentation indentation2 = new Indentation(){ Start = "864", End = "864" };
            Justification justification2 = new Justification(){ Val = JustificationValues.Center };

            styleParagraphProperties14.Append(paragraphBorders1);
            styleParagraphProperties14.Append(spacingBetweenLines13);
            styleParagraphProperties14.Append(indentation2);
            styleParagraphProperties14.Append(justification2);

            StyleRunProperties styleRunProperties26 = new StyleRunProperties();
            Italic italic10 = new Italic();
            ItalicComplexScript italicComplexScript10 = new ItalicComplexScript();
            Color color24 = new Color(){ Val = "0F4761", ThemeColor = ThemeColorValues.Accent1, ThemeShade = "BF" };

            styleRunProperties26.Append(italic10);
            styleRunProperties26.Append(italicComplexScript10);
            styleRunProperties26.Append(color24);

            style31.Append(styleName31);
            style31.Append(basedOn27);
            style31.Append(nextParagraphStyle13);
            style31.Append(linkedStyle25);
            style31.Append(uIPriority30);
            style31.Append(primaryStyle16);
            style31.Append(rsid31);
            style31.Append(styleParagraphProperties14);
            style31.Append(styleRunProperties26);

            Style style32 = new Style(){ Type = StyleValues.Character, StyleId = "IntenseQuoteChar", CustomStyle = true };
            StyleName styleName32 = new StyleName(){ Val = "Intense Quote Char" };
            BasedOn basedOn28 = new BasedOn(){ Val = "DefaultParagraphFont" };
            LinkedStyle linkedStyle26 = new LinkedStyle(){ Val = "IntenseQuote" };
            UIPriority uIPriority31 = new UIPriority(){ Val = 30 };
            Rsid rsid32 = new Rsid(){ Val = "00440271" };

            StyleRunProperties styleRunProperties27 = new StyleRunProperties();
            Italic italic11 = new Italic();
            ItalicComplexScript italicComplexScript11 = new ItalicComplexScript();
            Color color25 = new Color(){ Val = "0F4761", ThemeColor = ThemeColorValues.Accent1, ThemeShade = "BF" };

            styleRunProperties27.Append(italic11);
            styleRunProperties27.Append(italicComplexScript11);
            styleRunProperties27.Append(color25);

            style32.Append(styleName32);
            style32.Append(basedOn28);
            style32.Append(linkedStyle26);
            style32.Append(uIPriority31);
            style32.Append(rsid32);
            style32.Append(styleRunProperties27);

            Style style33 = new Style(){ Type = StyleValues.Character, StyleId = "IntenseReference" };
            StyleName styleName33 = new StyleName(){ Val = "Intense Reference" };
            BasedOn basedOn29 = new BasedOn(){ Val = "DefaultParagraphFont" };
            UIPriority uIPriority32 = new UIPriority(){ Val = 32 };
            PrimaryStyle primaryStyle17 = new PrimaryStyle();
            Rsid rsid33 = new Rsid(){ Val = "00440271" };

            StyleRunProperties styleRunProperties28 = new StyleRunProperties();
            Bold bold1 = new Bold();
            BoldComplexScript boldComplexScript1 = new BoldComplexScript();
            SmallCaps smallCaps1 = new SmallCaps();
            Color color26 = new Color(){ Val = "0F4761", ThemeColor = ThemeColorValues.Accent1, ThemeShade = "BF" };
            Spacing spacing5 = new Spacing(){ Val = 5 };

            styleRunProperties28.Append(bold1);
            styleRunProperties28.Append(boldComplexScript1);
            styleRunProperties28.Append(smallCaps1);
            styleRunProperties28.Append(color26);
            styleRunProperties28.Append(spacing5);

            style33.Append(styleName33);
            style33.Append(basedOn29);
            style33.Append(uIPriority32);
            style33.Append(primaryStyle17);
            style33.Append(rsid33);
            style33.Append(styleRunProperties28);

            styles1.Append(docDefaults1);
            styles1.Append(latentStyles1);
            styles1.Append(style1);
            styles1.Append(style2);
            styles1.Append(style3);
            styles1.Append(style4);
            styles1.Append(style5);
            styles1.Append(style6);
            styles1.Append(style7);
            styles1.Append(style8);
            styles1.Append(style9);
            styles1.Append(style10);
            styles1.Append(style11);
            styles1.Append(style12);
            styles1.Append(style13);
            styles1.Append(style14);
            styles1.Append(style15);
            styles1.Append(style16);
            styles1.Append(style17);
            styles1.Append(style18);
            styles1.Append(style19);
            styles1.Append(style20);
            styles1.Append(style21);
            styles1.Append(style22);
            styles1.Append(style23);
            styles1.Append(style24);
            styles1.Append(style25);
            styles1.Append(style26);
            styles1.Append(style27);
            styles1.Append(style28);
            styles1.Append(style29);
            styles1.Append(style30);
            styles1.Append(style31);
            styles1.Append(style32);
            styles1.Append(style33);

            styleDefinitionsPart1.Styles = styles1;
        }

        // Generates content of diagramStylePart1.
        private void GenerateDiagramStylePart1Content(DiagramStylePart diagramStylePart1)
        {
            Dgm.StyleDefinition styleDefinition1 = new Dgm.StyleDefinition(){ UniqueId = "urn:microsoft.com/office/officeart/2005/8/quickstyle/simple1" };
            styleDefinition1.AddNamespaceDeclaration("dgm", "http://schemas.openxmlformats.org/drawingml/2006/diagram");
            styleDefinition1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            Dgm.StyleDefinitionTitle styleDefinitionTitle1 = new Dgm.StyleDefinitionTitle(){ Val = "" };
            Dgm.StyleLabelDescription styleLabelDescription1 = new Dgm.StyleLabelDescription(){ Val = "" };

            Dgm.StyleDisplayCategories styleDisplayCategories1 = new Dgm.StyleDisplayCategories();
            Dgm.StyleDisplayCategory styleDisplayCategory1 = new Dgm.StyleDisplayCategory(){ Type = "simple", Priority = (UInt32Value)10100U };

            styleDisplayCategories1.Append(styleDisplayCategory1);

            Dgm.Scene3D scene3D1 = new Dgm.Scene3D();
            A.Camera camera1 = new A.Camera(){ Preset = A.PresetCameraValues.OrthographicFront };
            A.LightRig lightRig1 = new A.LightRig(){ Rig = A.LightRigValues.ThreePoints, Direction = A.LightRigDirectionValues.Top };

            scene3D1.Append(camera1);
            scene3D1.Append(lightRig1);

            Dgm.StyleLabel styleLabel1 = new Dgm.StyleLabel(){ Name = "node0" };

            Dgm.Scene3D scene3D2 = new Dgm.Scene3D();
            A.Camera camera2 = new A.Camera(){ Preset = A.PresetCameraValues.OrthographicFront };
            A.LightRig lightRig2 = new A.LightRig(){ Rig = A.LightRigValues.ThreePoints, Direction = A.LightRigDirectionValues.Top };

            scene3D2.Append(camera2);
            scene3D2.Append(lightRig2);
            Dgm.Shape3D shape3D1 = new Dgm.Shape3D();
            Dgm.TextProperties textProperties1 = new Dgm.TextProperties();

            Dgm.Style style34 = new Dgm.Style();

            A.LineReference lineReference6 = new A.LineReference(){ Index = (UInt32Value)2U };
            A.RgbColorModelPercentage rgbColorModelPercentage16 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            lineReference6.Append(rgbColorModelPercentage16);

            A.FillReference fillReference6 = new A.FillReference(){ Index = (UInt32Value)1U };
            A.RgbColorModelPercentage rgbColorModelPercentage17 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            fillReference6.Append(rgbColorModelPercentage17);

            A.EffectReference effectReference6 = new A.EffectReference(){ Index = (UInt32Value)0U };
            A.RgbColorModelPercentage rgbColorModelPercentage18 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            effectReference6.Append(rgbColorModelPercentage18);

            A.FontReference fontReference6 = new A.FontReference(){ Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor147 = new A.SchemeColor(){ Val = A.SchemeColorValues.Light1 };

            fontReference6.Append(schemeColor147);

            style34.Append(lineReference6);
            style34.Append(fillReference6);
            style34.Append(effectReference6);
            style34.Append(fontReference6);

            styleLabel1.Append(scene3D2);
            styleLabel1.Append(shape3D1);
            styleLabel1.Append(textProperties1);
            styleLabel1.Append(style34);

            Dgm.StyleLabel styleLabel2 = new Dgm.StyleLabel(){ Name = "lnNode1" };

            Dgm.Scene3D scene3D3 = new Dgm.Scene3D();
            A.Camera camera3 = new A.Camera(){ Preset = A.PresetCameraValues.OrthographicFront };
            A.LightRig lightRig3 = new A.LightRig(){ Rig = A.LightRigValues.ThreePoints, Direction = A.LightRigDirectionValues.Top };

            scene3D3.Append(camera3);
            scene3D3.Append(lightRig3);
            Dgm.Shape3D shape3D2 = new Dgm.Shape3D();
            Dgm.TextProperties textProperties2 = new Dgm.TextProperties();

            Dgm.Style style35 = new Dgm.Style();

            A.LineReference lineReference7 = new A.LineReference(){ Index = (UInt32Value)2U };
            A.RgbColorModelPercentage rgbColorModelPercentage19 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            lineReference7.Append(rgbColorModelPercentage19);

            A.FillReference fillReference7 = new A.FillReference(){ Index = (UInt32Value)1U };
            A.RgbColorModelPercentage rgbColorModelPercentage20 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            fillReference7.Append(rgbColorModelPercentage20);

            A.EffectReference effectReference7 = new A.EffectReference(){ Index = (UInt32Value)0U };
            A.RgbColorModelPercentage rgbColorModelPercentage21 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            effectReference7.Append(rgbColorModelPercentage21);

            A.FontReference fontReference7 = new A.FontReference(){ Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor148 = new A.SchemeColor(){ Val = A.SchemeColorValues.Light1 };

            fontReference7.Append(schemeColor148);

            style35.Append(lineReference7);
            style35.Append(fillReference7);
            style35.Append(effectReference7);
            style35.Append(fontReference7);

            styleLabel2.Append(scene3D3);
            styleLabel2.Append(shape3D2);
            styleLabel2.Append(textProperties2);
            styleLabel2.Append(style35);

            Dgm.StyleLabel styleLabel3 = new Dgm.StyleLabel(){ Name = "vennNode1" };

            Dgm.Scene3D scene3D4 = new Dgm.Scene3D();
            A.Camera camera4 = new A.Camera(){ Preset = A.PresetCameraValues.OrthographicFront };
            A.LightRig lightRig4 = new A.LightRig(){ Rig = A.LightRigValues.ThreePoints, Direction = A.LightRigDirectionValues.Top };

            scene3D4.Append(camera4);
            scene3D4.Append(lightRig4);
            Dgm.Shape3D shape3D3 = new Dgm.Shape3D();
            Dgm.TextProperties textProperties3 = new Dgm.TextProperties();

            Dgm.Style style36 = new Dgm.Style();

            A.LineReference lineReference8 = new A.LineReference(){ Index = (UInt32Value)2U };
            A.RgbColorModelPercentage rgbColorModelPercentage22 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            lineReference8.Append(rgbColorModelPercentage22);

            A.FillReference fillReference8 = new A.FillReference(){ Index = (UInt32Value)1U };
            A.RgbColorModelPercentage rgbColorModelPercentage23 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            fillReference8.Append(rgbColorModelPercentage23);

            A.EffectReference effectReference8 = new A.EffectReference(){ Index = (UInt32Value)0U };
            A.RgbColorModelPercentage rgbColorModelPercentage24 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            effectReference8.Append(rgbColorModelPercentage24);

            A.FontReference fontReference8 = new A.FontReference(){ Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor149 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            fontReference8.Append(schemeColor149);

            style36.Append(lineReference8);
            style36.Append(fillReference8);
            style36.Append(effectReference8);
            style36.Append(fontReference8);

            styleLabel3.Append(scene3D4);
            styleLabel3.Append(shape3D3);
            styleLabel3.Append(textProperties3);
            styleLabel3.Append(style36);

            Dgm.StyleLabel styleLabel4 = new Dgm.StyleLabel(){ Name = "alignNode1" };

            Dgm.Scene3D scene3D5 = new Dgm.Scene3D();
            A.Camera camera5 = new A.Camera(){ Preset = A.PresetCameraValues.OrthographicFront };
            A.LightRig lightRig5 = new A.LightRig(){ Rig = A.LightRigValues.ThreePoints, Direction = A.LightRigDirectionValues.Top };

            scene3D5.Append(camera5);
            scene3D5.Append(lightRig5);
            Dgm.Shape3D shape3D4 = new Dgm.Shape3D();
            Dgm.TextProperties textProperties4 = new Dgm.TextProperties();

            Dgm.Style style37 = new Dgm.Style();

            A.LineReference lineReference9 = new A.LineReference(){ Index = (UInt32Value)2U };
            A.RgbColorModelPercentage rgbColorModelPercentage25 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            lineReference9.Append(rgbColorModelPercentage25);

            A.FillReference fillReference9 = new A.FillReference(){ Index = (UInt32Value)1U };
            A.RgbColorModelPercentage rgbColorModelPercentage26 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            fillReference9.Append(rgbColorModelPercentage26);

            A.EffectReference effectReference9 = new A.EffectReference(){ Index = (UInt32Value)0U };
            A.RgbColorModelPercentage rgbColorModelPercentage27 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            effectReference9.Append(rgbColorModelPercentage27);

            A.FontReference fontReference9 = new A.FontReference(){ Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor150 = new A.SchemeColor(){ Val = A.SchemeColorValues.Light1 };

            fontReference9.Append(schemeColor150);

            style37.Append(lineReference9);
            style37.Append(fillReference9);
            style37.Append(effectReference9);
            style37.Append(fontReference9);

            styleLabel4.Append(scene3D5);
            styleLabel4.Append(shape3D4);
            styleLabel4.Append(textProperties4);
            styleLabel4.Append(style37);

            Dgm.StyleLabel styleLabel5 = new Dgm.StyleLabel(){ Name = "node1" };

            Dgm.Scene3D scene3D6 = new Dgm.Scene3D();
            A.Camera camera6 = new A.Camera(){ Preset = A.PresetCameraValues.OrthographicFront };
            A.LightRig lightRig6 = new A.LightRig(){ Rig = A.LightRigValues.ThreePoints, Direction = A.LightRigDirectionValues.Top };

            scene3D6.Append(camera6);
            scene3D6.Append(lightRig6);
            Dgm.Shape3D shape3D5 = new Dgm.Shape3D();
            Dgm.TextProperties textProperties5 = new Dgm.TextProperties();

            Dgm.Style style38 = new Dgm.Style();

            A.LineReference lineReference10 = new A.LineReference(){ Index = (UInt32Value)2U };
            A.RgbColorModelPercentage rgbColorModelPercentage28 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            lineReference10.Append(rgbColorModelPercentage28);

            A.FillReference fillReference10 = new A.FillReference(){ Index = (UInt32Value)1U };
            A.RgbColorModelPercentage rgbColorModelPercentage29 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            fillReference10.Append(rgbColorModelPercentage29);

            A.EffectReference effectReference10 = new A.EffectReference(){ Index = (UInt32Value)0U };
            A.RgbColorModelPercentage rgbColorModelPercentage30 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            effectReference10.Append(rgbColorModelPercentage30);

            A.FontReference fontReference10 = new A.FontReference(){ Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor151 = new A.SchemeColor(){ Val = A.SchemeColorValues.Light1 };

            fontReference10.Append(schemeColor151);

            style38.Append(lineReference10);
            style38.Append(fillReference10);
            style38.Append(effectReference10);
            style38.Append(fontReference10);

            styleLabel5.Append(scene3D6);
            styleLabel5.Append(shape3D5);
            styleLabel5.Append(textProperties5);
            styleLabel5.Append(style38);

            Dgm.StyleLabel styleLabel6 = new Dgm.StyleLabel(){ Name = "node2" };

            Dgm.Scene3D scene3D7 = new Dgm.Scene3D();
            A.Camera camera7 = new A.Camera(){ Preset = A.PresetCameraValues.OrthographicFront };
            A.LightRig lightRig7 = new A.LightRig(){ Rig = A.LightRigValues.ThreePoints, Direction = A.LightRigDirectionValues.Top };

            scene3D7.Append(camera7);
            scene3D7.Append(lightRig7);
            Dgm.Shape3D shape3D6 = new Dgm.Shape3D();
            Dgm.TextProperties textProperties6 = new Dgm.TextProperties();

            Dgm.Style style39 = new Dgm.Style();

            A.LineReference lineReference11 = new A.LineReference(){ Index = (UInt32Value)2U };
            A.RgbColorModelPercentage rgbColorModelPercentage31 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            lineReference11.Append(rgbColorModelPercentage31);

            A.FillReference fillReference11 = new A.FillReference(){ Index = (UInt32Value)1U };
            A.RgbColorModelPercentage rgbColorModelPercentage32 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            fillReference11.Append(rgbColorModelPercentage32);

            A.EffectReference effectReference11 = new A.EffectReference(){ Index = (UInt32Value)0U };
            A.RgbColorModelPercentage rgbColorModelPercentage33 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            effectReference11.Append(rgbColorModelPercentage33);

            A.FontReference fontReference11 = new A.FontReference(){ Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor152 = new A.SchemeColor(){ Val = A.SchemeColorValues.Light1 };

            fontReference11.Append(schemeColor152);

            style39.Append(lineReference11);
            style39.Append(fillReference11);
            style39.Append(effectReference11);
            style39.Append(fontReference11);

            styleLabel6.Append(scene3D7);
            styleLabel6.Append(shape3D6);
            styleLabel6.Append(textProperties6);
            styleLabel6.Append(style39);

            Dgm.StyleLabel styleLabel7 = new Dgm.StyleLabel(){ Name = "node3" };

            Dgm.Scene3D scene3D8 = new Dgm.Scene3D();
            A.Camera camera8 = new A.Camera(){ Preset = A.PresetCameraValues.OrthographicFront };
            A.LightRig lightRig8 = new A.LightRig(){ Rig = A.LightRigValues.ThreePoints, Direction = A.LightRigDirectionValues.Top };

            scene3D8.Append(camera8);
            scene3D8.Append(lightRig8);
            Dgm.Shape3D shape3D7 = new Dgm.Shape3D();
            Dgm.TextProperties textProperties7 = new Dgm.TextProperties();

            Dgm.Style style40 = new Dgm.Style();

            A.LineReference lineReference12 = new A.LineReference(){ Index = (UInt32Value)2U };
            A.RgbColorModelPercentage rgbColorModelPercentage34 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            lineReference12.Append(rgbColorModelPercentage34);

            A.FillReference fillReference12 = new A.FillReference(){ Index = (UInt32Value)1U };
            A.RgbColorModelPercentage rgbColorModelPercentage35 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            fillReference12.Append(rgbColorModelPercentage35);

            A.EffectReference effectReference12 = new A.EffectReference(){ Index = (UInt32Value)0U };
            A.RgbColorModelPercentage rgbColorModelPercentage36 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            effectReference12.Append(rgbColorModelPercentage36);

            A.FontReference fontReference12 = new A.FontReference(){ Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor153 = new A.SchemeColor(){ Val = A.SchemeColorValues.Light1 };

            fontReference12.Append(schemeColor153);

            style40.Append(lineReference12);
            style40.Append(fillReference12);
            style40.Append(effectReference12);
            style40.Append(fontReference12);

            styleLabel7.Append(scene3D8);
            styleLabel7.Append(shape3D7);
            styleLabel7.Append(textProperties7);
            styleLabel7.Append(style40);

            Dgm.StyleLabel styleLabel8 = new Dgm.StyleLabel(){ Name = "node4" };

            Dgm.Scene3D scene3D9 = new Dgm.Scene3D();
            A.Camera camera9 = new A.Camera(){ Preset = A.PresetCameraValues.OrthographicFront };
            A.LightRig lightRig9 = new A.LightRig(){ Rig = A.LightRigValues.ThreePoints, Direction = A.LightRigDirectionValues.Top };

            scene3D9.Append(camera9);
            scene3D9.Append(lightRig9);
            Dgm.Shape3D shape3D8 = new Dgm.Shape3D();
            Dgm.TextProperties textProperties8 = new Dgm.TextProperties();

            Dgm.Style style41 = new Dgm.Style();

            A.LineReference lineReference13 = new A.LineReference(){ Index = (UInt32Value)2U };
            A.RgbColorModelPercentage rgbColorModelPercentage37 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            lineReference13.Append(rgbColorModelPercentage37);

            A.FillReference fillReference13 = new A.FillReference(){ Index = (UInt32Value)1U };
            A.RgbColorModelPercentage rgbColorModelPercentage38 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            fillReference13.Append(rgbColorModelPercentage38);

            A.EffectReference effectReference13 = new A.EffectReference(){ Index = (UInt32Value)0U };
            A.RgbColorModelPercentage rgbColorModelPercentage39 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            effectReference13.Append(rgbColorModelPercentage39);

            A.FontReference fontReference13 = new A.FontReference(){ Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor154 = new A.SchemeColor(){ Val = A.SchemeColorValues.Light1 };

            fontReference13.Append(schemeColor154);

            style41.Append(lineReference13);
            style41.Append(fillReference13);
            style41.Append(effectReference13);
            style41.Append(fontReference13);

            styleLabel8.Append(scene3D9);
            styleLabel8.Append(shape3D8);
            styleLabel8.Append(textProperties8);
            styleLabel8.Append(style41);

            Dgm.StyleLabel styleLabel9 = new Dgm.StyleLabel(){ Name = "fgImgPlace1" };

            Dgm.Scene3D scene3D10 = new Dgm.Scene3D();
            A.Camera camera10 = new A.Camera(){ Preset = A.PresetCameraValues.OrthographicFront };
            A.LightRig lightRig10 = new A.LightRig(){ Rig = A.LightRigValues.ThreePoints, Direction = A.LightRigDirectionValues.Top };

            scene3D10.Append(camera10);
            scene3D10.Append(lightRig10);
            Dgm.Shape3D shape3D9 = new Dgm.Shape3D();
            Dgm.TextProperties textProperties9 = new Dgm.TextProperties();

            Dgm.Style style42 = new Dgm.Style();

            A.LineReference lineReference14 = new A.LineReference(){ Index = (UInt32Value)2U };
            A.RgbColorModelPercentage rgbColorModelPercentage40 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            lineReference14.Append(rgbColorModelPercentage40);

            A.FillReference fillReference14 = new A.FillReference(){ Index = (UInt32Value)1U };
            A.RgbColorModelPercentage rgbColorModelPercentage41 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            fillReference14.Append(rgbColorModelPercentage41);

            A.EffectReference effectReference14 = new A.EffectReference(){ Index = (UInt32Value)0U };
            A.RgbColorModelPercentage rgbColorModelPercentage42 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            effectReference14.Append(rgbColorModelPercentage42);
            A.FontReference fontReference14 = new A.FontReference(){ Index = A.FontCollectionIndexValues.Minor };

            style42.Append(lineReference14);
            style42.Append(fillReference14);
            style42.Append(effectReference14);
            style42.Append(fontReference14);

            styleLabel9.Append(scene3D10);
            styleLabel9.Append(shape3D9);
            styleLabel9.Append(textProperties9);
            styleLabel9.Append(style42);

            Dgm.StyleLabel styleLabel10 = new Dgm.StyleLabel(){ Name = "alignImgPlace1" };

            Dgm.Scene3D scene3D11 = new Dgm.Scene3D();
            A.Camera camera11 = new A.Camera(){ Preset = A.PresetCameraValues.OrthographicFront };
            A.LightRig lightRig11 = new A.LightRig(){ Rig = A.LightRigValues.ThreePoints, Direction = A.LightRigDirectionValues.Top };

            scene3D11.Append(camera11);
            scene3D11.Append(lightRig11);
            Dgm.Shape3D shape3D10 = new Dgm.Shape3D();
            Dgm.TextProperties textProperties10 = new Dgm.TextProperties();

            Dgm.Style style43 = new Dgm.Style();

            A.LineReference lineReference15 = new A.LineReference(){ Index = (UInt32Value)2U };
            A.RgbColorModelPercentage rgbColorModelPercentage43 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            lineReference15.Append(rgbColorModelPercentage43);

            A.FillReference fillReference15 = new A.FillReference(){ Index = (UInt32Value)1U };
            A.RgbColorModelPercentage rgbColorModelPercentage44 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            fillReference15.Append(rgbColorModelPercentage44);

            A.EffectReference effectReference15 = new A.EffectReference(){ Index = (UInt32Value)0U };
            A.RgbColorModelPercentage rgbColorModelPercentage45 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            effectReference15.Append(rgbColorModelPercentage45);
            A.FontReference fontReference15 = new A.FontReference(){ Index = A.FontCollectionIndexValues.Minor };

            style43.Append(lineReference15);
            style43.Append(fillReference15);
            style43.Append(effectReference15);
            style43.Append(fontReference15);

            styleLabel10.Append(scene3D11);
            styleLabel10.Append(shape3D10);
            styleLabel10.Append(textProperties10);
            styleLabel10.Append(style43);

            Dgm.StyleLabel styleLabel11 = new Dgm.StyleLabel(){ Name = "bgImgPlace1" };

            Dgm.Scene3D scene3D12 = new Dgm.Scene3D();
            A.Camera camera12 = new A.Camera(){ Preset = A.PresetCameraValues.OrthographicFront };
            A.LightRig lightRig12 = new A.LightRig(){ Rig = A.LightRigValues.ThreePoints, Direction = A.LightRigDirectionValues.Top };

            scene3D12.Append(camera12);
            scene3D12.Append(lightRig12);
            Dgm.Shape3D shape3D11 = new Dgm.Shape3D();
            Dgm.TextProperties textProperties11 = new Dgm.TextProperties();

            Dgm.Style style44 = new Dgm.Style();

            A.LineReference lineReference16 = new A.LineReference(){ Index = (UInt32Value)2U };
            A.RgbColorModelPercentage rgbColorModelPercentage46 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            lineReference16.Append(rgbColorModelPercentage46);

            A.FillReference fillReference16 = new A.FillReference(){ Index = (UInt32Value)1U };
            A.RgbColorModelPercentage rgbColorModelPercentage47 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            fillReference16.Append(rgbColorModelPercentage47);

            A.EffectReference effectReference16 = new A.EffectReference(){ Index = (UInt32Value)0U };
            A.RgbColorModelPercentage rgbColorModelPercentage48 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            effectReference16.Append(rgbColorModelPercentage48);
            A.FontReference fontReference16 = new A.FontReference(){ Index = A.FontCollectionIndexValues.Minor };

            style44.Append(lineReference16);
            style44.Append(fillReference16);
            style44.Append(effectReference16);
            style44.Append(fontReference16);

            styleLabel11.Append(scene3D12);
            styleLabel11.Append(shape3D11);
            styleLabel11.Append(textProperties11);
            styleLabel11.Append(style44);

            Dgm.StyleLabel styleLabel12 = new Dgm.StyleLabel(){ Name = "sibTrans2D1" };

            Dgm.Scene3D scene3D13 = new Dgm.Scene3D();
            A.Camera camera13 = new A.Camera(){ Preset = A.PresetCameraValues.OrthographicFront };
            A.LightRig lightRig13 = new A.LightRig(){ Rig = A.LightRigValues.ThreePoints, Direction = A.LightRigDirectionValues.Top };

            scene3D13.Append(camera13);
            scene3D13.Append(lightRig13);
            Dgm.Shape3D shape3D12 = new Dgm.Shape3D();
            Dgm.TextProperties textProperties12 = new Dgm.TextProperties();

            Dgm.Style style45 = new Dgm.Style();

            A.LineReference lineReference17 = new A.LineReference(){ Index = (UInt32Value)0U };
            A.RgbColorModelPercentage rgbColorModelPercentage49 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            lineReference17.Append(rgbColorModelPercentage49);

            A.FillReference fillReference17 = new A.FillReference(){ Index = (UInt32Value)1U };
            A.RgbColorModelPercentage rgbColorModelPercentage50 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            fillReference17.Append(rgbColorModelPercentage50);

            A.EffectReference effectReference17 = new A.EffectReference(){ Index = (UInt32Value)0U };
            A.RgbColorModelPercentage rgbColorModelPercentage51 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            effectReference17.Append(rgbColorModelPercentage51);

            A.FontReference fontReference17 = new A.FontReference(){ Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor155 = new A.SchemeColor(){ Val = A.SchemeColorValues.Light1 };

            fontReference17.Append(schemeColor155);

            style45.Append(lineReference17);
            style45.Append(fillReference17);
            style45.Append(effectReference17);
            style45.Append(fontReference17);

            styleLabel12.Append(scene3D13);
            styleLabel12.Append(shape3D12);
            styleLabel12.Append(textProperties12);
            styleLabel12.Append(style45);

            Dgm.StyleLabel styleLabel13 = new Dgm.StyleLabel(){ Name = "fgSibTrans2D1" };

            Dgm.Scene3D scene3D14 = new Dgm.Scene3D();
            A.Camera camera14 = new A.Camera(){ Preset = A.PresetCameraValues.OrthographicFront };
            A.LightRig lightRig14 = new A.LightRig(){ Rig = A.LightRigValues.ThreePoints, Direction = A.LightRigDirectionValues.Top };

            scene3D14.Append(camera14);
            scene3D14.Append(lightRig14);
            Dgm.Shape3D shape3D13 = new Dgm.Shape3D();
            Dgm.TextProperties textProperties13 = new Dgm.TextProperties();

            Dgm.Style style46 = new Dgm.Style();

            A.LineReference lineReference18 = new A.LineReference(){ Index = (UInt32Value)0U };
            A.RgbColorModelPercentage rgbColorModelPercentage52 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            lineReference18.Append(rgbColorModelPercentage52);

            A.FillReference fillReference18 = new A.FillReference(){ Index = (UInt32Value)1U };
            A.RgbColorModelPercentage rgbColorModelPercentage53 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            fillReference18.Append(rgbColorModelPercentage53);

            A.EffectReference effectReference18 = new A.EffectReference(){ Index = (UInt32Value)0U };
            A.RgbColorModelPercentage rgbColorModelPercentage54 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            effectReference18.Append(rgbColorModelPercentage54);

            A.FontReference fontReference18 = new A.FontReference(){ Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor156 = new A.SchemeColor(){ Val = A.SchemeColorValues.Light1 };

            fontReference18.Append(schemeColor156);

            style46.Append(lineReference18);
            style46.Append(fillReference18);
            style46.Append(effectReference18);
            style46.Append(fontReference18);

            styleLabel13.Append(scene3D14);
            styleLabel13.Append(shape3D13);
            styleLabel13.Append(textProperties13);
            styleLabel13.Append(style46);

            Dgm.StyleLabel styleLabel14 = new Dgm.StyleLabel(){ Name = "bgSibTrans2D1" };

            Dgm.Scene3D scene3D15 = new Dgm.Scene3D();
            A.Camera camera15 = new A.Camera(){ Preset = A.PresetCameraValues.OrthographicFront };
            A.LightRig lightRig15 = new A.LightRig(){ Rig = A.LightRigValues.ThreePoints, Direction = A.LightRigDirectionValues.Top };

            scene3D15.Append(camera15);
            scene3D15.Append(lightRig15);
            Dgm.Shape3D shape3D14 = new Dgm.Shape3D();
            Dgm.TextProperties textProperties14 = new Dgm.TextProperties();

            Dgm.Style style47 = new Dgm.Style();

            A.LineReference lineReference19 = new A.LineReference(){ Index = (UInt32Value)0U };
            A.RgbColorModelPercentage rgbColorModelPercentage55 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            lineReference19.Append(rgbColorModelPercentage55);

            A.FillReference fillReference19 = new A.FillReference(){ Index = (UInt32Value)1U };
            A.RgbColorModelPercentage rgbColorModelPercentage56 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            fillReference19.Append(rgbColorModelPercentage56);

            A.EffectReference effectReference19 = new A.EffectReference(){ Index = (UInt32Value)0U };
            A.RgbColorModelPercentage rgbColorModelPercentage57 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            effectReference19.Append(rgbColorModelPercentage57);

            A.FontReference fontReference19 = new A.FontReference(){ Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor157 = new A.SchemeColor(){ Val = A.SchemeColorValues.Light1 };

            fontReference19.Append(schemeColor157);

            style47.Append(lineReference19);
            style47.Append(fillReference19);
            style47.Append(effectReference19);
            style47.Append(fontReference19);

            styleLabel14.Append(scene3D15);
            styleLabel14.Append(shape3D14);
            styleLabel14.Append(textProperties14);
            styleLabel14.Append(style47);

            Dgm.StyleLabel styleLabel15 = new Dgm.StyleLabel(){ Name = "sibTrans1D1" };

            Dgm.Scene3D scene3D16 = new Dgm.Scene3D();
            A.Camera camera16 = new A.Camera(){ Preset = A.PresetCameraValues.OrthographicFront };
            A.LightRig lightRig16 = new A.LightRig(){ Rig = A.LightRigValues.ThreePoints, Direction = A.LightRigDirectionValues.Top };

            scene3D16.Append(camera16);
            scene3D16.Append(lightRig16);
            Dgm.Shape3D shape3D15 = new Dgm.Shape3D();
            Dgm.TextProperties textProperties15 = new Dgm.TextProperties();

            Dgm.Style style48 = new Dgm.Style();

            A.LineReference lineReference20 = new A.LineReference(){ Index = (UInt32Value)1U };
            A.RgbColorModelPercentage rgbColorModelPercentage58 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            lineReference20.Append(rgbColorModelPercentage58);

            A.FillReference fillReference20 = new A.FillReference(){ Index = (UInt32Value)0U };
            A.RgbColorModelPercentage rgbColorModelPercentage59 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            fillReference20.Append(rgbColorModelPercentage59);

            A.EffectReference effectReference20 = new A.EffectReference(){ Index = (UInt32Value)0U };
            A.RgbColorModelPercentage rgbColorModelPercentage60 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            effectReference20.Append(rgbColorModelPercentage60);
            A.FontReference fontReference20 = new A.FontReference(){ Index = A.FontCollectionIndexValues.Minor };

            style48.Append(lineReference20);
            style48.Append(fillReference20);
            style48.Append(effectReference20);
            style48.Append(fontReference20);

            styleLabel15.Append(scene3D16);
            styleLabel15.Append(shape3D15);
            styleLabel15.Append(textProperties15);
            styleLabel15.Append(style48);

            Dgm.StyleLabel styleLabel16 = new Dgm.StyleLabel(){ Name = "callout" };

            Dgm.Scene3D scene3D17 = new Dgm.Scene3D();
            A.Camera camera17 = new A.Camera(){ Preset = A.PresetCameraValues.OrthographicFront };
            A.LightRig lightRig17 = new A.LightRig(){ Rig = A.LightRigValues.ThreePoints, Direction = A.LightRigDirectionValues.Top };

            scene3D17.Append(camera17);
            scene3D17.Append(lightRig17);
            Dgm.Shape3D shape3D16 = new Dgm.Shape3D();
            Dgm.TextProperties textProperties16 = new Dgm.TextProperties();

            Dgm.Style style49 = new Dgm.Style();

            A.LineReference lineReference21 = new A.LineReference(){ Index = (UInt32Value)2U };
            A.RgbColorModelPercentage rgbColorModelPercentage61 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            lineReference21.Append(rgbColorModelPercentage61);

            A.FillReference fillReference21 = new A.FillReference(){ Index = (UInt32Value)1U };
            A.RgbColorModelPercentage rgbColorModelPercentage62 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            fillReference21.Append(rgbColorModelPercentage62);

            A.EffectReference effectReference21 = new A.EffectReference(){ Index = (UInt32Value)0U };
            A.RgbColorModelPercentage rgbColorModelPercentage63 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            effectReference21.Append(rgbColorModelPercentage63);
            A.FontReference fontReference21 = new A.FontReference(){ Index = A.FontCollectionIndexValues.Minor };

            style49.Append(lineReference21);
            style49.Append(fillReference21);
            style49.Append(effectReference21);
            style49.Append(fontReference21);

            styleLabel16.Append(scene3D17);
            styleLabel16.Append(shape3D16);
            styleLabel16.Append(textProperties16);
            styleLabel16.Append(style49);

            Dgm.StyleLabel styleLabel17 = new Dgm.StyleLabel(){ Name = "asst0" };

            Dgm.Scene3D scene3D18 = new Dgm.Scene3D();
            A.Camera camera18 = new A.Camera(){ Preset = A.PresetCameraValues.OrthographicFront };
            A.LightRig lightRig18 = new A.LightRig(){ Rig = A.LightRigValues.ThreePoints, Direction = A.LightRigDirectionValues.Top };

            scene3D18.Append(camera18);
            scene3D18.Append(lightRig18);
            Dgm.Shape3D shape3D17 = new Dgm.Shape3D();
            Dgm.TextProperties textProperties17 = new Dgm.TextProperties();

            Dgm.Style style50 = new Dgm.Style();

            A.LineReference lineReference22 = new A.LineReference(){ Index = (UInt32Value)2U };
            A.RgbColorModelPercentage rgbColorModelPercentage64 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            lineReference22.Append(rgbColorModelPercentage64);

            A.FillReference fillReference22 = new A.FillReference(){ Index = (UInt32Value)1U };
            A.RgbColorModelPercentage rgbColorModelPercentage65 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            fillReference22.Append(rgbColorModelPercentage65);

            A.EffectReference effectReference22 = new A.EffectReference(){ Index = (UInt32Value)0U };
            A.RgbColorModelPercentage rgbColorModelPercentage66 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            effectReference22.Append(rgbColorModelPercentage66);

            A.FontReference fontReference22 = new A.FontReference(){ Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor158 = new A.SchemeColor(){ Val = A.SchemeColorValues.Light1 };

            fontReference22.Append(schemeColor158);

            style50.Append(lineReference22);
            style50.Append(fillReference22);
            style50.Append(effectReference22);
            style50.Append(fontReference22);

            styleLabel17.Append(scene3D18);
            styleLabel17.Append(shape3D17);
            styleLabel17.Append(textProperties17);
            styleLabel17.Append(style50);

            Dgm.StyleLabel styleLabel18 = new Dgm.StyleLabel(){ Name = "asst1" };

            Dgm.Scene3D scene3D19 = new Dgm.Scene3D();
            A.Camera camera19 = new A.Camera(){ Preset = A.PresetCameraValues.OrthographicFront };
            A.LightRig lightRig19 = new A.LightRig(){ Rig = A.LightRigValues.ThreePoints, Direction = A.LightRigDirectionValues.Top };

            scene3D19.Append(camera19);
            scene3D19.Append(lightRig19);
            Dgm.Shape3D shape3D18 = new Dgm.Shape3D();
            Dgm.TextProperties textProperties18 = new Dgm.TextProperties();

            Dgm.Style style51 = new Dgm.Style();

            A.LineReference lineReference23 = new A.LineReference(){ Index = (UInt32Value)2U };
            A.RgbColorModelPercentage rgbColorModelPercentage67 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            lineReference23.Append(rgbColorModelPercentage67);

            A.FillReference fillReference23 = new A.FillReference(){ Index = (UInt32Value)1U };
            A.RgbColorModelPercentage rgbColorModelPercentage68 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            fillReference23.Append(rgbColorModelPercentage68);

            A.EffectReference effectReference23 = new A.EffectReference(){ Index = (UInt32Value)0U };
            A.RgbColorModelPercentage rgbColorModelPercentage69 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            effectReference23.Append(rgbColorModelPercentage69);

            A.FontReference fontReference23 = new A.FontReference(){ Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor159 = new A.SchemeColor(){ Val = A.SchemeColorValues.Light1 };

            fontReference23.Append(schemeColor159);

            style51.Append(lineReference23);
            style51.Append(fillReference23);
            style51.Append(effectReference23);
            style51.Append(fontReference23);

            styleLabel18.Append(scene3D19);
            styleLabel18.Append(shape3D18);
            styleLabel18.Append(textProperties18);
            styleLabel18.Append(style51);

            Dgm.StyleLabel styleLabel19 = new Dgm.StyleLabel(){ Name = "asst2" };

            Dgm.Scene3D scene3D20 = new Dgm.Scene3D();
            A.Camera camera20 = new A.Camera(){ Preset = A.PresetCameraValues.OrthographicFront };
            A.LightRig lightRig20 = new A.LightRig(){ Rig = A.LightRigValues.ThreePoints, Direction = A.LightRigDirectionValues.Top };

            scene3D20.Append(camera20);
            scene3D20.Append(lightRig20);
            Dgm.Shape3D shape3D19 = new Dgm.Shape3D();
            Dgm.TextProperties textProperties19 = new Dgm.TextProperties();

            Dgm.Style style52 = new Dgm.Style();

            A.LineReference lineReference24 = new A.LineReference(){ Index = (UInt32Value)2U };
            A.RgbColorModelPercentage rgbColorModelPercentage70 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            lineReference24.Append(rgbColorModelPercentage70);

            A.FillReference fillReference24 = new A.FillReference(){ Index = (UInt32Value)1U };
            A.RgbColorModelPercentage rgbColorModelPercentage71 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            fillReference24.Append(rgbColorModelPercentage71);

            A.EffectReference effectReference24 = new A.EffectReference(){ Index = (UInt32Value)0U };
            A.RgbColorModelPercentage rgbColorModelPercentage72 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            effectReference24.Append(rgbColorModelPercentage72);

            A.FontReference fontReference24 = new A.FontReference(){ Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor160 = new A.SchemeColor(){ Val = A.SchemeColorValues.Light1 };

            fontReference24.Append(schemeColor160);

            style52.Append(lineReference24);
            style52.Append(fillReference24);
            style52.Append(effectReference24);
            style52.Append(fontReference24);

            styleLabel19.Append(scene3D20);
            styleLabel19.Append(shape3D19);
            styleLabel19.Append(textProperties19);
            styleLabel19.Append(style52);

            Dgm.StyleLabel styleLabel20 = new Dgm.StyleLabel(){ Name = "asst3" };

            Dgm.Scene3D scene3D21 = new Dgm.Scene3D();
            A.Camera camera21 = new A.Camera(){ Preset = A.PresetCameraValues.OrthographicFront };
            A.LightRig lightRig21 = new A.LightRig(){ Rig = A.LightRigValues.ThreePoints, Direction = A.LightRigDirectionValues.Top };

            scene3D21.Append(camera21);
            scene3D21.Append(lightRig21);
            Dgm.Shape3D shape3D20 = new Dgm.Shape3D();
            Dgm.TextProperties textProperties20 = new Dgm.TextProperties();

            Dgm.Style style53 = new Dgm.Style();

            A.LineReference lineReference25 = new A.LineReference(){ Index = (UInt32Value)2U };
            A.RgbColorModelPercentage rgbColorModelPercentage73 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            lineReference25.Append(rgbColorModelPercentage73);

            A.FillReference fillReference25 = new A.FillReference(){ Index = (UInt32Value)1U };
            A.RgbColorModelPercentage rgbColorModelPercentage74 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            fillReference25.Append(rgbColorModelPercentage74);

            A.EffectReference effectReference25 = new A.EffectReference(){ Index = (UInt32Value)0U };
            A.RgbColorModelPercentage rgbColorModelPercentage75 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            effectReference25.Append(rgbColorModelPercentage75);

            A.FontReference fontReference25 = new A.FontReference(){ Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor161 = new A.SchemeColor(){ Val = A.SchemeColorValues.Light1 };

            fontReference25.Append(schemeColor161);

            style53.Append(lineReference25);
            style53.Append(fillReference25);
            style53.Append(effectReference25);
            style53.Append(fontReference25);

            styleLabel20.Append(scene3D21);
            styleLabel20.Append(shape3D20);
            styleLabel20.Append(textProperties20);
            styleLabel20.Append(style53);

            Dgm.StyleLabel styleLabel21 = new Dgm.StyleLabel(){ Name = "asst4" };

            Dgm.Scene3D scene3D22 = new Dgm.Scene3D();
            A.Camera camera22 = new A.Camera(){ Preset = A.PresetCameraValues.OrthographicFront };
            A.LightRig lightRig22 = new A.LightRig(){ Rig = A.LightRigValues.ThreePoints, Direction = A.LightRigDirectionValues.Top };

            scene3D22.Append(camera22);
            scene3D22.Append(lightRig22);
            Dgm.Shape3D shape3D21 = new Dgm.Shape3D();
            Dgm.TextProperties textProperties21 = new Dgm.TextProperties();

            Dgm.Style style54 = new Dgm.Style();

            A.LineReference lineReference26 = new A.LineReference(){ Index = (UInt32Value)2U };
            A.RgbColorModelPercentage rgbColorModelPercentage76 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            lineReference26.Append(rgbColorModelPercentage76);

            A.FillReference fillReference26 = new A.FillReference(){ Index = (UInt32Value)1U };
            A.RgbColorModelPercentage rgbColorModelPercentage77 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            fillReference26.Append(rgbColorModelPercentage77);

            A.EffectReference effectReference26 = new A.EffectReference(){ Index = (UInt32Value)0U };
            A.RgbColorModelPercentage rgbColorModelPercentage78 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            effectReference26.Append(rgbColorModelPercentage78);

            A.FontReference fontReference26 = new A.FontReference(){ Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor162 = new A.SchemeColor(){ Val = A.SchemeColorValues.Light1 };

            fontReference26.Append(schemeColor162);

            style54.Append(lineReference26);
            style54.Append(fillReference26);
            style54.Append(effectReference26);
            style54.Append(fontReference26);

            styleLabel21.Append(scene3D22);
            styleLabel21.Append(shape3D21);
            styleLabel21.Append(textProperties21);
            styleLabel21.Append(style54);

            Dgm.StyleLabel styleLabel22 = new Dgm.StyleLabel(){ Name = "parChTrans2D1" };

            Dgm.Scene3D scene3D23 = new Dgm.Scene3D();
            A.Camera camera23 = new A.Camera(){ Preset = A.PresetCameraValues.OrthographicFront };
            A.LightRig lightRig23 = new A.LightRig(){ Rig = A.LightRigValues.ThreePoints, Direction = A.LightRigDirectionValues.Top };

            scene3D23.Append(camera23);
            scene3D23.Append(lightRig23);
            Dgm.Shape3D shape3D22 = new Dgm.Shape3D();
            Dgm.TextProperties textProperties22 = new Dgm.TextProperties();

            Dgm.Style style55 = new Dgm.Style();

            A.LineReference lineReference27 = new A.LineReference(){ Index = (UInt32Value)2U };
            A.RgbColorModelPercentage rgbColorModelPercentage79 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            lineReference27.Append(rgbColorModelPercentage79);

            A.FillReference fillReference27 = new A.FillReference(){ Index = (UInt32Value)1U };
            A.RgbColorModelPercentage rgbColorModelPercentage80 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            fillReference27.Append(rgbColorModelPercentage80);

            A.EffectReference effectReference27 = new A.EffectReference(){ Index = (UInt32Value)0U };
            A.RgbColorModelPercentage rgbColorModelPercentage81 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            effectReference27.Append(rgbColorModelPercentage81);

            A.FontReference fontReference27 = new A.FontReference(){ Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor163 = new A.SchemeColor(){ Val = A.SchemeColorValues.Light1 };

            fontReference27.Append(schemeColor163);

            style55.Append(lineReference27);
            style55.Append(fillReference27);
            style55.Append(effectReference27);
            style55.Append(fontReference27);

            styleLabel22.Append(scene3D23);
            styleLabel22.Append(shape3D22);
            styleLabel22.Append(textProperties22);
            styleLabel22.Append(style55);

            Dgm.StyleLabel styleLabel23 = new Dgm.StyleLabel(){ Name = "parChTrans2D2" };

            Dgm.Scene3D scene3D24 = new Dgm.Scene3D();
            A.Camera camera24 = new A.Camera(){ Preset = A.PresetCameraValues.OrthographicFront };
            A.LightRig lightRig24 = new A.LightRig(){ Rig = A.LightRigValues.ThreePoints, Direction = A.LightRigDirectionValues.Top };

            scene3D24.Append(camera24);
            scene3D24.Append(lightRig24);
            Dgm.Shape3D shape3D23 = new Dgm.Shape3D();
            Dgm.TextProperties textProperties23 = new Dgm.TextProperties();

            Dgm.Style style56 = new Dgm.Style();

            A.LineReference lineReference28 = new A.LineReference(){ Index = (UInt32Value)2U };
            A.RgbColorModelPercentage rgbColorModelPercentage82 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            lineReference28.Append(rgbColorModelPercentage82);

            A.FillReference fillReference28 = new A.FillReference(){ Index = (UInt32Value)1U };
            A.RgbColorModelPercentage rgbColorModelPercentage83 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            fillReference28.Append(rgbColorModelPercentage83);

            A.EffectReference effectReference28 = new A.EffectReference(){ Index = (UInt32Value)0U };
            A.RgbColorModelPercentage rgbColorModelPercentage84 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            effectReference28.Append(rgbColorModelPercentage84);

            A.FontReference fontReference28 = new A.FontReference(){ Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor164 = new A.SchemeColor(){ Val = A.SchemeColorValues.Light1 };

            fontReference28.Append(schemeColor164);

            style56.Append(lineReference28);
            style56.Append(fillReference28);
            style56.Append(effectReference28);
            style56.Append(fontReference28);

            styleLabel23.Append(scene3D24);
            styleLabel23.Append(shape3D23);
            styleLabel23.Append(textProperties23);
            styleLabel23.Append(style56);

            Dgm.StyleLabel styleLabel24 = new Dgm.StyleLabel(){ Name = "parChTrans2D3" };

            Dgm.Scene3D scene3D25 = new Dgm.Scene3D();
            A.Camera camera25 = new A.Camera(){ Preset = A.PresetCameraValues.OrthographicFront };
            A.LightRig lightRig25 = new A.LightRig(){ Rig = A.LightRigValues.ThreePoints, Direction = A.LightRigDirectionValues.Top };

            scene3D25.Append(camera25);
            scene3D25.Append(lightRig25);
            Dgm.Shape3D shape3D24 = new Dgm.Shape3D();
            Dgm.TextProperties textProperties24 = new Dgm.TextProperties();

            Dgm.Style style57 = new Dgm.Style();

            A.LineReference lineReference29 = new A.LineReference(){ Index = (UInt32Value)2U };
            A.RgbColorModelPercentage rgbColorModelPercentage85 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            lineReference29.Append(rgbColorModelPercentage85);

            A.FillReference fillReference29 = new A.FillReference(){ Index = (UInt32Value)1U };
            A.RgbColorModelPercentage rgbColorModelPercentage86 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            fillReference29.Append(rgbColorModelPercentage86);

            A.EffectReference effectReference29 = new A.EffectReference(){ Index = (UInt32Value)0U };
            A.RgbColorModelPercentage rgbColorModelPercentage87 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            effectReference29.Append(rgbColorModelPercentage87);

            A.FontReference fontReference29 = new A.FontReference(){ Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor165 = new A.SchemeColor(){ Val = A.SchemeColorValues.Light1 };

            fontReference29.Append(schemeColor165);

            style57.Append(lineReference29);
            style57.Append(fillReference29);
            style57.Append(effectReference29);
            style57.Append(fontReference29);

            styleLabel24.Append(scene3D25);
            styleLabel24.Append(shape3D24);
            styleLabel24.Append(textProperties24);
            styleLabel24.Append(style57);

            Dgm.StyleLabel styleLabel25 = new Dgm.StyleLabel(){ Name = "parChTrans2D4" };

            Dgm.Scene3D scene3D26 = new Dgm.Scene3D();
            A.Camera camera26 = new A.Camera(){ Preset = A.PresetCameraValues.OrthographicFront };
            A.LightRig lightRig26 = new A.LightRig(){ Rig = A.LightRigValues.ThreePoints, Direction = A.LightRigDirectionValues.Top };

            scene3D26.Append(camera26);
            scene3D26.Append(lightRig26);
            Dgm.Shape3D shape3D25 = new Dgm.Shape3D();
            Dgm.TextProperties textProperties25 = new Dgm.TextProperties();

            Dgm.Style style58 = new Dgm.Style();

            A.LineReference lineReference30 = new A.LineReference(){ Index = (UInt32Value)2U };
            A.RgbColorModelPercentage rgbColorModelPercentage88 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            lineReference30.Append(rgbColorModelPercentage88);

            A.FillReference fillReference30 = new A.FillReference(){ Index = (UInt32Value)1U };
            A.RgbColorModelPercentage rgbColorModelPercentage89 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            fillReference30.Append(rgbColorModelPercentage89);

            A.EffectReference effectReference30 = new A.EffectReference(){ Index = (UInt32Value)0U };
            A.RgbColorModelPercentage rgbColorModelPercentage90 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            effectReference30.Append(rgbColorModelPercentage90);

            A.FontReference fontReference30 = new A.FontReference(){ Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor166 = new A.SchemeColor(){ Val = A.SchemeColorValues.Light1 };

            fontReference30.Append(schemeColor166);

            style58.Append(lineReference30);
            style58.Append(fillReference30);
            style58.Append(effectReference30);
            style58.Append(fontReference30);

            styleLabel25.Append(scene3D26);
            styleLabel25.Append(shape3D25);
            styleLabel25.Append(textProperties25);
            styleLabel25.Append(style58);

            Dgm.StyleLabel styleLabel26 = new Dgm.StyleLabel(){ Name = "parChTrans1D1" };

            Dgm.Scene3D scene3D27 = new Dgm.Scene3D();
            A.Camera camera27 = new A.Camera(){ Preset = A.PresetCameraValues.OrthographicFront };
            A.LightRig lightRig27 = new A.LightRig(){ Rig = A.LightRigValues.ThreePoints, Direction = A.LightRigDirectionValues.Top };

            scene3D27.Append(camera27);
            scene3D27.Append(lightRig27);
            Dgm.Shape3D shape3D26 = new Dgm.Shape3D();
            Dgm.TextProperties textProperties26 = new Dgm.TextProperties();

            Dgm.Style style59 = new Dgm.Style();

            A.LineReference lineReference31 = new A.LineReference(){ Index = (UInt32Value)2U };
            A.RgbColorModelPercentage rgbColorModelPercentage91 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            lineReference31.Append(rgbColorModelPercentage91);

            A.FillReference fillReference31 = new A.FillReference(){ Index = (UInt32Value)0U };
            A.RgbColorModelPercentage rgbColorModelPercentage92 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            fillReference31.Append(rgbColorModelPercentage92);

            A.EffectReference effectReference31 = new A.EffectReference(){ Index = (UInt32Value)0U };
            A.RgbColorModelPercentage rgbColorModelPercentage93 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            effectReference31.Append(rgbColorModelPercentage93);
            A.FontReference fontReference31 = new A.FontReference(){ Index = A.FontCollectionIndexValues.Minor };

            style59.Append(lineReference31);
            style59.Append(fillReference31);
            style59.Append(effectReference31);
            style59.Append(fontReference31);

            styleLabel26.Append(scene3D27);
            styleLabel26.Append(shape3D26);
            styleLabel26.Append(textProperties26);
            styleLabel26.Append(style59);

            Dgm.StyleLabel styleLabel27 = new Dgm.StyleLabel(){ Name = "parChTrans1D2" };

            Dgm.Scene3D scene3D28 = new Dgm.Scene3D();
            A.Camera camera28 = new A.Camera(){ Preset = A.PresetCameraValues.OrthographicFront };
            A.LightRig lightRig28 = new A.LightRig(){ Rig = A.LightRigValues.ThreePoints, Direction = A.LightRigDirectionValues.Top };

            scene3D28.Append(camera28);
            scene3D28.Append(lightRig28);
            Dgm.Shape3D shape3D27 = new Dgm.Shape3D();
            Dgm.TextProperties textProperties27 = new Dgm.TextProperties();

            Dgm.Style style60 = new Dgm.Style();

            A.LineReference lineReference32 = new A.LineReference(){ Index = (UInt32Value)2U };
            A.RgbColorModelPercentage rgbColorModelPercentage94 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            lineReference32.Append(rgbColorModelPercentage94);

            A.FillReference fillReference32 = new A.FillReference(){ Index = (UInt32Value)0U };
            A.RgbColorModelPercentage rgbColorModelPercentage95 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            fillReference32.Append(rgbColorModelPercentage95);

            A.EffectReference effectReference32 = new A.EffectReference(){ Index = (UInt32Value)0U };
            A.RgbColorModelPercentage rgbColorModelPercentage96 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            effectReference32.Append(rgbColorModelPercentage96);
            A.FontReference fontReference32 = new A.FontReference(){ Index = A.FontCollectionIndexValues.Minor };

            style60.Append(lineReference32);
            style60.Append(fillReference32);
            style60.Append(effectReference32);
            style60.Append(fontReference32);

            styleLabel27.Append(scene3D28);
            styleLabel27.Append(shape3D27);
            styleLabel27.Append(textProperties27);
            styleLabel27.Append(style60);

            Dgm.StyleLabel styleLabel28 = new Dgm.StyleLabel(){ Name = "parChTrans1D3" };

            Dgm.Scene3D scene3D29 = new Dgm.Scene3D();
            A.Camera camera29 = new A.Camera(){ Preset = A.PresetCameraValues.OrthographicFront };
            A.LightRig lightRig29 = new A.LightRig(){ Rig = A.LightRigValues.ThreePoints, Direction = A.LightRigDirectionValues.Top };

            scene3D29.Append(camera29);
            scene3D29.Append(lightRig29);
            Dgm.Shape3D shape3D28 = new Dgm.Shape3D();
            Dgm.TextProperties textProperties28 = new Dgm.TextProperties();

            Dgm.Style style61 = new Dgm.Style();

            A.LineReference lineReference33 = new A.LineReference(){ Index = (UInt32Value)2U };
            A.RgbColorModelPercentage rgbColorModelPercentage97 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            lineReference33.Append(rgbColorModelPercentage97);

            A.FillReference fillReference33 = new A.FillReference(){ Index = (UInt32Value)0U };
            A.RgbColorModelPercentage rgbColorModelPercentage98 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            fillReference33.Append(rgbColorModelPercentage98);

            A.EffectReference effectReference33 = new A.EffectReference(){ Index = (UInt32Value)0U };
            A.RgbColorModelPercentage rgbColorModelPercentage99 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            effectReference33.Append(rgbColorModelPercentage99);
            A.FontReference fontReference33 = new A.FontReference(){ Index = A.FontCollectionIndexValues.Minor };

            style61.Append(lineReference33);
            style61.Append(fillReference33);
            style61.Append(effectReference33);
            style61.Append(fontReference33);

            styleLabel28.Append(scene3D29);
            styleLabel28.Append(shape3D28);
            styleLabel28.Append(textProperties28);
            styleLabel28.Append(style61);

            Dgm.StyleLabel styleLabel29 = new Dgm.StyleLabel(){ Name = "parChTrans1D4" };

            Dgm.Scene3D scene3D30 = new Dgm.Scene3D();
            A.Camera camera30 = new A.Camera(){ Preset = A.PresetCameraValues.OrthographicFront };
            A.LightRig lightRig30 = new A.LightRig(){ Rig = A.LightRigValues.ThreePoints, Direction = A.LightRigDirectionValues.Top };

            scene3D30.Append(camera30);
            scene3D30.Append(lightRig30);
            Dgm.Shape3D shape3D29 = new Dgm.Shape3D();
            Dgm.TextProperties textProperties29 = new Dgm.TextProperties();

            Dgm.Style style62 = new Dgm.Style();

            A.LineReference lineReference34 = new A.LineReference(){ Index = (UInt32Value)2U };
            A.RgbColorModelPercentage rgbColorModelPercentage100 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            lineReference34.Append(rgbColorModelPercentage100);

            A.FillReference fillReference34 = new A.FillReference(){ Index = (UInt32Value)0U };
            A.RgbColorModelPercentage rgbColorModelPercentage101 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            fillReference34.Append(rgbColorModelPercentage101);

            A.EffectReference effectReference34 = new A.EffectReference(){ Index = (UInt32Value)0U };
            A.RgbColorModelPercentage rgbColorModelPercentage102 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            effectReference34.Append(rgbColorModelPercentage102);
            A.FontReference fontReference34 = new A.FontReference(){ Index = A.FontCollectionIndexValues.Minor };

            style62.Append(lineReference34);
            style62.Append(fillReference34);
            style62.Append(effectReference34);
            style62.Append(fontReference34);

            styleLabel29.Append(scene3D30);
            styleLabel29.Append(shape3D29);
            styleLabel29.Append(textProperties29);
            styleLabel29.Append(style62);

            Dgm.StyleLabel styleLabel30 = new Dgm.StyleLabel(){ Name = "fgAcc1" };

            Dgm.Scene3D scene3D31 = new Dgm.Scene3D();
            A.Camera camera31 = new A.Camera(){ Preset = A.PresetCameraValues.OrthographicFront };
            A.LightRig lightRig31 = new A.LightRig(){ Rig = A.LightRigValues.ThreePoints, Direction = A.LightRigDirectionValues.Top };

            scene3D31.Append(camera31);
            scene3D31.Append(lightRig31);
            Dgm.Shape3D shape3D30 = new Dgm.Shape3D();
            Dgm.TextProperties textProperties30 = new Dgm.TextProperties();

            Dgm.Style style63 = new Dgm.Style();

            A.LineReference lineReference35 = new A.LineReference(){ Index = (UInt32Value)2U };
            A.RgbColorModelPercentage rgbColorModelPercentage103 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            lineReference35.Append(rgbColorModelPercentage103);

            A.FillReference fillReference35 = new A.FillReference(){ Index = (UInt32Value)1U };
            A.RgbColorModelPercentage rgbColorModelPercentage104 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            fillReference35.Append(rgbColorModelPercentage104);

            A.EffectReference effectReference35 = new A.EffectReference(){ Index = (UInt32Value)0U };
            A.RgbColorModelPercentage rgbColorModelPercentage105 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            effectReference35.Append(rgbColorModelPercentage105);
            A.FontReference fontReference35 = new A.FontReference(){ Index = A.FontCollectionIndexValues.Minor };

            style63.Append(lineReference35);
            style63.Append(fillReference35);
            style63.Append(effectReference35);
            style63.Append(fontReference35);

            styleLabel30.Append(scene3D31);
            styleLabel30.Append(shape3D30);
            styleLabel30.Append(textProperties30);
            styleLabel30.Append(style63);

            Dgm.StyleLabel styleLabel31 = new Dgm.StyleLabel(){ Name = "conFgAcc1" };

            Dgm.Scene3D scene3D32 = new Dgm.Scene3D();
            A.Camera camera32 = new A.Camera(){ Preset = A.PresetCameraValues.OrthographicFront };
            A.LightRig lightRig32 = new A.LightRig(){ Rig = A.LightRigValues.ThreePoints, Direction = A.LightRigDirectionValues.Top };

            scene3D32.Append(camera32);
            scene3D32.Append(lightRig32);
            Dgm.Shape3D shape3D31 = new Dgm.Shape3D();
            Dgm.TextProperties textProperties31 = new Dgm.TextProperties();

            Dgm.Style style64 = new Dgm.Style();

            A.LineReference lineReference36 = new A.LineReference(){ Index = (UInt32Value)2U };
            A.RgbColorModelPercentage rgbColorModelPercentage106 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            lineReference36.Append(rgbColorModelPercentage106);

            A.FillReference fillReference36 = new A.FillReference(){ Index = (UInt32Value)1U };
            A.RgbColorModelPercentage rgbColorModelPercentage107 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            fillReference36.Append(rgbColorModelPercentage107);

            A.EffectReference effectReference36 = new A.EffectReference(){ Index = (UInt32Value)0U };
            A.RgbColorModelPercentage rgbColorModelPercentage108 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            effectReference36.Append(rgbColorModelPercentage108);
            A.FontReference fontReference36 = new A.FontReference(){ Index = A.FontCollectionIndexValues.Minor };

            style64.Append(lineReference36);
            style64.Append(fillReference36);
            style64.Append(effectReference36);
            style64.Append(fontReference36);

            styleLabel31.Append(scene3D32);
            styleLabel31.Append(shape3D31);
            styleLabel31.Append(textProperties31);
            styleLabel31.Append(style64);

            Dgm.StyleLabel styleLabel32 = new Dgm.StyleLabel(){ Name = "alignAcc1" };

            Dgm.Scene3D scene3D33 = new Dgm.Scene3D();
            A.Camera camera33 = new A.Camera(){ Preset = A.PresetCameraValues.OrthographicFront };
            A.LightRig lightRig33 = new A.LightRig(){ Rig = A.LightRigValues.ThreePoints, Direction = A.LightRigDirectionValues.Top };

            scene3D33.Append(camera33);
            scene3D33.Append(lightRig33);
            Dgm.Shape3D shape3D32 = new Dgm.Shape3D();
            Dgm.TextProperties textProperties32 = new Dgm.TextProperties();

            Dgm.Style style65 = new Dgm.Style();

            A.LineReference lineReference37 = new A.LineReference(){ Index = (UInt32Value)2U };
            A.RgbColorModelPercentage rgbColorModelPercentage109 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            lineReference37.Append(rgbColorModelPercentage109);

            A.FillReference fillReference37 = new A.FillReference(){ Index = (UInt32Value)1U };
            A.RgbColorModelPercentage rgbColorModelPercentage110 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            fillReference37.Append(rgbColorModelPercentage110);

            A.EffectReference effectReference37 = new A.EffectReference(){ Index = (UInt32Value)0U };
            A.RgbColorModelPercentage rgbColorModelPercentage111 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            effectReference37.Append(rgbColorModelPercentage111);
            A.FontReference fontReference37 = new A.FontReference(){ Index = A.FontCollectionIndexValues.Minor };

            style65.Append(lineReference37);
            style65.Append(fillReference37);
            style65.Append(effectReference37);
            style65.Append(fontReference37);

            styleLabel32.Append(scene3D33);
            styleLabel32.Append(shape3D32);
            styleLabel32.Append(textProperties32);
            styleLabel32.Append(style65);

            Dgm.StyleLabel styleLabel33 = new Dgm.StyleLabel(){ Name = "trAlignAcc1" };

            Dgm.Scene3D scene3D34 = new Dgm.Scene3D();
            A.Camera camera34 = new A.Camera(){ Preset = A.PresetCameraValues.OrthographicFront };
            A.LightRig lightRig34 = new A.LightRig(){ Rig = A.LightRigValues.ThreePoints, Direction = A.LightRigDirectionValues.Top };

            scene3D34.Append(camera34);
            scene3D34.Append(lightRig34);
            Dgm.Shape3D shape3D33 = new Dgm.Shape3D();
            Dgm.TextProperties textProperties33 = new Dgm.TextProperties();

            Dgm.Style style66 = new Dgm.Style();

            A.LineReference lineReference38 = new A.LineReference(){ Index = (UInt32Value)1U };
            A.RgbColorModelPercentage rgbColorModelPercentage112 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            lineReference38.Append(rgbColorModelPercentage112);

            A.FillReference fillReference38 = new A.FillReference(){ Index = (UInt32Value)1U };
            A.RgbColorModelPercentage rgbColorModelPercentage113 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            fillReference38.Append(rgbColorModelPercentage113);

            A.EffectReference effectReference38 = new A.EffectReference(){ Index = (UInt32Value)0U };
            A.RgbColorModelPercentage rgbColorModelPercentage114 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            effectReference38.Append(rgbColorModelPercentage114);
            A.FontReference fontReference38 = new A.FontReference(){ Index = A.FontCollectionIndexValues.Minor };

            style66.Append(lineReference38);
            style66.Append(fillReference38);
            style66.Append(effectReference38);
            style66.Append(fontReference38);

            styleLabel33.Append(scene3D34);
            styleLabel33.Append(shape3D33);
            styleLabel33.Append(textProperties33);
            styleLabel33.Append(style66);

            Dgm.StyleLabel styleLabel34 = new Dgm.StyleLabel(){ Name = "bgAcc1" };

            Dgm.Scene3D scene3D35 = new Dgm.Scene3D();
            A.Camera camera35 = new A.Camera(){ Preset = A.PresetCameraValues.OrthographicFront };
            A.LightRig lightRig35 = new A.LightRig(){ Rig = A.LightRigValues.ThreePoints, Direction = A.LightRigDirectionValues.Top };

            scene3D35.Append(camera35);
            scene3D35.Append(lightRig35);
            Dgm.Shape3D shape3D34 = new Dgm.Shape3D();
            Dgm.TextProperties textProperties34 = new Dgm.TextProperties();

            Dgm.Style style67 = new Dgm.Style();

            A.LineReference lineReference39 = new A.LineReference(){ Index = (UInt32Value)2U };
            A.RgbColorModelPercentage rgbColorModelPercentage115 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            lineReference39.Append(rgbColorModelPercentage115);

            A.FillReference fillReference39 = new A.FillReference(){ Index = (UInt32Value)1U };
            A.RgbColorModelPercentage rgbColorModelPercentage116 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            fillReference39.Append(rgbColorModelPercentage116);

            A.EffectReference effectReference39 = new A.EffectReference(){ Index = (UInt32Value)0U };
            A.RgbColorModelPercentage rgbColorModelPercentage117 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            effectReference39.Append(rgbColorModelPercentage117);
            A.FontReference fontReference39 = new A.FontReference(){ Index = A.FontCollectionIndexValues.Minor };

            style67.Append(lineReference39);
            style67.Append(fillReference39);
            style67.Append(effectReference39);
            style67.Append(fontReference39);

            styleLabel34.Append(scene3D35);
            styleLabel34.Append(shape3D34);
            styleLabel34.Append(textProperties34);
            styleLabel34.Append(style67);

            Dgm.StyleLabel styleLabel35 = new Dgm.StyleLabel(){ Name = "solidFgAcc1" };

            Dgm.Scene3D scene3D36 = new Dgm.Scene3D();
            A.Camera camera36 = new A.Camera(){ Preset = A.PresetCameraValues.OrthographicFront };
            A.LightRig lightRig36 = new A.LightRig(){ Rig = A.LightRigValues.ThreePoints, Direction = A.LightRigDirectionValues.Top };

            scene3D36.Append(camera36);
            scene3D36.Append(lightRig36);
            Dgm.Shape3D shape3D35 = new Dgm.Shape3D();
            Dgm.TextProperties textProperties35 = new Dgm.TextProperties();

            Dgm.Style style68 = new Dgm.Style();

            A.LineReference lineReference40 = new A.LineReference(){ Index = (UInt32Value)2U };
            A.RgbColorModelPercentage rgbColorModelPercentage118 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            lineReference40.Append(rgbColorModelPercentage118);

            A.FillReference fillReference40 = new A.FillReference(){ Index = (UInt32Value)1U };
            A.RgbColorModelPercentage rgbColorModelPercentage119 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            fillReference40.Append(rgbColorModelPercentage119);

            A.EffectReference effectReference40 = new A.EffectReference(){ Index = (UInt32Value)0U };
            A.RgbColorModelPercentage rgbColorModelPercentage120 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            effectReference40.Append(rgbColorModelPercentage120);
            A.FontReference fontReference40 = new A.FontReference(){ Index = A.FontCollectionIndexValues.Minor };

            style68.Append(lineReference40);
            style68.Append(fillReference40);
            style68.Append(effectReference40);
            style68.Append(fontReference40);

            styleLabel35.Append(scene3D36);
            styleLabel35.Append(shape3D35);
            styleLabel35.Append(textProperties35);
            styleLabel35.Append(style68);

            Dgm.StyleLabel styleLabel36 = new Dgm.StyleLabel(){ Name = "solidAlignAcc1" };

            Dgm.Scene3D scene3D37 = new Dgm.Scene3D();
            A.Camera camera37 = new A.Camera(){ Preset = A.PresetCameraValues.OrthographicFront };
            A.LightRig lightRig37 = new A.LightRig(){ Rig = A.LightRigValues.ThreePoints, Direction = A.LightRigDirectionValues.Top };

            scene3D37.Append(camera37);
            scene3D37.Append(lightRig37);
            Dgm.Shape3D shape3D36 = new Dgm.Shape3D();
            Dgm.TextProperties textProperties36 = new Dgm.TextProperties();

            Dgm.Style style69 = new Dgm.Style();

            A.LineReference lineReference41 = new A.LineReference(){ Index = (UInt32Value)2U };
            A.RgbColorModelPercentage rgbColorModelPercentage121 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            lineReference41.Append(rgbColorModelPercentage121);

            A.FillReference fillReference41 = new A.FillReference(){ Index = (UInt32Value)1U };
            A.RgbColorModelPercentage rgbColorModelPercentage122 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            fillReference41.Append(rgbColorModelPercentage122);

            A.EffectReference effectReference41 = new A.EffectReference(){ Index = (UInt32Value)0U };
            A.RgbColorModelPercentage rgbColorModelPercentage123 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            effectReference41.Append(rgbColorModelPercentage123);
            A.FontReference fontReference41 = new A.FontReference(){ Index = A.FontCollectionIndexValues.Minor };

            style69.Append(lineReference41);
            style69.Append(fillReference41);
            style69.Append(effectReference41);
            style69.Append(fontReference41);

            styleLabel36.Append(scene3D37);
            styleLabel36.Append(shape3D36);
            styleLabel36.Append(textProperties36);
            styleLabel36.Append(style69);

            Dgm.StyleLabel styleLabel37 = new Dgm.StyleLabel(){ Name = "solidBgAcc1" };

            Dgm.Scene3D scene3D38 = new Dgm.Scene3D();
            A.Camera camera38 = new A.Camera(){ Preset = A.PresetCameraValues.OrthographicFront };
            A.LightRig lightRig38 = new A.LightRig(){ Rig = A.LightRigValues.ThreePoints, Direction = A.LightRigDirectionValues.Top };

            scene3D38.Append(camera38);
            scene3D38.Append(lightRig38);
            Dgm.Shape3D shape3D37 = new Dgm.Shape3D();
            Dgm.TextProperties textProperties37 = new Dgm.TextProperties();

            Dgm.Style style70 = new Dgm.Style();

            A.LineReference lineReference42 = new A.LineReference(){ Index = (UInt32Value)2U };
            A.RgbColorModelPercentage rgbColorModelPercentage124 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            lineReference42.Append(rgbColorModelPercentage124);

            A.FillReference fillReference42 = new A.FillReference(){ Index = (UInt32Value)1U };
            A.RgbColorModelPercentage rgbColorModelPercentage125 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            fillReference42.Append(rgbColorModelPercentage125);

            A.EffectReference effectReference42 = new A.EffectReference(){ Index = (UInt32Value)0U };
            A.RgbColorModelPercentage rgbColorModelPercentage126 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            effectReference42.Append(rgbColorModelPercentage126);
            A.FontReference fontReference42 = new A.FontReference(){ Index = A.FontCollectionIndexValues.Minor };

            style70.Append(lineReference42);
            style70.Append(fillReference42);
            style70.Append(effectReference42);
            style70.Append(fontReference42);

            styleLabel37.Append(scene3D38);
            styleLabel37.Append(shape3D37);
            styleLabel37.Append(textProperties37);
            styleLabel37.Append(style70);

            Dgm.StyleLabel styleLabel38 = new Dgm.StyleLabel(){ Name = "fgAccFollowNode1" };

            Dgm.Scene3D scene3D39 = new Dgm.Scene3D();
            A.Camera camera39 = new A.Camera(){ Preset = A.PresetCameraValues.OrthographicFront };
            A.LightRig lightRig39 = new A.LightRig(){ Rig = A.LightRigValues.ThreePoints, Direction = A.LightRigDirectionValues.Top };

            scene3D39.Append(camera39);
            scene3D39.Append(lightRig39);
            Dgm.Shape3D shape3D38 = new Dgm.Shape3D();
            Dgm.TextProperties textProperties38 = new Dgm.TextProperties();

            Dgm.Style style71 = new Dgm.Style();

            A.LineReference lineReference43 = new A.LineReference(){ Index = (UInt32Value)2U };
            A.RgbColorModelPercentage rgbColorModelPercentage127 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            lineReference43.Append(rgbColorModelPercentage127);

            A.FillReference fillReference43 = new A.FillReference(){ Index = (UInt32Value)1U };
            A.RgbColorModelPercentage rgbColorModelPercentage128 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            fillReference43.Append(rgbColorModelPercentage128);

            A.EffectReference effectReference43 = new A.EffectReference(){ Index = (UInt32Value)0U };
            A.RgbColorModelPercentage rgbColorModelPercentage129 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            effectReference43.Append(rgbColorModelPercentage129);
            A.FontReference fontReference43 = new A.FontReference(){ Index = A.FontCollectionIndexValues.Minor };

            style71.Append(lineReference43);
            style71.Append(fillReference43);
            style71.Append(effectReference43);
            style71.Append(fontReference43);

            styleLabel38.Append(scene3D39);
            styleLabel38.Append(shape3D38);
            styleLabel38.Append(textProperties38);
            styleLabel38.Append(style71);

            Dgm.StyleLabel styleLabel39 = new Dgm.StyleLabel(){ Name = "alignAccFollowNode1" };

            Dgm.Scene3D scene3D40 = new Dgm.Scene3D();
            A.Camera camera40 = new A.Camera(){ Preset = A.PresetCameraValues.OrthographicFront };
            A.LightRig lightRig40 = new A.LightRig(){ Rig = A.LightRigValues.ThreePoints, Direction = A.LightRigDirectionValues.Top };

            scene3D40.Append(camera40);
            scene3D40.Append(lightRig40);
            Dgm.Shape3D shape3D39 = new Dgm.Shape3D();
            Dgm.TextProperties textProperties39 = new Dgm.TextProperties();

            Dgm.Style style72 = new Dgm.Style();

            A.LineReference lineReference44 = new A.LineReference(){ Index = (UInt32Value)2U };
            A.RgbColorModelPercentage rgbColorModelPercentage130 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            lineReference44.Append(rgbColorModelPercentage130);

            A.FillReference fillReference44 = new A.FillReference(){ Index = (UInt32Value)1U };
            A.RgbColorModelPercentage rgbColorModelPercentage131 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            fillReference44.Append(rgbColorModelPercentage131);

            A.EffectReference effectReference44 = new A.EffectReference(){ Index = (UInt32Value)0U };
            A.RgbColorModelPercentage rgbColorModelPercentage132 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            effectReference44.Append(rgbColorModelPercentage132);
            A.FontReference fontReference44 = new A.FontReference(){ Index = A.FontCollectionIndexValues.Minor };

            style72.Append(lineReference44);
            style72.Append(fillReference44);
            style72.Append(effectReference44);
            style72.Append(fontReference44);

            styleLabel39.Append(scene3D40);
            styleLabel39.Append(shape3D39);
            styleLabel39.Append(textProperties39);
            styleLabel39.Append(style72);

            Dgm.StyleLabel styleLabel40 = new Dgm.StyleLabel(){ Name = "bgAccFollowNode1" };

            Dgm.Scene3D scene3D41 = new Dgm.Scene3D();
            A.Camera camera41 = new A.Camera(){ Preset = A.PresetCameraValues.OrthographicFront };
            A.LightRig lightRig41 = new A.LightRig(){ Rig = A.LightRigValues.ThreePoints, Direction = A.LightRigDirectionValues.Top };

            scene3D41.Append(camera41);
            scene3D41.Append(lightRig41);
            Dgm.Shape3D shape3D40 = new Dgm.Shape3D();
            Dgm.TextProperties textProperties40 = new Dgm.TextProperties();

            Dgm.Style style73 = new Dgm.Style();

            A.LineReference lineReference45 = new A.LineReference(){ Index = (UInt32Value)2U };
            A.RgbColorModelPercentage rgbColorModelPercentage133 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            lineReference45.Append(rgbColorModelPercentage133);

            A.FillReference fillReference45 = new A.FillReference(){ Index = (UInt32Value)1U };
            A.RgbColorModelPercentage rgbColorModelPercentage134 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            fillReference45.Append(rgbColorModelPercentage134);

            A.EffectReference effectReference45 = new A.EffectReference(){ Index = (UInt32Value)0U };
            A.RgbColorModelPercentage rgbColorModelPercentage135 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            effectReference45.Append(rgbColorModelPercentage135);
            A.FontReference fontReference45 = new A.FontReference(){ Index = A.FontCollectionIndexValues.Minor };

            style73.Append(lineReference45);
            style73.Append(fillReference45);
            style73.Append(effectReference45);
            style73.Append(fontReference45);

            styleLabel40.Append(scene3D41);
            styleLabel40.Append(shape3D40);
            styleLabel40.Append(textProperties40);
            styleLabel40.Append(style73);

            Dgm.StyleLabel styleLabel41 = new Dgm.StyleLabel(){ Name = "fgAcc0" };

            Dgm.Scene3D scene3D42 = new Dgm.Scene3D();
            A.Camera camera42 = new A.Camera(){ Preset = A.PresetCameraValues.OrthographicFront };
            A.LightRig lightRig42 = new A.LightRig(){ Rig = A.LightRigValues.ThreePoints, Direction = A.LightRigDirectionValues.Top };

            scene3D42.Append(camera42);
            scene3D42.Append(lightRig42);
            Dgm.Shape3D shape3D41 = new Dgm.Shape3D();
            Dgm.TextProperties textProperties41 = new Dgm.TextProperties();

            Dgm.Style style74 = new Dgm.Style();

            A.LineReference lineReference46 = new A.LineReference(){ Index = (UInt32Value)2U };
            A.RgbColorModelPercentage rgbColorModelPercentage136 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            lineReference46.Append(rgbColorModelPercentage136);

            A.FillReference fillReference46 = new A.FillReference(){ Index = (UInt32Value)1U };
            A.RgbColorModelPercentage rgbColorModelPercentage137 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            fillReference46.Append(rgbColorModelPercentage137);

            A.EffectReference effectReference46 = new A.EffectReference(){ Index = (UInt32Value)0U };
            A.RgbColorModelPercentage rgbColorModelPercentage138 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            effectReference46.Append(rgbColorModelPercentage138);
            A.FontReference fontReference46 = new A.FontReference(){ Index = A.FontCollectionIndexValues.Minor };

            style74.Append(lineReference46);
            style74.Append(fillReference46);
            style74.Append(effectReference46);
            style74.Append(fontReference46);

            styleLabel41.Append(scene3D42);
            styleLabel41.Append(shape3D41);
            styleLabel41.Append(textProperties41);
            styleLabel41.Append(style74);

            Dgm.StyleLabel styleLabel42 = new Dgm.StyleLabel(){ Name = "fgAcc2" };

            Dgm.Scene3D scene3D43 = new Dgm.Scene3D();
            A.Camera camera43 = new A.Camera(){ Preset = A.PresetCameraValues.OrthographicFront };
            A.LightRig lightRig43 = new A.LightRig(){ Rig = A.LightRigValues.ThreePoints, Direction = A.LightRigDirectionValues.Top };

            scene3D43.Append(camera43);
            scene3D43.Append(lightRig43);
            Dgm.Shape3D shape3D42 = new Dgm.Shape3D();
            Dgm.TextProperties textProperties42 = new Dgm.TextProperties();

            Dgm.Style style75 = new Dgm.Style();

            A.LineReference lineReference47 = new A.LineReference(){ Index = (UInt32Value)2U };
            A.RgbColorModelPercentage rgbColorModelPercentage139 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            lineReference47.Append(rgbColorModelPercentage139);

            A.FillReference fillReference47 = new A.FillReference(){ Index = (UInt32Value)1U };
            A.RgbColorModelPercentage rgbColorModelPercentage140 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            fillReference47.Append(rgbColorModelPercentage140);

            A.EffectReference effectReference47 = new A.EffectReference(){ Index = (UInt32Value)0U };
            A.RgbColorModelPercentage rgbColorModelPercentage141 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            effectReference47.Append(rgbColorModelPercentage141);
            A.FontReference fontReference47 = new A.FontReference(){ Index = A.FontCollectionIndexValues.Minor };

            style75.Append(lineReference47);
            style75.Append(fillReference47);
            style75.Append(effectReference47);
            style75.Append(fontReference47);

            styleLabel42.Append(scene3D43);
            styleLabel42.Append(shape3D42);
            styleLabel42.Append(textProperties42);
            styleLabel42.Append(style75);

            Dgm.StyleLabel styleLabel43 = new Dgm.StyleLabel(){ Name = "fgAcc3" };

            Dgm.Scene3D scene3D44 = new Dgm.Scene3D();
            A.Camera camera44 = new A.Camera(){ Preset = A.PresetCameraValues.OrthographicFront };
            A.LightRig lightRig44 = new A.LightRig(){ Rig = A.LightRigValues.ThreePoints, Direction = A.LightRigDirectionValues.Top };

            scene3D44.Append(camera44);
            scene3D44.Append(lightRig44);
            Dgm.Shape3D shape3D43 = new Dgm.Shape3D();
            Dgm.TextProperties textProperties43 = new Dgm.TextProperties();

            Dgm.Style style76 = new Dgm.Style();

            A.LineReference lineReference48 = new A.LineReference(){ Index = (UInt32Value)2U };
            A.RgbColorModelPercentage rgbColorModelPercentage142 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            lineReference48.Append(rgbColorModelPercentage142);

            A.FillReference fillReference48 = new A.FillReference(){ Index = (UInt32Value)1U };
            A.RgbColorModelPercentage rgbColorModelPercentage143 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            fillReference48.Append(rgbColorModelPercentage143);

            A.EffectReference effectReference48 = new A.EffectReference(){ Index = (UInt32Value)0U };
            A.RgbColorModelPercentage rgbColorModelPercentage144 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            effectReference48.Append(rgbColorModelPercentage144);
            A.FontReference fontReference48 = new A.FontReference(){ Index = A.FontCollectionIndexValues.Minor };

            style76.Append(lineReference48);
            style76.Append(fillReference48);
            style76.Append(effectReference48);
            style76.Append(fontReference48);

            styleLabel43.Append(scene3D44);
            styleLabel43.Append(shape3D43);
            styleLabel43.Append(textProperties43);
            styleLabel43.Append(style76);

            Dgm.StyleLabel styleLabel44 = new Dgm.StyleLabel(){ Name = "fgAcc4" };

            Dgm.Scene3D scene3D45 = new Dgm.Scene3D();
            A.Camera camera45 = new A.Camera(){ Preset = A.PresetCameraValues.OrthographicFront };
            A.LightRig lightRig45 = new A.LightRig(){ Rig = A.LightRigValues.ThreePoints, Direction = A.LightRigDirectionValues.Top };

            scene3D45.Append(camera45);
            scene3D45.Append(lightRig45);
            Dgm.Shape3D shape3D44 = new Dgm.Shape3D();
            Dgm.TextProperties textProperties44 = new Dgm.TextProperties();

            Dgm.Style style77 = new Dgm.Style();

            A.LineReference lineReference49 = new A.LineReference(){ Index = (UInt32Value)2U };
            A.RgbColorModelPercentage rgbColorModelPercentage145 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            lineReference49.Append(rgbColorModelPercentage145);

            A.FillReference fillReference49 = new A.FillReference(){ Index = (UInt32Value)1U };
            A.RgbColorModelPercentage rgbColorModelPercentage146 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            fillReference49.Append(rgbColorModelPercentage146);

            A.EffectReference effectReference49 = new A.EffectReference(){ Index = (UInt32Value)0U };
            A.RgbColorModelPercentage rgbColorModelPercentage147 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            effectReference49.Append(rgbColorModelPercentage147);
            A.FontReference fontReference49 = new A.FontReference(){ Index = A.FontCollectionIndexValues.Minor };

            style77.Append(lineReference49);
            style77.Append(fillReference49);
            style77.Append(effectReference49);
            style77.Append(fontReference49);

            styleLabel44.Append(scene3D45);
            styleLabel44.Append(shape3D44);
            styleLabel44.Append(textProperties44);
            styleLabel44.Append(style77);

            Dgm.StyleLabel styleLabel45 = new Dgm.StyleLabel(){ Name = "bgShp" };

            Dgm.Scene3D scene3D46 = new Dgm.Scene3D();
            A.Camera camera46 = new A.Camera(){ Preset = A.PresetCameraValues.OrthographicFront };
            A.LightRig lightRig46 = new A.LightRig(){ Rig = A.LightRigValues.ThreePoints, Direction = A.LightRigDirectionValues.Top };

            scene3D46.Append(camera46);
            scene3D46.Append(lightRig46);
            Dgm.Shape3D shape3D45 = new Dgm.Shape3D();
            Dgm.TextProperties textProperties45 = new Dgm.TextProperties();

            Dgm.Style style78 = new Dgm.Style();

            A.LineReference lineReference50 = new A.LineReference(){ Index = (UInt32Value)0U };
            A.RgbColorModelPercentage rgbColorModelPercentage148 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            lineReference50.Append(rgbColorModelPercentage148);

            A.FillReference fillReference50 = new A.FillReference(){ Index = (UInt32Value)1U };
            A.RgbColorModelPercentage rgbColorModelPercentage149 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            fillReference50.Append(rgbColorModelPercentage149);

            A.EffectReference effectReference50 = new A.EffectReference(){ Index = (UInt32Value)0U };
            A.RgbColorModelPercentage rgbColorModelPercentage150 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            effectReference50.Append(rgbColorModelPercentage150);
            A.FontReference fontReference50 = new A.FontReference(){ Index = A.FontCollectionIndexValues.Minor };

            style78.Append(lineReference50);
            style78.Append(fillReference50);
            style78.Append(effectReference50);
            style78.Append(fontReference50);

            styleLabel45.Append(scene3D46);
            styleLabel45.Append(shape3D45);
            styleLabel45.Append(textProperties45);
            styleLabel45.Append(style78);

            Dgm.StyleLabel styleLabel46 = new Dgm.StyleLabel(){ Name = "dkBgShp" };

            Dgm.Scene3D scene3D47 = new Dgm.Scene3D();
            A.Camera camera47 = new A.Camera(){ Preset = A.PresetCameraValues.OrthographicFront };
            A.LightRig lightRig47 = new A.LightRig(){ Rig = A.LightRigValues.ThreePoints, Direction = A.LightRigDirectionValues.Top };

            scene3D47.Append(camera47);
            scene3D47.Append(lightRig47);
            Dgm.Shape3D shape3D46 = new Dgm.Shape3D();
            Dgm.TextProperties textProperties46 = new Dgm.TextProperties();

            Dgm.Style style79 = new Dgm.Style();

            A.LineReference lineReference51 = new A.LineReference(){ Index = (UInt32Value)0U };
            A.RgbColorModelPercentage rgbColorModelPercentage151 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            lineReference51.Append(rgbColorModelPercentage151);

            A.FillReference fillReference51 = new A.FillReference(){ Index = (UInt32Value)1U };
            A.RgbColorModelPercentage rgbColorModelPercentage152 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            fillReference51.Append(rgbColorModelPercentage152);

            A.EffectReference effectReference51 = new A.EffectReference(){ Index = (UInt32Value)0U };
            A.RgbColorModelPercentage rgbColorModelPercentage153 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            effectReference51.Append(rgbColorModelPercentage153);
            A.FontReference fontReference51 = new A.FontReference(){ Index = A.FontCollectionIndexValues.Minor };

            style79.Append(lineReference51);
            style79.Append(fillReference51);
            style79.Append(effectReference51);
            style79.Append(fontReference51);

            styleLabel46.Append(scene3D47);
            styleLabel46.Append(shape3D46);
            styleLabel46.Append(textProperties46);
            styleLabel46.Append(style79);

            Dgm.StyleLabel styleLabel47 = new Dgm.StyleLabel(){ Name = "trBgShp" };

            Dgm.Scene3D scene3D48 = new Dgm.Scene3D();
            A.Camera camera48 = new A.Camera(){ Preset = A.PresetCameraValues.OrthographicFront };
            A.LightRig lightRig48 = new A.LightRig(){ Rig = A.LightRigValues.ThreePoints, Direction = A.LightRigDirectionValues.Top };

            scene3D48.Append(camera48);
            scene3D48.Append(lightRig48);
            Dgm.Shape3D shape3D47 = new Dgm.Shape3D();
            Dgm.TextProperties textProperties47 = new Dgm.TextProperties();

            Dgm.Style style80 = new Dgm.Style();

            A.LineReference lineReference52 = new A.LineReference(){ Index = (UInt32Value)0U };
            A.RgbColorModelPercentage rgbColorModelPercentage154 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            lineReference52.Append(rgbColorModelPercentage154);

            A.FillReference fillReference52 = new A.FillReference(){ Index = (UInt32Value)1U };
            A.RgbColorModelPercentage rgbColorModelPercentage155 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            fillReference52.Append(rgbColorModelPercentage155);

            A.EffectReference effectReference52 = new A.EffectReference(){ Index = (UInt32Value)0U };
            A.RgbColorModelPercentage rgbColorModelPercentage156 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            effectReference52.Append(rgbColorModelPercentage156);
            A.FontReference fontReference52 = new A.FontReference(){ Index = A.FontCollectionIndexValues.Minor };

            style80.Append(lineReference52);
            style80.Append(fillReference52);
            style80.Append(effectReference52);
            style80.Append(fontReference52);

            styleLabel47.Append(scene3D48);
            styleLabel47.Append(shape3D47);
            styleLabel47.Append(textProperties47);
            styleLabel47.Append(style80);

            Dgm.StyleLabel styleLabel48 = new Dgm.StyleLabel(){ Name = "fgShp" };

            Dgm.Scene3D scene3D49 = new Dgm.Scene3D();
            A.Camera camera49 = new A.Camera(){ Preset = A.PresetCameraValues.OrthographicFront };
            A.LightRig lightRig49 = new A.LightRig(){ Rig = A.LightRigValues.ThreePoints, Direction = A.LightRigDirectionValues.Top };

            scene3D49.Append(camera49);
            scene3D49.Append(lightRig49);
            Dgm.Shape3D shape3D48 = new Dgm.Shape3D();
            Dgm.TextProperties textProperties48 = new Dgm.TextProperties();

            Dgm.Style style81 = new Dgm.Style();

            A.LineReference lineReference53 = new A.LineReference(){ Index = (UInt32Value)2U };
            A.RgbColorModelPercentage rgbColorModelPercentage157 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            lineReference53.Append(rgbColorModelPercentage157);

            A.FillReference fillReference53 = new A.FillReference(){ Index = (UInt32Value)1U };
            A.RgbColorModelPercentage rgbColorModelPercentage158 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            fillReference53.Append(rgbColorModelPercentage158);

            A.EffectReference effectReference53 = new A.EffectReference(){ Index = (UInt32Value)0U };
            A.RgbColorModelPercentage rgbColorModelPercentage159 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            effectReference53.Append(rgbColorModelPercentage159);
            A.FontReference fontReference53 = new A.FontReference(){ Index = A.FontCollectionIndexValues.Minor };

            style81.Append(lineReference53);
            style81.Append(fillReference53);
            style81.Append(effectReference53);
            style81.Append(fontReference53);

            styleLabel48.Append(scene3D49);
            styleLabel48.Append(shape3D48);
            styleLabel48.Append(textProperties48);
            styleLabel48.Append(style81);

            Dgm.StyleLabel styleLabel49 = new Dgm.StyleLabel(){ Name = "revTx" };

            Dgm.Scene3D scene3D50 = new Dgm.Scene3D();
            A.Camera camera50 = new A.Camera(){ Preset = A.PresetCameraValues.OrthographicFront };
            A.LightRig lightRig50 = new A.LightRig(){ Rig = A.LightRigValues.ThreePoints, Direction = A.LightRigDirectionValues.Top };

            scene3D50.Append(camera50);
            scene3D50.Append(lightRig50);
            Dgm.Shape3D shape3D49 = new Dgm.Shape3D();
            Dgm.TextProperties textProperties49 = new Dgm.TextProperties();

            Dgm.Style style82 = new Dgm.Style();

            A.LineReference lineReference54 = new A.LineReference(){ Index = (UInt32Value)0U };
            A.RgbColorModelPercentage rgbColorModelPercentage160 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            lineReference54.Append(rgbColorModelPercentage160);

            A.FillReference fillReference54 = new A.FillReference(){ Index = (UInt32Value)0U };
            A.RgbColorModelPercentage rgbColorModelPercentage161 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            fillReference54.Append(rgbColorModelPercentage161);

            A.EffectReference effectReference54 = new A.EffectReference(){ Index = (UInt32Value)0U };
            A.RgbColorModelPercentage rgbColorModelPercentage162 = new A.RgbColorModelPercentage(){ RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

            effectReference54.Append(rgbColorModelPercentage162);
            A.FontReference fontReference54 = new A.FontReference(){ Index = A.FontCollectionIndexValues.Minor };

            style82.Append(lineReference54);
            style82.Append(fillReference54);
            style82.Append(effectReference54);
            style82.Append(fontReference54);

            styleLabel49.Append(scene3D50);
            styleLabel49.Append(shape3D49);
            styleLabel49.Append(textProperties49);
            styleLabel49.Append(style82);

            styleDefinition1.Append(styleDefinitionTitle1);
            styleDefinition1.Append(styleLabelDescription1);
            styleDefinition1.Append(styleDisplayCategories1);
            styleDefinition1.Append(scene3D1);
            styleDefinition1.Append(styleLabel1);
            styleDefinition1.Append(styleLabel2);
            styleDefinition1.Append(styleLabel3);
            styleDefinition1.Append(styleLabel4);
            styleDefinition1.Append(styleLabel5);
            styleDefinition1.Append(styleLabel6);
            styleDefinition1.Append(styleLabel7);
            styleDefinition1.Append(styleLabel8);
            styleDefinition1.Append(styleLabel9);
            styleDefinition1.Append(styleLabel10);
            styleDefinition1.Append(styleLabel11);
            styleDefinition1.Append(styleLabel12);
            styleDefinition1.Append(styleLabel13);
            styleDefinition1.Append(styleLabel14);
            styleDefinition1.Append(styleLabel15);
            styleDefinition1.Append(styleLabel16);
            styleDefinition1.Append(styleLabel17);
            styleDefinition1.Append(styleLabel18);
            styleDefinition1.Append(styleLabel19);
            styleDefinition1.Append(styleLabel20);
            styleDefinition1.Append(styleLabel21);
            styleDefinition1.Append(styleLabel22);
            styleDefinition1.Append(styleLabel23);
            styleDefinition1.Append(styleLabel24);
            styleDefinition1.Append(styleLabel25);
            styleDefinition1.Append(styleLabel26);
            styleDefinition1.Append(styleLabel27);
            styleDefinition1.Append(styleLabel28);
            styleDefinition1.Append(styleLabel29);
            styleDefinition1.Append(styleLabel30);
            styleDefinition1.Append(styleLabel31);
            styleDefinition1.Append(styleLabel32);
            styleDefinition1.Append(styleLabel33);
            styleDefinition1.Append(styleLabel34);
            styleDefinition1.Append(styleLabel35);
            styleDefinition1.Append(styleLabel36);
            styleDefinition1.Append(styleLabel37);
            styleDefinition1.Append(styleLabel38);
            styleDefinition1.Append(styleLabel39);
            styleDefinition1.Append(styleLabel40);
            styleDefinition1.Append(styleLabel41);
            styleDefinition1.Append(styleLabel42);
            styleDefinition1.Append(styleLabel43);
            styleDefinition1.Append(styleLabel44);
            styleDefinition1.Append(styleLabel45);
            styleDefinition1.Append(styleLabel46);
            styleDefinition1.Append(styleLabel47);
            styleDefinition1.Append(styleLabel48);
            styleDefinition1.Append(styleLabel49);

            diagramStylePart1.StyleDefinition = styleDefinition1;
        }

        // Generates content of diagramLayoutDefinitionPart1.
        private void GenerateDiagramLayoutDefinitionPart1Content(DiagramLayoutDefinitionPart diagramLayoutDefinitionPart1)
        {
            Dgm.LayoutDefinition layoutDefinition1 = new Dgm.LayoutDefinition(){ UniqueId = "urn:microsoft.com/office/officeart/2005/8/layout/default" };
            layoutDefinition1.AddNamespaceDeclaration("dgm", "http://schemas.openxmlformats.org/drawingml/2006/diagram");
            layoutDefinition1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            Dgm.Title title1 = new Dgm.Title(){ Val = "" };
            Dgm.Description description1 = new Dgm.Description(){ Val = "" };

            Dgm.CategoryList categoryList1 = new Dgm.CategoryList();
            Dgm.Category category1 = new Dgm.Category(){ Type = "list", Priority = (UInt32Value)400U };

            categoryList1.Append(category1);

            Dgm.SampleData sampleData1 = new Dgm.SampleData();

            Dgm.DataModel dataModel1 = new Dgm.DataModel();

            Dgm.PointList pointList1 = new Dgm.PointList();
            Dgm.Point point1 = new Dgm.Point(){ ModelId = "0", Type = Dgm.PointValues.Document };

            Dgm.Point point2 = new Dgm.Point(){ ModelId = "1" };
            Dgm.PropertySet propertySet1 = new Dgm.PropertySet(){ Placeholder = true };

            point2.Append(propertySet1);

            Dgm.Point point3 = new Dgm.Point(){ ModelId = "2" };
            Dgm.PropertySet propertySet2 = new Dgm.PropertySet(){ Placeholder = true };

            point3.Append(propertySet2);

            Dgm.Point point4 = new Dgm.Point(){ ModelId = "3" };
            Dgm.PropertySet propertySet3 = new Dgm.PropertySet(){ Placeholder = true };

            point4.Append(propertySet3);

            Dgm.Point point5 = new Dgm.Point(){ ModelId = "4" };
            Dgm.PropertySet propertySet4 = new Dgm.PropertySet(){ Placeholder = true };

            point5.Append(propertySet4);

            Dgm.Point point6 = new Dgm.Point(){ ModelId = "5" };
            Dgm.PropertySet propertySet5 = new Dgm.PropertySet(){ Placeholder = true };

            point6.Append(propertySet5);

            pointList1.Append(point1);
            pointList1.Append(point2);
            pointList1.Append(point3);
            pointList1.Append(point4);
            pointList1.Append(point5);
            pointList1.Append(point6);

            Dgm.ConnectionList connectionList1 = new Dgm.ConnectionList();
            Dgm.Connection connection1 = new Dgm.Connection(){ ModelId = "6", SourceId = "0", DestinationId = "1", SourcePosition = (UInt32Value)0U, DestinationPosition = (UInt32Value)0U };
            Dgm.Connection connection2 = new Dgm.Connection(){ ModelId = "7", SourceId = "0", DestinationId = "2", SourcePosition = (UInt32Value)1U, DestinationPosition = (UInt32Value)0U };
            Dgm.Connection connection3 = new Dgm.Connection(){ ModelId = "8", SourceId = "0", DestinationId = "3", SourcePosition = (UInt32Value)2U, DestinationPosition = (UInt32Value)0U };
            Dgm.Connection connection4 = new Dgm.Connection(){ ModelId = "9", SourceId = "0", DestinationId = "4", SourcePosition = (UInt32Value)3U, DestinationPosition = (UInt32Value)0U };
            Dgm.Connection connection5 = new Dgm.Connection(){ ModelId = "10", SourceId = "0", DestinationId = "5", SourcePosition = (UInt32Value)4U, DestinationPosition = (UInt32Value)0U };

            connectionList1.Append(connection1);
            connectionList1.Append(connection2);
            connectionList1.Append(connection3);
            connectionList1.Append(connection4);
            connectionList1.Append(connection5);
            Dgm.Background background1 = new Dgm.Background();
            Dgm.Whole whole1 = new Dgm.Whole();

            dataModel1.Append(pointList1);
            dataModel1.Append(connectionList1);
            dataModel1.Append(background1);
            dataModel1.Append(whole1);

            sampleData1.Append(dataModel1);

            Dgm.StyleData styleData1 = new Dgm.StyleData();

            Dgm.DataModel dataModel2 = new Dgm.DataModel();

            Dgm.PointList pointList2 = new Dgm.PointList();
            Dgm.Point point7 = new Dgm.Point(){ ModelId = "0", Type = Dgm.PointValues.Document };
            Dgm.Point point8 = new Dgm.Point(){ ModelId = "1" };
            Dgm.Point point9 = new Dgm.Point(){ ModelId = "2" };

            pointList2.Append(point7);
            pointList2.Append(point8);
            pointList2.Append(point9);

            Dgm.ConnectionList connectionList2 = new Dgm.ConnectionList();
            Dgm.Connection connection6 = new Dgm.Connection(){ ModelId = "3", SourceId = "0", DestinationId = "1", SourcePosition = (UInt32Value)0U, DestinationPosition = (UInt32Value)0U };
            Dgm.Connection connection7 = new Dgm.Connection(){ ModelId = "4", SourceId = "0", DestinationId = "2", SourcePosition = (UInt32Value)1U, DestinationPosition = (UInt32Value)0U };

            connectionList2.Append(connection6);
            connectionList2.Append(connection7);
            Dgm.Background background2 = new Dgm.Background();
            Dgm.Whole whole2 = new Dgm.Whole();

            dataModel2.Append(pointList2);
            dataModel2.Append(connectionList2);
            dataModel2.Append(background2);
            dataModel2.Append(whole2);

            styleData1.Append(dataModel2);

            Dgm.ColorData colorData1 = new Dgm.ColorData();

            Dgm.DataModel dataModel3 = new Dgm.DataModel();

            Dgm.PointList pointList3 = new Dgm.PointList();
            Dgm.Point point10 = new Dgm.Point(){ ModelId = "0", Type = Dgm.PointValues.Document };
            Dgm.Point point11 = new Dgm.Point(){ ModelId = "1" };
            Dgm.Point point12 = new Dgm.Point(){ ModelId = "2" };
            Dgm.Point point13 = new Dgm.Point(){ ModelId = "3" };
            Dgm.Point point14 = new Dgm.Point(){ ModelId = "4" };
            Dgm.Point point15 = new Dgm.Point(){ ModelId = "5" };
            Dgm.Point point16 = new Dgm.Point(){ ModelId = "6" };

            pointList3.Append(point10);
            pointList3.Append(point11);
            pointList3.Append(point12);
            pointList3.Append(point13);
            pointList3.Append(point14);
            pointList3.Append(point15);
            pointList3.Append(point16);

            Dgm.ConnectionList connectionList3 = new Dgm.ConnectionList();
            Dgm.Connection connection8 = new Dgm.Connection(){ ModelId = "7", SourceId = "0", DestinationId = "1", SourcePosition = (UInt32Value)0U, DestinationPosition = (UInt32Value)0U };
            Dgm.Connection connection9 = new Dgm.Connection(){ ModelId = "8", SourceId = "0", DestinationId = "2", SourcePosition = (UInt32Value)1U, DestinationPosition = (UInt32Value)0U };
            Dgm.Connection connection10 = new Dgm.Connection(){ ModelId = "9", SourceId = "0", DestinationId = "3", SourcePosition = (UInt32Value)2U, DestinationPosition = (UInt32Value)0U };
            Dgm.Connection connection11 = new Dgm.Connection(){ ModelId = "10", SourceId = "0", DestinationId = "4", SourcePosition = (UInt32Value)3U, DestinationPosition = (UInt32Value)0U };
            Dgm.Connection connection12 = new Dgm.Connection(){ ModelId = "11", SourceId = "0", DestinationId = "5", SourcePosition = (UInt32Value)4U, DestinationPosition = (UInt32Value)0U };
            Dgm.Connection connection13 = new Dgm.Connection(){ ModelId = "12", SourceId = "0", DestinationId = "6", SourcePosition = (UInt32Value)5U, DestinationPosition = (UInt32Value)0U };

            connectionList3.Append(connection8);
            connectionList3.Append(connection9);
            connectionList3.Append(connection10);
            connectionList3.Append(connection11);
            connectionList3.Append(connection12);
            connectionList3.Append(connection13);
            Dgm.Background background3 = new Dgm.Background();
            Dgm.Whole whole3 = new Dgm.Whole();

            dataModel3.Append(pointList3);
            dataModel3.Append(connectionList3);
            dataModel3.Append(background3);
            dataModel3.Append(whole3);

            colorData1.Append(dataModel3);

            Dgm.LayoutNode layoutNode1 = new Dgm.LayoutNode(){ Name = "diagram" };

            Dgm.VariableList variableList1 = new Dgm.VariableList();
            Dgm.Direction direction1 = new Dgm.Direction();
            Dgm.ResizeHandles resizeHandles1 = new Dgm.ResizeHandles(){ Val = Dgm.ResizeHandlesStringValues.Exact };

            variableList1.Append(direction1);
            variableList1.Append(resizeHandles1);

            Dgm.Choose choose1 = new Dgm.Choose(){ Name = "Name0" };

            Dgm.DiagramChooseIf diagramChooseIf1 = new Dgm.DiagramChooseIf(){ Name = "Name1", Function = Dgm.FunctionValues.Variable, Argument = "dir", Operator = Dgm.FunctionOperatorValues.Equal, Val = "norm" };

            Dgm.Algorithm algorithm1 = new Dgm.Algorithm(){ Type = Dgm.AlgorithmValues.Snake };
            Dgm.Parameter parameter1 = new Dgm.Parameter(){ Type = Dgm.ParameterIdValues.GrowDirection, Val = "tL" };
            Dgm.Parameter parameter2 = new Dgm.Parameter(){ Type = Dgm.ParameterIdValues.FlowDirection, Val = "row" };
            Dgm.Parameter parameter3 = new Dgm.Parameter(){ Type = Dgm.ParameterIdValues.ContinueDirection, Val = "sameDir" };
            Dgm.Parameter parameter4 = new Dgm.Parameter(){ Type = Dgm.ParameterIdValues.Offset, Val = "ctr" };

            algorithm1.Append(parameter1);
            algorithm1.Append(parameter2);
            algorithm1.Append(parameter3);
            algorithm1.Append(parameter4);

            diagramChooseIf1.Append(algorithm1);

            Dgm.DiagramChooseElse diagramChooseElse1 = new Dgm.DiagramChooseElse(){ Name = "Name2" };

            Dgm.Algorithm algorithm2 = new Dgm.Algorithm(){ Type = Dgm.AlgorithmValues.Snake };
            Dgm.Parameter parameter5 = new Dgm.Parameter(){ Type = Dgm.ParameterIdValues.GrowDirection, Val = "tR" };
            Dgm.Parameter parameter6 = new Dgm.Parameter(){ Type = Dgm.ParameterIdValues.FlowDirection, Val = "row" };
            Dgm.Parameter parameter7 = new Dgm.Parameter(){ Type = Dgm.ParameterIdValues.ContinueDirection, Val = "sameDir" };
            Dgm.Parameter parameter8 = new Dgm.Parameter(){ Type = Dgm.ParameterIdValues.Offset, Val = "ctr" };

            algorithm2.Append(parameter5);
            algorithm2.Append(parameter6);
            algorithm2.Append(parameter7);
            algorithm2.Append(parameter8);

            diagramChooseElse1.Append(algorithm2);

            choose1.Append(diagramChooseIf1);
            choose1.Append(diagramChooseElse1);

            Dgm.Shape shape6 = new Dgm.Shape(){ Blip = "" };
            shape6.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            Dgm.AdjustList adjustList1 = new Dgm.AdjustList();

            shape6.Append(adjustList1);
            Dgm.PresentationOf presentationOf1 = new Dgm.PresentationOf();

            Dgm.Constraints constraints1 = new Dgm.Constraints();
            Dgm.Constraint constraint1 = new Dgm.Constraint(){ Type = Dgm.ConstraintValues.Width, For = Dgm.ConstraintRelationshipValues.Child, ForName = "node", ReferenceType = Dgm.ConstraintValues.Width };
            Dgm.Constraint constraint2 = new Dgm.Constraint(){ Type = Dgm.ConstraintValues.Height, For = Dgm.ConstraintRelationshipValues.Child, ForName = "node", ReferenceType = Dgm.ConstraintValues.Width, ReferenceFor = Dgm.ConstraintRelationshipValues.Child, ReferenceForName = "node", Fact = 0.6D };
            Dgm.Constraint constraint3 = new Dgm.Constraint(){ Type = Dgm.ConstraintValues.Width, For = Dgm.ConstraintRelationshipValues.Child, ForName = "sibTrans", ReferenceType = Dgm.ConstraintValues.Width, ReferenceFor = Dgm.ConstraintRelationshipValues.Child, ReferenceForName = "node", Fact = 0.1D };
            Dgm.Constraint constraint4 = new Dgm.Constraint(){ Type = Dgm.ConstraintValues.Spacing, ReferenceType = Dgm.ConstraintValues.Width, ReferenceFor = Dgm.ConstraintRelationshipValues.Child, ReferenceForName = "sibTrans" };
            Dgm.Constraint constraint5 = new Dgm.Constraint(){ Type = Dgm.ConstraintValues.PrimaryFontSize, For = Dgm.ConstraintRelationshipValues.Child, ForName = "node", Operator = Dgm.BoolOperatorValues.Equal, Val = 65D };

            constraints1.Append(constraint1);
            constraints1.Append(constraint2);
            constraints1.Append(constraint3);
            constraints1.Append(constraint4);
            constraints1.Append(constraint5);
            Dgm.RuleList ruleList1 = new Dgm.RuleList();

            Dgm.ForEach forEach1 = new Dgm.ForEach(){ Name = "Name3", Axis = new ListValue<EnumValue<DocumentFormat.OpenXml.Drawing.Diagrams.AxisValues>> { InnerText = "ch" }, PointType = new ListValue<EnumValue<DocumentFormat.OpenXml.Drawing.Diagrams.ElementValues>> { InnerText = "node" } };

            Dgm.LayoutNode layoutNode2 = new Dgm.LayoutNode(){ Name = "node" };

            Dgm.VariableList variableList2 = new Dgm.VariableList();
            Dgm.BulletEnabled bulletEnabled1 = new Dgm.BulletEnabled(){ Val = true };

            variableList2.Append(bulletEnabled1);
            Dgm.Algorithm algorithm3 = new Dgm.Algorithm(){ Type = Dgm.AlgorithmValues.Text };

            Dgm.Shape shape7 = new Dgm.Shape(){ Type = "rect", Blip = "" };
            shape7.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            Dgm.AdjustList adjustList2 = new Dgm.AdjustList();

            shape7.Append(adjustList2);
            Dgm.PresentationOf presentationOf2 = new Dgm.PresentationOf(){ Axis = new ListValue<EnumValue<DocumentFormat.OpenXml.Drawing.Diagrams.AxisValues>> { InnerText = "desOrSelf" }, PointType = new ListValue<EnumValue<DocumentFormat.OpenXml.Drawing.Diagrams.ElementValues>> { InnerText = "node" } };

            Dgm.Constraints constraints2 = new Dgm.Constraints();
            Dgm.Constraint constraint6 = new Dgm.Constraint(){ Type = Dgm.ConstraintValues.LeftMargin, ReferenceType = Dgm.ConstraintValues.PrimaryFontSize, Fact = 0.3D };
            Dgm.Constraint constraint7 = new Dgm.Constraint(){ Type = Dgm.ConstraintValues.RightMargin, ReferenceType = Dgm.ConstraintValues.PrimaryFontSize, Fact = 0.3D };
            Dgm.Constraint constraint8 = new Dgm.Constraint(){ Type = Dgm.ConstraintValues.TopMargin, ReferenceType = Dgm.ConstraintValues.PrimaryFontSize, Fact = 0.3D };
            Dgm.Constraint constraint9 = new Dgm.Constraint(){ Type = Dgm.ConstraintValues.BottomMargin, ReferenceType = Dgm.ConstraintValues.PrimaryFontSize, Fact = 0.3D };

            constraints2.Append(constraint6);
            constraints2.Append(constraint7);
            constraints2.Append(constraint8);
            constraints2.Append(constraint9);

            Dgm.RuleList ruleList2 = new Dgm.RuleList();
            Dgm.Rule rule1 = new Dgm.Rule(){ Type = Dgm.ConstraintValues.PrimaryFontSize, Val = 5D, Fact = new DoubleValue() { InnerText = "NaN" }, Max = new DoubleValue() { InnerText = "NaN" } };

            ruleList2.Append(rule1);

            layoutNode2.Append(variableList2);
            layoutNode2.Append(algorithm3);
            layoutNode2.Append(shape7);
            layoutNode2.Append(presentationOf2);
            layoutNode2.Append(constraints2);
            layoutNode2.Append(ruleList2);

            Dgm.ForEach forEach2 = new Dgm.ForEach(){ Name = "Name4", Axis = new ListValue<EnumValue<DocumentFormat.OpenXml.Drawing.Diagrams.AxisValues>> { InnerText = "followSib" }, PointType = new ListValue<EnumValue<DocumentFormat.OpenXml.Drawing.Diagrams.ElementValues>> { InnerText = "sibTrans" }, Count = new ListValue<UInt32Value>() { InnerText = "1" } };

            Dgm.LayoutNode layoutNode3 = new Dgm.LayoutNode(){ Name = "sibTrans" };
            Dgm.Algorithm algorithm4 = new Dgm.Algorithm(){ Type = Dgm.AlgorithmValues.Space };

            Dgm.Shape shape8 = new Dgm.Shape(){ Blip = "" };
            shape8.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            Dgm.AdjustList adjustList3 = new Dgm.AdjustList();

            shape8.Append(adjustList3);
            Dgm.PresentationOf presentationOf3 = new Dgm.PresentationOf();
            Dgm.Constraints constraints3 = new Dgm.Constraints();
            Dgm.RuleList ruleList3 = new Dgm.RuleList();

            layoutNode3.Append(algorithm4);
            layoutNode3.Append(shape8);
            layoutNode3.Append(presentationOf3);
            layoutNode3.Append(constraints3);
            layoutNode3.Append(ruleList3);

            forEach2.Append(layoutNode3);

            forEach1.Append(layoutNode2);
            forEach1.Append(forEach2);

            layoutNode1.Append(variableList1);
            layoutNode1.Append(choose1);
            layoutNode1.Append(shape6);
            layoutNode1.Append(presentationOf1);
            layoutNode1.Append(constraints1);
            layoutNode1.Append(ruleList1);
            layoutNode1.Append(forEach1);

            layoutDefinition1.Append(title1);
            layoutDefinition1.Append(description1);
            layoutDefinition1.Append(categoryList1);
            layoutDefinition1.Append(sampleData1);
            layoutDefinition1.Append(styleData1);
            layoutDefinition1.Append(colorData1);
            layoutDefinition1.Append(layoutNode1);

            diagramLayoutDefinitionPart1.LayoutDefinition = layoutDefinition1;
        }

        // Generates content of themePart1.
        private void GenerateThemePart1Content(ThemePart themePart1)
        {
            A.Theme theme1 = new A.Theme(){ Name = "Office Theme" };
            theme1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            A.ThemeElements themeElements1 = new A.ThemeElements();

            A.ColorScheme colorScheme1 = new A.ColorScheme(){ Name = "Office" };

            A.Dark1Color dark1Color1 = new A.Dark1Color();
            A.SystemColor systemColor1 = new A.SystemColor(){ Val = A.SystemColorValues.WindowText, LastColor = "000000" };

            dark1Color1.Append(systemColor1);

            A.Light1Color light1Color1 = new A.Light1Color();
            A.SystemColor systemColor2 = new A.SystemColor(){ Val = A.SystemColorValues.Window, LastColor = "FFFFFF" };

            light1Color1.Append(systemColor2);

            A.Dark2Color dark2Color1 = new A.Dark2Color();
            A.RgbColorModelHex rgbColorModelHex1 = new A.RgbColorModelHex(){ Val = "0E2841" };

            dark2Color1.Append(rgbColorModelHex1);

            A.Light2Color light2Color1 = new A.Light2Color();
            A.RgbColorModelHex rgbColorModelHex2 = new A.RgbColorModelHex(){ Val = "E8E8E8" };

            light2Color1.Append(rgbColorModelHex2);

            A.Accent1Color accent1Color1 = new A.Accent1Color();
            A.RgbColorModelHex rgbColorModelHex3 = new A.RgbColorModelHex(){ Val = "156082" };

            accent1Color1.Append(rgbColorModelHex3);

            A.Accent2Color accent2Color1 = new A.Accent2Color();
            A.RgbColorModelHex rgbColorModelHex4 = new A.RgbColorModelHex(){ Val = "E97132" };

            accent2Color1.Append(rgbColorModelHex4);

            A.Accent3Color accent3Color1 = new A.Accent3Color();
            A.RgbColorModelHex rgbColorModelHex5 = new A.RgbColorModelHex(){ Val = "196B24" };

            accent3Color1.Append(rgbColorModelHex5);

            A.Accent4Color accent4Color1 = new A.Accent4Color();
            A.RgbColorModelHex rgbColorModelHex6 = new A.RgbColorModelHex(){ Val = "0F9ED5" };

            accent4Color1.Append(rgbColorModelHex6);

            A.Accent5Color accent5Color1 = new A.Accent5Color();
            A.RgbColorModelHex rgbColorModelHex7 = new A.RgbColorModelHex(){ Val = "A02B93" };

            accent5Color1.Append(rgbColorModelHex7);

            A.Accent6Color accent6Color1 = new A.Accent6Color();
            A.RgbColorModelHex rgbColorModelHex8 = new A.RgbColorModelHex(){ Val = "4EA72E" };

            accent6Color1.Append(rgbColorModelHex8);

            A.Hyperlink hyperlink1 = new A.Hyperlink();
            A.RgbColorModelHex rgbColorModelHex9 = new A.RgbColorModelHex(){ Val = "467886" };

            hyperlink1.Append(rgbColorModelHex9);

            A.FollowedHyperlinkColor followedHyperlinkColor1 = new A.FollowedHyperlinkColor();
            A.RgbColorModelHex rgbColorModelHex10 = new A.RgbColorModelHex(){ Val = "96607D" };

            followedHyperlinkColor1.Append(rgbColorModelHex10);

            colorScheme1.Append(dark1Color1);
            colorScheme1.Append(light1Color1);
            colorScheme1.Append(dark2Color1);
            colorScheme1.Append(light2Color1);
            colorScheme1.Append(accent1Color1);
            colorScheme1.Append(accent2Color1);
            colorScheme1.Append(accent3Color1);
            colorScheme1.Append(accent4Color1);
            colorScheme1.Append(accent5Color1);
            colorScheme1.Append(accent6Color1);
            colorScheme1.Append(hyperlink1);
            colorScheme1.Append(followedHyperlinkColor1);

            A.FontScheme fontScheme1 = new A.FontScheme(){ Name = "Office" };

            A.MajorFont majorFont1 = new A.MajorFont();
            A.LatinFont latinFont1 = new A.LatinFont(){ Typeface = "Aptos Display", Panose = "02110004020202020204" };
            A.EastAsianFont eastAsianFont1 = new A.EastAsianFont(){ Typeface = "" };
            A.ComplexScriptFont complexScriptFont1 = new A.ComplexScriptFont(){ Typeface = "" };
            A.SupplementalFont supplementalFont1 = new A.SupplementalFont(){ Script = "Jpan", Typeface = "游ゴシック Light" };
            A.SupplementalFont supplementalFont2 = new A.SupplementalFont(){ Script = "Hang", Typeface = "맑은 고딕" };
            A.SupplementalFont supplementalFont3 = new A.SupplementalFont(){ Script = "Hans", Typeface = "等线 Light" };
            A.SupplementalFont supplementalFont4 = new A.SupplementalFont(){ Script = "Hant", Typeface = "新細明體" };
            A.SupplementalFont supplementalFont5 = new A.SupplementalFont(){ Script = "Arab", Typeface = "Times New Roman" };
            A.SupplementalFont supplementalFont6 = new A.SupplementalFont(){ Script = "Hebr", Typeface = "Times New Roman" };
            A.SupplementalFont supplementalFont7 = new A.SupplementalFont(){ Script = "Thai", Typeface = "Angsana New" };
            A.SupplementalFont supplementalFont8 = new A.SupplementalFont(){ Script = "Ethi", Typeface = "Nyala" };
            A.SupplementalFont supplementalFont9 = new A.SupplementalFont(){ Script = "Beng", Typeface = "Vrinda" };
            A.SupplementalFont supplementalFont10 = new A.SupplementalFont(){ Script = "Gujr", Typeface = "Shruti" };
            A.SupplementalFont supplementalFont11 = new A.SupplementalFont(){ Script = "Khmr", Typeface = "MoolBoran" };
            A.SupplementalFont supplementalFont12 = new A.SupplementalFont(){ Script = "Knda", Typeface = "Tunga" };
            A.SupplementalFont supplementalFont13 = new A.SupplementalFont(){ Script = "Guru", Typeface = "Raavi" };
            A.SupplementalFont supplementalFont14 = new A.SupplementalFont(){ Script = "Cans", Typeface = "Euphemia" };
            A.SupplementalFont supplementalFont15 = new A.SupplementalFont(){ Script = "Cher", Typeface = "Plantagenet Cherokee" };
            A.SupplementalFont supplementalFont16 = new A.SupplementalFont(){ Script = "Yiii", Typeface = "Microsoft Yi Baiti" };
            A.SupplementalFont supplementalFont17 = new A.SupplementalFont(){ Script = "Tibt", Typeface = "Microsoft Himalaya" };
            A.SupplementalFont supplementalFont18 = new A.SupplementalFont(){ Script = "Thaa", Typeface = "MV Boli" };
            A.SupplementalFont supplementalFont19 = new A.SupplementalFont(){ Script = "Deva", Typeface = "Mangal" };
            A.SupplementalFont supplementalFont20 = new A.SupplementalFont(){ Script = "Telu", Typeface = "Gautami" };
            A.SupplementalFont supplementalFont21 = new A.SupplementalFont(){ Script = "Taml", Typeface = "Latha" };
            A.SupplementalFont supplementalFont22 = new A.SupplementalFont(){ Script = "Syrc", Typeface = "Estrangelo Edessa" };
            A.SupplementalFont supplementalFont23 = new A.SupplementalFont(){ Script = "Orya", Typeface = "Kalinga" };
            A.SupplementalFont supplementalFont24 = new A.SupplementalFont(){ Script = "Mlym", Typeface = "Kartika" };
            A.SupplementalFont supplementalFont25 = new A.SupplementalFont(){ Script = "Laoo", Typeface = "DokChampa" };
            A.SupplementalFont supplementalFont26 = new A.SupplementalFont(){ Script = "Sinh", Typeface = "Iskoola Pota" };
            A.SupplementalFont supplementalFont27 = new A.SupplementalFont(){ Script = "Mong", Typeface = "Mongolian Baiti" };
            A.SupplementalFont supplementalFont28 = new A.SupplementalFont(){ Script = "Viet", Typeface = "Times New Roman" };
            A.SupplementalFont supplementalFont29 = new A.SupplementalFont(){ Script = "Uigh", Typeface = "Microsoft Uighur" };
            A.SupplementalFont supplementalFont30 = new A.SupplementalFont(){ Script = "Geor", Typeface = "Sylfaen" };
            A.SupplementalFont supplementalFont31 = new A.SupplementalFont(){ Script = "Armn", Typeface = "Arial" };
            A.SupplementalFont supplementalFont32 = new A.SupplementalFont(){ Script = "Bugi", Typeface = "Leelawadee UI" };
            A.SupplementalFont supplementalFont33 = new A.SupplementalFont(){ Script = "Bopo", Typeface = "Microsoft JhengHei" };
            A.SupplementalFont supplementalFont34 = new A.SupplementalFont(){ Script = "Java", Typeface = "Javanese Text" };
            A.SupplementalFont supplementalFont35 = new A.SupplementalFont(){ Script = "Lisu", Typeface = "Segoe UI" };
            A.SupplementalFont supplementalFont36 = new A.SupplementalFont(){ Script = "Mymr", Typeface = "Myanmar Text" };
            A.SupplementalFont supplementalFont37 = new A.SupplementalFont(){ Script = "Nkoo", Typeface = "Ebrima" };
            A.SupplementalFont supplementalFont38 = new A.SupplementalFont(){ Script = "Olck", Typeface = "Nirmala UI" };
            A.SupplementalFont supplementalFont39 = new A.SupplementalFont(){ Script = "Osma", Typeface = "Ebrima" };
            A.SupplementalFont supplementalFont40 = new A.SupplementalFont(){ Script = "Phag", Typeface = "Phagspa" };
            A.SupplementalFont supplementalFont41 = new A.SupplementalFont(){ Script = "Syrn", Typeface = "Estrangelo Edessa" };
            A.SupplementalFont supplementalFont42 = new A.SupplementalFont(){ Script = "Syrj", Typeface = "Estrangelo Edessa" };
            A.SupplementalFont supplementalFont43 = new A.SupplementalFont(){ Script = "Syre", Typeface = "Estrangelo Edessa" };
            A.SupplementalFont supplementalFont44 = new A.SupplementalFont(){ Script = "Sora", Typeface = "Nirmala UI" };
            A.SupplementalFont supplementalFont45 = new A.SupplementalFont(){ Script = "Tale", Typeface = "Microsoft Tai Le" };
            A.SupplementalFont supplementalFont46 = new A.SupplementalFont(){ Script = "Talu", Typeface = "Microsoft New Tai Lue" };
            A.SupplementalFont supplementalFont47 = new A.SupplementalFont(){ Script = "Tfng", Typeface = "Ebrima" };

            majorFont1.Append(latinFont1);
            majorFont1.Append(eastAsianFont1);
            majorFont1.Append(complexScriptFont1);
            majorFont1.Append(supplementalFont1);
            majorFont1.Append(supplementalFont2);
            majorFont1.Append(supplementalFont3);
            majorFont1.Append(supplementalFont4);
            majorFont1.Append(supplementalFont5);
            majorFont1.Append(supplementalFont6);
            majorFont1.Append(supplementalFont7);
            majorFont1.Append(supplementalFont8);
            majorFont1.Append(supplementalFont9);
            majorFont1.Append(supplementalFont10);
            majorFont1.Append(supplementalFont11);
            majorFont1.Append(supplementalFont12);
            majorFont1.Append(supplementalFont13);
            majorFont1.Append(supplementalFont14);
            majorFont1.Append(supplementalFont15);
            majorFont1.Append(supplementalFont16);
            majorFont1.Append(supplementalFont17);
            majorFont1.Append(supplementalFont18);
            majorFont1.Append(supplementalFont19);
            majorFont1.Append(supplementalFont20);
            majorFont1.Append(supplementalFont21);
            majorFont1.Append(supplementalFont22);
            majorFont1.Append(supplementalFont23);
            majorFont1.Append(supplementalFont24);
            majorFont1.Append(supplementalFont25);
            majorFont1.Append(supplementalFont26);
            majorFont1.Append(supplementalFont27);
            majorFont1.Append(supplementalFont28);
            majorFont1.Append(supplementalFont29);
            majorFont1.Append(supplementalFont30);
            majorFont1.Append(supplementalFont31);
            majorFont1.Append(supplementalFont32);
            majorFont1.Append(supplementalFont33);
            majorFont1.Append(supplementalFont34);
            majorFont1.Append(supplementalFont35);
            majorFont1.Append(supplementalFont36);
            majorFont1.Append(supplementalFont37);
            majorFont1.Append(supplementalFont38);
            majorFont1.Append(supplementalFont39);
            majorFont1.Append(supplementalFont40);
            majorFont1.Append(supplementalFont41);
            majorFont1.Append(supplementalFont42);
            majorFont1.Append(supplementalFont43);
            majorFont1.Append(supplementalFont44);
            majorFont1.Append(supplementalFont45);
            majorFont1.Append(supplementalFont46);
            majorFont1.Append(supplementalFont47);

            A.MinorFont minorFont1 = new A.MinorFont();
            A.LatinFont latinFont2 = new A.LatinFont(){ Typeface = "Aptos", Panose = "02110004020202020204" };
            A.EastAsianFont eastAsianFont2 = new A.EastAsianFont(){ Typeface = "" };
            A.ComplexScriptFont complexScriptFont2 = new A.ComplexScriptFont(){ Typeface = "" };
            A.SupplementalFont supplementalFont48 = new A.SupplementalFont(){ Script = "Jpan", Typeface = "游明朝" };
            A.SupplementalFont supplementalFont49 = new A.SupplementalFont(){ Script = "Hang", Typeface = "맑은 고딕" };
            A.SupplementalFont supplementalFont50 = new A.SupplementalFont(){ Script = "Hans", Typeface = "等线" };
            A.SupplementalFont supplementalFont51 = new A.SupplementalFont(){ Script = "Hant", Typeface = "新細明體" };
            A.SupplementalFont supplementalFont52 = new A.SupplementalFont(){ Script = "Arab", Typeface = "Arial" };
            A.SupplementalFont supplementalFont53 = new A.SupplementalFont(){ Script = "Hebr", Typeface = "Arial" };
            A.SupplementalFont supplementalFont54 = new A.SupplementalFont(){ Script = "Thai", Typeface = "Cordia New" };
            A.SupplementalFont supplementalFont55 = new A.SupplementalFont(){ Script = "Ethi", Typeface = "Nyala" };
            A.SupplementalFont supplementalFont56 = new A.SupplementalFont(){ Script = "Beng", Typeface = "Vrinda" };
            A.SupplementalFont supplementalFont57 = new A.SupplementalFont(){ Script = "Gujr", Typeface = "Shruti" };
            A.SupplementalFont supplementalFont58 = new A.SupplementalFont(){ Script = "Khmr", Typeface = "DaunPenh" };
            A.SupplementalFont supplementalFont59 = new A.SupplementalFont(){ Script = "Knda", Typeface = "Tunga" };
            A.SupplementalFont supplementalFont60 = new A.SupplementalFont(){ Script = "Guru", Typeface = "Raavi" };
            A.SupplementalFont supplementalFont61 = new A.SupplementalFont(){ Script = "Cans", Typeface = "Euphemia" };
            A.SupplementalFont supplementalFont62 = new A.SupplementalFont(){ Script = "Cher", Typeface = "Plantagenet Cherokee" };
            A.SupplementalFont supplementalFont63 = new A.SupplementalFont(){ Script = "Yiii", Typeface = "Microsoft Yi Baiti" };
            A.SupplementalFont supplementalFont64 = new A.SupplementalFont(){ Script = "Tibt", Typeface = "Microsoft Himalaya" };
            A.SupplementalFont supplementalFont65 = new A.SupplementalFont(){ Script = "Thaa", Typeface = "MV Boli" };
            A.SupplementalFont supplementalFont66 = new A.SupplementalFont(){ Script = "Deva", Typeface = "Mangal" };
            A.SupplementalFont supplementalFont67 = new A.SupplementalFont(){ Script = "Telu", Typeface = "Gautami" };
            A.SupplementalFont supplementalFont68 = new A.SupplementalFont(){ Script = "Taml", Typeface = "Latha" };
            A.SupplementalFont supplementalFont69 = new A.SupplementalFont(){ Script = "Syrc", Typeface = "Estrangelo Edessa" };
            A.SupplementalFont supplementalFont70 = new A.SupplementalFont(){ Script = "Orya", Typeface = "Kalinga" };
            A.SupplementalFont supplementalFont71 = new A.SupplementalFont(){ Script = "Mlym", Typeface = "Kartika" };
            A.SupplementalFont supplementalFont72 = new A.SupplementalFont(){ Script = "Laoo", Typeface = "DokChampa" };
            A.SupplementalFont supplementalFont73 = new A.SupplementalFont(){ Script = "Sinh", Typeface = "Iskoola Pota" };
            A.SupplementalFont supplementalFont74 = new A.SupplementalFont(){ Script = "Mong", Typeface = "Mongolian Baiti" };
            A.SupplementalFont supplementalFont75 = new A.SupplementalFont(){ Script = "Viet", Typeface = "Arial" };
            A.SupplementalFont supplementalFont76 = new A.SupplementalFont(){ Script = "Uigh", Typeface = "Microsoft Uighur" };
            A.SupplementalFont supplementalFont77 = new A.SupplementalFont(){ Script = "Geor", Typeface = "Sylfaen" };
            A.SupplementalFont supplementalFont78 = new A.SupplementalFont(){ Script = "Armn", Typeface = "Arial" };
            A.SupplementalFont supplementalFont79 = new A.SupplementalFont(){ Script = "Bugi", Typeface = "Leelawadee UI" };
            A.SupplementalFont supplementalFont80 = new A.SupplementalFont(){ Script = "Bopo", Typeface = "Microsoft JhengHei" };
            A.SupplementalFont supplementalFont81 = new A.SupplementalFont(){ Script = "Java", Typeface = "Javanese Text" };
            A.SupplementalFont supplementalFont82 = new A.SupplementalFont(){ Script = "Lisu", Typeface = "Segoe UI" };
            A.SupplementalFont supplementalFont83 = new A.SupplementalFont(){ Script = "Mymr", Typeface = "Myanmar Text" };
            A.SupplementalFont supplementalFont84 = new A.SupplementalFont(){ Script = "Nkoo", Typeface = "Ebrima" };
            A.SupplementalFont supplementalFont85 = new A.SupplementalFont(){ Script = "Olck", Typeface = "Nirmala UI" };
            A.SupplementalFont supplementalFont86 = new A.SupplementalFont(){ Script = "Osma", Typeface = "Ebrima" };
            A.SupplementalFont supplementalFont87 = new A.SupplementalFont(){ Script = "Phag", Typeface = "Phagspa" };
            A.SupplementalFont supplementalFont88 = new A.SupplementalFont(){ Script = "Syrn", Typeface = "Estrangelo Edessa" };
            A.SupplementalFont supplementalFont89 = new A.SupplementalFont(){ Script = "Syrj", Typeface = "Estrangelo Edessa" };
            A.SupplementalFont supplementalFont90 = new A.SupplementalFont(){ Script = "Syre", Typeface = "Estrangelo Edessa" };
            A.SupplementalFont supplementalFont91 = new A.SupplementalFont(){ Script = "Sora", Typeface = "Nirmala UI" };
            A.SupplementalFont supplementalFont92 = new A.SupplementalFont(){ Script = "Tale", Typeface = "Microsoft Tai Le" };
            A.SupplementalFont supplementalFont93 = new A.SupplementalFont(){ Script = "Talu", Typeface = "Microsoft New Tai Lue" };
            A.SupplementalFont supplementalFont94 = new A.SupplementalFont(){ Script = "Tfng", Typeface = "Ebrima" };

            minorFont1.Append(latinFont2);
            minorFont1.Append(eastAsianFont2);
            minorFont1.Append(complexScriptFont2);
            minorFont1.Append(supplementalFont48);
            minorFont1.Append(supplementalFont49);
            minorFont1.Append(supplementalFont50);
            minorFont1.Append(supplementalFont51);
            minorFont1.Append(supplementalFont52);
            minorFont1.Append(supplementalFont53);
            minorFont1.Append(supplementalFont54);
            minorFont1.Append(supplementalFont55);
            minorFont1.Append(supplementalFont56);
            minorFont1.Append(supplementalFont57);
            minorFont1.Append(supplementalFont58);
            minorFont1.Append(supplementalFont59);
            minorFont1.Append(supplementalFont60);
            minorFont1.Append(supplementalFont61);
            minorFont1.Append(supplementalFont62);
            minorFont1.Append(supplementalFont63);
            minorFont1.Append(supplementalFont64);
            minorFont1.Append(supplementalFont65);
            minorFont1.Append(supplementalFont66);
            minorFont1.Append(supplementalFont67);
            minorFont1.Append(supplementalFont68);
            minorFont1.Append(supplementalFont69);
            minorFont1.Append(supplementalFont70);
            minorFont1.Append(supplementalFont71);
            minorFont1.Append(supplementalFont72);
            minorFont1.Append(supplementalFont73);
            minorFont1.Append(supplementalFont74);
            minorFont1.Append(supplementalFont75);
            minorFont1.Append(supplementalFont76);
            minorFont1.Append(supplementalFont77);
            minorFont1.Append(supplementalFont78);
            minorFont1.Append(supplementalFont79);
            minorFont1.Append(supplementalFont80);
            minorFont1.Append(supplementalFont81);
            minorFont1.Append(supplementalFont82);
            minorFont1.Append(supplementalFont83);
            minorFont1.Append(supplementalFont84);
            minorFont1.Append(supplementalFont85);
            minorFont1.Append(supplementalFont86);
            minorFont1.Append(supplementalFont87);
            minorFont1.Append(supplementalFont88);
            minorFont1.Append(supplementalFont89);
            minorFont1.Append(supplementalFont90);
            minorFont1.Append(supplementalFont91);
            minorFont1.Append(supplementalFont92);
            minorFont1.Append(supplementalFont93);
            minorFont1.Append(supplementalFont94);

            fontScheme1.Append(majorFont1);
            fontScheme1.Append(minorFont1);

            A.FormatScheme formatScheme1 = new A.FormatScheme(){ Name = "Office" };

            A.FillStyleList fillStyleList1 = new A.FillStyleList();

            A.SolidFill solidFill11 = new A.SolidFill();
            A.SchemeColor schemeColor167 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };

            solidFill11.Append(schemeColor167);

            A.GradientFill gradientFill1 = new A.GradientFill(){ RotateWithShape = true };

            A.GradientStopList gradientStopList1 = new A.GradientStopList();

            A.GradientStop gradientStop1 = new A.GradientStop(){ Position = 0 };

            A.SchemeColor schemeColor168 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };
            A.LuminanceModulation luminanceModulation1 = new A.LuminanceModulation(){ Val = 110000 };
            A.SaturationModulation saturationModulation1 = new A.SaturationModulation(){ Val = 105000 };
            A.Tint tint22 = new A.Tint(){ Val = 67000 };

            schemeColor168.Append(luminanceModulation1);
            schemeColor168.Append(saturationModulation1);
            schemeColor168.Append(tint22);

            gradientStop1.Append(schemeColor168);

            A.GradientStop gradientStop2 = new A.GradientStop(){ Position = 50000 };

            A.SchemeColor schemeColor169 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };
            A.LuminanceModulation luminanceModulation2 = new A.LuminanceModulation(){ Val = 105000 };
            A.SaturationModulation saturationModulation2 = new A.SaturationModulation(){ Val = 103000 };
            A.Tint tint23 = new A.Tint(){ Val = 73000 };

            schemeColor169.Append(luminanceModulation2);
            schemeColor169.Append(saturationModulation2);
            schemeColor169.Append(tint23);

            gradientStop2.Append(schemeColor169);

            A.GradientStop gradientStop3 = new A.GradientStop(){ Position = 100000 };

            A.SchemeColor schemeColor170 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };
            A.LuminanceModulation luminanceModulation3 = new A.LuminanceModulation(){ Val = 105000 };
            A.SaturationModulation saturationModulation3 = new A.SaturationModulation(){ Val = 109000 };
            A.Tint tint24 = new A.Tint(){ Val = 81000 };

            schemeColor170.Append(luminanceModulation3);
            schemeColor170.Append(saturationModulation3);
            schemeColor170.Append(tint24);

            gradientStop3.Append(schemeColor170);

            gradientStopList1.Append(gradientStop1);
            gradientStopList1.Append(gradientStop2);
            gradientStopList1.Append(gradientStop3);
            A.LinearGradientFill linearGradientFill1 = new A.LinearGradientFill(){ Angle = 5400000, Scaled = false };

            gradientFill1.Append(gradientStopList1);
            gradientFill1.Append(linearGradientFill1);

            A.GradientFill gradientFill2 = new A.GradientFill(){ RotateWithShape = true };

            A.GradientStopList gradientStopList2 = new A.GradientStopList();

            A.GradientStop gradientStop4 = new A.GradientStop(){ Position = 0 };

            A.SchemeColor schemeColor171 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };
            A.SaturationModulation saturationModulation4 = new A.SaturationModulation(){ Val = 103000 };
            A.LuminanceModulation luminanceModulation4 = new A.LuminanceModulation(){ Val = 102000 };
            A.Tint tint25 = new A.Tint(){ Val = 94000 };

            schemeColor171.Append(saturationModulation4);
            schemeColor171.Append(luminanceModulation4);
            schemeColor171.Append(tint25);

            gradientStop4.Append(schemeColor171);

            A.GradientStop gradientStop5 = new A.GradientStop(){ Position = 50000 };

            A.SchemeColor schemeColor172 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };
            A.SaturationModulation saturationModulation5 = new A.SaturationModulation(){ Val = 110000 };
            A.LuminanceModulation luminanceModulation5 = new A.LuminanceModulation(){ Val = 100000 };
            A.Shade shade6 = new A.Shade(){ Val = 100000 };

            schemeColor172.Append(saturationModulation5);
            schemeColor172.Append(luminanceModulation5);
            schemeColor172.Append(shade6);

            gradientStop5.Append(schemeColor172);

            A.GradientStop gradientStop6 = new A.GradientStop(){ Position = 100000 };

            A.SchemeColor schemeColor173 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };
            A.LuminanceModulation luminanceModulation6 = new A.LuminanceModulation(){ Val = 99000 };
            A.SaturationModulation saturationModulation6 = new A.SaturationModulation(){ Val = 120000 };
            A.Shade shade7 = new A.Shade(){ Val = 78000 };

            schemeColor173.Append(luminanceModulation6);
            schemeColor173.Append(saturationModulation6);
            schemeColor173.Append(shade7);

            gradientStop6.Append(schemeColor173);

            gradientStopList2.Append(gradientStop4);
            gradientStopList2.Append(gradientStop5);
            gradientStopList2.Append(gradientStop6);
            A.LinearGradientFill linearGradientFill2 = new A.LinearGradientFill(){ Angle = 5400000, Scaled = false };

            gradientFill2.Append(gradientStopList2);
            gradientFill2.Append(linearGradientFill2);

            fillStyleList1.Append(solidFill11);
            fillStyleList1.Append(gradientFill1);
            fillStyleList1.Append(gradientFill2);

            A.LineStyleList lineStyleList1 = new A.LineStyleList();

            A.Outline outline6 = new A.Outline(){ Width = 12700, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill12 = new A.SolidFill();
            A.SchemeColor schemeColor174 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };

            solidFill12.Append(schemeColor174);
            A.PresetDash presetDash6 = new A.PresetDash(){ Val = A.PresetLineDashValues.Solid };
            A.Miter miter6 = new A.Miter(){ Limit = 800000 };

            outline6.Append(solidFill12);
            outline6.Append(presetDash6);
            outline6.Append(miter6);

            A.Outline outline7 = new A.Outline(){ Width = 19050, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill13 = new A.SolidFill();
            A.SchemeColor schemeColor175 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };

            solidFill13.Append(schemeColor175);
            A.PresetDash presetDash7 = new A.PresetDash(){ Val = A.PresetLineDashValues.Solid };
            A.Miter miter7 = new A.Miter(){ Limit = 800000 };

            outline7.Append(solidFill13);
            outline7.Append(presetDash7);
            outline7.Append(miter7);

            A.Outline outline8 = new A.Outline(){ Width = 25400, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill14 = new A.SolidFill();
            A.SchemeColor schemeColor176 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };

            solidFill14.Append(schemeColor176);
            A.PresetDash presetDash8 = new A.PresetDash(){ Val = A.PresetLineDashValues.Solid };
            A.Miter miter8 = new A.Miter(){ Limit = 800000 };

            outline8.Append(solidFill14);
            outline8.Append(presetDash8);
            outline8.Append(miter8);

            lineStyleList1.Append(outline6);
            lineStyleList1.Append(outline7);
            lineStyleList1.Append(outline8);

            A.EffectStyleList effectStyleList1 = new A.EffectStyleList();

            A.EffectStyle effectStyle1 = new A.EffectStyle();
            A.EffectList effectList6 = new A.EffectList();

            effectStyle1.Append(effectList6);

            A.EffectStyle effectStyle2 = new A.EffectStyle();
            A.EffectList effectList7 = new A.EffectList();

            effectStyle2.Append(effectList7);

            A.EffectStyle effectStyle3 = new A.EffectStyle();

            A.EffectList effectList8 = new A.EffectList();

            A.OuterShadow outerShadow1 = new A.OuterShadow(){ BlurRadius = 57150L, Distance = 19050L, Direction = 5400000, Alignment = A.RectangleAlignmentValues.Center, RotateWithShape = false };

            A.RgbColorModelHex rgbColorModelHex11 = new A.RgbColorModelHex(){ Val = "000000" };
            A.Alpha alpha20 = new A.Alpha(){ Val = 63000 };

            rgbColorModelHex11.Append(alpha20);

            outerShadow1.Append(rgbColorModelHex11);

            effectList8.Append(outerShadow1);

            effectStyle3.Append(effectList8);

            effectStyleList1.Append(effectStyle1);
            effectStyleList1.Append(effectStyle2);
            effectStyleList1.Append(effectStyle3);

            A.BackgroundFillStyleList backgroundFillStyleList1 = new A.BackgroundFillStyleList();

            A.SolidFill solidFill15 = new A.SolidFill();
            A.SchemeColor schemeColor177 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };

            solidFill15.Append(schemeColor177);

            A.SolidFill solidFill16 = new A.SolidFill();

            A.SchemeColor schemeColor178 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };
            A.Tint tint26 = new A.Tint(){ Val = 95000 };
            A.SaturationModulation saturationModulation7 = new A.SaturationModulation(){ Val = 170000 };

            schemeColor178.Append(tint26);
            schemeColor178.Append(saturationModulation7);

            solidFill16.Append(schemeColor178);

            A.GradientFill gradientFill3 = new A.GradientFill(){ RotateWithShape = true };

            A.GradientStopList gradientStopList3 = new A.GradientStopList();

            A.GradientStop gradientStop7 = new A.GradientStop(){ Position = 0 };

            A.SchemeColor schemeColor179 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };
            A.Tint tint27 = new A.Tint(){ Val = 93000 };
            A.SaturationModulation saturationModulation8 = new A.SaturationModulation(){ Val = 150000 };
            A.Shade shade8 = new A.Shade(){ Val = 98000 };
            A.LuminanceModulation luminanceModulation7 = new A.LuminanceModulation(){ Val = 102000 };

            schemeColor179.Append(tint27);
            schemeColor179.Append(saturationModulation8);
            schemeColor179.Append(shade8);
            schemeColor179.Append(luminanceModulation7);

            gradientStop7.Append(schemeColor179);

            A.GradientStop gradientStop8 = new A.GradientStop(){ Position = 50000 };

            A.SchemeColor schemeColor180 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };
            A.Tint tint28 = new A.Tint(){ Val = 98000 };
            A.SaturationModulation saturationModulation9 = new A.SaturationModulation(){ Val = 130000 };
            A.Shade shade9 = new A.Shade(){ Val = 90000 };
            A.LuminanceModulation luminanceModulation8 = new A.LuminanceModulation(){ Val = 103000 };

            schemeColor180.Append(tint28);
            schemeColor180.Append(saturationModulation9);
            schemeColor180.Append(shade9);
            schemeColor180.Append(luminanceModulation8);

            gradientStop8.Append(schemeColor180);

            A.GradientStop gradientStop9 = new A.GradientStop(){ Position = 100000 };

            A.SchemeColor schemeColor181 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };
            A.Shade shade10 = new A.Shade(){ Val = 63000 };
            A.SaturationModulation saturationModulation10 = new A.SaturationModulation(){ Val = 120000 };

            schemeColor181.Append(shade10);
            schemeColor181.Append(saturationModulation10);

            gradientStop9.Append(schemeColor181);

            gradientStopList3.Append(gradientStop7);
            gradientStopList3.Append(gradientStop8);
            gradientStopList3.Append(gradientStop9);
            A.LinearGradientFill linearGradientFill3 = new A.LinearGradientFill(){ Angle = 5400000, Scaled = false };

            gradientFill3.Append(gradientStopList3);
            gradientFill3.Append(linearGradientFill3);

            backgroundFillStyleList1.Append(solidFill15);
            backgroundFillStyleList1.Append(solidFill16);
            backgroundFillStyleList1.Append(gradientFill3);

            formatScheme1.Append(fillStyleList1);
            formatScheme1.Append(lineStyleList1);
            formatScheme1.Append(effectStyleList1);
            formatScheme1.Append(backgroundFillStyleList1);

            themeElements1.Append(colorScheme1);
            themeElements1.Append(fontScheme1);
            themeElements1.Append(formatScheme1);

            A.ObjectDefaults objectDefaults1 = new A.ObjectDefaults();

            A.LineDefault lineDefault1 = new A.LineDefault();
            A.ShapeProperties shapeProperties6 = new A.ShapeProperties();
            A.BodyProperties bodyProperties6 = new A.BodyProperties();
            A.ListStyle listStyle6 = new A.ListStyle();

            A.ShapeStyle shapeStyle6 = new A.ShapeStyle();

            A.LineReference lineReference55 = new A.LineReference(){ Index = (UInt32Value)2U };
            A.SchemeColor schemeColor182 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent1 };

            lineReference55.Append(schemeColor182);

            A.FillReference fillReference55 = new A.FillReference(){ Index = (UInt32Value)0U };
            A.SchemeColor schemeColor183 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent1 };

            fillReference55.Append(schemeColor183);

            A.EffectReference effectReference55 = new A.EffectReference(){ Index = (UInt32Value)1U };
            A.SchemeColor schemeColor184 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent1 };

            effectReference55.Append(schemeColor184);

            A.FontReference fontReference55 = new A.FontReference(){ Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor185 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            fontReference55.Append(schemeColor185);

            shapeStyle6.Append(lineReference55);
            shapeStyle6.Append(fillReference55);
            shapeStyle6.Append(effectReference55);
            shapeStyle6.Append(fontReference55);

            lineDefault1.Append(shapeProperties6);
            lineDefault1.Append(bodyProperties6);
            lineDefault1.Append(listStyle6);
            lineDefault1.Append(shapeStyle6);

            objectDefaults1.Append(lineDefault1);
            A.ExtraColorSchemeList extraColorSchemeList1 = new A.ExtraColorSchemeList();

            A.OfficeStyleSheetExtensionList officeStyleSheetExtensionList1 = new A.OfficeStyleSheetExtensionList();

            A.OfficeStyleSheetExtension officeStyleSheetExtension1 = new A.OfficeStyleSheetExtension(){ Uri = "{05A4C25C-085E-4340-85A3-A5531E510DB2}" };

            Thm15.ThemeFamily themeFamily1 = new Thm15.ThemeFamily(){ Name = "Office Theme", Id = "{2E142A2C-CD16-42D6-873A-C26D2A0506FA}", Vid = "{1BDDFF52-6CD6-40A5-AB3C-68EB2F1E4D0A}" };
            themeFamily1.AddNamespaceDeclaration("thm15", "http://schemas.microsoft.com/office/thememl/2012/main");

            officeStyleSheetExtension1.Append(themeFamily1);

            officeStyleSheetExtensionList1.Append(officeStyleSheetExtension1);

            theme1.Append(themeElements1);
            theme1.Append(objectDefaults1);
            theme1.Append(extraColorSchemeList1);
            theme1.Append(officeStyleSheetExtensionList1);

            themePart1.Theme = theme1;
        }

        // Generates content of diagramDataPart1.
        private void GenerateDiagramDataPart1Content(DiagramDataPart diagramDataPart1)
        {
            Dgm.DataModelRoot dataModelRoot1 = new Dgm.DataModelRoot();
            dataModelRoot1.AddNamespaceDeclaration("dgm", "http://schemas.openxmlformats.org/drawingml/2006/diagram");
            dataModelRoot1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            Dgm.PointList pointList4 = new Dgm.PointList();

            Dgm.Point point17 = new Dgm.Point(){ ModelId = "{050810F5-6B93-4502-92CB-06E52385CAF3}", Type = Dgm.PointValues.Document };
            Dgm.PropertySet propertySet6 = new Dgm.PropertySet(){ LayoutTypeId = "urn:microsoft.com/office/officeart/2005/8/layout/default", LayoutCategoryId = "list", QuickStyleTypeId = "urn:microsoft.com/office/officeart/2005/8/quickstyle/simple1", QuickStyleCategoryId = "simple", ColorType = "urn:microsoft.com/office/officeart/2005/8/colors/accent1_2", ColorCategoryId = "accent1", Placeholder = false };
            Dgm.ShapeProperties shapeProperties7 = new Dgm.ShapeProperties();

            Dgm.TextBody textBody6 = new Dgm.TextBody();
            A.BodyProperties bodyProperties7 = new A.BodyProperties();
            A.ListStyle listStyle7 = new A.ListStyle();

            A.Paragraph paragraph7 = new A.Paragraph();
            A.EndParagraphRunProperties endParagraphRunProperties6 = new A.EndParagraphRunProperties(){ Language = "pl-PL" };

            paragraph7.Append(endParagraphRunProperties6);

            textBody6.Append(bodyProperties7);
            textBody6.Append(listStyle7);
            textBody6.Append(paragraph7);

            point17.Append(propertySet6);
            point17.Append(shapeProperties7);
            point17.Append(textBody6);

            Dgm.Point point18 = new Dgm.Point(){ ModelId = "{2EEF2A58-A2D7-4991-8A78-E26574C46C74}" };
            Dgm.PropertySet propertySet7 = new Dgm.PropertySet(){ PlaceholderText = "[Text]", Placeholder = true };
            Dgm.ShapeProperties shapeProperties8 = new Dgm.ShapeProperties();

            Dgm.TextBody textBody7 = new Dgm.TextBody();
            A.BodyProperties bodyProperties8 = new A.BodyProperties();
            A.ListStyle listStyle8 = new A.ListStyle();

            A.Paragraph paragraph8 = new A.Paragraph();
            A.EndParagraphRunProperties endParagraphRunProperties7 = new A.EndParagraphRunProperties(){ Language = "pl-PL" };

            paragraph8.Append(endParagraphRunProperties7);

            textBody7.Append(bodyProperties8);
            textBody7.Append(listStyle8);
            textBody7.Append(paragraph8);

            point18.Append(propertySet7);
            point18.Append(shapeProperties8);
            point18.Append(textBody7);

            Dgm.Point point19 = new Dgm.Point(){ ModelId = "{34BEA7A6-9CDE-4021-B5E2-BA989ECE9AE2}", Type = Dgm.PointValues.ParentTransition, ConnectionId = "{53DE9593-1885-4455-BA51-60A7F54355EE}" };
            Dgm.PropertySet propertySet8 = new Dgm.PropertySet();
            Dgm.ShapeProperties shapeProperties9 = new Dgm.ShapeProperties();

            Dgm.TextBody textBody8 = new Dgm.TextBody();
            A.BodyProperties bodyProperties9 = new A.BodyProperties();
            A.ListStyle listStyle9 = new A.ListStyle();

            A.Paragraph paragraph9 = new A.Paragraph();
            A.EndParagraphRunProperties endParagraphRunProperties8 = new A.EndParagraphRunProperties(){ Language = "pl-PL" };

            paragraph9.Append(endParagraphRunProperties8);

            textBody8.Append(bodyProperties9);
            textBody8.Append(listStyle9);
            textBody8.Append(paragraph9);

            point19.Append(propertySet8);
            point19.Append(shapeProperties9);
            point19.Append(textBody8);

            Dgm.Point point20 = new Dgm.Point(){ ModelId = "{D1398D45-A4D5-4AEC-A3DC-A6C3D843196D}", Type = Dgm.PointValues.SiblingTransition, ConnectionId = "{53DE9593-1885-4455-BA51-60A7F54355EE}" };
            Dgm.PropertySet propertySet9 = new Dgm.PropertySet();
            Dgm.ShapeProperties shapeProperties10 = new Dgm.ShapeProperties();

            Dgm.TextBody textBody9 = new Dgm.TextBody();
            A.BodyProperties bodyProperties10 = new A.BodyProperties();
            A.ListStyle listStyle10 = new A.ListStyle();

            A.Paragraph paragraph10 = new A.Paragraph();
            A.EndParagraphRunProperties endParagraphRunProperties9 = new A.EndParagraphRunProperties(){ Language = "pl-PL" };

            paragraph10.Append(endParagraphRunProperties9);

            textBody9.Append(bodyProperties10);
            textBody9.Append(listStyle10);
            textBody9.Append(paragraph10);

            point20.Append(propertySet9);
            point20.Append(shapeProperties10);
            point20.Append(textBody9);

            Dgm.Point point21 = new Dgm.Point(){ ModelId = "{68641FAB-77F7-4312-BEB5-72B80B86845C}" };
            Dgm.PropertySet propertySet10 = new Dgm.PropertySet(){ PlaceholderText = "[Text]", Placeholder = true };
            Dgm.ShapeProperties shapeProperties11 = new Dgm.ShapeProperties();

            Dgm.TextBody textBody10 = new Dgm.TextBody();
            A.BodyProperties bodyProperties11 = new A.BodyProperties();
            A.ListStyle listStyle11 = new A.ListStyle();

            A.Paragraph paragraph11 = new A.Paragraph();
            A.EndParagraphRunProperties endParagraphRunProperties10 = new A.EndParagraphRunProperties(){ Language = "pl-PL" };

            paragraph11.Append(endParagraphRunProperties10);

            textBody10.Append(bodyProperties11);
            textBody10.Append(listStyle11);
            textBody10.Append(paragraph11);

            point21.Append(propertySet10);
            point21.Append(shapeProperties11);
            point21.Append(textBody10);

            Dgm.Point point22 = new Dgm.Point(){ ModelId = "{3D52860C-F7E4-494B-8C26-803853184C4F}", Type = Dgm.PointValues.ParentTransition, ConnectionId = "{20801D82-1CB0-4B72-A2A5-2604C9CC9D7E}" };
            Dgm.PropertySet propertySet11 = new Dgm.PropertySet();
            Dgm.ShapeProperties shapeProperties12 = new Dgm.ShapeProperties();

            Dgm.TextBody textBody11 = new Dgm.TextBody();
            A.BodyProperties bodyProperties12 = new A.BodyProperties();
            A.ListStyle listStyle12 = new A.ListStyle();

            A.Paragraph paragraph12 = new A.Paragraph();
            A.EndParagraphRunProperties endParagraphRunProperties11 = new A.EndParagraphRunProperties(){ Language = "pl-PL" };

            paragraph12.Append(endParagraphRunProperties11);

            textBody11.Append(bodyProperties12);
            textBody11.Append(listStyle12);
            textBody11.Append(paragraph12);

            point22.Append(propertySet11);
            point22.Append(shapeProperties12);
            point22.Append(textBody11);

            Dgm.Point point23 = new Dgm.Point(){ ModelId = "{93C814DC-C96A-464D-A307-B751C432E31D}", Type = Dgm.PointValues.SiblingTransition, ConnectionId = "{20801D82-1CB0-4B72-A2A5-2604C9CC9D7E}" };
            Dgm.PropertySet propertySet12 = new Dgm.PropertySet();
            Dgm.ShapeProperties shapeProperties13 = new Dgm.ShapeProperties();

            Dgm.TextBody textBody12 = new Dgm.TextBody();
            A.BodyProperties bodyProperties13 = new A.BodyProperties();
            A.ListStyle listStyle13 = new A.ListStyle();

            A.Paragraph paragraph13 = new A.Paragraph();
            A.EndParagraphRunProperties endParagraphRunProperties12 = new A.EndParagraphRunProperties(){ Language = "pl-PL" };

            paragraph13.Append(endParagraphRunProperties12);

            textBody12.Append(bodyProperties13);
            textBody12.Append(listStyle13);
            textBody12.Append(paragraph13);

            point23.Append(propertySet12);
            point23.Append(shapeProperties13);
            point23.Append(textBody12);

            Dgm.Point point24 = new Dgm.Point(){ ModelId = "{89391C13-C504-4B29-8FBB-561C84CC10C1}" };
            Dgm.PropertySet propertySet13 = new Dgm.PropertySet(){ PlaceholderText = "[Text]", Placeholder = true };
            Dgm.ShapeProperties shapeProperties14 = new Dgm.ShapeProperties();

            Dgm.TextBody textBody13 = new Dgm.TextBody();
            A.BodyProperties bodyProperties14 = new A.BodyProperties();
            A.ListStyle listStyle14 = new A.ListStyle();

            A.Paragraph paragraph14 = new A.Paragraph();
            A.EndParagraphRunProperties endParagraphRunProperties13 = new A.EndParagraphRunProperties(){ Language = "pl-PL" };

            paragraph14.Append(endParagraphRunProperties13);

            textBody13.Append(bodyProperties14);
            textBody13.Append(listStyle14);
            textBody13.Append(paragraph14);

            point24.Append(propertySet13);
            point24.Append(shapeProperties14);
            point24.Append(textBody13);

            Dgm.Point point25 = new Dgm.Point(){ ModelId = "{EA6D4184-8096-4199-B209-0A82B4347DDE}", Type = Dgm.PointValues.ParentTransition, ConnectionId = "{192BD615-74D4-4535-B946-F0127C2DAB21}" };
            Dgm.PropertySet propertySet14 = new Dgm.PropertySet();
            Dgm.ShapeProperties shapeProperties15 = new Dgm.ShapeProperties();

            Dgm.TextBody textBody14 = new Dgm.TextBody();
            A.BodyProperties bodyProperties15 = new A.BodyProperties();
            A.ListStyle listStyle15 = new A.ListStyle();

            A.Paragraph paragraph15 = new A.Paragraph();
            A.EndParagraphRunProperties endParagraphRunProperties14 = new A.EndParagraphRunProperties(){ Language = "pl-PL" };

            paragraph15.Append(endParagraphRunProperties14);

            textBody14.Append(bodyProperties15);
            textBody14.Append(listStyle15);
            textBody14.Append(paragraph15);

            point25.Append(propertySet14);
            point25.Append(shapeProperties15);
            point25.Append(textBody14);

            Dgm.Point point26 = new Dgm.Point(){ ModelId = "{839AC1D3-AFF4-44E4-AD2D-D35D5C61E303}", Type = Dgm.PointValues.SiblingTransition, ConnectionId = "{192BD615-74D4-4535-B946-F0127C2DAB21}" };
            Dgm.PropertySet propertySet15 = new Dgm.PropertySet();
            Dgm.ShapeProperties shapeProperties16 = new Dgm.ShapeProperties();

            Dgm.TextBody textBody15 = new Dgm.TextBody();
            A.BodyProperties bodyProperties16 = new A.BodyProperties();
            A.ListStyle listStyle16 = new A.ListStyle();

            A.Paragraph paragraph16 = new A.Paragraph();
            A.EndParagraphRunProperties endParagraphRunProperties15 = new A.EndParagraphRunProperties(){ Language = "pl-PL" };

            paragraph16.Append(endParagraphRunProperties15);

            textBody15.Append(bodyProperties16);
            textBody15.Append(listStyle16);
            textBody15.Append(paragraph16);

            point26.Append(propertySet15);
            point26.Append(shapeProperties16);
            point26.Append(textBody15);

            Dgm.Point point27 = new Dgm.Point(){ ModelId = "{575655FE-BCCE-4492-8A81-F5E8A3A92F30}" };
            Dgm.PropertySet propertySet16 = new Dgm.PropertySet(){ PlaceholderText = "[Text]", Placeholder = true };
            Dgm.ShapeProperties shapeProperties17 = new Dgm.ShapeProperties();

            Dgm.TextBody textBody16 = new Dgm.TextBody();
            A.BodyProperties bodyProperties17 = new A.BodyProperties();
            A.ListStyle listStyle17 = new A.ListStyle();

            A.Paragraph paragraph17 = new A.Paragraph();
            A.EndParagraphRunProperties endParagraphRunProperties16 = new A.EndParagraphRunProperties(){ Language = "pl-PL" };

            paragraph17.Append(endParagraphRunProperties16);

            textBody16.Append(bodyProperties17);
            textBody16.Append(listStyle17);
            textBody16.Append(paragraph17);

            point27.Append(propertySet16);
            point27.Append(shapeProperties17);
            point27.Append(textBody16);

            Dgm.Point point28 = new Dgm.Point(){ ModelId = "{5DCAA827-5E19-4785-BA24-36EDD3800D76}", Type = Dgm.PointValues.ParentTransition, ConnectionId = "{5B0B9687-F94F-40E3-BC77-4E2284D67B82}" };
            Dgm.PropertySet propertySet17 = new Dgm.PropertySet();
            Dgm.ShapeProperties shapeProperties18 = new Dgm.ShapeProperties();

            Dgm.TextBody textBody17 = new Dgm.TextBody();
            A.BodyProperties bodyProperties18 = new A.BodyProperties();
            A.ListStyle listStyle18 = new A.ListStyle();

            A.Paragraph paragraph18 = new A.Paragraph();
            A.EndParagraphRunProperties endParagraphRunProperties17 = new A.EndParagraphRunProperties(){ Language = "pl-PL" };

            paragraph18.Append(endParagraphRunProperties17);

            textBody17.Append(bodyProperties18);
            textBody17.Append(listStyle18);
            textBody17.Append(paragraph18);

            point28.Append(propertySet17);
            point28.Append(shapeProperties18);
            point28.Append(textBody17);

            Dgm.Point point29 = new Dgm.Point(){ ModelId = "{D9D42E39-AEF2-4B6C-BD25-361A6F0D8CB2}", Type = Dgm.PointValues.SiblingTransition, ConnectionId = "{5B0B9687-F94F-40E3-BC77-4E2284D67B82}" };
            Dgm.PropertySet propertySet18 = new Dgm.PropertySet();
            Dgm.ShapeProperties shapeProperties19 = new Dgm.ShapeProperties();

            Dgm.TextBody textBody18 = new Dgm.TextBody();
            A.BodyProperties bodyProperties19 = new A.BodyProperties();
            A.ListStyle listStyle19 = new A.ListStyle();

            A.Paragraph paragraph19 = new A.Paragraph();
            A.EndParagraphRunProperties endParagraphRunProperties18 = new A.EndParagraphRunProperties(){ Language = "pl-PL" };

            paragraph19.Append(endParagraphRunProperties18);

            textBody18.Append(bodyProperties19);
            textBody18.Append(listStyle19);
            textBody18.Append(paragraph19);

            point29.Append(propertySet18);
            point29.Append(shapeProperties19);
            point29.Append(textBody18);

            Dgm.Point point30 = new Dgm.Point(){ ModelId = "{17A72E10-5B50-4595-A1C0-8605F16EB379}" };
            Dgm.PropertySet propertySet19 = new Dgm.PropertySet(){ PlaceholderText = "[Text]", Placeholder = true };
            Dgm.ShapeProperties shapeProperties20 = new Dgm.ShapeProperties();

            Dgm.TextBody textBody19 = new Dgm.TextBody();
            A.BodyProperties bodyProperties20 = new A.BodyProperties();
            A.ListStyle listStyle20 = new A.ListStyle();

            A.Paragraph paragraph20 = new A.Paragraph();
            A.EndParagraphRunProperties endParagraphRunProperties19 = new A.EndParagraphRunProperties(){ Language = "pl-PL" };

            paragraph20.Append(endParagraphRunProperties19);

            textBody19.Append(bodyProperties20);
            textBody19.Append(listStyle20);
            textBody19.Append(paragraph20);

            point30.Append(propertySet19);
            point30.Append(shapeProperties20);
            point30.Append(textBody19);

            Dgm.Point point31 = new Dgm.Point(){ ModelId = "{55D66658-965D-44CC-ABAE-6638BD3536C3}", Type = Dgm.PointValues.ParentTransition, ConnectionId = "{BD2FFBF0-81A7-4151-B2A4-72895B9DAC55}" };
            Dgm.PropertySet propertySet20 = new Dgm.PropertySet();
            Dgm.ShapeProperties shapeProperties21 = new Dgm.ShapeProperties();

            Dgm.TextBody textBody20 = new Dgm.TextBody();
            A.BodyProperties bodyProperties21 = new A.BodyProperties();
            A.ListStyle listStyle21 = new A.ListStyle();

            A.Paragraph paragraph21 = new A.Paragraph();
            A.EndParagraphRunProperties endParagraphRunProperties20 = new A.EndParagraphRunProperties(){ Language = "pl-PL" };

            paragraph21.Append(endParagraphRunProperties20);

            textBody20.Append(bodyProperties21);
            textBody20.Append(listStyle21);
            textBody20.Append(paragraph21);

            point31.Append(propertySet20);
            point31.Append(shapeProperties21);
            point31.Append(textBody20);

            Dgm.Point point32 = new Dgm.Point(){ ModelId = "{0ADD6561-125E-4ED0-A0C2-6BC4566DCBA6}", Type = Dgm.PointValues.SiblingTransition, ConnectionId = "{BD2FFBF0-81A7-4151-B2A4-72895B9DAC55}" };
            Dgm.PropertySet propertySet21 = new Dgm.PropertySet();
            Dgm.ShapeProperties shapeProperties22 = new Dgm.ShapeProperties();

            Dgm.TextBody textBody21 = new Dgm.TextBody();
            A.BodyProperties bodyProperties22 = new A.BodyProperties();
            A.ListStyle listStyle22 = new A.ListStyle();

            A.Paragraph paragraph22 = new A.Paragraph();
            A.EndParagraphRunProperties endParagraphRunProperties21 = new A.EndParagraphRunProperties(){ Language = "pl-PL" };

            paragraph22.Append(endParagraphRunProperties21);

            textBody21.Append(bodyProperties22);
            textBody21.Append(listStyle22);
            textBody21.Append(paragraph22);

            point32.Append(propertySet21);
            point32.Append(shapeProperties22);
            point32.Append(textBody21);

            Dgm.Point point33 = new Dgm.Point(){ ModelId = "{32306AFD-D27D-4BBE-847F-5EACAA1E64CC}", Type = Dgm.PointValues.Presentation };

            Dgm.PropertySet propertySet22 = new Dgm.PropertySet(){ PresentationElementId = "{050810F5-6B93-4502-92CB-06E52385CAF3}", PresentationName = "diagram", PresentationStyleCount = 0 };

            Dgm.PresentationLayoutVariables presentationLayoutVariables1 = new Dgm.PresentationLayoutVariables();
            Dgm.Direction direction2 = new Dgm.Direction();
            Dgm.ResizeHandles resizeHandles2 = new Dgm.ResizeHandles(){ Val = Dgm.ResizeHandlesStringValues.Exact };

            presentationLayoutVariables1.Append(direction2);
            presentationLayoutVariables1.Append(resizeHandles2);

            propertySet22.Append(presentationLayoutVariables1);
            Dgm.ShapeProperties shapeProperties23 = new Dgm.ShapeProperties();

            point33.Append(propertySet22);
            point33.Append(shapeProperties23);

            Dgm.Point point34 = new Dgm.Point(){ ModelId = "{068D8084-B4F1-4349-BF7B-3A540F7ACE9A}", Type = Dgm.PointValues.Presentation };

            Dgm.PropertySet propertySet23 = new Dgm.PropertySet(){ PresentationElementId = "{2EEF2A58-A2D7-4991-8A78-E26574C46C74}", PresentationName = "node", PresentationStyleLabel = "node1", PresentationStyleIndex = 0, PresentationStyleCount = 5 };

            Dgm.PresentationLayoutVariables presentationLayoutVariables2 = new Dgm.PresentationLayoutVariables();
            Dgm.BulletEnabled bulletEnabled2 = new Dgm.BulletEnabled(){ Val = true };

            presentationLayoutVariables2.Append(bulletEnabled2);

            propertySet23.Append(presentationLayoutVariables2);
            Dgm.ShapeProperties shapeProperties24 = new Dgm.ShapeProperties();

            point34.Append(propertySet23);
            point34.Append(shapeProperties24);

            Dgm.Point point35 = new Dgm.Point(){ ModelId = "{BB3DC0D3-305A-444F-A1F6-B180204939B5}", Type = Dgm.PointValues.Presentation };
            Dgm.PropertySet propertySet24 = new Dgm.PropertySet(){ PresentationElementId = "{D1398D45-A4D5-4AEC-A3DC-A6C3D843196D}", PresentationName = "sibTrans", PresentationStyleCount = 0 };
            Dgm.ShapeProperties shapeProperties25 = new Dgm.ShapeProperties();

            point35.Append(propertySet24);
            point35.Append(shapeProperties25);

            Dgm.Point point36 = new Dgm.Point(){ ModelId = "{DF06976E-7188-463E-AE39-BDA19617EFC4}", Type = Dgm.PointValues.Presentation };

            Dgm.PropertySet propertySet25 = new Dgm.PropertySet(){ PresentationElementId = "{68641FAB-77F7-4312-BEB5-72B80B86845C}", PresentationName = "node", PresentationStyleLabel = "node1", PresentationStyleIndex = 1, PresentationStyleCount = 5 };

            Dgm.PresentationLayoutVariables presentationLayoutVariables3 = new Dgm.PresentationLayoutVariables();
            Dgm.BulletEnabled bulletEnabled3 = new Dgm.BulletEnabled(){ Val = true };

            presentationLayoutVariables3.Append(bulletEnabled3);

            propertySet25.Append(presentationLayoutVariables3);
            Dgm.ShapeProperties shapeProperties26 = new Dgm.ShapeProperties();

            point36.Append(propertySet25);
            point36.Append(shapeProperties26);

            Dgm.Point point37 = new Dgm.Point(){ ModelId = "{2AD38D92-7E4A-4864-A84B-BD1C732A96A3}", Type = Dgm.PointValues.Presentation };
            Dgm.PropertySet propertySet26 = new Dgm.PropertySet(){ PresentationElementId = "{93C814DC-C96A-464D-A307-B751C432E31D}", PresentationName = "sibTrans", PresentationStyleCount = 0 };
            Dgm.ShapeProperties shapeProperties27 = new Dgm.ShapeProperties();

            point37.Append(propertySet26);
            point37.Append(shapeProperties27);

            Dgm.Point point38 = new Dgm.Point(){ ModelId = "{73C7BCEA-927D-4615-95E9-BB89F1A66540}", Type = Dgm.PointValues.Presentation };

            Dgm.PropertySet propertySet27 = new Dgm.PropertySet(){ PresentationElementId = "{89391C13-C504-4B29-8FBB-561C84CC10C1}", PresentationName = "node", PresentationStyleLabel = "node1", PresentationStyleIndex = 2, PresentationStyleCount = 5 };

            Dgm.PresentationLayoutVariables presentationLayoutVariables4 = new Dgm.PresentationLayoutVariables();
            Dgm.BulletEnabled bulletEnabled4 = new Dgm.BulletEnabled(){ Val = true };

            presentationLayoutVariables4.Append(bulletEnabled4);

            propertySet27.Append(presentationLayoutVariables4);
            Dgm.ShapeProperties shapeProperties28 = new Dgm.ShapeProperties();

            point38.Append(propertySet27);
            point38.Append(shapeProperties28);

            Dgm.Point point39 = new Dgm.Point(){ ModelId = "{0052E64C-8063-4397-B08E-4468F28ADC78}", Type = Dgm.PointValues.Presentation };
            Dgm.PropertySet propertySet28 = new Dgm.PropertySet(){ PresentationElementId = "{839AC1D3-AFF4-44E4-AD2D-D35D5C61E303}", PresentationName = "sibTrans", PresentationStyleCount = 0 };
            Dgm.ShapeProperties shapeProperties29 = new Dgm.ShapeProperties();

            point39.Append(propertySet28);
            point39.Append(shapeProperties29);

            Dgm.Point point40 = new Dgm.Point(){ ModelId = "{72A7A719-1E8E-46AB-8256-BB41F6817818}", Type = Dgm.PointValues.Presentation };

            Dgm.PropertySet propertySet29 = new Dgm.PropertySet(){ PresentationElementId = "{575655FE-BCCE-4492-8A81-F5E8A3A92F30}", PresentationName = "node", PresentationStyleLabel = "node1", PresentationStyleIndex = 3, PresentationStyleCount = 5 };

            Dgm.PresentationLayoutVariables presentationLayoutVariables5 = new Dgm.PresentationLayoutVariables();
            Dgm.BulletEnabled bulletEnabled5 = new Dgm.BulletEnabled(){ Val = true };

            presentationLayoutVariables5.Append(bulletEnabled5);

            propertySet29.Append(presentationLayoutVariables5);
            Dgm.ShapeProperties shapeProperties30 = new Dgm.ShapeProperties();

            point40.Append(propertySet29);
            point40.Append(shapeProperties30);

            Dgm.Point point41 = new Dgm.Point(){ ModelId = "{465E4E3A-A29E-4172-B02E-CDF8A89566C6}", Type = Dgm.PointValues.Presentation };
            Dgm.PropertySet propertySet30 = new Dgm.PropertySet(){ PresentationElementId = "{D9D42E39-AEF2-4B6C-BD25-361A6F0D8CB2}", PresentationName = "sibTrans", PresentationStyleCount = 0 };
            Dgm.ShapeProperties shapeProperties31 = new Dgm.ShapeProperties();

            point41.Append(propertySet30);
            point41.Append(shapeProperties31);

            Dgm.Point point42 = new Dgm.Point(){ ModelId = "{2B5DA877-174D-4060-8E30-014EB5090235}", Type = Dgm.PointValues.Presentation };

            Dgm.PropertySet propertySet31 = new Dgm.PropertySet(){ PresentationElementId = "{17A72E10-5B50-4595-A1C0-8605F16EB379}", PresentationName = "node", PresentationStyleLabel = "node1", PresentationStyleIndex = 4, PresentationStyleCount = 5 };

            Dgm.PresentationLayoutVariables presentationLayoutVariables6 = new Dgm.PresentationLayoutVariables();
            Dgm.BulletEnabled bulletEnabled6 = new Dgm.BulletEnabled(){ Val = true };

            presentationLayoutVariables6.Append(bulletEnabled6);

            propertySet31.Append(presentationLayoutVariables6);
            Dgm.ShapeProperties shapeProperties32 = new Dgm.ShapeProperties();

            point42.Append(propertySet31);
            point42.Append(shapeProperties32);

            pointList4.Append(point17);
            pointList4.Append(point18);
            pointList4.Append(point19);
            pointList4.Append(point20);
            pointList4.Append(point21);
            pointList4.Append(point22);
            pointList4.Append(point23);
            pointList4.Append(point24);
            pointList4.Append(point25);
            pointList4.Append(point26);
            pointList4.Append(point27);
            pointList4.Append(point28);
            pointList4.Append(point29);
            pointList4.Append(point30);
            pointList4.Append(point31);
            pointList4.Append(point32);
            pointList4.Append(point33);
            pointList4.Append(point34);
            pointList4.Append(point35);
            pointList4.Append(point36);
            pointList4.Append(point37);
            pointList4.Append(point38);
            pointList4.Append(point39);
            pointList4.Append(point40);
            pointList4.Append(point41);
            pointList4.Append(point42);

            Dgm.ConnectionList connectionList4 = new Dgm.ConnectionList();
            Dgm.Connection connection14 = new Dgm.Connection(){ ModelId = "{192BD615-74D4-4535-B946-F0127C2DAB21}", SourceId = "{050810F5-6B93-4502-92CB-06E52385CAF3}", DestinationId = "{89391C13-C504-4B29-8FBB-561C84CC10C1}", SourcePosition = (UInt32Value)2U, DestinationPosition = (UInt32Value)0U, ParentTransitionId = "{EA6D4184-8096-4199-B209-0A82B4347DDE}", SiblingTransitionId = "{839AC1D3-AFF4-44E4-AD2D-D35D5C61E303}" };
            Dgm.Connection connection15 = new Dgm.Connection(){ ModelId = "{C7D7632F-5F3B-4615-BED0-E4F654AF3976}", Type = Dgm.ConnectionValues.PresentationOf, SourceId = "{68641FAB-77F7-4312-BEB5-72B80B86845C}", DestinationId = "{DF06976E-7188-463E-AE39-BDA19617EFC4}", SourcePosition = (UInt32Value)0U, DestinationPosition = (UInt32Value)0U, PresentationId = "urn:microsoft.com/office/officeart/2005/8/layout/default" };
            Dgm.Connection connection16 = new Dgm.Connection(){ ModelId = "{3CB88063-C5B8-452B-952D-69316D1CF1E1}", Type = Dgm.ConnectionValues.PresentationOf, SourceId = "{89391C13-C504-4B29-8FBB-561C84CC10C1}", DestinationId = "{73C7BCEA-927D-4615-95E9-BB89F1A66540}", SourcePosition = (UInt32Value)0U, DestinationPosition = (UInt32Value)0U, PresentationId = "urn:microsoft.com/office/officeart/2005/8/layout/default" };
            Dgm.Connection connection17 = new Dgm.Connection(){ ModelId = "{A0235465-ACC9-4B56-A649-2C0B585E7AC2}", Type = Dgm.ConnectionValues.PresentationOf, SourceId = "{17A72E10-5B50-4595-A1C0-8605F16EB379}", DestinationId = "{2B5DA877-174D-4060-8E30-014EB5090235}", SourcePosition = (UInt32Value)0U, DestinationPosition = (UInt32Value)0U, PresentationId = "urn:microsoft.com/office/officeart/2005/8/layout/default" };
            Dgm.Connection connection18 = new Dgm.Connection(){ ModelId = "{20801D82-1CB0-4B72-A2A5-2604C9CC9D7E}", SourceId = "{050810F5-6B93-4502-92CB-06E52385CAF3}", DestinationId = "{68641FAB-77F7-4312-BEB5-72B80B86845C}", SourcePosition = (UInt32Value)1U, DestinationPosition = (UInt32Value)0U, ParentTransitionId = "{3D52860C-F7E4-494B-8C26-803853184C4F}", SiblingTransitionId = "{93C814DC-C96A-464D-A307-B751C432E31D}" };
            Dgm.Connection connection19 = new Dgm.Connection(){ ModelId = "{5B0B9687-F94F-40E3-BC77-4E2284D67B82}", SourceId = "{050810F5-6B93-4502-92CB-06E52385CAF3}", DestinationId = "{575655FE-BCCE-4492-8A81-F5E8A3A92F30}", SourcePosition = (UInt32Value)3U, DestinationPosition = (UInt32Value)0U, ParentTransitionId = "{5DCAA827-5E19-4785-BA24-36EDD3800D76}", SiblingTransitionId = "{D9D42E39-AEF2-4B6C-BD25-361A6F0D8CB2}" };
            Dgm.Connection connection20 = new Dgm.Connection(){ ModelId = "{53DE9593-1885-4455-BA51-60A7F54355EE}", SourceId = "{050810F5-6B93-4502-92CB-06E52385CAF3}", DestinationId = "{2EEF2A58-A2D7-4991-8A78-E26574C46C74}", SourcePosition = (UInt32Value)0U, DestinationPosition = (UInt32Value)0U, ParentTransitionId = "{34BEA7A6-9CDE-4021-B5E2-BA989ECE9AE2}", SiblingTransitionId = "{D1398D45-A4D5-4AEC-A3DC-A6C3D843196D}" };
            Dgm.Connection connection21 = new Dgm.Connection(){ ModelId = "{94F827CB-37AE-4FCF-AF0D-07027C70578E}", Type = Dgm.ConnectionValues.PresentationOf, SourceId = "{050810F5-6B93-4502-92CB-06E52385CAF3}", DestinationId = "{32306AFD-D27D-4BBE-847F-5EACAA1E64CC}", SourcePosition = (UInt32Value)0U, DestinationPosition = (UInt32Value)0U, PresentationId = "urn:microsoft.com/office/officeart/2005/8/layout/default" };
            Dgm.Connection connection22 = new Dgm.Connection(){ ModelId = "{804E90DD-8FAB-4442-AFA0-589FCD05937D}", Type = Dgm.ConnectionValues.PresentationOf, SourceId = "{2EEF2A58-A2D7-4991-8A78-E26574C46C74}", DestinationId = "{068D8084-B4F1-4349-BF7B-3A540F7ACE9A}", SourcePosition = (UInt32Value)0U, DestinationPosition = (UInt32Value)0U, PresentationId = "urn:microsoft.com/office/officeart/2005/8/layout/default" };
            Dgm.Connection connection23 = new Dgm.Connection(){ ModelId = "{F1CC3DEA-5D9E-4559-9F44-D957E04D3B0C}", Type = Dgm.ConnectionValues.PresentationOf, SourceId = "{575655FE-BCCE-4492-8A81-F5E8A3A92F30}", DestinationId = "{72A7A719-1E8E-46AB-8256-BB41F6817818}", SourcePosition = (UInt32Value)0U, DestinationPosition = (UInt32Value)0U, PresentationId = "urn:microsoft.com/office/officeart/2005/8/layout/default" };
            Dgm.Connection connection24 = new Dgm.Connection(){ ModelId = "{BD2FFBF0-81A7-4151-B2A4-72895B9DAC55}", SourceId = "{050810F5-6B93-4502-92CB-06E52385CAF3}", DestinationId = "{17A72E10-5B50-4595-A1C0-8605F16EB379}", SourcePosition = (UInt32Value)4U, DestinationPosition = (UInt32Value)0U, ParentTransitionId = "{55D66658-965D-44CC-ABAE-6638BD3536C3}", SiblingTransitionId = "{0ADD6561-125E-4ED0-A0C2-6BC4566DCBA6}" };
            Dgm.Connection connection25 = new Dgm.Connection(){ ModelId = "{3AD65271-84D5-4902-B189-D664DAD6CE2C}", Type = Dgm.ConnectionValues.PresentationParentOf, SourceId = "{32306AFD-D27D-4BBE-847F-5EACAA1E64CC}", DestinationId = "{068D8084-B4F1-4349-BF7B-3A540F7ACE9A}", SourcePosition = (UInt32Value)0U, DestinationPosition = (UInt32Value)0U, PresentationId = "urn:microsoft.com/office/officeart/2005/8/layout/default" };
            Dgm.Connection connection26 = new Dgm.Connection(){ ModelId = "{392E2E2C-F8C4-44C7-969F-94074A656118}", Type = Dgm.ConnectionValues.PresentationParentOf, SourceId = "{32306AFD-D27D-4BBE-847F-5EACAA1E64CC}", DestinationId = "{BB3DC0D3-305A-444F-A1F6-B180204939B5}", SourcePosition = (UInt32Value)1U, DestinationPosition = (UInt32Value)0U, PresentationId = "urn:microsoft.com/office/officeart/2005/8/layout/default" };
            Dgm.Connection connection27 = new Dgm.Connection(){ ModelId = "{C02074ED-D2E7-48FE-88AB-AB6203803B84}", Type = Dgm.ConnectionValues.PresentationParentOf, SourceId = "{32306AFD-D27D-4BBE-847F-5EACAA1E64CC}", DestinationId = "{DF06976E-7188-463E-AE39-BDA19617EFC4}", SourcePosition = (UInt32Value)2U, DestinationPosition = (UInt32Value)0U, PresentationId = "urn:microsoft.com/office/officeart/2005/8/layout/default" };
            Dgm.Connection connection28 = new Dgm.Connection(){ ModelId = "{BF3C97F1-587E-400F-8FD5-55B2040136D5}", Type = Dgm.ConnectionValues.PresentationParentOf, SourceId = "{32306AFD-D27D-4BBE-847F-5EACAA1E64CC}", DestinationId = "{2AD38D92-7E4A-4864-A84B-BD1C732A96A3}", SourcePosition = (UInt32Value)3U, DestinationPosition = (UInt32Value)0U, PresentationId = "urn:microsoft.com/office/officeart/2005/8/layout/default" };
            Dgm.Connection connection29 = new Dgm.Connection(){ ModelId = "{5AABD672-643B-4279-86F2-A5FCBC7C7E8E}", Type = Dgm.ConnectionValues.PresentationParentOf, SourceId = "{32306AFD-D27D-4BBE-847F-5EACAA1E64CC}", DestinationId = "{73C7BCEA-927D-4615-95E9-BB89F1A66540}", SourcePosition = (UInt32Value)4U, DestinationPosition = (UInt32Value)0U, PresentationId = "urn:microsoft.com/office/officeart/2005/8/layout/default" };
            Dgm.Connection connection30 = new Dgm.Connection(){ ModelId = "{EEC3D173-2022-4A1A-AF01-1B5182E31CDB}", Type = Dgm.ConnectionValues.PresentationParentOf, SourceId = "{32306AFD-D27D-4BBE-847F-5EACAA1E64CC}", DestinationId = "{0052E64C-8063-4397-B08E-4468F28ADC78}", SourcePosition = (UInt32Value)5U, DestinationPosition = (UInt32Value)0U, PresentationId = "urn:microsoft.com/office/officeart/2005/8/layout/default" };
            Dgm.Connection connection31 = new Dgm.Connection(){ ModelId = "{A5E31740-A371-4B04-B1BC-9D96F83815BD}", Type = Dgm.ConnectionValues.PresentationParentOf, SourceId = "{32306AFD-D27D-4BBE-847F-5EACAA1E64CC}", DestinationId = "{72A7A719-1E8E-46AB-8256-BB41F6817818}", SourcePosition = (UInt32Value)6U, DestinationPosition = (UInt32Value)0U, PresentationId = "urn:microsoft.com/office/officeart/2005/8/layout/default" };
            Dgm.Connection connection32 = new Dgm.Connection(){ ModelId = "{167CF112-754F-4EC8-BD10-A11D3BA8283D}", Type = Dgm.ConnectionValues.PresentationParentOf, SourceId = "{32306AFD-D27D-4BBE-847F-5EACAA1E64CC}", DestinationId = "{465E4E3A-A29E-4172-B02E-CDF8A89566C6}", SourcePosition = (UInt32Value)7U, DestinationPosition = (UInt32Value)0U, PresentationId = "urn:microsoft.com/office/officeart/2005/8/layout/default" };
            Dgm.Connection connection33 = new Dgm.Connection(){ ModelId = "{E9F009DD-E549-41DA-832C-6317973486CD}", Type = Dgm.ConnectionValues.PresentationParentOf, SourceId = "{32306AFD-D27D-4BBE-847F-5EACAA1E64CC}", DestinationId = "{2B5DA877-174D-4060-8E30-014EB5090235}", SourcePosition = (UInt32Value)8U, DestinationPosition = (UInt32Value)0U, PresentationId = "urn:microsoft.com/office/officeart/2005/8/layout/default" };

            connectionList4.Append(connection14);
            connectionList4.Append(connection15);
            connectionList4.Append(connection16);
            connectionList4.Append(connection17);
            connectionList4.Append(connection18);
            connectionList4.Append(connection19);
            connectionList4.Append(connection20);
            connectionList4.Append(connection21);
            connectionList4.Append(connection22);
            connectionList4.Append(connection23);
            connectionList4.Append(connection24);
            connectionList4.Append(connection25);
            connectionList4.Append(connection26);
            connectionList4.Append(connection27);
            connectionList4.Append(connection28);
            connectionList4.Append(connection29);
            connectionList4.Append(connection30);
            connectionList4.Append(connection31);
            connectionList4.Append(connection32);
            connectionList4.Append(connection33);
            Dgm.Background background4 = new Dgm.Background();
            Dgm.Whole whole4 = new Dgm.Whole();

            Dgm.DataModelExtensionList dataModelExtensionList1 = new Dgm.DataModelExtensionList();

            A.DataModelExtension dataModelExtension1 = new A.DataModelExtension(){ Uri = "http://schemas.microsoft.com/office/drawing/2008/diagram" };

            Dsp.DataModelExtensionBlock dataModelExtensionBlock1 = new Dsp.DataModelExtensionBlock(){ RelId = "rId8", MinVer = "http://schemas.openxmlformats.org/drawingml/2006/diagram" };
            dataModelExtensionBlock1.AddNamespaceDeclaration("dsp", "http://schemas.microsoft.com/office/drawing/2008/diagram");

            dataModelExtension1.Append(dataModelExtensionBlock1);

            dataModelExtensionList1.Append(dataModelExtension1);

            dataModelRoot1.Append(pointList4);
            dataModelRoot1.Append(connectionList4);
            dataModelRoot1.Append(background4);
            dataModelRoot1.Append(whole4);
            dataModelRoot1.Append(dataModelExtensionList1);

            diagramDataPart1.DataModelRoot = dataModelRoot1;
        }

        // Generates content of fontTablePart1.
        private void GenerateFontTablePart1Content(FontTablePart fontTablePart1)
        {
            Fonts fonts1 = new Fonts(){ MCAttributes = new MarkupCompatibilityAttributes(){ Ignorable = "w14 w15 w16se w16cid w16 w16cex w16sdtdh w16sdtfl w16du" }  };
            fonts1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            fonts1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            fonts1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            fonts1.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            fonts1.AddNamespaceDeclaration("w15", "http://schemas.microsoft.com/office/word/2012/wordml");
            fonts1.AddNamespaceDeclaration("w16cex", "http://schemas.microsoft.com/office/word/2018/wordml/cex");
            fonts1.AddNamespaceDeclaration("w16cid", "http://schemas.microsoft.com/office/word/2016/wordml/cid");
            fonts1.AddNamespaceDeclaration("w16", "http://schemas.microsoft.com/office/word/2018/wordml");
            fonts1.AddNamespaceDeclaration("w16du", "http://schemas.microsoft.com/office/word/2023/wordml/word16du");
            fonts1.AddNamespaceDeclaration("w16sdtdh", "http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash");
            fonts1.AddNamespaceDeclaration("w16sdtfl", "http://schemas.microsoft.com/office/word/2024/wordml/sdtformatlock");
            fonts1.AddNamespaceDeclaration("w16se", "http://schemas.microsoft.com/office/word/2015/wordml/symex");

            Font font1 = new Font(){ Name = "Aptos" };
            FontCharSet fontCharSet1 = new FontCharSet(){ Val = "00" };
            FontFamily fontFamily1 = new FontFamily(){ Val = FontFamilyValues.Swiss };
            Pitch pitch1 = new Pitch(){ Val = FontPitchValues.Variable };
            FontSignature fontSignature1 = new FontSignature(){ UnicodeSignature0 = "20000287", UnicodeSignature1 = "00000003", UnicodeSignature2 = "00000000", UnicodeSignature3 = "00000000", CodePageSignature0 = "0000019F", CodePageSignature1 = "00000000" };

            font1.Append(fontCharSet1);
            font1.Append(fontFamily1);
            font1.Append(pitch1);
            font1.Append(fontSignature1);

            Font font2 = new Font(){ Name = "Times New Roman" };
            Panose1Number panose1Number1 = new Panose1Number(){ Val = "02020603050405020304" };
            FontCharSet fontCharSet2 = new FontCharSet(){ Val = "00" };
            FontFamily fontFamily2 = new FontFamily(){ Val = FontFamilyValues.Roman };
            Pitch pitch2 = new Pitch(){ Val = FontPitchValues.Variable };
            FontSignature fontSignature2 = new FontSignature(){ UnicodeSignature0 = "E0002EFF", UnicodeSignature1 = "C000785B", UnicodeSignature2 = "00000009", UnicodeSignature3 = "00000000", CodePageSignature0 = "000001FF", CodePageSignature1 = "00000000" };

            font2.Append(panose1Number1);
            font2.Append(fontCharSet2);
            font2.Append(fontFamily2);
            font2.Append(pitch2);
            font2.Append(fontSignature2);

            Font font3 = new Font(){ Name = "Aptos Display" };
            FontCharSet fontCharSet3 = new FontCharSet(){ Val = "00" };
            FontFamily fontFamily3 = new FontFamily(){ Val = FontFamilyValues.Swiss };
            Pitch pitch3 = new Pitch(){ Val = FontPitchValues.Variable };
            FontSignature fontSignature3 = new FontSignature(){ UnicodeSignature0 = "20000287", UnicodeSignature1 = "00000003", UnicodeSignature2 = "00000000", UnicodeSignature3 = "00000000", CodePageSignature0 = "0000019F", CodePageSignature1 = "00000000" };

            font3.Append(fontCharSet3);
            font3.Append(fontFamily3);
            font3.Append(pitch3);
            font3.Append(fontSignature3);

            fonts1.Append(font1);
            fonts1.Append(font2);
            fonts1.Append(font3);

            fontTablePart1.Fonts = fonts1;
        }

        private void SetPackageProperties(OpenXmlPackage document)
        {
            document.PackageProperties.Creator = "Przemysław Kłys";
            document.PackageProperties.Title = "";
            document.PackageProperties.Subject = "";
            document.PackageProperties.Keywords = "";
            document.PackageProperties.Description = "";
            document.PackageProperties.Revision = "1";
            document.PackageProperties.Created = System.Xml.XmlConvert.ToDateTime("2025-09-01T07:02:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
            document.PackageProperties.Modified = System.Xml.XmlConvert.ToDateTime("2025-09-01T07:03:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
            document.PackageProperties.LastModifiedBy = "Przemysław Kłys";
        }


    }
}
