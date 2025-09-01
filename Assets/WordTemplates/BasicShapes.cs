using DocumentFormat.OpenXml.Packaging;
using Ap = DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using Wp = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using Wps = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;
using V = DocumentFormat.OpenXml.Vml;
using Ovml = DocumentFormat.OpenXml.Vml.Office;
using M = DocumentFormat.OpenXml.Math;
using W14 = DocumentFormat.OpenXml.Office2010.Word;
using W15 = DocumentFormat.OpenXml.Office2013.Word;
using Ds = DocumentFormat.OpenXml.CustomXmlDataProperties;
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

            DocumentSettingsPart documentSettingsPart1 = mainDocumentPart1.AddNewPart<DocumentSettingsPart>("rId3");
            GenerateDocumentSettingsPart1Content(documentSettingsPart1);

            StyleDefinitionsPart styleDefinitionsPart1 = mainDocumentPart1.AddNewPart<StyleDefinitionsPart>("rId2");
            GenerateStyleDefinitionsPart1Content(styleDefinitionsPart1);

            CustomXmlPart customXmlPart1 = mainDocumentPart1.AddNewPart<CustomXmlPart>("application/xml", "rId1");
            GenerateCustomXmlPart1Content(customXmlPart1);

            CustomXmlPropertiesPart customXmlPropertiesPart1 = customXmlPart1.AddNewPart<CustomXmlPropertiesPart>("rId1");
            GenerateCustomXmlPropertiesPart1Content(customXmlPropertiesPart1);

            ThemePart themePart1 = mainDocumentPart1.AddNewPart<ThemePart>("rId6");
            GenerateThemePart1Content(themePart1);

            FontTablePart fontTablePart1 = mainDocumentPart1.AddNewPart<FontTablePart>("rId5");
            GenerateFontTablePart1Content(fontTablePart1);

            WebSettingsPart webSettingsPart1 = mainDocumentPart1.AddNewPart<WebSettingsPart>("rId4");
            GenerateWebSettingsPart1Content(webSettingsPart1);

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
            characters1.Text = "3";
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
            charactersWithSpaces1.Text = "3";
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

            Paragraph paragraph1 = new Paragraph(){ RsidParagraphAddition = "0018596D", RsidRunAdditionDefault = "007719D7", ParagraphId = "544A6021", TextId = "2B801335" };

            Run run1 = new Run();

            RunProperties runProperties1 = new RunProperties();
            NoProof noProof1 = new NoProof();

            runProperties1.Append(noProof1);

            AlternateContent alternateContent1 = new AlternateContent();

            AlternateContentChoice alternateContentChoice1 = new AlternateContentChoice(){ Requires = "wps" };

            Drawing drawing1 = new Drawing();

            Wp.Anchor anchor1 = new Wp.Anchor(){ DistanceFromTop = (UInt32Value)0U, DistanceFromBottom = (UInt32Value)0U, DistanceFromLeft = (UInt32Value)114300U, DistanceFromRight = (UInt32Value)114300U, SimplePos = false, RelativeHeight = (UInt32Value)251661312U, BehindDoc = false, Locked = false, LayoutInCell = true, AllowOverlap = true, EditId = "6DD2C0E8", AnchorId = "6B6DB0BE" };
            Wp.SimplePosition simplePosition1 = new Wp.SimplePosition(){ X = 0L, Y = 0L };

            Wp.HorizontalPosition horizontalPosition1 = new Wp.HorizontalPosition(){ RelativeFrom = Wp.HorizontalRelativePositionValues.Column };
            Wp.PositionOffset positionOffset1 = new Wp.PositionOffset();
            positionOffset1.Text = "3619500";

            horizontalPosition1.Append(positionOffset1);

            Wp.VerticalPosition verticalPosition1 = new Wp.VerticalPosition(){ RelativeFrom = Wp.VerticalRelativePositionValues.Paragraph };
            Wp.PositionOffset positionOffset2 = new Wp.PositionOffset();
            positionOffset2.Text = "2609850";

            verticalPosition1.Append(positionOffset2);
            Wp.Extent extent1 = new Wp.Extent(){ Cx = 1060704L, Cy = 914400L };
            Wp.EffectExtent effectExtent1 = new Wp.EffectExtent(){ LeftEdge = 19050L, TopEdge = 19050L, RightEdge = 44450L, BottomEdge = 19050L };
            Wp.WrapNone wrapNone1 = new Wp.WrapNone();
            Wp.DocProperties docProperties1 = new Wp.DocProperties(){ Id = (UInt32Value)1857184239U, Name = "Isosceles Triangle 3" };
            Wp.NonVisualGraphicFrameDrawingProperties nonVisualGraphicFrameDrawingProperties1 = new Wp.NonVisualGraphicFrameDrawingProperties();

            A.Graphic graphic1 = new A.Graphic();
            graphic1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            A.GraphicData graphicData1 = new A.GraphicData(){ Uri = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape" };

            Wps.WordprocessingShape wordprocessingShape1 = new Wps.WordprocessingShape();
            Wps.NonVisualDrawingShapeProperties nonVisualDrawingShapeProperties1 = new Wps.NonVisualDrawingShapeProperties();

            Wps.ShapeProperties shapeProperties1 = new Wps.ShapeProperties();

            A.Transform2D transform2D1 = new A.Transform2D();
            A.Offset offset1 = new A.Offset(){ X = 0L, Y = 0L };
            A.Extents extents1 = new A.Extents(){ Cx = 1060704L, Cy = 914400L };

            transform2D1.Append(offset1);
            transform2D1.Append(extents1);

            A.PresetGeometry presetGeometry1 = new A.PresetGeometry(){ Preset = A.ShapeTypeValues.Triangle };
            A.AdjustValueList adjustValueList1 = new A.AdjustValueList();

            presetGeometry1.Append(adjustValueList1);

            shapeProperties1.Append(transform2D1);
            shapeProperties1.Append(presetGeometry1);

            Wps.ShapeStyle shapeStyle1 = new Wps.ShapeStyle();

            A.LineReference lineReference1 = new A.LineReference(){ Index = (UInt32Value)2U };

            A.SchemeColor schemeColor1 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent1 };
            A.Shade shade1 = new A.Shade(){ Val = 15000 };

            schemeColor1.Append(shade1);

            lineReference1.Append(schemeColor1);

            A.FillReference fillReference1 = new A.FillReference(){ Index = (UInt32Value)1U };
            A.SchemeColor schemeColor2 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent1 };

            fillReference1.Append(schemeColor2);

            A.EffectReference effectReference1 = new A.EffectReference(){ Index = (UInt32Value)0U };
            A.SchemeColor schemeColor3 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent1 };

            effectReference1.Append(schemeColor3);

            A.FontReference fontReference1 = new A.FontReference(){ Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor4 = new A.SchemeColor(){ Val = A.SchemeColorValues.Light1 };

            fontReference1.Append(schemeColor4);

            shapeStyle1.Append(lineReference1);
            shapeStyle1.Append(fillReference1);
            shapeStyle1.Append(effectReference1);
            shapeStyle1.Append(fontReference1);

            Wps.TextBodyProperties textBodyProperties1 = new Wps.TextBodyProperties(){ Rotation = 0, UseParagraphSpacing = false, VerticalOverflow = A.TextVerticalOverflowValues.Overflow, HorizontalOverflow = A.TextHorizontalOverflowValues.Overflow, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, LeftInset = 91440, TopInset = 45720, RightInset = 91440, BottomInset = 45720, ColumnCount = 1, ColumnSpacing = 0, RightToLeftColumns = false, FromWordArt = false, Anchor = A.TextAnchoringTypeValues.Center, AnchorCenter = false, ForceAntiAlias = false, CompatibleLineSpacing = true };

            A.PresetTextWrap presetTextWrap1 = new A.PresetTextWrap(){ Preset = A.TextShapeValues.TextNoShape };
            A.AdjustValueList adjustValueList2 = new A.AdjustValueList();

            presetTextWrap1.Append(adjustValueList2);
            A.NoAutoFit noAutoFit1 = new A.NoAutoFit();

            textBodyProperties1.Append(presetTextWrap1);
            textBodyProperties1.Append(noAutoFit1);

            wordprocessingShape1.Append(nonVisualDrawingShapeProperties1);
            wordprocessingShape1.Append(shapeProperties1);
            wordprocessingShape1.Append(shapeStyle1);
            wordprocessingShape1.Append(textBodyProperties1);

            graphicData1.Append(wordprocessingShape1);

            graphic1.Append(graphicData1);

            anchor1.Append(simplePosition1);
            anchor1.Append(horizontalPosition1);
            anchor1.Append(verticalPosition1);
            anchor1.Append(extent1);
            anchor1.Append(effectExtent1);
            anchor1.Append(wrapNone1);
            anchor1.Append(docProperties1);
            anchor1.Append(nonVisualGraphicFrameDrawingProperties1);
            anchor1.Append(graphic1);

            drawing1.Append(anchor1);

            alternateContentChoice1.Append(drawing1);

            AlternateContentFallback alternateContentFallback1 = new AlternateContentFallback();

            Picture picture1 = new Picture();

            V.Shapetype shapetype1 = new V.Shapetype(){ Id = "_x0000_t5", CoordinateSize = "21600,21600", OptionalNumber = 5, Adjustment = "10800", EdgePath = "m@0,l,21600r21600,xe" };
            shapetype1.SetAttribute(new OpenXmlAttribute("w14", "anchorId", "http://schemas.microsoft.com/office/word/2010/wordml", "43BF660B"));
            V.Stroke stroke1 = new V.Stroke(){ JoinStyle = V.StrokeJoinStyleValues.Miter };

            V.Formulas formulas1 = new V.Formulas();
            V.Formula formula1 = new V.Formula(){ Equation = "val #0" };
            V.Formula formula2 = new V.Formula(){ Equation = "prod #0 1 2" };
            V.Formula formula3 = new V.Formula(){ Equation = "sum @1 10800 0" };

            formulas1.Append(formula1);
            formulas1.Append(formula2);
            formulas1.Append(formula3);
            V.Path path1 = new V.Path(){ TextboxRectangle = "0,10800,10800,18000;5400,10800,16200,18000;10800,10800,21600,18000;0,7200,7200,21600;7200,7200,14400,21600;14400,7200,21600,21600", AllowGradientShape = true, ConnectionPointType = Ovml.ConnectValues.Custom, ConnectionPoints = "@0,0;@1,10800;0,21600;10800,21600;21600,21600;@2,10800" };

            V.ShapeHandles shapeHandles1 = new V.ShapeHandles();
            V.ShapeHandle shapeHandle1 = new V.ShapeHandle(){ Position = "#0,topLeft", XRange = "0,21600" };

            shapeHandles1.Append(shapeHandle1);

            shapetype1.Append(stroke1);
            shapetype1.Append(formulas1);
            shapetype1.Append(path1);
            shapetype1.Append(shapeHandles1);
            V.Shape shape1 = new V.Shape(){ Id = "Isosceles Triangle 3", Style = "position:absolute;margin-left:285pt;margin-top:205.5pt;width:83.5pt;height:1in;z-index:251661312;visibility:visible;mso-wrap-style:square;mso-wrap-distance-left:9pt;mso-wrap-distance-top:0;mso-wrap-distance-right:9pt;mso-wrap-distance-bottom:0;mso-position-horizontal:absolute;mso-position-horizontal-relative:text;mso-position-vertical:absolute;mso-position-vertical-relative:text;v-text-anchor:middle", OptionalString = "_x0000_s1026", FillColor = "#156082 [3204]", StrokeColor = "#030e13 [484]", StrokeWeight = "1.5pt", Type = "#_x0000_t5", EncodedPackage = "UEsDBBQABgAIAAAAIQC2gziS/gAAAOEBAAATAAAAW0NvbnRlbnRfVHlwZXNdLnhtbJSRQU7DMBBF\n90jcwfIWJU67QAgl6YK0S0CoHGBkTxKLZGx5TGhvj5O2G0SRWNoz/78nu9wcxkFMGNg6quQqL6RA\n0s5Y6ir5vt9lD1JwBDIwOMJKHpHlpr69KfdHjyxSmriSfYz+USnWPY7AufNIadK6MEJMx9ApD/oD\nOlTrorhX2lFEilmcO2RdNtjC5xDF9pCuTyYBB5bi6bQ4syoJ3g9WQ0ymaiLzg5KdCXlKLjvcW893\nSUOqXwnz5DrgnHtJTxOsQfEKIT7DmDSUCaxw7Rqn8787ZsmRM9e2VmPeBN4uqYvTtW7jvijg9N/y\nJsXecLq0q+WD6m8AAAD//wMAUEsDBBQABgAIAAAAIQA4/SH/1gAAAJQBAAALAAAAX3JlbHMvLnJl\nbHOkkMFqwzAMhu+DvYPRfXGawxijTi+j0GvpHsDYimMaW0Yy2fr2M4PBMnrbUb/Q94l/f/hMi1qR\nJVI2sOt6UJgd+ZiDgffL8ekFlFSbvV0oo4EbChzGx4f9GRdb25HMsYhqlCwG5lrLq9biZkxWOiqY\n22YiTra2kYMu1l1tQD30/bPm3wwYN0x18gb45AdQl1tp5j/sFB2T0FQ7R0nTNEV3j6o9feQzro1i\nOWA14Fm+Q8a1a8+Bvu/d/dMb2JY5uiPbhG/ktn4cqGU/er3pcvwCAAD//wMAUEsDBBQABgAIAAAA\nIQBxUZqBWwIAABcFAAAOAAAAZHJzL2Uyb0RvYy54bWysVMFu2zAMvQ/YPwi6r7aDtN2COkXQosOA\noi3WDj2rshQLkEWNUuJkXz9KdpxiLXYYdpElkXwknx91cbnrLNsqDAZczauTkjPlJDTGrWv+4+nm\n02fOQhSuERacqvleBX65/PjhovcLNYMWbKOQEYgLi97XvI3RL4oiyFZ1IpyAV46MGrATkY64LhoU\nPaF3tpiV5VnRAzYeQaoQ6PZ6MPJlxtdayXivdVCR2ZpTbTGvmNeXtBbLC7FYo/CtkWMZ4h+q6IRx\nlHSCuhZRsA2aN1CdkQgBdDyR0BWgtZEq90DdVOUf3Ty2wqvcC5ET/ERT+H+w8m776B+QaOh9WATa\npi52Grv0pfrYLpO1n8hSu8gkXVblWXlezjmTZPtSzedlZrM4RnsM8auCjqVNzSMa4dY2NSQWYnsb\nImUl94MbHY415F3cW5WcrfuuNDMNZZ3l6CwPdWWRbQX9WCGlcrEaTK1o1HBdnZZTTVNETpkBE7I2\n1k7YI0CS3lvsodbRP4WqrK4puPxbYUPwFJEzg4tTcGcc4HsAlroaMw/+B5IGahJLL9DsH5AhDNoO\nXt4Y4vtWhPggkMRMsqcBjfe0aAt9zWHccdYC/nrvPvmTxsjKWU/DUfPwcyNQcWa/OVJf/t00Tfkw\nPz2fUQ58bXl5bXGb7groN1X0FHiZt8k/2sNWI3TPNMerlJVMwknKXXMZ8XC4isPQ0ksg1WqV3WiC\nvIi37tHLBJ5YTVp62j0L9AfRkVzv4DBIb3Q3+KZIB6tNBG2yKI+8jnzT9GXhjC9FGu/X5+x1fM+W\nvwEAAP//AwBQSwMEFAAGAAgAAAAhAGzk0o3eAAAACwEAAA8AAABkcnMvZG93bnJldi54bWxMj8FO\nwzAQRO9I/IO1SNyoHWhICXGqCgGiN1r6AU5skoh4HWy3cf+e5QS3Wc1o9k21TnZkJ+PD4FBCthDA\nDLZOD9hJOHy83KyAhahQq9GhkXA2Adb15UWlSu1m3JnTPnaMSjCUSkIf41RyHtreWBUWbjJI3qfz\nVkU6fce1VzOV25HfCnHPrRqQPvRqMk+9ab/2Ryvh+fU7+ge3TGkudv79LTTnzbaR8voqbR6BRZPi\nXxh+8QkdamJq3BF1YKOEvBC0JUpYZhkJShR3BYmGrDwXwOuK/99Q/wAAAP//AwBQSwECLQAUAAYA\nCAAAACEAtoM4kv4AAADhAQAAEwAAAAAAAAAAAAAAAAAAAAAAW0NvbnRlbnRfVHlwZXNdLnhtbFBL\nAQItABQABgAIAAAAIQA4/SH/1gAAAJQBAAALAAAAAAAAAAAAAAAAAC8BAABfcmVscy8ucmVsc1BL\nAQItABQABgAIAAAAIQBxUZqBWwIAABcFAAAOAAAAAAAAAAAAAAAAAC4CAABkcnMvZTJvRG9jLnht\nbFBLAQItABQABgAIAAAAIQBs5NKN3gAAAAsBAAAPAAAAAAAAAAAAAAAAALUEAABkcnMvZG93bnJl\ndi54bWxQSwUGAAAAAAQABADzAAAAwAUAAAAA\n" };

            picture1.Append(shapetype1);
            picture1.Append(shape1);

            alternateContentFallback1.Append(picture1);

            alternateContent1.Append(alternateContentChoice1);
            alternateContent1.Append(alternateContentFallback1);

            run1.Append(runProperties1);
            run1.Append(alternateContent1);

            Run run2 = new Run();

            RunProperties runProperties2 = new RunProperties();
            NoProof noProof2 = new NoProof();

            runProperties2.Append(noProof2);

            AlternateContent alternateContent2 = new AlternateContent();

            AlternateContentChoice alternateContentChoice2 = new AlternateContentChoice(){ Requires = "wps" };

            Drawing drawing2 = new Drawing();

            Wp.Anchor anchor2 = new Wp.Anchor(){ DistanceFromTop = (UInt32Value)0U, DistanceFromBottom = (UInt32Value)0U, DistanceFromLeft = (UInt32Value)114300U, DistanceFromRight = (UInt32Value)114300U, SimplePos = false, RelativeHeight = (UInt32Value)251660288U, BehindDoc = false, Locked = false, LayoutInCell = true, AllowOverlap = true, EditId = "4A6E00FD", AnchorId = "2D1CF8B2" };
            Wp.SimplePosition simplePosition2 = new Wp.SimplePosition(){ X = 0L, Y = 0L };

            Wp.HorizontalPosition horizontalPosition2 = new Wp.HorizontalPosition(){ RelativeFrom = Wp.HorizontalRelativePositionValues.Column };
            Wp.PositionOffset positionOffset3 = new Wp.PositionOffset();
            positionOffset3.Text = "1162050";

            horizontalPosition2.Append(positionOffset3);

            Wp.VerticalPosition verticalPosition2 = new Wp.VerticalPosition(){ RelativeFrom = Wp.VerticalRelativePositionValues.Paragraph };
            Wp.PositionOffset positionOffset4 = new Wp.PositionOffset();
            positionOffset4.Text = "2559050";

            verticalPosition2.Append(positionOffset4);
            Wp.Extent extent2 = new Wp.Extent(){ Cx = 914400L, Cy = 914400L };
            Wp.EffectExtent effectExtent2 = new Wp.EffectExtent(){ LeftEdge = 0L, TopEdge = 0L, RightEdge = 19050L, BottomEdge = 19050L };
            Wp.WrapNone wrapNone2 = new Wp.WrapNone();
            Wp.DocProperties docProperties2 = new Wp.DocProperties(){ Id = (UInt32Value)1864490150U, Name = "Oval 2" };
            Wp.NonVisualGraphicFrameDrawingProperties nonVisualGraphicFrameDrawingProperties2 = new Wp.NonVisualGraphicFrameDrawingProperties();

            A.Graphic graphic2 = new A.Graphic();
            graphic2.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            A.GraphicData graphicData2 = new A.GraphicData(){ Uri = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape" };

            Wps.WordprocessingShape wordprocessingShape2 = new Wps.WordprocessingShape();
            Wps.NonVisualDrawingShapeProperties nonVisualDrawingShapeProperties2 = new Wps.NonVisualDrawingShapeProperties();

            Wps.ShapeProperties shapeProperties2 = new Wps.ShapeProperties();

            A.Transform2D transform2D2 = new A.Transform2D();
            A.Offset offset2 = new A.Offset(){ X = 0L, Y = 0L };
            A.Extents extents2 = new A.Extents(){ Cx = 914400L, Cy = 914400L };

            transform2D2.Append(offset2);
            transform2D2.Append(extents2);

            A.PresetGeometry presetGeometry2 = new A.PresetGeometry(){ Preset = A.ShapeTypeValues.Ellipse };
            A.AdjustValueList adjustValueList3 = new A.AdjustValueList();

            presetGeometry2.Append(adjustValueList3);

            shapeProperties2.Append(transform2D2);
            shapeProperties2.Append(presetGeometry2);

            Wps.ShapeStyle shapeStyle2 = new Wps.ShapeStyle();

            A.LineReference lineReference2 = new A.LineReference(){ Index = (UInt32Value)2U };

            A.SchemeColor schemeColor5 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent1 };
            A.Shade shade2 = new A.Shade(){ Val = 15000 };

            schemeColor5.Append(shade2);

            lineReference2.Append(schemeColor5);

            A.FillReference fillReference2 = new A.FillReference(){ Index = (UInt32Value)1U };
            A.SchemeColor schemeColor6 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent1 };

            fillReference2.Append(schemeColor6);

            A.EffectReference effectReference2 = new A.EffectReference(){ Index = (UInt32Value)0U };
            A.SchemeColor schemeColor7 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent1 };

            effectReference2.Append(schemeColor7);

            A.FontReference fontReference2 = new A.FontReference(){ Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor8 = new A.SchemeColor(){ Val = A.SchemeColorValues.Light1 };

            fontReference2.Append(schemeColor8);

            shapeStyle2.Append(lineReference2);
            shapeStyle2.Append(fillReference2);
            shapeStyle2.Append(effectReference2);
            shapeStyle2.Append(fontReference2);

            Wps.TextBodyProperties textBodyProperties2 = new Wps.TextBodyProperties(){ Rotation = 0, UseParagraphSpacing = false, VerticalOverflow = A.TextVerticalOverflowValues.Overflow, HorizontalOverflow = A.TextHorizontalOverflowValues.Overflow, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, LeftInset = 91440, TopInset = 45720, RightInset = 91440, BottomInset = 45720, ColumnCount = 1, ColumnSpacing = 0, RightToLeftColumns = false, FromWordArt = false, Anchor = A.TextAnchoringTypeValues.Center, AnchorCenter = false, ForceAntiAlias = false, CompatibleLineSpacing = true };

            A.PresetTextWrap presetTextWrap2 = new A.PresetTextWrap(){ Preset = A.TextShapeValues.TextNoShape };
            A.AdjustValueList adjustValueList4 = new A.AdjustValueList();

            presetTextWrap2.Append(adjustValueList4);
            A.NoAutoFit noAutoFit2 = new A.NoAutoFit();

            textBodyProperties2.Append(presetTextWrap2);
            textBodyProperties2.Append(noAutoFit2);

            wordprocessingShape2.Append(nonVisualDrawingShapeProperties2);
            wordprocessingShape2.Append(shapeProperties2);
            wordprocessingShape2.Append(shapeStyle2);
            wordprocessingShape2.Append(textBodyProperties2);

            graphicData2.Append(wordprocessingShape2);

            graphic2.Append(graphicData2);

            anchor2.Append(simplePosition2);
            anchor2.Append(horizontalPosition2);
            anchor2.Append(verticalPosition2);
            anchor2.Append(extent2);
            anchor2.Append(effectExtent2);
            anchor2.Append(wrapNone2);
            anchor2.Append(docProperties2);
            anchor2.Append(nonVisualGraphicFrameDrawingProperties2);
            anchor2.Append(graphic2);

            drawing2.Append(anchor2);

            alternateContentChoice2.Append(drawing2);

            AlternateContentFallback alternateContentFallback2 = new AlternateContentFallback();

            Picture picture2 = new Picture();

            V.Oval oval1 = new V.Oval(){ Id = "Oval 2", Style = "position:absolute;margin-left:91.5pt;margin-top:201.5pt;width:1in;height:1in;z-index:251660288;visibility:visible;mso-wrap-style:square;mso-wrap-distance-left:9pt;mso-wrap-distance-top:0;mso-wrap-distance-right:9pt;mso-wrap-distance-bottom:0;mso-position-horizontal:absolute;mso-position-horizontal-relative:text;mso-position-vertical:absolute;mso-position-vertical-relative:text;v-text-anchor:middle", OptionalString = "_x0000_s1026", FillColor = "#156082 [3204]", StrokeColor = "#030e13 [484]", StrokeWeight = "1.5pt" };
            oval1.SetAttribute(new OpenXmlAttribute("w14", "anchorId", "http://schemas.microsoft.com/office/word/2010/wordml", "13C095C5"));
            oval1.SetAttribute(new OpenXmlAttribute("o", "gfxdata", "urn:schemas-microsoft-com:office:office", "UEsDBBQABgAIAAAAIQC2gziS/gAAAOEBAAATAAAAW0NvbnRlbnRfVHlwZXNdLnhtbJSRQU7DMBBF\n90jcwfIWJU67QAgl6YK0S0CoHGBkTxKLZGx5TGhvj5O2G0SRWNoz/78nu9wcxkFMGNg6quQqL6RA\n0s5Y6ir5vt9lD1JwBDIwOMJKHpHlpr69KfdHjyxSmriSfYz+USnWPY7AufNIadK6MEJMx9ApD/oD\nOlTrorhX2lFEilmcO2RdNtjC5xDF9pCuTyYBB5bi6bQ4syoJ3g9WQ0ymaiLzg5KdCXlKLjvcW893\nSUOqXwnz5DrgnHtJTxOsQfEKIT7DmDSUCaxw7Rqn8787ZsmRM9e2VmPeBN4uqYvTtW7jvijg9N/y\nJsXecLq0q+WD6m8AAAD//wMAUEsDBBQABgAIAAAAIQA4/SH/1gAAAJQBAAALAAAAX3JlbHMvLnJl\nbHOkkMFqwzAMhu+DvYPRfXGawxijTi+j0GvpHsDYimMaW0Yy2fr2M4PBMnrbUb/Q94l/f/hMi1qR\nJVI2sOt6UJgd+ZiDgffL8ekFlFSbvV0oo4EbChzGx4f9GRdb25HMsYhqlCwG5lrLq9biZkxWOiqY\n22YiTra2kYMu1l1tQD30/bPm3wwYN0x18gb45AdQl1tp5j/sFB2T0FQ7R0nTNEV3j6o9feQzro1i\nOWA14Fm+Q8a1a8+Bvu/d/dMb2JY5uiPbhG/ktn4cqGU/er3pcvwCAAD//wMAUEsDBBQABgAIAAAA\nIQCGXH1IVwIAABUFAAAOAAAAZHJzL2Uyb0RvYy54bWysVE1v2zAMvQ/YfxB0X+wE6T6COEXQosOA\noi2WDj2rslQLkEWNUuJkv36U7DjBWuww7CJLIvlIPj9qeblvLdspDAZcxaeTkjPlJNTGvVT8x+PN\nh8+chShcLSw4VfGDCvxy9f7dsvMLNYMGbK2QEYgLi85XvInRL4oiyEa1IkzAK0dGDdiKSEd8KWoU\nHaG3tpiV5ceiA6w9glQh0O11b+SrjK+1kvFe66AisxWn2mJeMa/PaS1WS7F4QeEbI4cyxD9U0Qrj\nKOkIdS2iYFs0r6BaIxEC6DiR0BagtZEq90DdTMs/utk0wqvcC5ET/EhT+H+w8m638Q9INHQ+LAJt\nUxd7jW36Un1sn8k6jGSpfWSSLr9M5/OSKJVkGvaEUpyCPYb4VUHL0qbiylrjQ2pHLMTuNsTe++hF\noacK8i4erErO1n1Xmpmacs5ydBaHurLIdoJ+q5BSuTjtTY2oVX89vSipvD7JGJELzIAJWRtrR+wB\nIAnvNXYPM/inUJW1NQaXfyusDx4jcmZwcQxujQN8C8BSV0Pm3v9IUk9NYukZ6sMDMoRe2cHLG0N0\n34oQHwSSlOkP0XjGe1q0ha7iMOw4awB/vXWf/ElhZOWso9GoePi5Fag4s98caS//bZqlfJhffJpR\nDjy3PJ9b3La9AvpNU3oIvMzb5B/tcasR2iea4nXKSibhJOWuuIx4PFzFfmTpHZBqvc5uND9exFu3\n8TKBJ1aTlh73TwL9oLlIYr2D4xi90l3vmyIdrLcRtMmiPPE68E2zl4UzvBNpuM/P2ev0mq1+AwAA\n//8DAFBLAwQUAAYACAAAACEAhMzWWN4AAAALAQAADwAAAGRycy9kb3ducmV2LnhtbEyPzU7DMBCE\n70i8g7VIXFDr9AeIQpwqAiFOqKL0Adx4idPE6yh22sDTsz3BbUY7mv0m30yuEyccQuNJwWKegECq\nvGmoVrD/fJ2lIELUZHTnCRV8Y4BNcX2V68z4M33gaRdrwSUUMq3AxthnUobKotNh7nskvn35wenI\ndqilGfSZy10nl0nyIJ1uiD9Y3eOzxardjU7BunT2Zyxf7rbt++Lt2FI4prFS6vZmKp9ARJziXxgu\n+IwOBTMd/EgmiI59uuItkcuSi+DEavnI4qDgfs1CFrn8v6H4BQAA//8DAFBLAQItABQABgAIAAAA\nIQC2gziS/gAAAOEBAAATAAAAAAAAAAAAAAAAAAAAAABbQ29udGVudF9UeXBlc10ueG1sUEsBAi0A\nFAAGAAgAAAAhADj9If/WAAAAlAEAAAsAAAAAAAAAAAAAAAAALwEAAF9yZWxzLy5yZWxzUEsBAi0A\nFAAGAAgAAAAhAIZcfUhXAgAAFQUAAA4AAAAAAAAAAAAAAAAALgIAAGRycy9lMm9Eb2MueG1sUEsB\nAi0AFAAGAAgAAAAhAITM1ljeAAAACwEAAA8AAAAAAAAAAAAAAAAAsQQAAGRycy9kb3ducmV2Lnht\nbFBLBQYAAAAABAAEAPMAAAC8BQAAAAA=\n"));
            V.Stroke stroke2 = new V.Stroke(){ JoinStyle = V.StrokeJoinStyleValues.Miter };

            oval1.Append(stroke2);

            picture2.Append(oval1);

            alternateContentFallback2.Append(picture2);

            alternateContent2.Append(alternateContentChoice2);
            alternateContent2.Append(alternateContentFallback2);

            run2.Append(runProperties2);
            run2.Append(alternateContent2);

            Run run3 = new Run();

            RunProperties runProperties3 = new RunProperties();
            NoProof noProof3 = new NoProof();

            runProperties3.Append(noProof3);

            AlternateContent alternateContent3 = new AlternateContent();

            AlternateContentChoice alternateContentChoice3 = new AlternateContentChoice(){ Requires = "wps" };

            Drawing drawing3 = new Drawing();

            Wp.Anchor anchor3 = new Wp.Anchor(){ DistanceFromTop = (UInt32Value)0U, DistanceFromBottom = (UInt32Value)0U, DistanceFromLeft = (UInt32Value)114300U, DistanceFromRight = (UInt32Value)114300U, SimplePos = false, RelativeHeight = (UInt32Value)251659264U, BehindDoc = false, Locked = false, LayoutInCell = true, AllowOverlap = true, EditId = "55018DA5", AnchorId = "43894DD6" };
            Wp.SimplePosition simplePosition3 = new Wp.SimplePosition(){ X = 0L, Y = 0L };

            Wp.HorizontalPosition horizontalPosition3 = new Wp.HorizontalPosition(){ RelativeFrom = Wp.HorizontalRelativePositionValues.Column };
            Wp.PositionOffset positionOffset5 = new Wp.PositionOffset();
            positionOffset5.Text = "552450";

            horizontalPosition3.Append(positionOffset5);

            Wp.VerticalPosition verticalPosition3 = new Wp.VerticalPosition(){ RelativeFrom = Wp.VerticalRelativePositionValues.Paragraph };
            Wp.PositionOffset positionOffset6 = new Wp.PositionOffset();
            positionOffset6.Text = "342900";

            verticalPosition3.Append(positionOffset6);
            Wp.Extent extent3 = new Wp.Extent(){ Cx = 914400L, Cy = 914400L };
            Wp.EffectExtent effectExtent3 = new Wp.EffectExtent(){ LeftEdge = 0L, TopEdge = 0L, RightEdge = 19050L, BottomEdge = 19050L };
            Wp.WrapNone wrapNone3 = new Wp.WrapNone();
            Wp.DocProperties docProperties3 = new Wp.DocProperties(){ Id = (UInt32Value)2139055263U, Name = "Rectangle 1" };
            Wp.NonVisualGraphicFrameDrawingProperties nonVisualGraphicFrameDrawingProperties3 = new Wp.NonVisualGraphicFrameDrawingProperties();

            A.Graphic graphic3 = new A.Graphic();
            graphic3.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            A.GraphicData graphicData3 = new A.GraphicData(){ Uri = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape" };

            Wps.WordprocessingShape wordprocessingShape3 = new Wps.WordprocessingShape();
            Wps.NonVisualDrawingShapeProperties nonVisualDrawingShapeProperties3 = new Wps.NonVisualDrawingShapeProperties();

            Wps.ShapeProperties shapeProperties3 = new Wps.ShapeProperties();

            A.Transform2D transform2D3 = new A.Transform2D();
            A.Offset offset3 = new A.Offset(){ X = 0L, Y = 0L };
            A.Extents extents3 = new A.Extents(){ Cx = 914400L, Cy = 914400L };

            transform2D3.Append(offset3);
            transform2D3.Append(extents3);

            A.PresetGeometry presetGeometry3 = new A.PresetGeometry(){ Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList5 = new A.AdjustValueList();

            presetGeometry3.Append(adjustValueList5);

            shapeProperties3.Append(transform2D3);
            shapeProperties3.Append(presetGeometry3);

            Wps.ShapeStyle shapeStyle3 = new Wps.ShapeStyle();

            A.LineReference lineReference3 = new A.LineReference(){ Index = (UInt32Value)2U };

            A.SchemeColor schemeColor9 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent1 };
            A.Shade shade3 = new A.Shade(){ Val = 15000 };

            schemeColor9.Append(shade3);

            lineReference3.Append(schemeColor9);

            A.FillReference fillReference3 = new A.FillReference(){ Index = (UInt32Value)1U };
            A.SchemeColor schemeColor10 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent1 };

            fillReference3.Append(schemeColor10);

            A.EffectReference effectReference3 = new A.EffectReference(){ Index = (UInt32Value)0U };
            A.SchemeColor schemeColor11 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent1 };

            effectReference3.Append(schemeColor11);

            A.FontReference fontReference3 = new A.FontReference(){ Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor12 = new A.SchemeColor(){ Val = A.SchemeColorValues.Light1 };

            fontReference3.Append(schemeColor12);

            shapeStyle3.Append(lineReference3);
            shapeStyle3.Append(fillReference3);
            shapeStyle3.Append(effectReference3);
            shapeStyle3.Append(fontReference3);

            Wps.TextBodyProperties textBodyProperties3 = new Wps.TextBodyProperties(){ Rotation = 0, UseParagraphSpacing = false, VerticalOverflow = A.TextVerticalOverflowValues.Overflow, HorizontalOverflow = A.TextHorizontalOverflowValues.Overflow, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, LeftInset = 91440, TopInset = 45720, RightInset = 91440, BottomInset = 45720, ColumnCount = 1, ColumnSpacing = 0, RightToLeftColumns = false, FromWordArt = false, Anchor = A.TextAnchoringTypeValues.Center, AnchorCenter = false, ForceAntiAlias = false, CompatibleLineSpacing = true };

            A.PresetTextWrap presetTextWrap3 = new A.PresetTextWrap(){ Preset = A.TextShapeValues.TextNoShape };
            A.AdjustValueList adjustValueList6 = new A.AdjustValueList();

            presetTextWrap3.Append(adjustValueList6);
            A.NoAutoFit noAutoFit3 = new A.NoAutoFit();

            textBodyProperties3.Append(presetTextWrap3);
            textBodyProperties3.Append(noAutoFit3);

            wordprocessingShape3.Append(nonVisualDrawingShapeProperties3);
            wordprocessingShape3.Append(shapeProperties3);
            wordprocessingShape3.Append(shapeStyle3);
            wordprocessingShape3.Append(textBodyProperties3);

            graphicData3.Append(wordprocessingShape3);

            graphic3.Append(graphicData3);

            anchor3.Append(simplePosition3);
            anchor3.Append(horizontalPosition3);
            anchor3.Append(verticalPosition3);
            anchor3.Append(extent3);
            anchor3.Append(effectExtent3);
            anchor3.Append(wrapNone3);
            anchor3.Append(docProperties3);
            anchor3.Append(nonVisualGraphicFrameDrawingProperties3);
            anchor3.Append(graphic3);

            drawing3.Append(anchor3);

            alternateContentChoice3.Append(drawing3);

            AlternateContentFallback alternateContentFallback3 = new AlternateContentFallback();

            Picture picture3 = new Picture();

            V.Rectangle rectangle1 = new V.Rectangle(){ Id = "Rectangle 1", Style = "position:absolute;margin-left:43.5pt;margin-top:27pt;width:1in;height:1in;z-index:251659264;visibility:visible;mso-wrap-style:square;mso-wrap-distance-left:9pt;mso-wrap-distance-top:0;mso-wrap-distance-right:9pt;mso-wrap-distance-bottom:0;mso-position-horizontal:absolute;mso-position-horizontal-relative:text;mso-position-vertical:absolute;mso-position-vertical-relative:text;v-text-anchor:middle", OptionalString = "_x0000_s1026", FillColor = "#156082 [3204]", StrokeColor = "#030e13 [484]", StrokeWeight = "1.5pt" };
            rectangle1.SetAttribute(new OpenXmlAttribute("w14", "anchorId", "http://schemas.microsoft.com/office/word/2010/wordml", "57A4C86B"));
            rectangle1.SetAttribute(new OpenXmlAttribute("o", "gfxdata", "urn:schemas-microsoft-com:office:office", "UEsDBBQABgAIAAAAIQC2gziS/gAAAOEBAAATAAAAW0NvbnRlbnRfVHlwZXNdLnhtbJSRQU7DMBBF\n90jcwfIWJU67QAgl6YK0S0CoHGBkTxKLZGx5TGhvj5O2G0SRWNoz/78nu9wcxkFMGNg6quQqL6RA\n0s5Y6ir5vt9lD1JwBDIwOMJKHpHlpr69KfdHjyxSmriSfYz+USnWPY7AufNIadK6MEJMx9ApD/oD\nOlTrorhX2lFEilmcO2RdNtjC5xDF9pCuTyYBB5bi6bQ4syoJ3g9WQ0ymaiLzg5KdCXlKLjvcW893\nSUOqXwnz5DrgnHtJTxOsQfEKIT7DmDSUCaxw7Rqn8787ZsmRM9e2VmPeBN4uqYvTtW7jvijg9N/y\nJsXecLq0q+WD6m8AAAD//wMAUEsDBBQABgAIAAAAIQA4/SH/1gAAAJQBAAALAAAAX3JlbHMvLnJl\nbHOkkMFqwzAMhu+DvYPRfXGawxijTi+j0GvpHsDYimMaW0Yy2fr2M4PBMnrbUb/Q94l/f/hMi1qR\nJVI2sOt6UJgd+ZiDgffL8ekFlFSbvV0oo4EbChzGx4f9GRdb25HMsYhqlCwG5lrLq9biZkxWOiqY\n22YiTra2kYMu1l1tQD30/bPm3wwYN0x18gb45AdQl1tp5j/sFB2T0FQ7R0nTNEV3j6o9feQzro1i\nOWA14Fm+Q8a1a8+Bvu/d/dMb2JY5uiPbhG/ktn4cqGU/er3pcvwCAAD//wMAUEsDBBQABgAIAAAA\nIQAb8MsLVgIAABIFAAAOAAAAZHJzL2Uyb0RvYy54bWysVE1v2zAMvQ/YfxB0X+wE6T6COEWQosOA\noA3aDj0rshQbkEWNUuJkv36U7DhFW+ww7GJTIvlIPT1qfn1sDDso9DXYgo9HOWfKSihruyv4z6fb\nT18580HYUhiwquAn5fn14uOHeetmagIVmFIhIxDrZ60reBWCm2WZl5VqhB+BU5acGrARgZa4y0oU\nLaE3Jpvk+eesBSwdglTe0+5N5+SLhK+1kuFea68CMwWn3kL6Yvpu4zdbzMVsh8JVtezbEP/QRSNq\nS0UHqBsRBNtj/QaqqSWCBx1GEpoMtK6lSmeg04zzV6d5rIRT6SxEjncDTf7/wcq7w6PbINHQOj/z\nZMZTHDU28U/9sWMi6zSQpY6BSdr8Np5Oc6JUkqu3CSW7JDv04buChkWj4Eh3kSgSh7UPXeg5hPIu\n5ZMVTkbFDox9UJrVJRWcpOykDLUyyA6C7lRIqWwYd65KlKrbHl/l1FtXZMhI3SXAiKxrYwbsHiCq\n7i12B9PHx1SVhDUk539rrEseMlJlsGFIbmoL+B6AoVP1lbv4M0kdNZGlLZSnDTKETtbeyduauF4L\nHzYCScd0PTSb4Z4+2kBbcOgtzirA3+/tx3iSF3k5a2kuCu5/7QUqzswPS8JLV02DlBbTqy8TqoEv\nPduXHrtvVkDXNKZXwMlkxvhgzqZGaJ5phJexKrmElVS74DLgebEK3bzSIyDVcpnCaHicCGv76GQE\nj6xGLT0dnwW6XnCBlHoH5xkSs1e662JjpoXlPoCukygvvPZ80+Al4fSPRJzsl+sUdXnKFn8AAAD/\n/wMAUEsDBBQABgAIAAAAIQA6H/li3wAAAAkBAAAPAAAAZHJzL2Rvd25yZXYueG1sTI9BT8MwDIXv\nSPyHyEi7IJauZVBK02kaGhIntIE4u03WVmucrkm38u8xJzjZ1nt6/l6+mmwnzmbwrSMFi3kEwlDl\ndEu1gs+P7V0KwgckjZ0jo+DbeFgV11c5ZtpdaGfO+1ALDiGfoYImhD6T0leNsejnrjfE2sENFgOf\nQy31gBcOt52Mo+hBWmyJPzTYm01jquN+tAqsPn2V4+vRJf3L+jY5bQ9vGL8rNbuZ1s8ggpnCnxl+\n8RkdCmYq3Ujai05B+shVgoLlPU/W42TBS8nGpzQCWeTyf4PiBwAA//8DAFBLAQItABQABgAIAAAA\nIQC2gziS/gAAAOEBAAATAAAAAAAAAAAAAAAAAAAAAABbQ29udGVudF9UeXBlc10ueG1sUEsBAi0A\nFAAGAAgAAAAhADj9If/WAAAAlAEAAAsAAAAAAAAAAAAAAAAALwEAAF9yZWxzLy5yZWxzUEsBAi0A\nFAAGAAgAAAAhABvwywtWAgAAEgUAAA4AAAAAAAAAAAAAAAAALgIAAGRycy9lMm9Eb2MueG1sUEsB\nAi0AFAAGAAgAAAAhADof+WLfAAAACQEAAA8AAAAAAAAAAAAAAAAAsAQAAGRycy9kb3ducmV2Lnht\nbFBLBQYAAAAABAAEAPMAAAC8BQAAAAA=\n"));

            picture3.Append(rectangle1);

            alternateContentFallback3.Append(picture3);

            alternateContent3.Append(alternateContentChoice3);
            alternateContent3.Append(alternateContentFallback3);

            run3.Append(runProperties3);
            run3.Append(alternateContent3);

            paragraph1.Append(run1);
            paragraph1.Append(run2);
            paragraph1.Append(run3);

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
            RsidRoot rsidRoot1 = new RsidRoot(){ Val = "007719D7" };
            Rsid rsid1 = new Rsid(){ Val = "0018596D" };
            Rsid rsid2 = new Rsid(){ Val = "007719D7" };
            Rsid rsid3 = new Rsid(){ Val = "009674C8" };
            Rsid rsid4 = new Rsid(){ Val = "00DD5CA4" };

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
            W14.DocumentId documentId1 = new W14.DocumentId(){ Val = "3E7EDBFA" };
            W15.ChartTrackingRefBased chartTrackingRefBased1 = new W15.ChartTrackingRefBased();
            W15.PersistentDocumentId persistentDocumentId1 = new W15.PersistentDocumentId(){ Val = "{C5460107-BD89-4D6C-93C3-E0A0E92324C0}" };

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
            Rsid rsid5 = new Rsid(){ Val = "007719D7" };

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
            Rsid rsid6 = new Rsid(){ Val = "007719D7" };

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
            Rsid rsid7 = new Rsid(){ Val = "007719D7" };

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
            Rsid rsid8 = new Rsid(){ Val = "007719D7" };

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
            Rsid rsid9 = new Rsid(){ Val = "007719D7" };

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
            Rsid rsid10 = new Rsid(){ Val = "007719D7" };

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
            Rsid rsid11 = new Rsid(){ Val = "007719D7" };

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
            Rsid rsid12 = new Rsid(){ Val = "007719D7" };

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
            Rsid rsid13 = new Rsid(){ Val = "007719D7" };

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
            Rsid rsid14 = new Rsid(){ Val = "007719D7" };

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
            Rsid rsid15 = new Rsid(){ Val = "007719D7" };

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
            Rsid rsid16 = new Rsid(){ Val = "007719D7" };

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
            Rsid rsid17 = new Rsid(){ Val = "007719D7" };

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
            Rsid rsid18 = new Rsid(){ Val = "007719D7" };

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
            Rsid rsid19 = new Rsid(){ Val = "007719D7" };

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
            Rsid rsid20 = new Rsid(){ Val = "007719D7" };

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
            Rsid rsid21 = new Rsid(){ Val = "007719D7" };

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
            Rsid rsid22 = new Rsid(){ Val = "007719D7" };

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
            Rsid rsid23 = new Rsid(){ Val = "007719D7" };

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
            Rsid rsid24 = new Rsid(){ Val = "007719D7" };

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
            Rsid rsid25 = new Rsid(){ Val = "007719D7" };

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
            Rsid rsid26 = new Rsid(){ Val = "007719D7" };

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
            Rsid rsid27 = new Rsid(){ Val = "007719D7" };

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
            Rsid rsid28 = new Rsid(){ Val = "007719D7" };

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
            Rsid rsid29 = new Rsid(){ Val = "007719D7" };

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
            Rsid rsid30 = new Rsid(){ Val = "007719D7" };

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
            Rsid rsid31 = new Rsid(){ Val = "007719D7" };

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
            Rsid rsid32 = new Rsid(){ Val = "007719D7" };

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
            Rsid rsid33 = new Rsid(){ Val = "007719D7" };

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

        // Generates content of customXmlPart1.
        private void GenerateCustomXmlPart1Content(CustomXmlPart customXmlPart1)
        {
            System.Xml.XmlTextWriter writer = new System.Xml.XmlTextWriter(customXmlPart1.GetStream(System.IO.FileMode.Create), System.Text.Encoding.UTF8);
            writer.WriteRaw("<?xml version=\"1.0\" standalone=\"no\"?><b:Sources xmlns:b=\"http://schemas.openxmlformats.org/officeDocument/2006/bibliography\" xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/bibliography\" SelectedStyle=\"\\APASixthEditionOfficeOnline.xsl\" StyleName=\"APA\" Version=\"6\"></b:Sources>");
            writer.Flush();
            writer.Close();
        }

        // Generates content of customXmlPropertiesPart1.
        private void GenerateCustomXmlPropertiesPart1Content(CustomXmlPropertiesPart customXmlPropertiesPart1)
        {
            Ds.DataStoreItem dataStoreItem1 = new Ds.DataStoreItem(){ ItemId = "{6C65F340-9D2D-467A-A929-A5652C474C7D}" };
            dataStoreItem1.AddNamespaceDeclaration("ds", "http://schemas.openxmlformats.org/officeDocument/2006/customXml");

            Ds.SchemaReferences schemaReferences1 = new Ds.SchemaReferences();
            Ds.SchemaReference schemaReference1 = new Ds.SchemaReference(){ Uri = "http://schemas.openxmlformats.org/officeDocument/2006/bibliography" };

            schemaReferences1.Append(schemaReference1);

            dataStoreItem1.Append(schemaReferences1);

            customXmlPropertiesPart1.DataStoreItem = dataStoreItem1;
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

            A.SolidFill solidFill1 = new A.SolidFill();
            A.SchemeColor schemeColor13 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };

            solidFill1.Append(schemeColor13);

            A.GradientFill gradientFill1 = new A.GradientFill(){ RotateWithShape = true };

            A.GradientStopList gradientStopList1 = new A.GradientStopList();

            A.GradientStop gradientStop1 = new A.GradientStop(){ Position = 0 };

            A.SchemeColor schemeColor14 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };
            A.LuminanceModulation luminanceModulation1 = new A.LuminanceModulation(){ Val = 110000 };
            A.SaturationModulation saturationModulation1 = new A.SaturationModulation(){ Val = 105000 };
            A.Tint tint1 = new A.Tint(){ Val = 67000 };

            schemeColor14.Append(luminanceModulation1);
            schemeColor14.Append(saturationModulation1);
            schemeColor14.Append(tint1);

            gradientStop1.Append(schemeColor14);

            A.GradientStop gradientStop2 = new A.GradientStop(){ Position = 50000 };

            A.SchemeColor schemeColor15 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };
            A.LuminanceModulation luminanceModulation2 = new A.LuminanceModulation(){ Val = 105000 };
            A.SaturationModulation saturationModulation2 = new A.SaturationModulation(){ Val = 103000 };
            A.Tint tint2 = new A.Tint(){ Val = 73000 };

            schemeColor15.Append(luminanceModulation2);
            schemeColor15.Append(saturationModulation2);
            schemeColor15.Append(tint2);

            gradientStop2.Append(schemeColor15);

            A.GradientStop gradientStop3 = new A.GradientStop(){ Position = 100000 };

            A.SchemeColor schemeColor16 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };
            A.LuminanceModulation luminanceModulation3 = new A.LuminanceModulation(){ Val = 105000 };
            A.SaturationModulation saturationModulation3 = new A.SaturationModulation(){ Val = 109000 };
            A.Tint tint3 = new A.Tint(){ Val = 81000 };

            schemeColor16.Append(luminanceModulation3);
            schemeColor16.Append(saturationModulation3);
            schemeColor16.Append(tint3);

            gradientStop3.Append(schemeColor16);

            gradientStopList1.Append(gradientStop1);
            gradientStopList1.Append(gradientStop2);
            gradientStopList1.Append(gradientStop3);
            A.LinearGradientFill linearGradientFill1 = new A.LinearGradientFill(){ Angle = 5400000, Scaled = false };

            gradientFill1.Append(gradientStopList1);
            gradientFill1.Append(linearGradientFill1);

            A.GradientFill gradientFill2 = new A.GradientFill(){ RotateWithShape = true };

            A.GradientStopList gradientStopList2 = new A.GradientStopList();

            A.GradientStop gradientStop4 = new A.GradientStop(){ Position = 0 };

            A.SchemeColor schemeColor17 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };
            A.SaturationModulation saturationModulation4 = new A.SaturationModulation(){ Val = 103000 };
            A.LuminanceModulation luminanceModulation4 = new A.LuminanceModulation(){ Val = 102000 };
            A.Tint tint4 = new A.Tint(){ Val = 94000 };

            schemeColor17.Append(saturationModulation4);
            schemeColor17.Append(luminanceModulation4);
            schemeColor17.Append(tint4);

            gradientStop4.Append(schemeColor17);

            A.GradientStop gradientStop5 = new A.GradientStop(){ Position = 50000 };

            A.SchemeColor schemeColor18 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };
            A.SaturationModulation saturationModulation5 = new A.SaturationModulation(){ Val = 110000 };
            A.LuminanceModulation luminanceModulation5 = new A.LuminanceModulation(){ Val = 100000 };
            A.Shade shade4 = new A.Shade(){ Val = 100000 };

            schemeColor18.Append(saturationModulation5);
            schemeColor18.Append(luminanceModulation5);
            schemeColor18.Append(shade4);

            gradientStop5.Append(schemeColor18);

            A.GradientStop gradientStop6 = new A.GradientStop(){ Position = 100000 };

            A.SchemeColor schemeColor19 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };
            A.LuminanceModulation luminanceModulation6 = new A.LuminanceModulation(){ Val = 99000 };
            A.SaturationModulation saturationModulation6 = new A.SaturationModulation(){ Val = 120000 };
            A.Shade shade5 = new A.Shade(){ Val = 78000 };

            schemeColor19.Append(luminanceModulation6);
            schemeColor19.Append(saturationModulation6);
            schemeColor19.Append(shade5);

            gradientStop6.Append(schemeColor19);

            gradientStopList2.Append(gradientStop4);
            gradientStopList2.Append(gradientStop5);
            gradientStopList2.Append(gradientStop6);
            A.LinearGradientFill linearGradientFill2 = new A.LinearGradientFill(){ Angle = 5400000, Scaled = false };

            gradientFill2.Append(gradientStopList2);
            gradientFill2.Append(linearGradientFill2);

            fillStyleList1.Append(solidFill1);
            fillStyleList1.Append(gradientFill1);
            fillStyleList1.Append(gradientFill2);

            A.LineStyleList lineStyleList1 = new A.LineStyleList();

            A.Outline outline1 = new A.Outline(){ Width = 12700, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill2 = new A.SolidFill();
            A.SchemeColor schemeColor20 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };

            solidFill2.Append(schemeColor20);
            A.PresetDash presetDash1 = new A.PresetDash(){ Val = A.PresetLineDashValues.Solid };
            A.Miter miter1 = new A.Miter(){ Limit = 800000 };

            outline1.Append(solidFill2);
            outline1.Append(presetDash1);
            outline1.Append(miter1);

            A.Outline outline2 = new A.Outline(){ Width = 19050, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill3 = new A.SolidFill();
            A.SchemeColor schemeColor21 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };

            solidFill3.Append(schemeColor21);
            A.PresetDash presetDash2 = new A.PresetDash(){ Val = A.PresetLineDashValues.Solid };
            A.Miter miter2 = new A.Miter(){ Limit = 800000 };

            outline2.Append(solidFill3);
            outline2.Append(presetDash2);
            outline2.Append(miter2);

            A.Outline outline3 = new A.Outline(){ Width = 25400, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill4 = new A.SolidFill();
            A.SchemeColor schemeColor22 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };

            solidFill4.Append(schemeColor22);
            A.PresetDash presetDash3 = new A.PresetDash(){ Val = A.PresetLineDashValues.Solid };
            A.Miter miter3 = new A.Miter(){ Limit = 800000 };

            outline3.Append(solidFill4);
            outline3.Append(presetDash3);
            outline3.Append(miter3);

            lineStyleList1.Append(outline1);
            lineStyleList1.Append(outline2);
            lineStyleList1.Append(outline3);

            A.EffectStyleList effectStyleList1 = new A.EffectStyleList();

            A.EffectStyle effectStyle1 = new A.EffectStyle();
            A.EffectList effectList1 = new A.EffectList();

            effectStyle1.Append(effectList1);

            A.EffectStyle effectStyle2 = new A.EffectStyle();
            A.EffectList effectList2 = new A.EffectList();

            effectStyle2.Append(effectList2);

            A.EffectStyle effectStyle3 = new A.EffectStyle();

            A.EffectList effectList3 = new A.EffectList();

            A.OuterShadow outerShadow1 = new A.OuterShadow(){ BlurRadius = 57150L, Distance = 19050L, Direction = 5400000, Alignment = A.RectangleAlignmentValues.Center, RotateWithShape = false };

            A.RgbColorModelHex rgbColorModelHex11 = new A.RgbColorModelHex(){ Val = "000000" };
            A.Alpha alpha1 = new A.Alpha(){ Val = 63000 };

            rgbColorModelHex11.Append(alpha1);

            outerShadow1.Append(rgbColorModelHex11);

            effectList3.Append(outerShadow1);

            effectStyle3.Append(effectList3);

            effectStyleList1.Append(effectStyle1);
            effectStyleList1.Append(effectStyle2);
            effectStyleList1.Append(effectStyle3);

            A.BackgroundFillStyleList backgroundFillStyleList1 = new A.BackgroundFillStyleList();

            A.SolidFill solidFill5 = new A.SolidFill();
            A.SchemeColor schemeColor23 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };

            solidFill5.Append(schemeColor23);

            A.SolidFill solidFill6 = new A.SolidFill();

            A.SchemeColor schemeColor24 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };
            A.Tint tint5 = new A.Tint(){ Val = 95000 };
            A.SaturationModulation saturationModulation7 = new A.SaturationModulation(){ Val = 170000 };

            schemeColor24.Append(tint5);
            schemeColor24.Append(saturationModulation7);

            solidFill6.Append(schemeColor24);

            A.GradientFill gradientFill3 = new A.GradientFill(){ RotateWithShape = true };

            A.GradientStopList gradientStopList3 = new A.GradientStopList();

            A.GradientStop gradientStop7 = new A.GradientStop(){ Position = 0 };

            A.SchemeColor schemeColor25 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };
            A.Tint tint6 = new A.Tint(){ Val = 93000 };
            A.SaturationModulation saturationModulation8 = new A.SaturationModulation(){ Val = 150000 };
            A.Shade shade6 = new A.Shade(){ Val = 98000 };
            A.LuminanceModulation luminanceModulation7 = new A.LuminanceModulation(){ Val = 102000 };

            schemeColor25.Append(tint6);
            schemeColor25.Append(saturationModulation8);
            schemeColor25.Append(shade6);
            schemeColor25.Append(luminanceModulation7);

            gradientStop7.Append(schemeColor25);

            A.GradientStop gradientStop8 = new A.GradientStop(){ Position = 50000 };

            A.SchemeColor schemeColor26 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };
            A.Tint tint7 = new A.Tint(){ Val = 98000 };
            A.SaturationModulation saturationModulation9 = new A.SaturationModulation(){ Val = 130000 };
            A.Shade shade7 = new A.Shade(){ Val = 90000 };
            A.LuminanceModulation luminanceModulation8 = new A.LuminanceModulation(){ Val = 103000 };

            schemeColor26.Append(tint7);
            schemeColor26.Append(saturationModulation9);
            schemeColor26.Append(shade7);
            schemeColor26.Append(luminanceModulation8);

            gradientStop8.Append(schemeColor26);

            A.GradientStop gradientStop9 = new A.GradientStop(){ Position = 100000 };

            A.SchemeColor schemeColor27 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };
            A.Shade shade8 = new A.Shade(){ Val = 63000 };
            A.SaturationModulation saturationModulation10 = new A.SaturationModulation(){ Val = 120000 };

            schemeColor27.Append(shade8);
            schemeColor27.Append(saturationModulation10);

            gradientStop9.Append(schemeColor27);

            gradientStopList3.Append(gradientStop7);
            gradientStopList3.Append(gradientStop8);
            gradientStopList3.Append(gradientStop9);
            A.LinearGradientFill linearGradientFill3 = new A.LinearGradientFill(){ Angle = 5400000, Scaled = false };

            gradientFill3.Append(gradientStopList3);
            gradientFill3.Append(linearGradientFill3);

            backgroundFillStyleList1.Append(solidFill5);
            backgroundFillStyleList1.Append(solidFill6);
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
            A.ShapeProperties shapeProperties4 = new A.ShapeProperties();
            A.BodyProperties bodyProperties1 = new A.BodyProperties();
            A.ListStyle listStyle1 = new A.ListStyle();

            A.ShapeStyle shapeStyle4 = new A.ShapeStyle();

            A.LineReference lineReference4 = new A.LineReference(){ Index = (UInt32Value)2U };
            A.SchemeColor schemeColor28 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent1 };

            lineReference4.Append(schemeColor28);

            A.FillReference fillReference4 = new A.FillReference(){ Index = (UInt32Value)0U };
            A.SchemeColor schemeColor29 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent1 };

            fillReference4.Append(schemeColor29);

            A.EffectReference effectReference4 = new A.EffectReference(){ Index = (UInt32Value)1U };
            A.SchemeColor schemeColor30 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent1 };

            effectReference4.Append(schemeColor30);

            A.FontReference fontReference4 = new A.FontReference(){ Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor31 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            fontReference4.Append(schemeColor31);

            shapeStyle4.Append(lineReference4);
            shapeStyle4.Append(fillReference4);
            shapeStyle4.Append(effectReference4);
            shapeStyle4.Append(fontReference4);

            lineDefault1.Append(shapeProperties4);
            lineDefault1.Append(bodyProperties1);
            lineDefault1.Append(listStyle1);
            lineDefault1.Append(shapeStyle4);

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

        private void SetPackageProperties(OpenXmlPackage document)
        {
            document.PackageProperties.Creator = "Przemysław Kłys";
            document.PackageProperties.Title = "";
            document.PackageProperties.Subject = "";
            document.PackageProperties.Keywords = "";
            document.PackageProperties.Description = "";
            document.PackageProperties.Revision = "1";
            document.PackageProperties.Created = System.Xml.XmlConvert.ToDateTime("2025-09-01T07:01:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
            document.PackageProperties.Modified = System.Xml.XmlConvert.ToDateTime("2025-09-01T07:02:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
            document.PackageProperties.LastModifiedBy = "Przemysław Kłys";
        }


    }
}
