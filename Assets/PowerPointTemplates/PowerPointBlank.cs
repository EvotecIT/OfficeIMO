using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;
using P15 = DocumentFormat.OpenXml.Office2013.PowerPoint;
using Thm15 = DocumentFormat.OpenXml.Office2013.Theme;
using Ap = DocumentFormat.OpenXml.ExtendedProperties;
using Vt = DocumentFormat.OpenXml.VariantTypes;

namespace GeneratedCode
{
    public class GeneratedClass
    {
        // Creates a PresentationDocument.
        public void CreatePackage(string filePath)
        {
            using(PresentationDocument package = PresentationDocument.Create(filePath, PresentationDocumentType.Presentation))
            {
                CreateParts(package);
            }
        }

        // Adds child parts and generates content of the specified part.
        private void CreateParts(PresentationDocument document)
        {
            ThumbnailPart thumbnailPart1 = document.AddNewPart<ThumbnailPart>("image/jpeg", "rId2");
            GenerateThumbnailPart1Content(thumbnailPart1);

            PresentationPart presentationPart1 = document.AddPresentationPart();
            GeneratePresentationPart1Content(presentationPart1);

            PresentationPropertiesPart presentationPropertiesPart1 = presentationPart1.AddNewPart<PresentationPropertiesPart>("rId3");
            GeneratePresentationPropertiesPart1Content(presentationPropertiesPart1);

            SlidePart slidePart1 = presentationPart1.AddNewPart<SlidePart>("rId2");
            GenerateSlidePart1Content(slidePart1);

            SlideLayoutPart slideLayoutPart1 = slidePart1.AddNewPart<SlideLayoutPart>("rId1");
            GenerateSlideLayoutPart1Content(slideLayoutPart1);

            SlideMasterPart slideMasterPart1 = slideLayoutPart1.AddNewPart<SlideMasterPart>("rId1");
            GenerateSlideMasterPart1Content(slideMasterPart1);

            SlideLayoutPart slideLayoutPart2 = slideMasterPart1.AddNewPart<SlideLayoutPart>("rId8");
            GenerateSlideLayoutPart2Content(slideLayoutPart2);

            slideLayoutPart2.AddPart(slideMasterPart1, "rId1");

            SlideLayoutPart slideLayoutPart3 = slideMasterPart1.AddNewPart<SlideLayoutPart>("rId3");
            GenerateSlideLayoutPart3Content(slideLayoutPart3);

            slideLayoutPart3.AddPart(slideMasterPart1, "rId1");

            SlideLayoutPart slideLayoutPart4 = slideMasterPart1.AddNewPart<SlideLayoutPart>("rId7");
            GenerateSlideLayoutPart4Content(slideLayoutPart4);

            slideLayoutPart4.AddPart(slideMasterPart1, "rId1");

            ThemePart themePart1 = slideMasterPart1.AddNewPart<ThemePart>("rId12");
            GenerateThemePart1Content(themePart1);

            SlideLayoutPart slideLayoutPart5 = slideMasterPart1.AddNewPart<SlideLayoutPart>("rId2");
            GenerateSlideLayoutPart5Content(slideLayoutPart5);

            slideLayoutPart5.AddPart(slideMasterPart1, "rId1");

            slideMasterPart1.AddPart(slideLayoutPart1, "rId1");

            SlideLayoutPart slideLayoutPart6 = slideMasterPart1.AddNewPart<SlideLayoutPart>("rId6");
            GenerateSlideLayoutPart6Content(slideLayoutPart6);

            slideLayoutPart6.AddPart(slideMasterPart1, "rId1");

            SlideLayoutPart slideLayoutPart7 = slideMasterPart1.AddNewPart<SlideLayoutPart>("rId11");
            GenerateSlideLayoutPart7Content(slideLayoutPart7);

            slideLayoutPart7.AddPart(slideMasterPart1, "rId1");

            SlideLayoutPart slideLayoutPart8 = slideMasterPart1.AddNewPart<SlideLayoutPart>("rId5");
            GenerateSlideLayoutPart8Content(slideLayoutPart8);

            slideLayoutPart8.AddPart(slideMasterPart1, "rId1");

            SlideLayoutPart slideLayoutPart9 = slideMasterPart1.AddNewPart<SlideLayoutPart>("rId10");
            GenerateSlideLayoutPart9Content(slideLayoutPart9);

            slideLayoutPart9.AddPart(slideMasterPart1, "rId1");

            SlideLayoutPart slideLayoutPart10 = slideMasterPart1.AddNewPart<SlideLayoutPart>("rId4");
            GenerateSlideLayoutPart10Content(slideLayoutPart10);

            slideLayoutPart10.AddPart(slideMasterPart1, "rId1");

            SlideLayoutPart slideLayoutPart11 = slideMasterPart1.AddNewPart<SlideLayoutPart>("rId9");
            GenerateSlideLayoutPart11Content(slideLayoutPart11);

            slideLayoutPart11.AddPart(slideMasterPart1, "rId1");

            presentationPart1.AddPart(slideMasterPart1, "rId1");

            TableStylesPart tableStylesPart1 = presentationPart1.AddNewPart<TableStylesPart>("rId6");
            GenerateTableStylesPart1Content(tableStylesPart1);

            presentationPart1.AddPart(themePart1, "rId5");

            ViewPropertiesPart viewPropertiesPart1 = presentationPart1.AddNewPart<ViewPropertiesPart>("rId4");
            GenerateViewPropertiesPart1Content(viewPropertiesPart1);

            ExtendedFilePropertiesPart extendedFilePropertiesPart1 = document.AddNewPart<ExtendedFilePropertiesPart>("rId4");
            GenerateExtendedFilePropertiesPart1Content(extendedFilePropertiesPart1);

            SetPackageProperties(document);
        }

        // Generates content of thumbnailPart1.
        private void GenerateThumbnailPart1Content(ThumbnailPart thumbnailPart1)
        {
            System.IO.Stream data = GetBinaryDataStream(thumbnailPart1Data);
            thumbnailPart1.FeedData(data);
            data.Close();
        }

        // Generates content of presentationPart1.
        private void GeneratePresentationPart1Content(PresentationPart presentationPart1)
        {
            Presentation presentation1 = new Presentation(){ SaveSubsetFonts = true };
            presentation1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            presentation1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            presentation1.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            SlideMasterIdList slideMasterIdList1 = new SlideMasterIdList();
            SlideMasterId slideMasterId1 = new SlideMasterId(){ Id = (UInt32Value)2147483648U, RelationshipId = "rId1" };

            slideMasterIdList1.Append(slideMasterId1);

            SlideIdList slideIdList1 = new SlideIdList();
            SlideId slideId1 = new SlideId(){ Id = (UInt32Value)256U, RelationshipId = "rId2" };

            slideIdList1.Append(slideId1);
            SlideSize slideSize1 = new SlideSize(){ Cx = 12192000, Cy = 6858000 };
            NotesSize notesSize1 = new NotesSize(){ Cx = 6858000L, Cy = 9144000L };

            DefaultTextStyle defaultTextStyle1 = new DefaultTextStyle();

            A.DefaultParagraphProperties defaultParagraphProperties1 = new A.DefaultParagraphProperties();
            A.DefaultRunProperties defaultRunProperties1 = new A.DefaultRunProperties(){ Language = "pl-PL" };

            defaultParagraphProperties1.Append(defaultRunProperties1);

            A.Level1ParagraphProperties level1ParagraphProperties1 = new A.Level1ParagraphProperties(){ LeftMargin = 0, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties2 = new A.DefaultRunProperties(){ FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill1 = new A.SolidFill();
            A.SchemeColor schemeColor1 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            solidFill1.Append(schemeColor1);
            A.LatinFont latinFont1 = new A.LatinFont(){ Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont1 = new A.EastAsianFont(){ Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont1 = new A.ComplexScriptFont(){ Typeface = "+mn-cs" };

            defaultRunProperties2.Append(solidFill1);
            defaultRunProperties2.Append(latinFont1);
            defaultRunProperties2.Append(eastAsianFont1);
            defaultRunProperties2.Append(complexScriptFont1);

            level1ParagraphProperties1.Append(defaultRunProperties2);

            A.Level2ParagraphProperties level2ParagraphProperties1 = new A.Level2ParagraphProperties(){ LeftMargin = 457200, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties3 = new A.DefaultRunProperties(){ FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill2 = new A.SolidFill();
            A.SchemeColor schemeColor2 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            solidFill2.Append(schemeColor2);
            A.LatinFont latinFont2 = new A.LatinFont(){ Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont2 = new A.EastAsianFont(){ Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont2 = new A.ComplexScriptFont(){ Typeface = "+mn-cs" };

            defaultRunProperties3.Append(solidFill2);
            defaultRunProperties3.Append(latinFont2);
            defaultRunProperties3.Append(eastAsianFont2);
            defaultRunProperties3.Append(complexScriptFont2);

            level2ParagraphProperties1.Append(defaultRunProperties3);

            A.Level3ParagraphProperties level3ParagraphProperties1 = new A.Level3ParagraphProperties(){ LeftMargin = 914400, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties4 = new A.DefaultRunProperties(){ FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill3 = new A.SolidFill();
            A.SchemeColor schemeColor3 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            solidFill3.Append(schemeColor3);
            A.LatinFont latinFont3 = new A.LatinFont(){ Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont3 = new A.EastAsianFont(){ Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont3 = new A.ComplexScriptFont(){ Typeface = "+mn-cs" };

            defaultRunProperties4.Append(solidFill3);
            defaultRunProperties4.Append(latinFont3);
            defaultRunProperties4.Append(eastAsianFont3);
            defaultRunProperties4.Append(complexScriptFont3);

            level3ParagraphProperties1.Append(defaultRunProperties4);

            A.Level4ParagraphProperties level4ParagraphProperties1 = new A.Level4ParagraphProperties(){ LeftMargin = 1371600, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties5 = new A.DefaultRunProperties(){ FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill4 = new A.SolidFill();
            A.SchemeColor schemeColor4 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            solidFill4.Append(schemeColor4);
            A.LatinFont latinFont4 = new A.LatinFont(){ Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont4 = new A.EastAsianFont(){ Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont4 = new A.ComplexScriptFont(){ Typeface = "+mn-cs" };

            defaultRunProperties5.Append(solidFill4);
            defaultRunProperties5.Append(latinFont4);
            defaultRunProperties5.Append(eastAsianFont4);
            defaultRunProperties5.Append(complexScriptFont4);

            level4ParagraphProperties1.Append(defaultRunProperties5);

            A.Level5ParagraphProperties level5ParagraphProperties1 = new A.Level5ParagraphProperties(){ LeftMargin = 1828800, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties6 = new A.DefaultRunProperties(){ FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill5 = new A.SolidFill();
            A.SchemeColor schemeColor5 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            solidFill5.Append(schemeColor5);
            A.LatinFont latinFont5 = new A.LatinFont(){ Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont5 = new A.EastAsianFont(){ Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont5 = new A.ComplexScriptFont(){ Typeface = "+mn-cs" };

            defaultRunProperties6.Append(solidFill5);
            defaultRunProperties6.Append(latinFont5);
            defaultRunProperties6.Append(eastAsianFont5);
            defaultRunProperties6.Append(complexScriptFont5);

            level5ParagraphProperties1.Append(defaultRunProperties6);

            A.Level6ParagraphProperties level6ParagraphProperties1 = new A.Level6ParagraphProperties(){ LeftMargin = 2286000, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties7 = new A.DefaultRunProperties(){ FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill6 = new A.SolidFill();
            A.SchemeColor schemeColor6 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            solidFill6.Append(schemeColor6);
            A.LatinFont latinFont6 = new A.LatinFont(){ Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont6 = new A.EastAsianFont(){ Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont6 = new A.ComplexScriptFont(){ Typeface = "+mn-cs" };

            defaultRunProperties7.Append(solidFill6);
            defaultRunProperties7.Append(latinFont6);
            defaultRunProperties7.Append(eastAsianFont6);
            defaultRunProperties7.Append(complexScriptFont6);

            level6ParagraphProperties1.Append(defaultRunProperties7);

            A.Level7ParagraphProperties level7ParagraphProperties1 = new A.Level7ParagraphProperties(){ LeftMargin = 2743200, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties8 = new A.DefaultRunProperties(){ FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill7 = new A.SolidFill();
            A.SchemeColor schemeColor7 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            solidFill7.Append(schemeColor7);
            A.LatinFont latinFont7 = new A.LatinFont(){ Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont7 = new A.EastAsianFont(){ Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont7 = new A.ComplexScriptFont(){ Typeface = "+mn-cs" };

            defaultRunProperties8.Append(solidFill7);
            defaultRunProperties8.Append(latinFont7);
            defaultRunProperties8.Append(eastAsianFont7);
            defaultRunProperties8.Append(complexScriptFont7);

            level7ParagraphProperties1.Append(defaultRunProperties8);

            A.Level8ParagraphProperties level8ParagraphProperties1 = new A.Level8ParagraphProperties(){ LeftMargin = 3200400, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties9 = new A.DefaultRunProperties(){ FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill8 = new A.SolidFill();
            A.SchemeColor schemeColor8 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            solidFill8.Append(schemeColor8);
            A.LatinFont latinFont8 = new A.LatinFont(){ Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont8 = new A.EastAsianFont(){ Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont8 = new A.ComplexScriptFont(){ Typeface = "+mn-cs" };

            defaultRunProperties9.Append(solidFill8);
            defaultRunProperties9.Append(latinFont8);
            defaultRunProperties9.Append(eastAsianFont8);
            defaultRunProperties9.Append(complexScriptFont8);

            level8ParagraphProperties1.Append(defaultRunProperties9);

            A.Level9ParagraphProperties level9ParagraphProperties1 = new A.Level9ParagraphProperties(){ LeftMargin = 3657600, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties10 = new A.DefaultRunProperties(){ FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill9 = new A.SolidFill();
            A.SchemeColor schemeColor9 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            solidFill9.Append(schemeColor9);
            A.LatinFont latinFont9 = new A.LatinFont(){ Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont9 = new A.EastAsianFont(){ Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont9 = new A.ComplexScriptFont(){ Typeface = "+mn-cs" };

            defaultRunProperties10.Append(solidFill9);
            defaultRunProperties10.Append(latinFont9);
            defaultRunProperties10.Append(eastAsianFont9);
            defaultRunProperties10.Append(complexScriptFont9);

            level9ParagraphProperties1.Append(defaultRunProperties10);

            defaultTextStyle1.Append(defaultParagraphProperties1);
            defaultTextStyle1.Append(level1ParagraphProperties1);
            defaultTextStyle1.Append(level2ParagraphProperties1);
            defaultTextStyle1.Append(level3ParagraphProperties1);
            defaultTextStyle1.Append(level4ParagraphProperties1);
            defaultTextStyle1.Append(level5ParagraphProperties1);
            defaultTextStyle1.Append(level6ParagraphProperties1);
            defaultTextStyle1.Append(level7ParagraphProperties1);
            defaultTextStyle1.Append(level8ParagraphProperties1);
            defaultTextStyle1.Append(level9ParagraphProperties1);

            presentation1.Append(slideMasterIdList1);
            presentation1.Append(slideIdList1);
            presentation1.Append(slideSize1);
            presentation1.Append(notesSize1);
            presentation1.Append(defaultTextStyle1);

            presentationPart1.Presentation = presentation1;
        }

        // Generates content of presentationPropertiesPart1.
        private void GeneratePresentationPropertiesPart1Content(PresentationPropertiesPart presentationPropertiesPart1)
        {
            PresentationProperties presentationProperties1 = new PresentationProperties();
            presentationProperties1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            presentationProperties1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            presentationProperties1.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            PresentationPropertiesExtensionList presentationPropertiesExtensionList1 = new PresentationPropertiesExtensionList();

            PresentationPropertiesExtension presentationPropertiesExtension1 = new PresentationPropertiesExtension(){ Uri = "{E76CE94A-603C-4142-B9EB-6D1370010A27}" };

            P14.DiscardImageEditData discardImageEditData1 = new P14.DiscardImageEditData(){ Val = false };
            discardImageEditData1.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            presentationPropertiesExtension1.Append(discardImageEditData1);

            PresentationPropertiesExtension presentationPropertiesExtension2 = new PresentationPropertiesExtension(){ Uri = "{D31A062A-798A-4329-ABDD-BBA856620510}" };

            P14.DefaultImageDpi defaultImageDpi1 = new P14.DefaultImageDpi(){ Val = (UInt32Value)32767U };
            defaultImageDpi1.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            presentationPropertiesExtension2.Append(defaultImageDpi1);

            PresentationPropertiesExtension presentationPropertiesExtension3 = new PresentationPropertiesExtension(){ Uri = "{FD5EFAAD-0ECE-453E-9831-46B23BE46B34}" };

            P15.ChartTrackingReferenceBased chartTrackingReferenceBased1 = new P15.ChartTrackingReferenceBased(){ Val = true };
            chartTrackingReferenceBased1.AddNamespaceDeclaration("p15", "http://schemas.microsoft.com/office/powerpoint/2012/main");

            presentationPropertiesExtension3.Append(chartTrackingReferenceBased1);

            presentationPropertiesExtensionList1.Append(presentationPropertiesExtension1);
            presentationPropertiesExtensionList1.Append(presentationPropertiesExtension2);
            presentationPropertiesExtensionList1.Append(presentationPropertiesExtension3);

            presentationProperties1.Append(presentationPropertiesExtensionList1);

            presentationPropertiesPart1.PresentationProperties = presentationProperties1;
        }

        // Generates content of slidePart1.
        private void GenerateSlidePart1Content(SlidePart slidePart1)
        {
            Slide slide1 = new Slide();
            slide1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            slide1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            slide1.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            CommonSlideData commonSlideData1 = new CommonSlideData();

            ShapeTree shapeTree1 = new ShapeTree();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties1 = new NonVisualGroupShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties1 = new NonVisualDrawingProperties(){ Id = (UInt32Value)1U, Name = "" };
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties1 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties1 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties1.Append(nonVisualDrawingProperties1);
            nonVisualGroupShapeProperties1.Append(nonVisualGroupShapeDrawingProperties1);
            nonVisualGroupShapeProperties1.Append(applicationNonVisualDrawingProperties1);

            GroupShapeProperties groupShapeProperties1 = new GroupShapeProperties();

            A.TransformGroup transformGroup1 = new A.TransformGroup();
            A.Offset offset1 = new A.Offset(){ X = 0L, Y = 0L };
            A.Extents extents1 = new A.Extents(){ Cx = 0L, Cy = 0L };
            A.ChildOffset childOffset1 = new A.ChildOffset(){ X = 0L, Y = 0L };
            A.ChildExtents childExtents1 = new A.ChildExtents(){ Cx = 0L, Cy = 0L };

            transformGroup1.Append(offset1);
            transformGroup1.Append(extents1);
            transformGroup1.Append(childOffset1);
            transformGroup1.Append(childExtents1);

            groupShapeProperties1.Append(transformGroup1);

            shapeTree1.Append(nonVisualGroupShapeProperties1);
            shapeTree1.Append(groupShapeProperties1);

            CommonSlideDataExtensionList commonSlideDataExtensionList1 = new CommonSlideDataExtensionList();

            CommonSlideDataExtension commonSlideDataExtension1 = new CommonSlideDataExtension(){ Uri = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}" };

            P14.CreationId creationId1 = new P14.CreationId(){ Val = (UInt32Value)234584052U };
            creationId1.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            commonSlideDataExtension1.Append(creationId1);

            commonSlideDataExtensionList1.Append(commonSlideDataExtension1);

            commonSlideData1.Append(shapeTree1);
            commonSlideData1.Append(commonSlideDataExtensionList1);

            ColorMapOverride colorMapOverride1 = new ColorMapOverride();
            A.MasterColorMapping masterColorMapping1 = new A.MasterColorMapping();

            colorMapOverride1.Append(masterColorMapping1);

            slide1.Append(commonSlideData1);
            slide1.Append(colorMapOverride1);

            slidePart1.Slide = slide1;
        }

        // Generates content of slideLayoutPart1.
        private void GenerateSlideLayoutPart1Content(SlideLayoutPart slideLayoutPart1)
        {
            SlideLayout slideLayout1 = new SlideLayout(){ Type = SlideLayoutValues.Title, Preserve = true };
            slideLayout1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            slideLayout1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            slideLayout1.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            CommonSlideData commonSlideData2 = new CommonSlideData(){ Name = "Title Slide" };

            ShapeTree shapeTree2 = new ShapeTree();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties2 = new NonVisualGroupShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties2 = new NonVisualDrawingProperties(){ Id = (UInt32Value)1U, Name = "" };
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties2 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties2 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties2.Append(nonVisualDrawingProperties2);
            nonVisualGroupShapeProperties2.Append(nonVisualGroupShapeDrawingProperties2);
            nonVisualGroupShapeProperties2.Append(applicationNonVisualDrawingProperties2);

            GroupShapeProperties groupShapeProperties2 = new GroupShapeProperties();

            A.TransformGroup transformGroup2 = new A.TransformGroup();
            A.Offset offset2 = new A.Offset(){ X = 0L, Y = 0L };
            A.Extents extents2 = new A.Extents(){ Cx = 0L, Cy = 0L };
            A.ChildOffset childOffset2 = new A.ChildOffset(){ X = 0L, Y = 0L };
            A.ChildExtents childExtents2 = new A.ChildExtents(){ Cx = 0L, Cy = 0L };

            transformGroup2.Append(offset2);
            transformGroup2.Append(extents2);
            transformGroup2.Append(childOffset2);
            transformGroup2.Append(childExtents2);

            groupShapeProperties2.Append(transformGroup2);

            Shape shape1 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties1 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties3 = new NonVisualDrawingProperties(){ Id = (UInt32Value)2U, Name = "Title 1" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList1 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension1 = new A.NonVisualDrawingPropertiesExtension(){ Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            OpenXmlUnknownElement openXmlUnknownElement1 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{8B34C2AD-2B89-D91D-BFF9-62C33D9EF755}\" />");

            nonVisualDrawingPropertiesExtension1.Append(openXmlUnknownElement1);

            nonVisualDrawingPropertiesExtensionList1.Append(nonVisualDrawingPropertiesExtension1);

            nonVisualDrawingProperties3.Append(nonVisualDrawingPropertiesExtensionList1);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties1 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks1 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties1.Append(shapeLocks1);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties3 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape1 = new PlaceholderShape(){ Type = PlaceholderValues.CenteredTitle };

            applicationNonVisualDrawingProperties3.Append(placeholderShape1);

            nonVisualShapeProperties1.Append(nonVisualDrawingProperties3);
            nonVisualShapeProperties1.Append(nonVisualShapeDrawingProperties1);
            nonVisualShapeProperties1.Append(applicationNonVisualDrawingProperties3);

            ShapeProperties shapeProperties1 = new ShapeProperties();

            A.Transform2D transform2D1 = new A.Transform2D();
            A.Offset offset3 = new A.Offset(){ X = 1524000L, Y = 1122363L };
            A.Extents extents3 = new A.Extents(){ Cx = 9144000L, Cy = 2387600L };

            transform2D1.Append(offset3);
            transform2D1.Append(extents3);

            shapeProperties1.Append(transform2D1);

            TextBody textBody1 = new TextBody();
            A.BodyProperties bodyProperties1 = new A.BodyProperties(){ Anchor = A.TextAnchoringTypeValues.Bottom };

            A.ListStyle listStyle1 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties2 = new A.Level1ParagraphProperties(){ Alignment = A.TextAlignmentTypeValues.Center };
            A.DefaultRunProperties defaultRunProperties11 = new A.DefaultRunProperties(){ FontSize = 6000 };

            level1ParagraphProperties2.Append(defaultRunProperties11);

            listStyle1.Append(level1ParagraphProperties2);

            A.Paragraph paragraph1 = new A.Paragraph();

            A.Run run1 = new A.Run();
            A.RunProperties runProperties1 = new A.RunProperties(){ Language = "en-US" };
            A.Text text1 = new A.Text();
            text1.Text = "Click to edit Master title style";

            run1.Append(runProperties1);
            run1.Append(text1);
            A.EndParagraphRunProperties endParagraphRunProperties1 = new A.EndParagraphRunProperties(){ Language = "pl-PL" };

            paragraph1.Append(run1);
            paragraph1.Append(endParagraphRunProperties1);

            textBody1.Append(bodyProperties1);
            textBody1.Append(listStyle1);
            textBody1.Append(paragraph1);

            shape1.Append(nonVisualShapeProperties1);
            shape1.Append(shapeProperties1);
            shape1.Append(textBody1);

            Shape shape2 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties2 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties4 = new NonVisualDrawingProperties(){ Id = (UInt32Value)3U, Name = "Subtitle 2" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList2 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension2 = new A.NonVisualDrawingPropertiesExtension(){ Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            OpenXmlUnknownElement openXmlUnknownElement2 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{82BF5914-4F8A-3ECA-543E-205F42B78D5A}\" />");

            nonVisualDrawingPropertiesExtension2.Append(openXmlUnknownElement2);

            nonVisualDrawingPropertiesExtensionList2.Append(nonVisualDrawingPropertiesExtension2);

            nonVisualDrawingProperties4.Append(nonVisualDrawingPropertiesExtensionList2);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties2 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks2 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties2.Append(shapeLocks2);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties4 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape2 = new PlaceholderShape(){ Type = PlaceholderValues.SubTitle, Index = (UInt32Value)1U };

            applicationNonVisualDrawingProperties4.Append(placeholderShape2);

            nonVisualShapeProperties2.Append(nonVisualDrawingProperties4);
            nonVisualShapeProperties2.Append(nonVisualShapeDrawingProperties2);
            nonVisualShapeProperties2.Append(applicationNonVisualDrawingProperties4);

            ShapeProperties shapeProperties2 = new ShapeProperties();

            A.Transform2D transform2D2 = new A.Transform2D();
            A.Offset offset4 = new A.Offset(){ X = 1524000L, Y = 3602038L };
            A.Extents extents4 = new A.Extents(){ Cx = 9144000L, Cy = 1655762L };

            transform2D2.Append(offset4);
            transform2D2.Append(extents4);

            shapeProperties2.Append(transform2D2);

            TextBody textBody2 = new TextBody();
            A.BodyProperties bodyProperties2 = new A.BodyProperties();

            A.ListStyle listStyle2 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties3 = new A.Level1ParagraphProperties(){ LeftMargin = 0, Indent = 0, Alignment = A.TextAlignmentTypeValues.Center };
            A.NoBullet noBullet1 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties12 = new A.DefaultRunProperties(){ FontSize = 2400 };

            level1ParagraphProperties3.Append(noBullet1);
            level1ParagraphProperties3.Append(defaultRunProperties12);

            A.Level2ParagraphProperties level2ParagraphProperties2 = new A.Level2ParagraphProperties(){ LeftMargin = 457200, Indent = 0, Alignment = A.TextAlignmentTypeValues.Center };
            A.NoBullet noBullet2 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties13 = new A.DefaultRunProperties(){ FontSize = 2000 };

            level2ParagraphProperties2.Append(noBullet2);
            level2ParagraphProperties2.Append(defaultRunProperties13);

            A.Level3ParagraphProperties level3ParagraphProperties2 = new A.Level3ParagraphProperties(){ LeftMargin = 914400, Indent = 0, Alignment = A.TextAlignmentTypeValues.Center };
            A.NoBullet noBullet3 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties14 = new A.DefaultRunProperties(){ FontSize = 1800 };

            level3ParagraphProperties2.Append(noBullet3);
            level3ParagraphProperties2.Append(defaultRunProperties14);

            A.Level4ParagraphProperties level4ParagraphProperties2 = new A.Level4ParagraphProperties(){ LeftMargin = 1371600, Indent = 0, Alignment = A.TextAlignmentTypeValues.Center };
            A.NoBullet noBullet4 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties15 = new A.DefaultRunProperties(){ FontSize = 1600 };

            level4ParagraphProperties2.Append(noBullet4);
            level4ParagraphProperties2.Append(defaultRunProperties15);

            A.Level5ParagraphProperties level5ParagraphProperties2 = new A.Level5ParagraphProperties(){ LeftMargin = 1828800, Indent = 0, Alignment = A.TextAlignmentTypeValues.Center };
            A.NoBullet noBullet5 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties16 = new A.DefaultRunProperties(){ FontSize = 1600 };

            level5ParagraphProperties2.Append(noBullet5);
            level5ParagraphProperties2.Append(defaultRunProperties16);

            A.Level6ParagraphProperties level6ParagraphProperties2 = new A.Level6ParagraphProperties(){ LeftMargin = 2286000, Indent = 0, Alignment = A.TextAlignmentTypeValues.Center };
            A.NoBullet noBullet6 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties17 = new A.DefaultRunProperties(){ FontSize = 1600 };

            level6ParagraphProperties2.Append(noBullet6);
            level6ParagraphProperties2.Append(defaultRunProperties17);

            A.Level7ParagraphProperties level7ParagraphProperties2 = new A.Level7ParagraphProperties(){ LeftMargin = 2743200, Indent = 0, Alignment = A.TextAlignmentTypeValues.Center };
            A.NoBullet noBullet7 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties18 = new A.DefaultRunProperties(){ FontSize = 1600 };

            level7ParagraphProperties2.Append(noBullet7);
            level7ParagraphProperties2.Append(defaultRunProperties18);

            A.Level8ParagraphProperties level8ParagraphProperties2 = new A.Level8ParagraphProperties(){ LeftMargin = 3200400, Indent = 0, Alignment = A.TextAlignmentTypeValues.Center };
            A.NoBullet noBullet8 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties19 = new A.DefaultRunProperties(){ FontSize = 1600 };

            level8ParagraphProperties2.Append(noBullet8);
            level8ParagraphProperties2.Append(defaultRunProperties19);

            A.Level9ParagraphProperties level9ParagraphProperties2 = new A.Level9ParagraphProperties(){ LeftMargin = 3657600, Indent = 0, Alignment = A.TextAlignmentTypeValues.Center };
            A.NoBullet noBullet9 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties20 = new A.DefaultRunProperties(){ FontSize = 1600 };

            level9ParagraphProperties2.Append(noBullet9);
            level9ParagraphProperties2.Append(defaultRunProperties20);

            listStyle2.Append(level1ParagraphProperties3);
            listStyle2.Append(level2ParagraphProperties2);
            listStyle2.Append(level3ParagraphProperties2);
            listStyle2.Append(level4ParagraphProperties2);
            listStyle2.Append(level5ParagraphProperties2);
            listStyle2.Append(level6ParagraphProperties2);
            listStyle2.Append(level7ParagraphProperties2);
            listStyle2.Append(level8ParagraphProperties2);
            listStyle2.Append(level9ParagraphProperties2);

            A.Paragraph paragraph2 = new A.Paragraph();

            A.Run run2 = new A.Run();
            A.RunProperties runProperties2 = new A.RunProperties(){ Language = "en-US" };
            A.Text text2 = new A.Text();
            text2.Text = "Click to edit Master subtitle style";

            run2.Append(runProperties2);
            run2.Append(text2);
            A.EndParagraphRunProperties endParagraphRunProperties2 = new A.EndParagraphRunProperties(){ Language = "pl-PL" };

            paragraph2.Append(run2);
            paragraph2.Append(endParagraphRunProperties2);

            textBody2.Append(bodyProperties2);
            textBody2.Append(listStyle2);
            textBody2.Append(paragraph2);

            shape2.Append(nonVisualShapeProperties2);
            shape2.Append(shapeProperties2);
            shape2.Append(textBody2);

            Shape shape3 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties3 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties5 = new NonVisualDrawingProperties(){ Id = (UInt32Value)4U, Name = "Date Placeholder 3" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList3 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension3 = new A.NonVisualDrawingPropertiesExtension(){ Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            OpenXmlUnknownElement openXmlUnknownElement3 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{8A3AC2FF-2F37-46B6-1EED-FA6325F80CF1}\" />");

            nonVisualDrawingPropertiesExtension3.Append(openXmlUnknownElement3);

            nonVisualDrawingPropertiesExtensionList3.Append(nonVisualDrawingPropertiesExtension3);

            nonVisualDrawingProperties5.Append(nonVisualDrawingPropertiesExtensionList3);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties3 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks3 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties3.Append(shapeLocks3);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties5 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape3 = new PlaceholderShape(){ Type = PlaceholderValues.DateAndTime, Size = PlaceholderSizeValues.Half, Index = (UInt32Value)10U };

            applicationNonVisualDrawingProperties5.Append(placeholderShape3);

            nonVisualShapeProperties3.Append(nonVisualDrawingProperties5);
            nonVisualShapeProperties3.Append(nonVisualShapeDrawingProperties3);
            nonVisualShapeProperties3.Append(applicationNonVisualDrawingProperties5);
            ShapeProperties shapeProperties3 = new ShapeProperties();

            TextBody textBody3 = new TextBody();
            A.BodyProperties bodyProperties3 = new A.BodyProperties();
            A.ListStyle listStyle3 = new A.ListStyle();

            A.Paragraph paragraph3 = new A.Paragraph();

            A.Field field1 = new A.Field(){ Id = "{F550F806-175B-4EE1-ABEA-A5F04E4FD72B}", Type = "datetimeFigureOut" };

            A.RunProperties runProperties3 = new A.RunProperties(){ Language = "pl-PL" };
            runProperties3.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text3 = new A.Text();
            text3.Text = "20.08.2025";

            field1.Append(runProperties3);
            field1.Append(text3);
            A.EndParagraphRunProperties endParagraphRunProperties3 = new A.EndParagraphRunProperties(){ Language = "pl-PL" };

            paragraph3.Append(field1);
            paragraph3.Append(endParagraphRunProperties3);

            textBody3.Append(bodyProperties3);
            textBody3.Append(listStyle3);
            textBody3.Append(paragraph3);

            shape3.Append(nonVisualShapeProperties3);
            shape3.Append(shapeProperties3);
            shape3.Append(textBody3);

            Shape shape4 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties4 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties6 = new NonVisualDrawingProperties(){ Id = (UInt32Value)5U, Name = "Footer Placeholder 4" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList4 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension4 = new A.NonVisualDrawingPropertiesExtension(){ Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            OpenXmlUnknownElement openXmlUnknownElement4 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{42488568-0327-012F-AFE1-5B034F8E109B}\" />");

            nonVisualDrawingPropertiesExtension4.Append(openXmlUnknownElement4);

            nonVisualDrawingPropertiesExtensionList4.Append(nonVisualDrawingPropertiesExtension4);

            nonVisualDrawingProperties6.Append(nonVisualDrawingPropertiesExtensionList4);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties4 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks4 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties4.Append(shapeLocks4);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties6 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape4 = new PlaceholderShape(){ Type = PlaceholderValues.Footer, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)11U };

            applicationNonVisualDrawingProperties6.Append(placeholderShape4);

            nonVisualShapeProperties4.Append(nonVisualDrawingProperties6);
            nonVisualShapeProperties4.Append(nonVisualShapeDrawingProperties4);
            nonVisualShapeProperties4.Append(applicationNonVisualDrawingProperties6);
            ShapeProperties shapeProperties4 = new ShapeProperties();

            TextBody textBody4 = new TextBody();
            A.BodyProperties bodyProperties4 = new A.BodyProperties();
            A.ListStyle listStyle4 = new A.ListStyle();

            A.Paragraph paragraph4 = new A.Paragraph();
            A.EndParagraphRunProperties endParagraphRunProperties4 = new A.EndParagraphRunProperties(){ Language = "pl-PL" };

            paragraph4.Append(endParagraphRunProperties4);

            textBody4.Append(bodyProperties4);
            textBody4.Append(listStyle4);
            textBody4.Append(paragraph4);

            shape4.Append(nonVisualShapeProperties4);
            shape4.Append(shapeProperties4);
            shape4.Append(textBody4);

            Shape shape5 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties5 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties7 = new NonVisualDrawingProperties(){ Id = (UInt32Value)6U, Name = "Slide Number Placeholder 5" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList5 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension5 = new A.NonVisualDrawingPropertiesExtension(){ Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            OpenXmlUnknownElement openXmlUnknownElement5 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{BD89448F-A6A5-7100-8437-1CE0B3C70555}\" />");

            nonVisualDrawingPropertiesExtension5.Append(openXmlUnknownElement5);

            nonVisualDrawingPropertiesExtensionList5.Append(nonVisualDrawingPropertiesExtension5);

            nonVisualDrawingProperties7.Append(nonVisualDrawingPropertiesExtensionList5);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties5 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks5 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties5.Append(shapeLocks5);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties7 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape5 = new PlaceholderShape(){ Type = PlaceholderValues.SlideNumber, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)12U };

            applicationNonVisualDrawingProperties7.Append(placeholderShape5);

            nonVisualShapeProperties5.Append(nonVisualDrawingProperties7);
            nonVisualShapeProperties5.Append(nonVisualShapeDrawingProperties5);
            nonVisualShapeProperties5.Append(applicationNonVisualDrawingProperties7);
            ShapeProperties shapeProperties5 = new ShapeProperties();

            TextBody textBody5 = new TextBody();
            A.BodyProperties bodyProperties5 = new A.BodyProperties();
            A.ListStyle listStyle5 = new A.ListStyle();

            A.Paragraph paragraph5 = new A.Paragraph();

            A.Field field2 = new A.Field(){ Id = "{5FAB34AF-B09A-4E2A-BAD3-CCD6E0AD81E6}", Type = "slidenum" };

            A.RunProperties runProperties4 = new A.RunProperties(){ Language = "pl-PL" };
            runProperties4.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text4 = new A.Text();
            text4.Text = "‹#›";

            field2.Append(runProperties4);
            field2.Append(text4);
            A.EndParagraphRunProperties endParagraphRunProperties5 = new A.EndParagraphRunProperties(){ Language = "pl-PL" };

            paragraph5.Append(field2);
            paragraph5.Append(endParagraphRunProperties5);

            textBody5.Append(bodyProperties5);
            textBody5.Append(listStyle5);
            textBody5.Append(paragraph5);

            shape5.Append(nonVisualShapeProperties5);
            shape5.Append(shapeProperties5);
            shape5.Append(textBody5);

            shapeTree2.Append(nonVisualGroupShapeProperties2);
            shapeTree2.Append(groupShapeProperties2);
            shapeTree2.Append(shape1);
            shapeTree2.Append(shape2);
            shapeTree2.Append(shape3);
            shapeTree2.Append(shape4);
            shapeTree2.Append(shape5);

            CommonSlideDataExtensionList commonSlideDataExtensionList2 = new CommonSlideDataExtensionList();

            CommonSlideDataExtension commonSlideDataExtension2 = new CommonSlideDataExtension(){ Uri = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}" };

            P14.CreationId creationId2 = new P14.CreationId(){ Val = (UInt32Value)82025322U };
            creationId2.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            commonSlideDataExtension2.Append(creationId2);

            commonSlideDataExtensionList2.Append(commonSlideDataExtension2);

            commonSlideData2.Append(shapeTree2);
            commonSlideData2.Append(commonSlideDataExtensionList2);

            ColorMapOverride colorMapOverride2 = new ColorMapOverride();
            A.MasterColorMapping masterColorMapping2 = new A.MasterColorMapping();

            colorMapOverride2.Append(masterColorMapping2);

            slideLayout1.Append(commonSlideData2);
            slideLayout1.Append(colorMapOverride2);

            slideLayoutPart1.SlideLayout = slideLayout1;
        }

        // Generates content of slideMasterPart1.
        private void GenerateSlideMasterPart1Content(SlideMasterPart slideMasterPart1)
        {
            SlideMaster slideMaster1 = new SlideMaster();
            slideMaster1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            slideMaster1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            slideMaster1.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            CommonSlideData commonSlideData3 = new CommonSlideData();

            Background background1 = new Background();

            BackgroundStyleReference backgroundStyleReference1 = new BackgroundStyleReference(){ Index = (UInt32Value)1001U };
            A.SchemeColor schemeColor10 = new A.SchemeColor(){ Val = A.SchemeColorValues.Background1 };

            backgroundStyleReference1.Append(schemeColor10);

            background1.Append(backgroundStyleReference1);

            ShapeTree shapeTree3 = new ShapeTree();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties3 = new NonVisualGroupShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties8 = new NonVisualDrawingProperties(){ Id = (UInt32Value)1U, Name = "" };
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties3 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties8 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties3.Append(nonVisualDrawingProperties8);
            nonVisualGroupShapeProperties3.Append(nonVisualGroupShapeDrawingProperties3);
            nonVisualGroupShapeProperties3.Append(applicationNonVisualDrawingProperties8);

            GroupShapeProperties groupShapeProperties3 = new GroupShapeProperties();

            A.TransformGroup transformGroup3 = new A.TransformGroup();
            A.Offset offset5 = new A.Offset(){ X = 0L, Y = 0L };
            A.Extents extents5 = new A.Extents(){ Cx = 0L, Cy = 0L };
            A.ChildOffset childOffset3 = new A.ChildOffset(){ X = 0L, Y = 0L };
            A.ChildExtents childExtents3 = new A.ChildExtents(){ Cx = 0L, Cy = 0L };

            transformGroup3.Append(offset5);
            transformGroup3.Append(extents5);
            transformGroup3.Append(childOffset3);
            transformGroup3.Append(childExtents3);

            groupShapeProperties3.Append(transformGroup3);

            Shape shape6 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties6 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties9 = new NonVisualDrawingProperties(){ Id = (UInt32Value)2U, Name = "Title Placeholder 1" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList6 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension6 = new A.NonVisualDrawingPropertiesExtension(){ Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            OpenXmlUnknownElement openXmlUnknownElement6 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{BCAD55D7-8407-282B-E405-D84A21921CC1}\" />");

            nonVisualDrawingPropertiesExtension6.Append(openXmlUnknownElement6);

            nonVisualDrawingPropertiesExtensionList6.Append(nonVisualDrawingPropertiesExtension6);

            nonVisualDrawingProperties9.Append(nonVisualDrawingPropertiesExtensionList6);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties6 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks6 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties6.Append(shapeLocks6);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties9 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape6 = new PlaceholderShape(){ Type = PlaceholderValues.Title };

            applicationNonVisualDrawingProperties9.Append(placeholderShape6);

            nonVisualShapeProperties6.Append(nonVisualDrawingProperties9);
            nonVisualShapeProperties6.Append(nonVisualShapeDrawingProperties6);
            nonVisualShapeProperties6.Append(applicationNonVisualDrawingProperties9);

            ShapeProperties shapeProperties6 = new ShapeProperties();

            A.Transform2D transform2D3 = new A.Transform2D();
            A.Offset offset6 = new A.Offset(){ X = 838200L, Y = 365125L };
            A.Extents extents6 = new A.Extents(){ Cx = 10515600L, Cy = 1325563L };

            transform2D3.Append(offset6);
            transform2D3.Append(extents6);

            A.PresetGeometry presetGeometry1 = new A.PresetGeometry(){ Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList1 = new A.AdjustValueList();

            presetGeometry1.Append(adjustValueList1);

            shapeProperties6.Append(transform2D3);
            shapeProperties6.Append(presetGeometry1);

            TextBody textBody6 = new TextBody();

            A.BodyProperties bodyProperties6 = new A.BodyProperties(){ Vertical = A.TextVerticalValues.Horizontal, LeftInset = 91440, TopInset = 45720, RightInset = 91440, BottomInset = 45720, RightToLeftColumns = false, Anchor = A.TextAnchoringTypeValues.Center };
            A.NormalAutoFit normalAutoFit1 = new A.NormalAutoFit();

            bodyProperties6.Append(normalAutoFit1);
            A.ListStyle listStyle6 = new A.ListStyle();

            A.Paragraph paragraph6 = new A.Paragraph();

            A.Run run3 = new A.Run();
            A.RunProperties runProperties5 = new A.RunProperties(){ Language = "en-US" };
            A.Text text5 = new A.Text();
            text5.Text = "Click to edit Master title style";

            run3.Append(runProperties5);
            run3.Append(text5);
            A.EndParagraphRunProperties endParagraphRunProperties6 = new A.EndParagraphRunProperties(){ Language = "pl-PL" };

            paragraph6.Append(run3);
            paragraph6.Append(endParagraphRunProperties6);

            textBody6.Append(bodyProperties6);
            textBody6.Append(listStyle6);
            textBody6.Append(paragraph6);

            shape6.Append(nonVisualShapeProperties6);
            shape6.Append(shapeProperties6);
            shape6.Append(textBody6);

            Shape shape7 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties7 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties10 = new NonVisualDrawingProperties(){ Id = (UInt32Value)3U, Name = "Text Placeholder 2" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList7 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension7 = new A.NonVisualDrawingPropertiesExtension(){ Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            OpenXmlUnknownElement openXmlUnknownElement7 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{49ED3347-F499-82C3-5008-A73716BFD37D}\" />");

            nonVisualDrawingPropertiesExtension7.Append(openXmlUnknownElement7);

            nonVisualDrawingPropertiesExtensionList7.Append(nonVisualDrawingPropertiesExtension7);

            nonVisualDrawingProperties10.Append(nonVisualDrawingPropertiesExtensionList7);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties7 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks7 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties7.Append(shapeLocks7);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties10 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape7 = new PlaceholderShape(){ Type = PlaceholderValues.Body, Index = (UInt32Value)1U };

            applicationNonVisualDrawingProperties10.Append(placeholderShape7);

            nonVisualShapeProperties7.Append(nonVisualDrawingProperties10);
            nonVisualShapeProperties7.Append(nonVisualShapeDrawingProperties7);
            nonVisualShapeProperties7.Append(applicationNonVisualDrawingProperties10);

            ShapeProperties shapeProperties7 = new ShapeProperties();

            A.Transform2D transform2D4 = new A.Transform2D();
            A.Offset offset7 = new A.Offset(){ X = 838200L, Y = 1825625L };
            A.Extents extents7 = new A.Extents(){ Cx = 10515600L, Cy = 4351338L };

            transform2D4.Append(offset7);
            transform2D4.Append(extents7);

            A.PresetGeometry presetGeometry2 = new A.PresetGeometry(){ Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList2 = new A.AdjustValueList();

            presetGeometry2.Append(adjustValueList2);

            shapeProperties7.Append(transform2D4);
            shapeProperties7.Append(presetGeometry2);

            TextBody textBody7 = new TextBody();

            A.BodyProperties bodyProperties7 = new A.BodyProperties(){ Vertical = A.TextVerticalValues.Horizontal, LeftInset = 91440, TopInset = 45720, RightInset = 91440, BottomInset = 45720, RightToLeftColumns = false };
            A.NormalAutoFit normalAutoFit2 = new A.NormalAutoFit();

            bodyProperties7.Append(normalAutoFit2);
            A.ListStyle listStyle7 = new A.ListStyle();

            A.Paragraph paragraph7 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties1 = new A.ParagraphProperties(){ Level = 0 };

            A.Run run4 = new A.Run();
            A.RunProperties runProperties6 = new A.RunProperties(){ Language = "en-US" };
            A.Text text6 = new A.Text();
            text6.Text = "Click to edit Master text styles";

            run4.Append(runProperties6);
            run4.Append(text6);

            paragraph7.Append(paragraphProperties1);
            paragraph7.Append(run4);

            A.Paragraph paragraph8 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties2 = new A.ParagraphProperties(){ Level = 1 };

            A.Run run5 = new A.Run();
            A.RunProperties runProperties7 = new A.RunProperties(){ Language = "en-US" };
            A.Text text7 = new A.Text();
            text7.Text = "Second level";

            run5.Append(runProperties7);
            run5.Append(text7);

            paragraph8.Append(paragraphProperties2);
            paragraph8.Append(run5);

            A.Paragraph paragraph9 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties3 = new A.ParagraphProperties(){ Level = 2 };

            A.Run run6 = new A.Run();
            A.RunProperties runProperties8 = new A.RunProperties(){ Language = "en-US" };
            A.Text text8 = new A.Text();
            text8.Text = "Third level";

            run6.Append(runProperties8);
            run6.Append(text8);

            paragraph9.Append(paragraphProperties3);
            paragraph9.Append(run6);

            A.Paragraph paragraph10 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties4 = new A.ParagraphProperties(){ Level = 3 };

            A.Run run7 = new A.Run();
            A.RunProperties runProperties9 = new A.RunProperties(){ Language = "en-US" };
            A.Text text9 = new A.Text();
            text9.Text = "Fourth level";

            run7.Append(runProperties9);
            run7.Append(text9);

            paragraph10.Append(paragraphProperties4);
            paragraph10.Append(run7);

            A.Paragraph paragraph11 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties5 = new A.ParagraphProperties(){ Level = 4 };

            A.Run run8 = new A.Run();
            A.RunProperties runProperties10 = new A.RunProperties(){ Language = "en-US" };
            A.Text text10 = new A.Text();
            text10.Text = "Fifth level";

            run8.Append(runProperties10);
            run8.Append(text10);
            A.EndParagraphRunProperties endParagraphRunProperties7 = new A.EndParagraphRunProperties(){ Language = "pl-PL" };

            paragraph11.Append(paragraphProperties5);
            paragraph11.Append(run8);
            paragraph11.Append(endParagraphRunProperties7);

            textBody7.Append(bodyProperties7);
            textBody7.Append(listStyle7);
            textBody7.Append(paragraph7);
            textBody7.Append(paragraph8);
            textBody7.Append(paragraph9);
            textBody7.Append(paragraph10);
            textBody7.Append(paragraph11);

            shape7.Append(nonVisualShapeProperties7);
            shape7.Append(shapeProperties7);
            shape7.Append(textBody7);

            Shape shape8 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties8 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties11 = new NonVisualDrawingProperties(){ Id = (UInt32Value)4U, Name = "Date Placeholder 3" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList8 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension8 = new A.NonVisualDrawingPropertiesExtension(){ Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            OpenXmlUnknownElement openXmlUnknownElement8 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{A8DB335C-48AC-6D23-FB30-15C4C4E8BA4A}\" />");

            nonVisualDrawingPropertiesExtension8.Append(openXmlUnknownElement8);

            nonVisualDrawingPropertiesExtensionList8.Append(nonVisualDrawingPropertiesExtension8);

            nonVisualDrawingProperties11.Append(nonVisualDrawingPropertiesExtensionList8);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties8 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks8 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties8.Append(shapeLocks8);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties11 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape8 = new PlaceholderShape(){ Type = PlaceholderValues.DateAndTime, Size = PlaceholderSizeValues.Half, Index = (UInt32Value)2U };

            applicationNonVisualDrawingProperties11.Append(placeholderShape8);

            nonVisualShapeProperties8.Append(nonVisualDrawingProperties11);
            nonVisualShapeProperties8.Append(nonVisualShapeDrawingProperties8);
            nonVisualShapeProperties8.Append(applicationNonVisualDrawingProperties11);

            ShapeProperties shapeProperties8 = new ShapeProperties();

            A.Transform2D transform2D5 = new A.Transform2D();
            A.Offset offset8 = new A.Offset(){ X = 838200L, Y = 6356350L };
            A.Extents extents8 = new A.Extents(){ Cx = 2743200L, Cy = 365125L };

            transform2D5.Append(offset8);
            transform2D5.Append(extents8);

            A.PresetGeometry presetGeometry3 = new A.PresetGeometry(){ Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList3 = new A.AdjustValueList();

            presetGeometry3.Append(adjustValueList3);

            shapeProperties8.Append(transform2D5);
            shapeProperties8.Append(presetGeometry3);

            TextBody textBody8 = new TextBody();
            A.BodyProperties bodyProperties8 = new A.BodyProperties(){ Vertical = A.TextVerticalValues.Horizontal, LeftInset = 91440, TopInset = 45720, RightInset = 91440, BottomInset = 45720, RightToLeftColumns = false, Anchor = A.TextAnchoringTypeValues.Center };

            A.ListStyle listStyle8 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties4 = new A.Level1ParagraphProperties(){ Alignment = A.TextAlignmentTypeValues.Left };

            A.DefaultRunProperties defaultRunProperties21 = new A.DefaultRunProperties(){ FontSize = 1200 };

            A.SolidFill solidFill10 = new A.SolidFill();

            A.SchemeColor schemeColor11 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };
            A.Tint tint1 = new A.Tint(){ Val = 82000 };

            schemeColor11.Append(tint1);

            solidFill10.Append(schemeColor11);

            defaultRunProperties21.Append(solidFill10);

            level1ParagraphProperties4.Append(defaultRunProperties21);

            listStyle8.Append(level1ParagraphProperties4);

            A.Paragraph paragraph12 = new A.Paragraph();

            A.Field field3 = new A.Field(){ Id = "{F550F806-175B-4EE1-ABEA-A5F04E4FD72B}", Type = "datetimeFigureOut" };

            A.RunProperties runProperties11 = new A.RunProperties(){ Language = "pl-PL" };
            runProperties11.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text11 = new A.Text();
            text11.Text = "20.08.2025";

            field3.Append(runProperties11);
            field3.Append(text11);
            A.EndParagraphRunProperties endParagraphRunProperties8 = new A.EndParagraphRunProperties(){ Language = "pl-PL" };

            paragraph12.Append(field3);
            paragraph12.Append(endParagraphRunProperties8);

            textBody8.Append(bodyProperties8);
            textBody8.Append(listStyle8);
            textBody8.Append(paragraph12);

            shape8.Append(nonVisualShapeProperties8);
            shape8.Append(shapeProperties8);
            shape8.Append(textBody8);

            Shape shape9 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties9 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties12 = new NonVisualDrawingProperties(){ Id = (UInt32Value)5U, Name = "Footer Placeholder 4" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList9 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension9 = new A.NonVisualDrawingPropertiesExtension(){ Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            OpenXmlUnknownElement openXmlUnknownElement9 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{0C56C7F4-DD84-1E1B-905A-BBAA0C9C1594}\" />");

            nonVisualDrawingPropertiesExtension9.Append(openXmlUnknownElement9);

            nonVisualDrawingPropertiesExtensionList9.Append(nonVisualDrawingPropertiesExtension9);

            nonVisualDrawingProperties12.Append(nonVisualDrawingPropertiesExtensionList9);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties9 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks9 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties9.Append(shapeLocks9);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties12 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape9 = new PlaceholderShape(){ Type = PlaceholderValues.Footer, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)3U };

            applicationNonVisualDrawingProperties12.Append(placeholderShape9);

            nonVisualShapeProperties9.Append(nonVisualDrawingProperties12);
            nonVisualShapeProperties9.Append(nonVisualShapeDrawingProperties9);
            nonVisualShapeProperties9.Append(applicationNonVisualDrawingProperties12);

            ShapeProperties shapeProperties9 = new ShapeProperties();

            A.Transform2D transform2D6 = new A.Transform2D();
            A.Offset offset9 = new A.Offset(){ X = 4038600L, Y = 6356350L };
            A.Extents extents9 = new A.Extents(){ Cx = 4114800L, Cy = 365125L };

            transform2D6.Append(offset9);
            transform2D6.Append(extents9);

            A.PresetGeometry presetGeometry4 = new A.PresetGeometry(){ Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList4 = new A.AdjustValueList();

            presetGeometry4.Append(adjustValueList4);

            shapeProperties9.Append(transform2D6);
            shapeProperties9.Append(presetGeometry4);

            TextBody textBody9 = new TextBody();
            A.BodyProperties bodyProperties9 = new A.BodyProperties(){ Vertical = A.TextVerticalValues.Horizontal, LeftInset = 91440, TopInset = 45720, RightInset = 91440, BottomInset = 45720, RightToLeftColumns = false, Anchor = A.TextAnchoringTypeValues.Center };

            A.ListStyle listStyle9 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties5 = new A.Level1ParagraphProperties(){ Alignment = A.TextAlignmentTypeValues.Center };

            A.DefaultRunProperties defaultRunProperties22 = new A.DefaultRunProperties(){ FontSize = 1200 };

            A.SolidFill solidFill11 = new A.SolidFill();

            A.SchemeColor schemeColor12 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };
            A.Tint tint2 = new A.Tint(){ Val = 82000 };

            schemeColor12.Append(tint2);

            solidFill11.Append(schemeColor12);

            defaultRunProperties22.Append(solidFill11);

            level1ParagraphProperties5.Append(defaultRunProperties22);

            listStyle9.Append(level1ParagraphProperties5);

            A.Paragraph paragraph13 = new A.Paragraph();
            A.EndParagraphRunProperties endParagraphRunProperties9 = new A.EndParagraphRunProperties(){ Language = "pl-PL" };

            paragraph13.Append(endParagraphRunProperties9);

            textBody9.Append(bodyProperties9);
            textBody9.Append(listStyle9);
            textBody9.Append(paragraph13);

            shape9.Append(nonVisualShapeProperties9);
            shape9.Append(shapeProperties9);
            shape9.Append(textBody9);

            Shape shape10 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties10 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties13 = new NonVisualDrawingProperties(){ Id = (UInt32Value)6U, Name = "Slide Number Placeholder 5" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList10 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension10 = new A.NonVisualDrawingPropertiesExtension(){ Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            OpenXmlUnknownElement openXmlUnknownElement10 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{1302F1BB-FA87-A34F-2419-C6ACD0D29928}\" />");

            nonVisualDrawingPropertiesExtension10.Append(openXmlUnknownElement10);

            nonVisualDrawingPropertiesExtensionList10.Append(nonVisualDrawingPropertiesExtension10);

            nonVisualDrawingProperties13.Append(nonVisualDrawingPropertiesExtensionList10);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties10 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks10 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties10.Append(shapeLocks10);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties13 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape10 = new PlaceholderShape(){ Type = PlaceholderValues.SlideNumber, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)4U };

            applicationNonVisualDrawingProperties13.Append(placeholderShape10);

            nonVisualShapeProperties10.Append(nonVisualDrawingProperties13);
            nonVisualShapeProperties10.Append(nonVisualShapeDrawingProperties10);
            nonVisualShapeProperties10.Append(applicationNonVisualDrawingProperties13);

            ShapeProperties shapeProperties10 = new ShapeProperties();

            A.Transform2D transform2D7 = new A.Transform2D();
            A.Offset offset10 = new A.Offset(){ X = 8610600L, Y = 6356350L };
            A.Extents extents10 = new A.Extents(){ Cx = 2743200L, Cy = 365125L };

            transform2D7.Append(offset10);
            transform2D7.Append(extents10);

            A.PresetGeometry presetGeometry5 = new A.PresetGeometry(){ Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList5 = new A.AdjustValueList();

            presetGeometry5.Append(adjustValueList5);

            shapeProperties10.Append(transform2D7);
            shapeProperties10.Append(presetGeometry5);

            TextBody textBody10 = new TextBody();
            A.BodyProperties bodyProperties10 = new A.BodyProperties(){ Vertical = A.TextVerticalValues.Horizontal, LeftInset = 91440, TopInset = 45720, RightInset = 91440, BottomInset = 45720, RightToLeftColumns = false, Anchor = A.TextAnchoringTypeValues.Center };

            A.ListStyle listStyle10 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties6 = new A.Level1ParagraphProperties(){ Alignment = A.TextAlignmentTypeValues.Right };

            A.DefaultRunProperties defaultRunProperties23 = new A.DefaultRunProperties(){ FontSize = 1200 };

            A.SolidFill solidFill12 = new A.SolidFill();

            A.SchemeColor schemeColor13 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };
            A.Tint tint3 = new A.Tint(){ Val = 82000 };

            schemeColor13.Append(tint3);

            solidFill12.Append(schemeColor13);

            defaultRunProperties23.Append(solidFill12);

            level1ParagraphProperties6.Append(defaultRunProperties23);

            listStyle10.Append(level1ParagraphProperties6);

            A.Paragraph paragraph14 = new A.Paragraph();

            A.Field field4 = new A.Field(){ Id = "{5FAB34AF-B09A-4E2A-BAD3-CCD6E0AD81E6}", Type = "slidenum" };

            A.RunProperties runProperties12 = new A.RunProperties(){ Language = "pl-PL" };
            runProperties12.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text12 = new A.Text();
            text12.Text = "‹#›";

            field4.Append(runProperties12);
            field4.Append(text12);
            A.EndParagraphRunProperties endParagraphRunProperties10 = new A.EndParagraphRunProperties(){ Language = "pl-PL" };

            paragraph14.Append(field4);
            paragraph14.Append(endParagraphRunProperties10);

            textBody10.Append(bodyProperties10);
            textBody10.Append(listStyle10);
            textBody10.Append(paragraph14);

            shape10.Append(nonVisualShapeProperties10);
            shape10.Append(shapeProperties10);
            shape10.Append(textBody10);

            shapeTree3.Append(nonVisualGroupShapeProperties3);
            shapeTree3.Append(groupShapeProperties3);
            shapeTree3.Append(shape6);
            shapeTree3.Append(shape7);
            shapeTree3.Append(shape8);
            shapeTree3.Append(shape9);
            shapeTree3.Append(shape10);

            CommonSlideDataExtensionList commonSlideDataExtensionList3 = new CommonSlideDataExtensionList();

            CommonSlideDataExtension commonSlideDataExtension3 = new CommonSlideDataExtension(){ Uri = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}" };

            P14.CreationId creationId3 = new P14.CreationId(){ Val = (UInt32Value)1568487410U };
            creationId3.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            commonSlideDataExtension3.Append(creationId3);

            commonSlideDataExtensionList3.Append(commonSlideDataExtension3);

            commonSlideData3.Append(background1);
            commonSlideData3.Append(shapeTree3);
            commonSlideData3.Append(commonSlideDataExtensionList3);
            ColorMap colorMap1 = new ColorMap(){ Background1 = A.ColorSchemeIndexValues.Light1, Text1 = A.ColorSchemeIndexValues.Dark1, Background2 = A.ColorSchemeIndexValues.Light2, Text2 = A.ColorSchemeIndexValues.Dark2, Accent1 = A.ColorSchemeIndexValues.Accent1, Accent2 = A.ColorSchemeIndexValues.Accent2, Accent3 = A.ColorSchemeIndexValues.Accent3, Accent4 = A.ColorSchemeIndexValues.Accent4, Accent5 = A.ColorSchemeIndexValues.Accent5, Accent6 = A.ColorSchemeIndexValues.Accent6, Hyperlink = A.ColorSchemeIndexValues.Hyperlink, FollowedHyperlink = A.ColorSchemeIndexValues.FollowedHyperlink };

            SlideLayoutIdList slideLayoutIdList1 = new SlideLayoutIdList();
            SlideLayoutId slideLayoutId1 = new SlideLayoutId(){ Id = (UInt32Value)2147483649U, RelationshipId = "rId1" };
            SlideLayoutId slideLayoutId2 = new SlideLayoutId(){ Id = (UInt32Value)2147483650U, RelationshipId = "rId2" };
            SlideLayoutId slideLayoutId3 = new SlideLayoutId(){ Id = (UInt32Value)2147483651U, RelationshipId = "rId3" };
            SlideLayoutId slideLayoutId4 = new SlideLayoutId(){ Id = (UInt32Value)2147483652U, RelationshipId = "rId4" };
            SlideLayoutId slideLayoutId5 = new SlideLayoutId(){ Id = (UInt32Value)2147483653U, RelationshipId = "rId5" };
            SlideLayoutId slideLayoutId6 = new SlideLayoutId(){ Id = (UInt32Value)2147483654U, RelationshipId = "rId6" };
            SlideLayoutId slideLayoutId7 = new SlideLayoutId(){ Id = (UInt32Value)2147483655U, RelationshipId = "rId7" };
            SlideLayoutId slideLayoutId8 = new SlideLayoutId(){ Id = (UInt32Value)2147483656U, RelationshipId = "rId8" };
            SlideLayoutId slideLayoutId9 = new SlideLayoutId(){ Id = (UInt32Value)2147483657U, RelationshipId = "rId9" };
            SlideLayoutId slideLayoutId10 = new SlideLayoutId(){ Id = (UInt32Value)2147483658U, RelationshipId = "rId10" };
            SlideLayoutId slideLayoutId11 = new SlideLayoutId(){ Id = (UInt32Value)2147483659U, RelationshipId = "rId11" };

            slideLayoutIdList1.Append(slideLayoutId1);
            slideLayoutIdList1.Append(slideLayoutId2);
            slideLayoutIdList1.Append(slideLayoutId3);
            slideLayoutIdList1.Append(slideLayoutId4);
            slideLayoutIdList1.Append(slideLayoutId5);
            slideLayoutIdList1.Append(slideLayoutId6);
            slideLayoutIdList1.Append(slideLayoutId7);
            slideLayoutIdList1.Append(slideLayoutId8);
            slideLayoutIdList1.Append(slideLayoutId9);
            slideLayoutIdList1.Append(slideLayoutId10);
            slideLayoutIdList1.Append(slideLayoutId11);

            TextStyles textStyles1 = new TextStyles();

            TitleStyle titleStyle1 = new TitleStyle();

            A.Level1ParagraphProperties level1ParagraphProperties7 = new A.Level1ParagraphProperties(){ Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.LineSpacing lineSpacing1 = new A.LineSpacing();
            A.SpacingPercent spacingPercent1 = new A.SpacingPercent(){ Val = 90000 };

            lineSpacing1.Append(spacingPercent1);

            A.SpaceBefore spaceBefore1 = new A.SpaceBefore();
            A.SpacingPercent spacingPercent2 = new A.SpacingPercent(){ Val = 0 };

            spaceBefore1.Append(spacingPercent2);
            A.NoBullet noBullet10 = new A.NoBullet();

            A.DefaultRunProperties defaultRunProperties24 = new A.DefaultRunProperties(){ FontSize = 4400, Kerning = 1200 };

            A.SolidFill solidFill13 = new A.SolidFill();
            A.SchemeColor schemeColor14 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            solidFill13.Append(schemeColor14);
            A.LatinFont latinFont10 = new A.LatinFont(){ Typeface = "+mj-lt" };
            A.EastAsianFont eastAsianFont10 = new A.EastAsianFont(){ Typeface = "+mj-ea" };
            A.ComplexScriptFont complexScriptFont10 = new A.ComplexScriptFont(){ Typeface = "+mj-cs" };

            defaultRunProperties24.Append(solidFill13);
            defaultRunProperties24.Append(latinFont10);
            defaultRunProperties24.Append(eastAsianFont10);
            defaultRunProperties24.Append(complexScriptFont10);

            level1ParagraphProperties7.Append(lineSpacing1);
            level1ParagraphProperties7.Append(spaceBefore1);
            level1ParagraphProperties7.Append(noBullet10);
            level1ParagraphProperties7.Append(defaultRunProperties24);

            titleStyle1.Append(level1ParagraphProperties7);

            BodyStyle bodyStyle1 = new BodyStyle();

            A.Level1ParagraphProperties level1ParagraphProperties8 = new A.Level1ParagraphProperties(){ LeftMargin = 228600, Indent = -228600, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.LineSpacing lineSpacing2 = new A.LineSpacing();
            A.SpacingPercent spacingPercent3 = new A.SpacingPercent(){ Val = 90000 };

            lineSpacing2.Append(spacingPercent3);

            A.SpaceBefore spaceBefore2 = new A.SpaceBefore();
            A.SpacingPoints spacingPoints1 = new A.SpacingPoints(){ Val = 1000 };

            spaceBefore2.Append(spacingPoints1);
            A.BulletFont bulletFont1 = new A.BulletFont(){ Typeface = "Arial", Panose = "020B0604020202020204", PitchFamily = 34, CharacterSet = 0 };
            A.CharacterBullet characterBullet1 = new A.CharacterBullet(){ Char = "•" };

            A.DefaultRunProperties defaultRunProperties25 = new A.DefaultRunProperties(){ FontSize = 2800, Kerning = 1200 };

            A.SolidFill solidFill14 = new A.SolidFill();
            A.SchemeColor schemeColor15 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            solidFill14.Append(schemeColor15);
            A.LatinFont latinFont11 = new A.LatinFont(){ Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont11 = new A.EastAsianFont(){ Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont11 = new A.ComplexScriptFont(){ Typeface = "+mn-cs" };

            defaultRunProperties25.Append(solidFill14);
            defaultRunProperties25.Append(latinFont11);
            defaultRunProperties25.Append(eastAsianFont11);
            defaultRunProperties25.Append(complexScriptFont11);

            level1ParagraphProperties8.Append(lineSpacing2);
            level1ParagraphProperties8.Append(spaceBefore2);
            level1ParagraphProperties8.Append(bulletFont1);
            level1ParagraphProperties8.Append(characterBullet1);
            level1ParagraphProperties8.Append(defaultRunProperties25);

            A.Level2ParagraphProperties level2ParagraphProperties3 = new A.Level2ParagraphProperties(){ LeftMargin = 685800, Indent = -228600, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.LineSpacing lineSpacing3 = new A.LineSpacing();
            A.SpacingPercent spacingPercent4 = new A.SpacingPercent(){ Val = 90000 };

            lineSpacing3.Append(spacingPercent4);

            A.SpaceBefore spaceBefore3 = new A.SpaceBefore();
            A.SpacingPoints spacingPoints2 = new A.SpacingPoints(){ Val = 500 };

            spaceBefore3.Append(spacingPoints2);
            A.BulletFont bulletFont2 = new A.BulletFont(){ Typeface = "Arial", Panose = "020B0604020202020204", PitchFamily = 34, CharacterSet = 0 };
            A.CharacterBullet characterBullet2 = new A.CharacterBullet(){ Char = "•" };

            A.DefaultRunProperties defaultRunProperties26 = new A.DefaultRunProperties(){ FontSize = 2400, Kerning = 1200 };

            A.SolidFill solidFill15 = new A.SolidFill();
            A.SchemeColor schemeColor16 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            solidFill15.Append(schemeColor16);
            A.LatinFont latinFont12 = new A.LatinFont(){ Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont12 = new A.EastAsianFont(){ Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont12 = new A.ComplexScriptFont(){ Typeface = "+mn-cs" };

            defaultRunProperties26.Append(solidFill15);
            defaultRunProperties26.Append(latinFont12);
            defaultRunProperties26.Append(eastAsianFont12);
            defaultRunProperties26.Append(complexScriptFont12);

            level2ParagraphProperties3.Append(lineSpacing3);
            level2ParagraphProperties3.Append(spaceBefore3);
            level2ParagraphProperties3.Append(bulletFont2);
            level2ParagraphProperties3.Append(characterBullet2);
            level2ParagraphProperties3.Append(defaultRunProperties26);

            A.Level3ParagraphProperties level3ParagraphProperties3 = new A.Level3ParagraphProperties(){ LeftMargin = 1143000, Indent = -228600, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.LineSpacing lineSpacing4 = new A.LineSpacing();
            A.SpacingPercent spacingPercent5 = new A.SpacingPercent(){ Val = 90000 };

            lineSpacing4.Append(spacingPercent5);

            A.SpaceBefore spaceBefore4 = new A.SpaceBefore();
            A.SpacingPoints spacingPoints3 = new A.SpacingPoints(){ Val = 500 };

            spaceBefore4.Append(spacingPoints3);
            A.BulletFont bulletFont3 = new A.BulletFont(){ Typeface = "Arial", Panose = "020B0604020202020204", PitchFamily = 34, CharacterSet = 0 };
            A.CharacterBullet characterBullet3 = new A.CharacterBullet(){ Char = "•" };

            A.DefaultRunProperties defaultRunProperties27 = new A.DefaultRunProperties(){ FontSize = 2000, Kerning = 1200 };

            A.SolidFill solidFill16 = new A.SolidFill();
            A.SchemeColor schemeColor17 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            solidFill16.Append(schemeColor17);
            A.LatinFont latinFont13 = new A.LatinFont(){ Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont13 = new A.EastAsianFont(){ Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont13 = new A.ComplexScriptFont(){ Typeface = "+mn-cs" };

            defaultRunProperties27.Append(solidFill16);
            defaultRunProperties27.Append(latinFont13);
            defaultRunProperties27.Append(eastAsianFont13);
            defaultRunProperties27.Append(complexScriptFont13);

            level3ParagraphProperties3.Append(lineSpacing4);
            level3ParagraphProperties3.Append(spaceBefore4);
            level3ParagraphProperties3.Append(bulletFont3);
            level3ParagraphProperties3.Append(characterBullet3);
            level3ParagraphProperties3.Append(defaultRunProperties27);

            A.Level4ParagraphProperties level4ParagraphProperties3 = new A.Level4ParagraphProperties(){ LeftMargin = 1600200, Indent = -228600, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.LineSpacing lineSpacing5 = new A.LineSpacing();
            A.SpacingPercent spacingPercent6 = new A.SpacingPercent(){ Val = 90000 };

            lineSpacing5.Append(spacingPercent6);

            A.SpaceBefore spaceBefore5 = new A.SpaceBefore();
            A.SpacingPoints spacingPoints4 = new A.SpacingPoints(){ Val = 500 };

            spaceBefore5.Append(spacingPoints4);
            A.BulletFont bulletFont4 = new A.BulletFont(){ Typeface = "Arial", Panose = "020B0604020202020204", PitchFamily = 34, CharacterSet = 0 };
            A.CharacterBullet characterBullet4 = new A.CharacterBullet(){ Char = "•" };

            A.DefaultRunProperties defaultRunProperties28 = new A.DefaultRunProperties(){ FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill17 = new A.SolidFill();
            A.SchemeColor schemeColor18 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            solidFill17.Append(schemeColor18);
            A.LatinFont latinFont14 = new A.LatinFont(){ Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont14 = new A.EastAsianFont(){ Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont14 = new A.ComplexScriptFont(){ Typeface = "+mn-cs" };

            defaultRunProperties28.Append(solidFill17);
            defaultRunProperties28.Append(latinFont14);
            defaultRunProperties28.Append(eastAsianFont14);
            defaultRunProperties28.Append(complexScriptFont14);

            level4ParagraphProperties3.Append(lineSpacing5);
            level4ParagraphProperties3.Append(spaceBefore5);
            level4ParagraphProperties3.Append(bulletFont4);
            level4ParagraphProperties3.Append(characterBullet4);
            level4ParagraphProperties3.Append(defaultRunProperties28);

            A.Level5ParagraphProperties level5ParagraphProperties3 = new A.Level5ParagraphProperties(){ LeftMargin = 2057400, Indent = -228600, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.LineSpacing lineSpacing6 = new A.LineSpacing();
            A.SpacingPercent spacingPercent7 = new A.SpacingPercent(){ Val = 90000 };

            lineSpacing6.Append(spacingPercent7);

            A.SpaceBefore spaceBefore6 = new A.SpaceBefore();
            A.SpacingPoints spacingPoints5 = new A.SpacingPoints(){ Val = 500 };

            spaceBefore6.Append(spacingPoints5);
            A.BulletFont bulletFont5 = new A.BulletFont(){ Typeface = "Arial", Panose = "020B0604020202020204", PitchFamily = 34, CharacterSet = 0 };
            A.CharacterBullet characterBullet5 = new A.CharacterBullet(){ Char = "•" };

            A.DefaultRunProperties defaultRunProperties29 = new A.DefaultRunProperties(){ FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill18 = new A.SolidFill();
            A.SchemeColor schemeColor19 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            solidFill18.Append(schemeColor19);
            A.LatinFont latinFont15 = new A.LatinFont(){ Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont15 = new A.EastAsianFont(){ Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont15 = new A.ComplexScriptFont(){ Typeface = "+mn-cs" };

            defaultRunProperties29.Append(solidFill18);
            defaultRunProperties29.Append(latinFont15);
            defaultRunProperties29.Append(eastAsianFont15);
            defaultRunProperties29.Append(complexScriptFont15);

            level5ParagraphProperties3.Append(lineSpacing6);
            level5ParagraphProperties3.Append(spaceBefore6);
            level5ParagraphProperties3.Append(bulletFont5);
            level5ParagraphProperties3.Append(characterBullet5);
            level5ParagraphProperties3.Append(defaultRunProperties29);

            A.Level6ParagraphProperties level6ParagraphProperties3 = new A.Level6ParagraphProperties(){ LeftMargin = 2514600, Indent = -228600, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.LineSpacing lineSpacing7 = new A.LineSpacing();
            A.SpacingPercent spacingPercent8 = new A.SpacingPercent(){ Val = 90000 };

            lineSpacing7.Append(spacingPercent8);

            A.SpaceBefore spaceBefore7 = new A.SpaceBefore();
            A.SpacingPoints spacingPoints6 = new A.SpacingPoints(){ Val = 500 };

            spaceBefore7.Append(spacingPoints6);
            A.BulletFont bulletFont6 = new A.BulletFont(){ Typeface = "Arial", Panose = "020B0604020202020204", PitchFamily = 34, CharacterSet = 0 };
            A.CharacterBullet characterBullet6 = new A.CharacterBullet(){ Char = "•" };

            A.DefaultRunProperties defaultRunProperties30 = new A.DefaultRunProperties(){ FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill19 = new A.SolidFill();
            A.SchemeColor schemeColor20 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            solidFill19.Append(schemeColor20);
            A.LatinFont latinFont16 = new A.LatinFont(){ Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont16 = new A.EastAsianFont(){ Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont16 = new A.ComplexScriptFont(){ Typeface = "+mn-cs" };

            defaultRunProperties30.Append(solidFill19);
            defaultRunProperties30.Append(latinFont16);
            defaultRunProperties30.Append(eastAsianFont16);
            defaultRunProperties30.Append(complexScriptFont16);

            level6ParagraphProperties3.Append(lineSpacing7);
            level6ParagraphProperties3.Append(spaceBefore7);
            level6ParagraphProperties3.Append(bulletFont6);
            level6ParagraphProperties3.Append(characterBullet6);
            level6ParagraphProperties3.Append(defaultRunProperties30);

            A.Level7ParagraphProperties level7ParagraphProperties3 = new A.Level7ParagraphProperties(){ LeftMargin = 2971800, Indent = -228600, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.LineSpacing lineSpacing8 = new A.LineSpacing();
            A.SpacingPercent spacingPercent9 = new A.SpacingPercent(){ Val = 90000 };

            lineSpacing8.Append(spacingPercent9);

            A.SpaceBefore spaceBefore8 = new A.SpaceBefore();
            A.SpacingPoints spacingPoints7 = new A.SpacingPoints(){ Val = 500 };

            spaceBefore8.Append(spacingPoints7);
            A.BulletFont bulletFont7 = new A.BulletFont(){ Typeface = "Arial", Panose = "020B0604020202020204", PitchFamily = 34, CharacterSet = 0 };
            A.CharacterBullet characterBullet7 = new A.CharacterBullet(){ Char = "•" };

            A.DefaultRunProperties defaultRunProperties31 = new A.DefaultRunProperties(){ FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill20 = new A.SolidFill();
            A.SchemeColor schemeColor21 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            solidFill20.Append(schemeColor21);
            A.LatinFont latinFont17 = new A.LatinFont(){ Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont17 = new A.EastAsianFont(){ Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont17 = new A.ComplexScriptFont(){ Typeface = "+mn-cs" };

            defaultRunProperties31.Append(solidFill20);
            defaultRunProperties31.Append(latinFont17);
            defaultRunProperties31.Append(eastAsianFont17);
            defaultRunProperties31.Append(complexScriptFont17);

            level7ParagraphProperties3.Append(lineSpacing8);
            level7ParagraphProperties3.Append(spaceBefore8);
            level7ParagraphProperties3.Append(bulletFont7);
            level7ParagraphProperties3.Append(characterBullet7);
            level7ParagraphProperties3.Append(defaultRunProperties31);

            A.Level8ParagraphProperties level8ParagraphProperties3 = new A.Level8ParagraphProperties(){ LeftMargin = 3429000, Indent = -228600, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.LineSpacing lineSpacing9 = new A.LineSpacing();
            A.SpacingPercent spacingPercent10 = new A.SpacingPercent(){ Val = 90000 };

            lineSpacing9.Append(spacingPercent10);

            A.SpaceBefore spaceBefore9 = new A.SpaceBefore();
            A.SpacingPoints spacingPoints8 = new A.SpacingPoints(){ Val = 500 };

            spaceBefore9.Append(spacingPoints8);
            A.BulletFont bulletFont8 = new A.BulletFont(){ Typeface = "Arial", Panose = "020B0604020202020204", PitchFamily = 34, CharacterSet = 0 };
            A.CharacterBullet characterBullet8 = new A.CharacterBullet(){ Char = "•" };

            A.DefaultRunProperties defaultRunProperties32 = new A.DefaultRunProperties(){ FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill21 = new A.SolidFill();
            A.SchemeColor schemeColor22 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            solidFill21.Append(schemeColor22);
            A.LatinFont latinFont18 = new A.LatinFont(){ Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont18 = new A.EastAsianFont(){ Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont18 = new A.ComplexScriptFont(){ Typeface = "+mn-cs" };

            defaultRunProperties32.Append(solidFill21);
            defaultRunProperties32.Append(latinFont18);
            defaultRunProperties32.Append(eastAsianFont18);
            defaultRunProperties32.Append(complexScriptFont18);

            level8ParagraphProperties3.Append(lineSpacing9);
            level8ParagraphProperties3.Append(spaceBefore9);
            level8ParagraphProperties3.Append(bulletFont8);
            level8ParagraphProperties3.Append(characterBullet8);
            level8ParagraphProperties3.Append(defaultRunProperties32);

            A.Level9ParagraphProperties level9ParagraphProperties3 = new A.Level9ParagraphProperties(){ LeftMargin = 3886200, Indent = -228600, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.LineSpacing lineSpacing10 = new A.LineSpacing();
            A.SpacingPercent spacingPercent11 = new A.SpacingPercent(){ Val = 90000 };

            lineSpacing10.Append(spacingPercent11);

            A.SpaceBefore spaceBefore10 = new A.SpaceBefore();
            A.SpacingPoints spacingPoints9 = new A.SpacingPoints(){ Val = 500 };

            spaceBefore10.Append(spacingPoints9);
            A.BulletFont bulletFont9 = new A.BulletFont(){ Typeface = "Arial", Panose = "020B0604020202020204", PitchFamily = 34, CharacterSet = 0 };
            A.CharacterBullet characterBullet9 = new A.CharacterBullet(){ Char = "•" };

            A.DefaultRunProperties defaultRunProperties33 = new A.DefaultRunProperties(){ FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill22 = new A.SolidFill();
            A.SchemeColor schemeColor23 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            solidFill22.Append(schemeColor23);
            A.LatinFont latinFont19 = new A.LatinFont(){ Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont19 = new A.EastAsianFont(){ Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont19 = new A.ComplexScriptFont(){ Typeface = "+mn-cs" };

            defaultRunProperties33.Append(solidFill22);
            defaultRunProperties33.Append(latinFont19);
            defaultRunProperties33.Append(eastAsianFont19);
            defaultRunProperties33.Append(complexScriptFont19);

            level9ParagraphProperties3.Append(lineSpacing10);
            level9ParagraphProperties3.Append(spaceBefore10);
            level9ParagraphProperties3.Append(bulletFont9);
            level9ParagraphProperties3.Append(characterBullet9);
            level9ParagraphProperties3.Append(defaultRunProperties33);

            bodyStyle1.Append(level1ParagraphProperties8);
            bodyStyle1.Append(level2ParagraphProperties3);
            bodyStyle1.Append(level3ParagraphProperties3);
            bodyStyle1.Append(level4ParagraphProperties3);
            bodyStyle1.Append(level5ParagraphProperties3);
            bodyStyle1.Append(level6ParagraphProperties3);
            bodyStyle1.Append(level7ParagraphProperties3);
            bodyStyle1.Append(level8ParagraphProperties3);
            bodyStyle1.Append(level9ParagraphProperties3);

            OtherStyle otherStyle1 = new OtherStyle();

            A.DefaultParagraphProperties defaultParagraphProperties2 = new A.DefaultParagraphProperties();
            A.DefaultRunProperties defaultRunProperties34 = new A.DefaultRunProperties(){ Language = "pl-PL" };

            defaultParagraphProperties2.Append(defaultRunProperties34);

            A.Level1ParagraphProperties level1ParagraphProperties9 = new A.Level1ParagraphProperties(){ LeftMargin = 0, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties35 = new A.DefaultRunProperties(){ FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill23 = new A.SolidFill();
            A.SchemeColor schemeColor24 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            solidFill23.Append(schemeColor24);
            A.LatinFont latinFont20 = new A.LatinFont(){ Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont20 = new A.EastAsianFont(){ Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont20 = new A.ComplexScriptFont(){ Typeface = "+mn-cs" };

            defaultRunProperties35.Append(solidFill23);
            defaultRunProperties35.Append(latinFont20);
            defaultRunProperties35.Append(eastAsianFont20);
            defaultRunProperties35.Append(complexScriptFont20);

            level1ParagraphProperties9.Append(defaultRunProperties35);

            A.Level2ParagraphProperties level2ParagraphProperties4 = new A.Level2ParagraphProperties(){ LeftMargin = 457200, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties36 = new A.DefaultRunProperties(){ FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill24 = new A.SolidFill();
            A.SchemeColor schemeColor25 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            solidFill24.Append(schemeColor25);
            A.LatinFont latinFont21 = new A.LatinFont(){ Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont21 = new A.EastAsianFont(){ Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont21 = new A.ComplexScriptFont(){ Typeface = "+mn-cs" };

            defaultRunProperties36.Append(solidFill24);
            defaultRunProperties36.Append(latinFont21);
            defaultRunProperties36.Append(eastAsianFont21);
            defaultRunProperties36.Append(complexScriptFont21);

            level2ParagraphProperties4.Append(defaultRunProperties36);

            A.Level3ParagraphProperties level3ParagraphProperties4 = new A.Level3ParagraphProperties(){ LeftMargin = 914400, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties37 = new A.DefaultRunProperties(){ FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill25 = new A.SolidFill();
            A.SchemeColor schemeColor26 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            solidFill25.Append(schemeColor26);
            A.LatinFont latinFont22 = new A.LatinFont(){ Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont22 = new A.EastAsianFont(){ Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont22 = new A.ComplexScriptFont(){ Typeface = "+mn-cs" };

            defaultRunProperties37.Append(solidFill25);
            defaultRunProperties37.Append(latinFont22);
            defaultRunProperties37.Append(eastAsianFont22);
            defaultRunProperties37.Append(complexScriptFont22);

            level3ParagraphProperties4.Append(defaultRunProperties37);

            A.Level4ParagraphProperties level4ParagraphProperties4 = new A.Level4ParagraphProperties(){ LeftMargin = 1371600, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties38 = new A.DefaultRunProperties(){ FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill26 = new A.SolidFill();
            A.SchemeColor schemeColor27 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            solidFill26.Append(schemeColor27);
            A.LatinFont latinFont23 = new A.LatinFont(){ Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont23 = new A.EastAsianFont(){ Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont23 = new A.ComplexScriptFont(){ Typeface = "+mn-cs" };

            defaultRunProperties38.Append(solidFill26);
            defaultRunProperties38.Append(latinFont23);
            defaultRunProperties38.Append(eastAsianFont23);
            defaultRunProperties38.Append(complexScriptFont23);

            level4ParagraphProperties4.Append(defaultRunProperties38);

            A.Level5ParagraphProperties level5ParagraphProperties4 = new A.Level5ParagraphProperties(){ LeftMargin = 1828800, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties39 = new A.DefaultRunProperties(){ FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill27 = new A.SolidFill();
            A.SchemeColor schemeColor28 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            solidFill27.Append(schemeColor28);
            A.LatinFont latinFont24 = new A.LatinFont(){ Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont24 = new A.EastAsianFont(){ Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont24 = new A.ComplexScriptFont(){ Typeface = "+mn-cs" };

            defaultRunProperties39.Append(solidFill27);
            defaultRunProperties39.Append(latinFont24);
            defaultRunProperties39.Append(eastAsianFont24);
            defaultRunProperties39.Append(complexScriptFont24);

            level5ParagraphProperties4.Append(defaultRunProperties39);

            A.Level6ParagraphProperties level6ParagraphProperties4 = new A.Level6ParagraphProperties(){ LeftMargin = 2286000, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties40 = new A.DefaultRunProperties(){ FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill28 = new A.SolidFill();
            A.SchemeColor schemeColor29 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            solidFill28.Append(schemeColor29);
            A.LatinFont latinFont25 = new A.LatinFont(){ Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont25 = new A.EastAsianFont(){ Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont25 = new A.ComplexScriptFont(){ Typeface = "+mn-cs" };

            defaultRunProperties40.Append(solidFill28);
            defaultRunProperties40.Append(latinFont25);
            defaultRunProperties40.Append(eastAsianFont25);
            defaultRunProperties40.Append(complexScriptFont25);

            level6ParagraphProperties4.Append(defaultRunProperties40);

            A.Level7ParagraphProperties level7ParagraphProperties4 = new A.Level7ParagraphProperties(){ LeftMargin = 2743200, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties41 = new A.DefaultRunProperties(){ FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill29 = new A.SolidFill();
            A.SchemeColor schemeColor30 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            solidFill29.Append(schemeColor30);
            A.LatinFont latinFont26 = new A.LatinFont(){ Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont26 = new A.EastAsianFont(){ Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont26 = new A.ComplexScriptFont(){ Typeface = "+mn-cs" };

            defaultRunProperties41.Append(solidFill29);
            defaultRunProperties41.Append(latinFont26);
            defaultRunProperties41.Append(eastAsianFont26);
            defaultRunProperties41.Append(complexScriptFont26);

            level7ParagraphProperties4.Append(defaultRunProperties41);

            A.Level8ParagraphProperties level8ParagraphProperties4 = new A.Level8ParagraphProperties(){ LeftMargin = 3200400, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties42 = new A.DefaultRunProperties(){ FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill30 = new A.SolidFill();
            A.SchemeColor schemeColor31 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            solidFill30.Append(schemeColor31);
            A.LatinFont latinFont27 = new A.LatinFont(){ Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont27 = new A.EastAsianFont(){ Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont27 = new A.ComplexScriptFont(){ Typeface = "+mn-cs" };

            defaultRunProperties42.Append(solidFill30);
            defaultRunProperties42.Append(latinFont27);
            defaultRunProperties42.Append(eastAsianFont27);
            defaultRunProperties42.Append(complexScriptFont27);

            level8ParagraphProperties4.Append(defaultRunProperties42);

            A.Level9ParagraphProperties level9ParagraphProperties4 = new A.Level9ParagraphProperties(){ LeftMargin = 3657600, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties43 = new A.DefaultRunProperties(){ FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill31 = new A.SolidFill();
            A.SchemeColor schemeColor32 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            solidFill31.Append(schemeColor32);
            A.LatinFont latinFont28 = new A.LatinFont(){ Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont28 = new A.EastAsianFont(){ Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont28 = new A.ComplexScriptFont(){ Typeface = "+mn-cs" };

            defaultRunProperties43.Append(solidFill31);
            defaultRunProperties43.Append(latinFont28);
            defaultRunProperties43.Append(eastAsianFont28);
            defaultRunProperties43.Append(complexScriptFont28);

            level9ParagraphProperties4.Append(defaultRunProperties43);

            otherStyle1.Append(defaultParagraphProperties2);
            otherStyle1.Append(level1ParagraphProperties9);
            otherStyle1.Append(level2ParagraphProperties4);
            otherStyle1.Append(level3ParagraphProperties4);
            otherStyle1.Append(level4ParagraphProperties4);
            otherStyle1.Append(level5ParagraphProperties4);
            otherStyle1.Append(level6ParagraphProperties4);
            otherStyle1.Append(level7ParagraphProperties4);
            otherStyle1.Append(level8ParagraphProperties4);
            otherStyle1.Append(level9ParagraphProperties4);

            textStyles1.Append(titleStyle1);
            textStyles1.Append(bodyStyle1);
            textStyles1.Append(otherStyle1);

            slideMaster1.Append(commonSlideData3);
            slideMaster1.Append(colorMap1);
            slideMaster1.Append(slideLayoutIdList1);
            slideMaster1.Append(textStyles1);

            slideMasterPart1.SlideMaster = slideMaster1;
        }

        // Generates content of slideLayoutPart2.
        private void GenerateSlideLayoutPart2Content(SlideLayoutPart slideLayoutPart2)
        {
            SlideLayout slideLayout2 = new SlideLayout(){ Type = SlideLayoutValues.ObjectText, Preserve = true };
            slideLayout2.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            slideLayout2.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            slideLayout2.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            CommonSlideData commonSlideData4 = new CommonSlideData(){ Name = "Content with Caption" };

            ShapeTree shapeTree4 = new ShapeTree();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties4 = new NonVisualGroupShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties14 = new NonVisualDrawingProperties(){ Id = (UInt32Value)1U, Name = "" };
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties4 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties14 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties4.Append(nonVisualDrawingProperties14);
            nonVisualGroupShapeProperties4.Append(nonVisualGroupShapeDrawingProperties4);
            nonVisualGroupShapeProperties4.Append(applicationNonVisualDrawingProperties14);

            GroupShapeProperties groupShapeProperties4 = new GroupShapeProperties();

            A.TransformGroup transformGroup4 = new A.TransformGroup();
            A.Offset offset11 = new A.Offset(){ X = 0L, Y = 0L };
            A.Extents extents11 = new A.Extents(){ Cx = 0L, Cy = 0L };
            A.ChildOffset childOffset4 = new A.ChildOffset(){ X = 0L, Y = 0L };
            A.ChildExtents childExtents4 = new A.ChildExtents(){ Cx = 0L, Cy = 0L };

            transformGroup4.Append(offset11);
            transformGroup4.Append(extents11);
            transformGroup4.Append(childOffset4);
            transformGroup4.Append(childExtents4);

            groupShapeProperties4.Append(transformGroup4);

            Shape shape11 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties11 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties15 = new NonVisualDrawingProperties(){ Id = (UInt32Value)2U, Name = "Title 1" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList11 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension11 = new A.NonVisualDrawingPropertiesExtension(){ Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            OpenXmlUnknownElement openXmlUnknownElement11 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{7A7110A2-DD46-EC0D-EB4A-D04827AFE4E6}\" />");

            nonVisualDrawingPropertiesExtension11.Append(openXmlUnknownElement11);

            nonVisualDrawingPropertiesExtensionList11.Append(nonVisualDrawingPropertiesExtension11);

            nonVisualDrawingProperties15.Append(nonVisualDrawingPropertiesExtensionList11);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties11 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks11 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties11.Append(shapeLocks11);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties15 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape11 = new PlaceholderShape(){ Type = PlaceholderValues.Title };

            applicationNonVisualDrawingProperties15.Append(placeholderShape11);

            nonVisualShapeProperties11.Append(nonVisualDrawingProperties15);
            nonVisualShapeProperties11.Append(nonVisualShapeDrawingProperties11);
            nonVisualShapeProperties11.Append(applicationNonVisualDrawingProperties15);

            ShapeProperties shapeProperties11 = new ShapeProperties();

            A.Transform2D transform2D8 = new A.Transform2D();
            A.Offset offset12 = new A.Offset(){ X = 839788L, Y = 457200L };
            A.Extents extents12 = new A.Extents(){ Cx = 3932237L, Cy = 1600200L };

            transform2D8.Append(offset12);
            transform2D8.Append(extents12);

            shapeProperties11.Append(transform2D8);

            TextBody textBody11 = new TextBody();
            A.BodyProperties bodyProperties11 = new A.BodyProperties(){ Anchor = A.TextAnchoringTypeValues.Bottom };

            A.ListStyle listStyle11 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties10 = new A.Level1ParagraphProperties();
            A.DefaultRunProperties defaultRunProperties44 = new A.DefaultRunProperties(){ FontSize = 3200 };

            level1ParagraphProperties10.Append(defaultRunProperties44);

            listStyle11.Append(level1ParagraphProperties10);

            A.Paragraph paragraph15 = new A.Paragraph();

            A.Run run9 = new A.Run();
            A.RunProperties runProperties13 = new A.RunProperties(){ Language = "en-US" };
            A.Text text13 = new A.Text();
            text13.Text = "Click to edit Master title style";

            run9.Append(runProperties13);
            run9.Append(text13);
            A.EndParagraphRunProperties endParagraphRunProperties11 = new A.EndParagraphRunProperties(){ Language = "pl-PL" };

            paragraph15.Append(run9);
            paragraph15.Append(endParagraphRunProperties11);

            textBody11.Append(bodyProperties11);
            textBody11.Append(listStyle11);
            textBody11.Append(paragraph15);

            shape11.Append(nonVisualShapeProperties11);
            shape11.Append(shapeProperties11);
            shape11.Append(textBody11);

            Shape shape12 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties12 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties16 = new NonVisualDrawingProperties(){ Id = (UInt32Value)3U, Name = "Content Placeholder 2" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList12 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension12 = new A.NonVisualDrawingPropertiesExtension(){ Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            OpenXmlUnknownElement openXmlUnknownElement12 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{299F80EA-C924-167B-61DF-001B20C95D2D}\" />");

            nonVisualDrawingPropertiesExtension12.Append(openXmlUnknownElement12);

            nonVisualDrawingPropertiesExtensionList12.Append(nonVisualDrawingPropertiesExtension12);

            nonVisualDrawingProperties16.Append(nonVisualDrawingPropertiesExtensionList12);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties12 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks12 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties12.Append(shapeLocks12);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties16 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape12 = new PlaceholderShape(){ Index = (UInt32Value)1U };

            applicationNonVisualDrawingProperties16.Append(placeholderShape12);

            nonVisualShapeProperties12.Append(nonVisualDrawingProperties16);
            nonVisualShapeProperties12.Append(nonVisualShapeDrawingProperties12);
            nonVisualShapeProperties12.Append(applicationNonVisualDrawingProperties16);

            ShapeProperties shapeProperties12 = new ShapeProperties();

            A.Transform2D transform2D9 = new A.Transform2D();
            A.Offset offset13 = new A.Offset(){ X = 5183188L, Y = 987425L };
            A.Extents extents13 = new A.Extents(){ Cx = 6172200L, Cy = 4873625L };

            transform2D9.Append(offset13);
            transform2D9.Append(extents13);

            shapeProperties12.Append(transform2D9);

            TextBody textBody12 = new TextBody();
            A.BodyProperties bodyProperties12 = new A.BodyProperties();

            A.ListStyle listStyle12 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties11 = new A.Level1ParagraphProperties();
            A.DefaultRunProperties defaultRunProperties45 = new A.DefaultRunProperties(){ FontSize = 3200 };

            level1ParagraphProperties11.Append(defaultRunProperties45);

            A.Level2ParagraphProperties level2ParagraphProperties5 = new A.Level2ParagraphProperties();
            A.DefaultRunProperties defaultRunProperties46 = new A.DefaultRunProperties(){ FontSize = 2800 };

            level2ParagraphProperties5.Append(defaultRunProperties46);

            A.Level3ParagraphProperties level3ParagraphProperties5 = new A.Level3ParagraphProperties();
            A.DefaultRunProperties defaultRunProperties47 = new A.DefaultRunProperties(){ FontSize = 2400 };

            level3ParagraphProperties5.Append(defaultRunProperties47);

            A.Level4ParagraphProperties level4ParagraphProperties5 = new A.Level4ParagraphProperties();
            A.DefaultRunProperties defaultRunProperties48 = new A.DefaultRunProperties(){ FontSize = 2000 };

            level4ParagraphProperties5.Append(defaultRunProperties48);

            A.Level5ParagraphProperties level5ParagraphProperties5 = new A.Level5ParagraphProperties();
            A.DefaultRunProperties defaultRunProperties49 = new A.DefaultRunProperties(){ FontSize = 2000 };

            level5ParagraphProperties5.Append(defaultRunProperties49);

            A.Level6ParagraphProperties level6ParagraphProperties5 = new A.Level6ParagraphProperties();
            A.DefaultRunProperties defaultRunProperties50 = new A.DefaultRunProperties(){ FontSize = 2000 };

            level6ParagraphProperties5.Append(defaultRunProperties50);

            A.Level7ParagraphProperties level7ParagraphProperties5 = new A.Level7ParagraphProperties();
            A.DefaultRunProperties defaultRunProperties51 = new A.DefaultRunProperties(){ FontSize = 2000 };

            level7ParagraphProperties5.Append(defaultRunProperties51);

            A.Level8ParagraphProperties level8ParagraphProperties5 = new A.Level8ParagraphProperties();
            A.DefaultRunProperties defaultRunProperties52 = new A.DefaultRunProperties(){ FontSize = 2000 };

            level8ParagraphProperties5.Append(defaultRunProperties52);

            A.Level9ParagraphProperties level9ParagraphProperties5 = new A.Level9ParagraphProperties();
            A.DefaultRunProperties defaultRunProperties53 = new A.DefaultRunProperties(){ FontSize = 2000 };

            level9ParagraphProperties5.Append(defaultRunProperties53);

            listStyle12.Append(level1ParagraphProperties11);
            listStyle12.Append(level2ParagraphProperties5);
            listStyle12.Append(level3ParagraphProperties5);
            listStyle12.Append(level4ParagraphProperties5);
            listStyle12.Append(level5ParagraphProperties5);
            listStyle12.Append(level6ParagraphProperties5);
            listStyle12.Append(level7ParagraphProperties5);
            listStyle12.Append(level8ParagraphProperties5);
            listStyle12.Append(level9ParagraphProperties5);

            A.Paragraph paragraph16 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties6 = new A.ParagraphProperties(){ Level = 0 };

            A.Run run10 = new A.Run();
            A.RunProperties runProperties14 = new A.RunProperties(){ Language = "en-US" };
            A.Text text14 = new A.Text();
            text14.Text = "Click to edit Master text styles";

            run10.Append(runProperties14);
            run10.Append(text14);

            paragraph16.Append(paragraphProperties6);
            paragraph16.Append(run10);

            A.Paragraph paragraph17 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties7 = new A.ParagraphProperties(){ Level = 1 };

            A.Run run11 = new A.Run();
            A.RunProperties runProperties15 = new A.RunProperties(){ Language = "en-US" };
            A.Text text15 = new A.Text();
            text15.Text = "Second level";

            run11.Append(runProperties15);
            run11.Append(text15);

            paragraph17.Append(paragraphProperties7);
            paragraph17.Append(run11);

            A.Paragraph paragraph18 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties8 = new A.ParagraphProperties(){ Level = 2 };

            A.Run run12 = new A.Run();
            A.RunProperties runProperties16 = new A.RunProperties(){ Language = "en-US" };
            A.Text text16 = new A.Text();
            text16.Text = "Third level";

            run12.Append(runProperties16);
            run12.Append(text16);

            paragraph18.Append(paragraphProperties8);
            paragraph18.Append(run12);

            A.Paragraph paragraph19 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties9 = new A.ParagraphProperties(){ Level = 3 };

            A.Run run13 = new A.Run();
            A.RunProperties runProperties17 = new A.RunProperties(){ Language = "en-US" };
            A.Text text17 = new A.Text();
            text17.Text = "Fourth level";

            run13.Append(runProperties17);
            run13.Append(text17);

            paragraph19.Append(paragraphProperties9);
            paragraph19.Append(run13);

            A.Paragraph paragraph20 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties10 = new A.ParagraphProperties(){ Level = 4 };

            A.Run run14 = new A.Run();
            A.RunProperties runProperties18 = new A.RunProperties(){ Language = "en-US" };
            A.Text text18 = new A.Text();
            text18.Text = "Fifth level";

            run14.Append(runProperties18);
            run14.Append(text18);
            A.EndParagraphRunProperties endParagraphRunProperties12 = new A.EndParagraphRunProperties(){ Language = "pl-PL" };

            paragraph20.Append(paragraphProperties10);
            paragraph20.Append(run14);
            paragraph20.Append(endParagraphRunProperties12);

            textBody12.Append(bodyProperties12);
            textBody12.Append(listStyle12);
            textBody12.Append(paragraph16);
            textBody12.Append(paragraph17);
            textBody12.Append(paragraph18);
            textBody12.Append(paragraph19);
            textBody12.Append(paragraph20);

            shape12.Append(nonVisualShapeProperties12);
            shape12.Append(shapeProperties12);
            shape12.Append(textBody12);

            Shape shape13 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties13 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties17 = new NonVisualDrawingProperties(){ Id = (UInt32Value)4U, Name = "Text Placeholder 3" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList13 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension13 = new A.NonVisualDrawingPropertiesExtension(){ Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            OpenXmlUnknownElement openXmlUnknownElement13 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{7FBAFD28-36D8-1D95-ECE1-29F9F0C63AB2}\" />");

            nonVisualDrawingPropertiesExtension13.Append(openXmlUnknownElement13);

            nonVisualDrawingPropertiesExtensionList13.Append(nonVisualDrawingPropertiesExtension13);

            nonVisualDrawingProperties17.Append(nonVisualDrawingPropertiesExtensionList13);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties13 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks13 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties13.Append(shapeLocks13);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties17 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape13 = new PlaceholderShape(){ Type = PlaceholderValues.Body, Size = PlaceholderSizeValues.Half, Index = (UInt32Value)2U };

            applicationNonVisualDrawingProperties17.Append(placeholderShape13);

            nonVisualShapeProperties13.Append(nonVisualDrawingProperties17);
            nonVisualShapeProperties13.Append(nonVisualShapeDrawingProperties13);
            nonVisualShapeProperties13.Append(applicationNonVisualDrawingProperties17);

            ShapeProperties shapeProperties13 = new ShapeProperties();

            A.Transform2D transform2D10 = new A.Transform2D();
            A.Offset offset14 = new A.Offset(){ X = 839788L, Y = 2057400L };
            A.Extents extents14 = new A.Extents(){ Cx = 3932237L, Cy = 3811588L };

            transform2D10.Append(offset14);
            transform2D10.Append(extents14);

            shapeProperties13.Append(transform2D10);

            TextBody textBody13 = new TextBody();
            A.BodyProperties bodyProperties13 = new A.BodyProperties();

            A.ListStyle listStyle13 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties12 = new A.Level1ParagraphProperties(){ LeftMargin = 0, Indent = 0 };
            A.NoBullet noBullet11 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties54 = new A.DefaultRunProperties(){ FontSize = 1600 };

            level1ParagraphProperties12.Append(noBullet11);
            level1ParagraphProperties12.Append(defaultRunProperties54);

            A.Level2ParagraphProperties level2ParagraphProperties6 = new A.Level2ParagraphProperties(){ LeftMargin = 457200, Indent = 0 };
            A.NoBullet noBullet12 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties55 = new A.DefaultRunProperties(){ FontSize = 1400 };

            level2ParagraphProperties6.Append(noBullet12);
            level2ParagraphProperties6.Append(defaultRunProperties55);

            A.Level3ParagraphProperties level3ParagraphProperties6 = new A.Level3ParagraphProperties(){ LeftMargin = 914400, Indent = 0 };
            A.NoBullet noBullet13 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties56 = new A.DefaultRunProperties(){ FontSize = 1200 };

            level3ParagraphProperties6.Append(noBullet13);
            level3ParagraphProperties6.Append(defaultRunProperties56);

            A.Level4ParagraphProperties level4ParagraphProperties6 = new A.Level4ParagraphProperties(){ LeftMargin = 1371600, Indent = 0 };
            A.NoBullet noBullet14 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties57 = new A.DefaultRunProperties(){ FontSize = 1000 };

            level4ParagraphProperties6.Append(noBullet14);
            level4ParagraphProperties6.Append(defaultRunProperties57);

            A.Level5ParagraphProperties level5ParagraphProperties6 = new A.Level5ParagraphProperties(){ LeftMargin = 1828800, Indent = 0 };
            A.NoBullet noBullet15 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties58 = new A.DefaultRunProperties(){ FontSize = 1000 };

            level5ParagraphProperties6.Append(noBullet15);
            level5ParagraphProperties6.Append(defaultRunProperties58);

            A.Level6ParagraphProperties level6ParagraphProperties6 = new A.Level6ParagraphProperties(){ LeftMargin = 2286000, Indent = 0 };
            A.NoBullet noBullet16 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties59 = new A.DefaultRunProperties(){ FontSize = 1000 };

            level6ParagraphProperties6.Append(noBullet16);
            level6ParagraphProperties6.Append(defaultRunProperties59);

            A.Level7ParagraphProperties level7ParagraphProperties6 = new A.Level7ParagraphProperties(){ LeftMargin = 2743200, Indent = 0 };
            A.NoBullet noBullet17 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties60 = new A.DefaultRunProperties(){ FontSize = 1000 };

            level7ParagraphProperties6.Append(noBullet17);
            level7ParagraphProperties6.Append(defaultRunProperties60);

            A.Level8ParagraphProperties level8ParagraphProperties6 = new A.Level8ParagraphProperties(){ LeftMargin = 3200400, Indent = 0 };
            A.NoBullet noBullet18 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties61 = new A.DefaultRunProperties(){ FontSize = 1000 };

            level8ParagraphProperties6.Append(noBullet18);
            level8ParagraphProperties6.Append(defaultRunProperties61);

            A.Level9ParagraphProperties level9ParagraphProperties6 = new A.Level9ParagraphProperties(){ LeftMargin = 3657600, Indent = 0 };
            A.NoBullet noBullet19 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties62 = new A.DefaultRunProperties(){ FontSize = 1000 };

            level9ParagraphProperties6.Append(noBullet19);
            level9ParagraphProperties6.Append(defaultRunProperties62);

            listStyle13.Append(level1ParagraphProperties12);
            listStyle13.Append(level2ParagraphProperties6);
            listStyle13.Append(level3ParagraphProperties6);
            listStyle13.Append(level4ParagraphProperties6);
            listStyle13.Append(level5ParagraphProperties6);
            listStyle13.Append(level6ParagraphProperties6);
            listStyle13.Append(level7ParagraphProperties6);
            listStyle13.Append(level8ParagraphProperties6);
            listStyle13.Append(level9ParagraphProperties6);

            A.Paragraph paragraph21 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties11 = new A.ParagraphProperties(){ Level = 0 };

            A.Run run15 = new A.Run();
            A.RunProperties runProperties19 = new A.RunProperties(){ Language = "en-US" };
            A.Text text19 = new A.Text();
            text19.Text = "Click to edit Master text styles";

            run15.Append(runProperties19);
            run15.Append(text19);

            paragraph21.Append(paragraphProperties11);
            paragraph21.Append(run15);

            textBody13.Append(bodyProperties13);
            textBody13.Append(listStyle13);
            textBody13.Append(paragraph21);

            shape13.Append(nonVisualShapeProperties13);
            shape13.Append(shapeProperties13);
            shape13.Append(textBody13);

            Shape shape14 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties14 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties18 = new NonVisualDrawingProperties(){ Id = (UInt32Value)5U, Name = "Date Placeholder 4" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList14 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension14 = new A.NonVisualDrawingPropertiesExtension(){ Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            OpenXmlUnknownElement openXmlUnknownElement14 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{DA58A854-C3C3-B609-1170-9A53519F91FE}\" />");

            nonVisualDrawingPropertiesExtension14.Append(openXmlUnknownElement14);

            nonVisualDrawingPropertiesExtensionList14.Append(nonVisualDrawingPropertiesExtension14);

            nonVisualDrawingProperties18.Append(nonVisualDrawingPropertiesExtensionList14);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties14 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks14 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties14.Append(shapeLocks14);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties18 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape14 = new PlaceholderShape(){ Type = PlaceholderValues.DateAndTime, Size = PlaceholderSizeValues.Half, Index = (UInt32Value)10U };

            applicationNonVisualDrawingProperties18.Append(placeholderShape14);

            nonVisualShapeProperties14.Append(nonVisualDrawingProperties18);
            nonVisualShapeProperties14.Append(nonVisualShapeDrawingProperties14);
            nonVisualShapeProperties14.Append(applicationNonVisualDrawingProperties18);
            ShapeProperties shapeProperties14 = new ShapeProperties();

            TextBody textBody14 = new TextBody();
            A.BodyProperties bodyProperties14 = new A.BodyProperties();
            A.ListStyle listStyle14 = new A.ListStyle();

            A.Paragraph paragraph22 = new A.Paragraph();

            A.Field field5 = new A.Field(){ Id = "{F550F806-175B-4EE1-ABEA-A5F04E4FD72B}", Type = "datetimeFigureOut" };

            A.RunProperties runProperties20 = new A.RunProperties(){ Language = "pl-PL" };
            runProperties20.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text20 = new A.Text();
            text20.Text = "20.08.2025";

            field5.Append(runProperties20);
            field5.Append(text20);
            A.EndParagraphRunProperties endParagraphRunProperties13 = new A.EndParagraphRunProperties(){ Language = "pl-PL" };

            paragraph22.Append(field5);
            paragraph22.Append(endParagraphRunProperties13);

            textBody14.Append(bodyProperties14);
            textBody14.Append(listStyle14);
            textBody14.Append(paragraph22);

            shape14.Append(nonVisualShapeProperties14);
            shape14.Append(shapeProperties14);
            shape14.Append(textBody14);

            Shape shape15 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties15 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties19 = new NonVisualDrawingProperties(){ Id = (UInt32Value)6U, Name = "Footer Placeholder 5" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList15 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension15 = new A.NonVisualDrawingPropertiesExtension(){ Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            OpenXmlUnknownElement openXmlUnknownElement15 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{77FDADC3-A993-FDB8-75A7-428B989CCE6B}\" />");

            nonVisualDrawingPropertiesExtension15.Append(openXmlUnknownElement15);

            nonVisualDrawingPropertiesExtensionList15.Append(nonVisualDrawingPropertiesExtension15);

            nonVisualDrawingProperties19.Append(nonVisualDrawingPropertiesExtensionList15);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties15 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks15 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties15.Append(shapeLocks15);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties19 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape15 = new PlaceholderShape(){ Type = PlaceholderValues.Footer, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)11U };

            applicationNonVisualDrawingProperties19.Append(placeholderShape15);

            nonVisualShapeProperties15.Append(nonVisualDrawingProperties19);
            nonVisualShapeProperties15.Append(nonVisualShapeDrawingProperties15);
            nonVisualShapeProperties15.Append(applicationNonVisualDrawingProperties19);
            ShapeProperties shapeProperties15 = new ShapeProperties();

            TextBody textBody15 = new TextBody();
            A.BodyProperties bodyProperties15 = new A.BodyProperties();
            A.ListStyle listStyle15 = new A.ListStyle();

            A.Paragraph paragraph23 = new A.Paragraph();
            A.EndParagraphRunProperties endParagraphRunProperties14 = new A.EndParagraphRunProperties(){ Language = "pl-PL" };

            paragraph23.Append(endParagraphRunProperties14);

            textBody15.Append(bodyProperties15);
            textBody15.Append(listStyle15);
            textBody15.Append(paragraph23);

            shape15.Append(nonVisualShapeProperties15);
            shape15.Append(shapeProperties15);
            shape15.Append(textBody15);

            Shape shape16 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties16 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties20 = new NonVisualDrawingProperties(){ Id = (UInt32Value)7U, Name = "Slide Number Placeholder 6" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList16 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension16 = new A.NonVisualDrawingPropertiesExtension(){ Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            OpenXmlUnknownElement openXmlUnknownElement16 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{C234B184-4DCC-07B8-3FFF-B051A5A3BBC2}\" />");

            nonVisualDrawingPropertiesExtension16.Append(openXmlUnknownElement16);

            nonVisualDrawingPropertiesExtensionList16.Append(nonVisualDrawingPropertiesExtension16);

            nonVisualDrawingProperties20.Append(nonVisualDrawingPropertiesExtensionList16);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties16 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks16 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties16.Append(shapeLocks16);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties20 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape16 = new PlaceholderShape(){ Type = PlaceholderValues.SlideNumber, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)12U };

            applicationNonVisualDrawingProperties20.Append(placeholderShape16);

            nonVisualShapeProperties16.Append(nonVisualDrawingProperties20);
            nonVisualShapeProperties16.Append(nonVisualShapeDrawingProperties16);
            nonVisualShapeProperties16.Append(applicationNonVisualDrawingProperties20);
            ShapeProperties shapeProperties16 = new ShapeProperties();

            TextBody textBody16 = new TextBody();
            A.BodyProperties bodyProperties16 = new A.BodyProperties();
            A.ListStyle listStyle16 = new A.ListStyle();

            A.Paragraph paragraph24 = new A.Paragraph();

            A.Field field6 = new A.Field(){ Id = "{5FAB34AF-B09A-4E2A-BAD3-CCD6E0AD81E6}", Type = "slidenum" };

            A.RunProperties runProperties21 = new A.RunProperties(){ Language = "pl-PL" };
            runProperties21.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text21 = new A.Text();
            text21.Text = "‹#›";

            field6.Append(runProperties21);
            field6.Append(text21);
            A.EndParagraphRunProperties endParagraphRunProperties15 = new A.EndParagraphRunProperties(){ Language = "pl-PL" };

            paragraph24.Append(field6);
            paragraph24.Append(endParagraphRunProperties15);

            textBody16.Append(bodyProperties16);
            textBody16.Append(listStyle16);
            textBody16.Append(paragraph24);

            shape16.Append(nonVisualShapeProperties16);
            shape16.Append(shapeProperties16);
            shape16.Append(textBody16);

            shapeTree4.Append(nonVisualGroupShapeProperties4);
            shapeTree4.Append(groupShapeProperties4);
            shapeTree4.Append(shape11);
            shapeTree4.Append(shape12);
            shapeTree4.Append(shape13);
            shapeTree4.Append(shape14);
            shapeTree4.Append(shape15);
            shapeTree4.Append(shape16);

            CommonSlideDataExtensionList commonSlideDataExtensionList4 = new CommonSlideDataExtensionList();

            CommonSlideDataExtension commonSlideDataExtension4 = new CommonSlideDataExtension(){ Uri = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}" };

            P14.CreationId creationId4 = new P14.CreationId(){ Val = (UInt32Value)2228906514U };
            creationId4.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            commonSlideDataExtension4.Append(creationId4);

            commonSlideDataExtensionList4.Append(commonSlideDataExtension4);

            commonSlideData4.Append(shapeTree4);
            commonSlideData4.Append(commonSlideDataExtensionList4);

            ColorMapOverride colorMapOverride3 = new ColorMapOverride();
            A.MasterColorMapping masterColorMapping3 = new A.MasterColorMapping();

            colorMapOverride3.Append(masterColorMapping3);

            slideLayout2.Append(commonSlideData4);
            slideLayout2.Append(colorMapOverride3);

            slideLayoutPart2.SlideLayout = slideLayout2;
        }

        // Generates content of slideLayoutPart3.
        private void GenerateSlideLayoutPart3Content(SlideLayoutPart slideLayoutPart3)
        {
            SlideLayout slideLayout3 = new SlideLayout(){ Type = SlideLayoutValues.SectionHeader, Preserve = true };
            slideLayout3.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            slideLayout3.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            slideLayout3.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            CommonSlideData commonSlideData5 = new CommonSlideData(){ Name = "Section Header" };

            ShapeTree shapeTree5 = new ShapeTree();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties5 = new NonVisualGroupShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties21 = new NonVisualDrawingProperties(){ Id = (UInt32Value)1U, Name = "" };
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties5 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties21 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties5.Append(nonVisualDrawingProperties21);
            nonVisualGroupShapeProperties5.Append(nonVisualGroupShapeDrawingProperties5);
            nonVisualGroupShapeProperties5.Append(applicationNonVisualDrawingProperties21);

            GroupShapeProperties groupShapeProperties5 = new GroupShapeProperties();

            A.TransformGroup transformGroup5 = new A.TransformGroup();
            A.Offset offset15 = new A.Offset(){ X = 0L, Y = 0L };
            A.Extents extents15 = new A.Extents(){ Cx = 0L, Cy = 0L };
            A.ChildOffset childOffset5 = new A.ChildOffset(){ X = 0L, Y = 0L };
            A.ChildExtents childExtents5 = new A.ChildExtents(){ Cx = 0L, Cy = 0L };

            transformGroup5.Append(offset15);
            transformGroup5.Append(extents15);
            transformGroup5.Append(childOffset5);
            transformGroup5.Append(childExtents5);

            groupShapeProperties5.Append(transformGroup5);

            Shape shape17 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties17 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties22 = new NonVisualDrawingProperties(){ Id = (UInt32Value)2U, Name = "Title 1" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList17 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension17 = new A.NonVisualDrawingPropertiesExtension(){ Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            OpenXmlUnknownElement openXmlUnknownElement17 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{20B3736A-67F0-A2C2-4BE7-027F8E1C3A4F}\" />");

            nonVisualDrawingPropertiesExtension17.Append(openXmlUnknownElement17);

            nonVisualDrawingPropertiesExtensionList17.Append(nonVisualDrawingPropertiesExtension17);

            nonVisualDrawingProperties22.Append(nonVisualDrawingPropertiesExtensionList17);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties17 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks17 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties17.Append(shapeLocks17);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties22 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape17 = new PlaceholderShape(){ Type = PlaceholderValues.Title };

            applicationNonVisualDrawingProperties22.Append(placeholderShape17);

            nonVisualShapeProperties17.Append(nonVisualDrawingProperties22);
            nonVisualShapeProperties17.Append(nonVisualShapeDrawingProperties17);
            nonVisualShapeProperties17.Append(applicationNonVisualDrawingProperties22);

            ShapeProperties shapeProperties17 = new ShapeProperties();

            A.Transform2D transform2D11 = new A.Transform2D();
            A.Offset offset16 = new A.Offset(){ X = 831850L, Y = 1709738L };
            A.Extents extents16 = new A.Extents(){ Cx = 10515600L, Cy = 2852737L };

            transform2D11.Append(offset16);
            transform2D11.Append(extents16);

            shapeProperties17.Append(transform2D11);

            TextBody textBody17 = new TextBody();
            A.BodyProperties bodyProperties17 = new A.BodyProperties(){ Anchor = A.TextAnchoringTypeValues.Bottom };

            A.ListStyle listStyle17 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties13 = new A.Level1ParagraphProperties();
            A.DefaultRunProperties defaultRunProperties63 = new A.DefaultRunProperties(){ FontSize = 6000 };

            level1ParagraphProperties13.Append(defaultRunProperties63);

            listStyle17.Append(level1ParagraphProperties13);

            A.Paragraph paragraph25 = new A.Paragraph();

            A.Run run16 = new A.Run();
            A.RunProperties runProperties22 = new A.RunProperties(){ Language = "en-US" };
            A.Text text22 = new A.Text();
            text22.Text = "Click to edit Master title style";

            run16.Append(runProperties22);
            run16.Append(text22);
            A.EndParagraphRunProperties endParagraphRunProperties16 = new A.EndParagraphRunProperties(){ Language = "pl-PL" };

            paragraph25.Append(run16);
            paragraph25.Append(endParagraphRunProperties16);

            textBody17.Append(bodyProperties17);
            textBody17.Append(listStyle17);
            textBody17.Append(paragraph25);

            shape17.Append(nonVisualShapeProperties17);
            shape17.Append(shapeProperties17);
            shape17.Append(textBody17);

            Shape shape18 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties18 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties23 = new NonVisualDrawingProperties(){ Id = (UInt32Value)3U, Name = "Text Placeholder 2" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList18 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension18 = new A.NonVisualDrawingPropertiesExtension(){ Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            OpenXmlUnknownElement openXmlUnknownElement18 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{BD716907-37A8-1558-7E02-FF55B2685DAA}\" />");

            nonVisualDrawingPropertiesExtension18.Append(openXmlUnknownElement18);

            nonVisualDrawingPropertiesExtensionList18.Append(nonVisualDrawingPropertiesExtension18);

            nonVisualDrawingProperties23.Append(nonVisualDrawingPropertiesExtensionList18);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties18 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks18 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties18.Append(shapeLocks18);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties23 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape18 = new PlaceholderShape(){ Type = PlaceholderValues.Body, Index = (UInt32Value)1U };

            applicationNonVisualDrawingProperties23.Append(placeholderShape18);

            nonVisualShapeProperties18.Append(nonVisualDrawingProperties23);
            nonVisualShapeProperties18.Append(nonVisualShapeDrawingProperties18);
            nonVisualShapeProperties18.Append(applicationNonVisualDrawingProperties23);

            ShapeProperties shapeProperties18 = new ShapeProperties();

            A.Transform2D transform2D12 = new A.Transform2D();
            A.Offset offset17 = new A.Offset(){ X = 831850L, Y = 4589463L };
            A.Extents extents17 = new A.Extents(){ Cx = 10515600L, Cy = 1500187L };

            transform2D12.Append(offset17);
            transform2D12.Append(extents17);

            shapeProperties18.Append(transform2D12);

            TextBody textBody18 = new TextBody();
            A.BodyProperties bodyProperties18 = new A.BodyProperties();

            A.ListStyle listStyle18 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties14 = new A.Level1ParagraphProperties(){ LeftMargin = 0, Indent = 0 };
            A.NoBullet noBullet20 = new A.NoBullet();

            A.DefaultRunProperties defaultRunProperties64 = new A.DefaultRunProperties(){ FontSize = 2400 };

            A.SolidFill solidFill32 = new A.SolidFill();

            A.SchemeColor schemeColor33 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };
            A.Tint tint4 = new A.Tint(){ Val = 82000 };

            schemeColor33.Append(tint4);

            solidFill32.Append(schemeColor33);

            defaultRunProperties64.Append(solidFill32);

            level1ParagraphProperties14.Append(noBullet20);
            level1ParagraphProperties14.Append(defaultRunProperties64);

            A.Level2ParagraphProperties level2ParagraphProperties7 = new A.Level2ParagraphProperties(){ LeftMargin = 457200, Indent = 0 };
            A.NoBullet noBullet21 = new A.NoBullet();

            A.DefaultRunProperties defaultRunProperties65 = new A.DefaultRunProperties(){ FontSize = 2000 };

            A.SolidFill solidFill33 = new A.SolidFill();

            A.SchemeColor schemeColor34 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };
            A.Tint tint5 = new A.Tint(){ Val = 82000 };

            schemeColor34.Append(tint5);

            solidFill33.Append(schemeColor34);

            defaultRunProperties65.Append(solidFill33);

            level2ParagraphProperties7.Append(noBullet21);
            level2ParagraphProperties7.Append(defaultRunProperties65);

            A.Level3ParagraphProperties level3ParagraphProperties7 = new A.Level3ParagraphProperties(){ LeftMargin = 914400, Indent = 0 };
            A.NoBullet noBullet22 = new A.NoBullet();

            A.DefaultRunProperties defaultRunProperties66 = new A.DefaultRunProperties(){ FontSize = 1800 };

            A.SolidFill solidFill34 = new A.SolidFill();

            A.SchemeColor schemeColor35 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };
            A.Tint tint6 = new A.Tint(){ Val = 82000 };

            schemeColor35.Append(tint6);

            solidFill34.Append(schemeColor35);

            defaultRunProperties66.Append(solidFill34);

            level3ParagraphProperties7.Append(noBullet22);
            level3ParagraphProperties7.Append(defaultRunProperties66);

            A.Level4ParagraphProperties level4ParagraphProperties7 = new A.Level4ParagraphProperties(){ LeftMargin = 1371600, Indent = 0 };
            A.NoBullet noBullet23 = new A.NoBullet();

            A.DefaultRunProperties defaultRunProperties67 = new A.DefaultRunProperties(){ FontSize = 1600 };

            A.SolidFill solidFill35 = new A.SolidFill();

            A.SchemeColor schemeColor36 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };
            A.Tint tint7 = new A.Tint(){ Val = 82000 };

            schemeColor36.Append(tint7);

            solidFill35.Append(schemeColor36);

            defaultRunProperties67.Append(solidFill35);

            level4ParagraphProperties7.Append(noBullet23);
            level4ParagraphProperties7.Append(defaultRunProperties67);

            A.Level5ParagraphProperties level5ParagraphProperties7 = new A.Level5ParagraphProperties(){ LeftMargin = 1828800, Indent = 0 };
            A.NoBullet noBullet24 = new A.NoBullet();

            A.DefaultRunProperties defaultRunProperties68 = new A.DefaultRunProperties(){ FontSize = 1600 };

            A.SolidFill solidFill36 = new A.SolidFill();

            A.SchemeColor schemeColor37 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };
            A.Tint tint8 = new A.Tint(){ Val = 82000 };

            schemeColor37.Append(tint8);

            solidFill36.Append(schemeColor37);

            defaultRunProperties68.Append(solidFill36);

            level5ParagraphProperties7.Append(noBullet24);
            level5ParagraphProperties7.Append(defaultRunProperties68);

            A.Level6ParagraphProperties level6ParagraphProperties7 = new A.Level6ParagraphProperties(){ LeftMargin = 2286000, Indent = 0 };
            A.NoBullet noBullet25 = new A.NoBullet();

            A.DefaultRunProperties defaultRunProperties69 = new A.DefaultRunProperties(){ FontSize = 1600 };

            A.SolidFill solidFill37 = new A.SolidFill();

            A.SchemeColor schemeColor38 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };
            A.Tint tint9 = new A.Tint(){ Val = 82000 };

            schemeColor38.Append(tint9);

            solidFill37.Append(schemeColor38);

            defaultRunProperties69.Append(solidFill37);

            level6ParagraphProperties7.Append(noBullet25);
            level6ParagraphProperties7.Append(defaultRunProperties69);

            A.Level7ParagraphProperties level7ParagraphProperties7 = new A.Level7ParagraphProperties(){ LeftMargin = 2743200, Indent = 0 };
            A.NoBullet noBullet26 = new A.NoBullet();

            A.DefaultRunProperties defaultRunProperties70 = new A.DefaultRunProperties(){ FontSize = 1600 };

            A.SolidFill solidFill38 = new A.SolidFill();

            A.SchemeColor schemeColor39 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };
            A.Tint tint10 = new A.Tint(){ Val = 82000 };

            schemeColor39.Append(tint10);

            solidFill38.Append(schemeColor39);

            defaultRunProperties70.Append(solidFill38);

            level7ParagraphProperties7.Append(noBullet26);
            level7ParagraphProperties7.Append(defaultRunProperties70);

            A.Level8ParagraphProperties level8ParagraphProperties7 = new A.Level8ParagraphProperties(){ LeftMargin = 3200400, Indent = 0 };
            A.NoBullet noBullet27 = new A.NoBullet();

            A.DefaultRunProperties defaultRunProperties71 = new A.DefaultRunProperties(){ FontSize = 1600 };

            A.SolidFill solidFill39 = new A.SolidFill();

            A.SchemeColor schemeColor40 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };
            A.Tint tint11 = new A.Tint(){ Val = 82000 };

            schemeColor40.Append(tint11);

            solidFill39.Append(schemeColor40);

            defaultRunProperties71.Append(solidFill39);

            level8ParagraphProperties7.Append(noBullet27);
            level8ParagraphProperties7.Append(defaultRunProperties71);

            A.Level9ParagraphProperties level9ParagraphProperties7 = new A.Level9ParagraphProperties(){ LeftMargin = 3657600, Indent = 0 };
            A.NoBullet noBullet28 = new A.NoBullet();

            A.DefaultRunProperties defaultRunProperties72 = new A.DefaultRunProperties(){ FontSize = 1600 };

            A.SolidFill solidFill40 = new A.SolidFill();

            A.SchemeColor schemeColor41 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };
            A.Tint tint12 = new A.Tint(){ Val = 82000 };

            schemeColor41.Append(tint12);

            solidFill40.Append(schemeColor41);

            defaultRunProperties72.Append(solidFill40);

            level9ParagraphProperties7.Append(noBullet28);
            level9ParagraphProperties7.Append(defaultRunProperties72);

            listStyle18.Append(level1ParagraphProperties14);
            listStyle18.Append(level2ParagraphProperties7);
            listStyle18.Append(level3ParagraphProperties7);
            listStyle18.Append(level4ParagraphProperties7);
            listStyle18.Append(level5ParagraphProperties7);
            listStyle18.Append(level6ParagraphProperties7);
            listStyle18.Append(level7ParagraphProperties7);
            listStyle18.Append(level8ParagraphProperties7);
            listStyle18.Append(level9ParagraphProperties7);

            A.Paragraph paragraph26 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties12 = new A.ParagraphProperties(){ Level = 0 };

            A.Run run17 = new A.Run();
            A.RunProperties runProperties23 = new A.RunProperties(){ Language = "en-US" };
            A.Text text23 = new A.Text();
            text23.Text = "Click to edit Master text styles";

            run17.Append(runProperties23);
            run17.Append(text23);

            paragraph26.Append(paragraphProperties12);
            paragraph26.Append(run17);

            textBody18.Append(bodyProperties18);
            textBody18.Append(listStyle18);
            textBody18.Append(paragraph26);

            shape18.Append(nonVisualShapeProperties18);
            shape18.Append(shapeProperties18);
            shape18.Append(textBody18);

            Shape shape19 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties19 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties24 = new NonVisualDrawingProperties(){ Id = (UInt32Value)4U, Name = "Date Placeholder 3" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList19 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension19 = new A.NonVisualDrawingPropertiesExtension(){ Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            OpenXmlUnknownElement openXmlUnknownElement19 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{DA9B272C-0100-1096-9A31-D71DDE9D0199}\" />");

            nonVisualDrawingPropertiesExtension19.Append(openXmlUnknownElement19);

            nonVisualDrawingPropertiesExtensionList19.Append(nonVisualDrawingPropertiesExtension19);

            nonVisualDrawingProperties24.Append(nonVisualDrawingPropertiesExtensionList19);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties19 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks19 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties19.Append(shapeLocks19);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties24 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape19 = new PlaceholderShape(){ Type = PlaceholderValues.DateAndTime, Size = PlaceholderSizeValues.Half, Index = (UInt32Value)10U };

            applicationNonVisualDrawingProperties24.Append(placeholderShape19);

            nonVisualShapeProperties19.Append(nonVisualDrawingProperties24);
            nonVisualShapeProperties19.Append(nonVisualShapeDrawingProperties19);
            nonVisualShapeProperties19.Append(applicationNonVisualDrawingProperties24);
            ShapeProperties shapeProperties19 = new ShapeProperties();

            TextBody textBody19 = new TextBody();
            A.BodyProperties bodyProperties19 = new A.BodyProperties();
            A.ListStyle listStyle19 = new A.ListStyle();

            A.Paragraph paragraph27 = new A.Paragraph();

            A.Field field7 = new A.Field(){ Id = "{F550F806-175B-4EE1-ABEA-A5F04E4FD72B}", Type = "datetimeFigureOut" };

            A.RunProperties runProperties24 = new A.RunProperties(){ Language = "pl-PL" };
            runProperties24.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text24 = new A.Text();
            text24.Text = "20.08.2025";

            field7.Append(runProperties24);
            field7.Append(text24);
            A.EndParagraphRunProperties endParagraphRunProperties17 = new A.EndParagraphRunProperties(){ Language = "pl-PL" };

            paragraph27.Append(field7);
            paragraph27.Append(endParagraphRunProperties17);

            textBody19.Append(bodyProperties19);
            textBody19.Append(listStyle19);
            textBody19.Append(paragraph27);

            shape19.Append(nonVisualShapeProperties19);
            shape19.Append(shapeProperties19);
            shape19.Append(textBody19);

            Shape shape20 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties20 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties25 = new NonVisualDrawingProperties(){ Id = (UInt32Value)5U, Name = "Footer Placeholder 4" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList20 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension20 = new A.NonVisualDrawingPropertiesExtension(){ Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            OpenXmlUnknownElement openXmlUnknownElement20 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{BDF9B897-809A-1719-4648-49B84910DBB9}\" />");

            nonVisualDrawingPropertiesExtension20.Append(openXmlUnknownElement20);

            nonVisualDrawingPropertiesExtensionList20.Append(nonVisualDrawingPropertiesExtension20);

            nonVisualDrawingProperties25.Append(nonVisualDrawingPropertiesExtensionList20);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties20 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks20 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties20.Append(shapeLocks20);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties25 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape20 = new PlaceholderShape(){ Type = PlaceholderValues.Footer, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)11U };

            applicationNonVisualDrawingProperties25.Append(placeholderShape20);

            nonVisualShapeProperties20.Append(nonVisualDrawingProperties25);
            nonVisualShapeProperties20.Append(nonVisualShapeDrawingProperties20);
            nonVisualShapeProperties20.Append(applicationNonVisualDrawingProperties25);
            ShapeProperties shapeProperties20 = new ShapeProperties();

            TextBody textBody20 = new TextBody();
            A.BodyProperties bodyProperties20 = new A.BodyProperties();
            A.ListStyle listStyle20 = new A.ListStyle();

            A.Paragraph paragraph28 = new A.Paragraph();
            A.EndParagraphRunProperties endParagraphRunProperties18 = new A.EndParagraphRunProperties(){ Language = "pl-PL" };

            paragraph28.Append(endParagraphRunProperties18);

            textBody20.Append(bodyProperties20);
            textBody20.Append(listStyle20);
            textBody20.Append(paragraph28);

            shape20.Append(nonVisualShapeProperties20);
            shape20.Append(shapeProperties20);
            shape20.Append(textBody20);

            Shape shape21 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties21 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties26 = new NonVisualDrawingProperties(){ Id = (UInt32Value)6U, Name = "Slide Number Placeholder 5" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList21 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension21 = new A.NonVisualDrawingPropertiesExtension(){ Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            OpenXmlUnknownElement openXmlUnknownElement21 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{0ACB2D60-068F-1150-CDE9-4511C6EB7EE4}\" />");

            nonVisualDrawingPropertiesExtension21.Append(openXmlUnknownElement21);

            nonVisualDrawingPropertiesExtensionList21.Append(nonVisualDrawingPropertiesExtension21);

            nonVisualDrawingProperties26.Append(nonVisualDrawingPropertiesExtensionList21);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties21 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks21 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties21.Append(shapeLocks21);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties26 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape21 = new PlaceholderShape(){ Type = PlaceholderValues.SlideNumber, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)12U };

            applicationNonVisualDrawingProperties26.Append(placeholderShape21);

            nonVisualShapeProperties21.Append(nonVisualDrawingProperties26);
            nonVisualShapeProperties21.Append(nonVisualShapeDrawingProperties21);
            nonVisualShapeProperties21.Append(applicationNonVisualDrawingProperties26);
            ShapeProperties shapeProperties21 = new ShapeProperties();

            TextBody textBody21 = new TextBody();
            A.BodyProperties bodyProperties21 = new A.BodyProperties();
            A.ListStyle listStyle21 = new A.ListStyle();

            A.Paragraph paragraph29 = new A.Paragraph();

            A.Field field8 = new A.Field(){ Id = "{5FAB34AF-B09A-4E2A-BAD3-CCD6E0AD81E6}", Type = "slidenum" };

            A.RunProperties runProperties25 = new A.RunProperties(){ Language = "pl-PL" };
            runProperties25.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text25 = new A.Text();
            text25.Text = "‹#›";

            field8.Append(runProperties25);
            field8.Append(text25);
            A.EndParagraphRunProperties endParagraphRunProperties19 = new A.EndParagraphRunProperties(){ Language = "pl-PL" };

            paragraph29.Append(field8);
            paragraph29.Append(endParagraphRunProperties19);

            textBody21.Append(bodyProperties21);
            textBody21.Append(listStyle21);
            textBody21.Append(paragraph29);

            shape21.Append(nonVisualShapeProperties21);
            shape21.Append(shapeProperties21);
            shape21.Append(textBody21);

            shapeTree5.Append(nonVisualGroupShapeProperties5);
            shapeTree5.Append(groupShapeProperties5);
            shapeTree5.Append(shape17);
            shapeTree5.Append(shape18);
            shapeTree5.Append(shape19);
            shapeTree5.Append(shape20);
            shapeTree5.Append(shape21);

            CommonSlideDataExtensionList commonSlideDataExtensionList5 = new CommonSlideDataExtensionList();

            CommonSlideDataExtension commonSlideDataExtension5 = new CommonSlideDataExtension(){ Uri = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}" };

            P14.CreationId creationId5 = new P14.CreationId(){ Val = (UInt32Value)2796734012U };
            creationId5.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            commonSlideDataExtension5.Append(creationId5);

            commonSlideDataExtensionList5.Append(commonSlideDataExtension5);

            commonSlideData5.Append(shapeTree5);
            commonSlideData5.Append(commonSlideDataExtensionList5);

            ColorMapOverride colorMapOverride4 = new ColorMapOverride();
            A.MasterColorMapping masterColorMapping4 = new A.MasterColorMapping();

            colorMapOverride4.Append(masterColorMapping4);

            slideLayout3.Append(commonSlideData5);
            slideLayout3.Append(colorMapOverride4);

            slideLayoutPart3.SlideLayout = slideLayout3;
        }

        // Generates content of slideLayoutPart4.
        private void GenerateSlideLayoutPart4Content(SlideLayoutPart slideLayoutPart4)
        {
            SlideLayout slideLayout4 = new SlideLayout(){ Type = SlideLayoutValues.Blank, Preserve = true };
            slideLayout4.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            slideLayout4.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            slideLayout4.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            CommonSlideData commonSlideData6 = new CommonSlideData(){ Name = "Blank" };

            ShapeTree shapeTree6 = new ShapeTree();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties6 = new NonVisualGroupShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties27 = new NonVisualDrawingProperties(){ Id = (UInt32Value)1U, Name = "" };
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties6 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties27 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties6.Append(nonVisualDrawingProperties27);
            nonVisualGroupShapeProperties6.Append(nonVisualGroupShapeDrawingProperties6);
            nonVisualGroupShapeProperties6.Append(applicationNonVisualDrawingProperties27);

            GroupShapeProperties groupShapeProperties6 = new GroupShapeProperties();

            A.TransformGroup transformGroup6 = new A.TransformGroup();
            A.Offset offset18 = new A.Offset(){ X = 0L, Y = 0L };
            A.Extents extents18 = new A.Extents(){ Cx = 0L, Cy = 0L };
            A.ChildOffset childOffset6 = new A.ChildOffset(){ X = 0L, Y = 0L };
            A.ChildExtents childExtents6 = new A.ChildExtents(){ Cx = 0L, Cy = 0L };

            transformGroup6.Append(offset18);
            transformGroup6.Append(extents18);
            transformGroup6.Append(childOffset6);
            transformGroup6.Append(childExtents6);

            groupShapeProperties6.Append(transformGroup6);

            Shape shape22 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties22 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties28 = new NonVisualDrawingProperties(){ Id = (UInt32Value)2U, Name = "Date Placeholder 1" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList22 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension22 = new A.NonVisualDrawingPropertiesExtension(){ Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            OpenXmlUnknownElement openXmlUnknownElement22 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{37685DD7-06D2-0546-70AA-9C13AB9DB767}\" />");

            nonVisualDrawingPropertiesExtension22.Append(openXmlUnknownElement22);

            nonVisualDrawingPropertiesExtensionList22.Append(nonVisualDrawingPropertiesExtension22);

            nonVisualDrawingProperties28.Append(nonVisualDrawingPropertiesExtensionList22);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties22 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks22 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties22.Append(shapeLocks22);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties28 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape22 = new PlaceholderShape(){ Type = PlaceholderValues.DateAndTime, Size = PlaceholderSizeValues.Half, Index = (UInt32Value)10U };

            applicationNonVisualDrawingProperties28.Append(placeholderShape22);

            nonVisualShapeProperties22.Append(nonVisualDrawingProperties28);
            nonVisualShapeProperties22.Append(nonVisualShapeDrawingProperties22);
            nonVisualShapeProperties22.Append(applicationNonVisualDrawingProperties28);
            ShapeProperties shapeProperties22 = new ShapeProperties();

            TextBody textBody22 = new TextBody();
            A.BodyProperties bodyProperties22 = new A.BodyProperties();
            A.ListStyle listStyle22 = new A.ListStyle();

            A.Paragraph paragraph30 = new A.Paragraph();

            A.Field field9 = new A.Field(){ Id = "{F550F806-175B-4EE1-ABEA-A5F04E4FD72B}", Type = "datetimeFigureOut" };

            A.RunProperties runProperties26 = new A.RunProperties(){ Language = "pl-PL" };
            runProperties26.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text26 = new A.Text();
            text26.Text = "20.08.2025";

            field9.Append(runProperties26);
            field9.Append(text26);
            A.EndParagraphRunProperties endParagraphRunProperties20 = new A.EndParagraphRunProperties(){ Language = "pl-PL" };

            paragraph30.Append(field9);
            paragraph30.Append(endParagraphRunProperties20);

            textBody22.Append(bodyProperties22);
            textBody22.Append(listStyle22);
            textBody22.Append(paragraph30);

            shape22.Append(nonVisualShapeProperties22);
            shape22.Append(shapeProperties22);
            shape22.Append(textBody22);

            Shape shape23 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties23 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties29 = new NonVisualDrawingProperties(){ Id = (UInt32Value)3U, Name = "Footer Placeholder 2" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList23 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension23 = new A.NonVisualDrawingPropertiesExtension(){ Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            OpenXmlUnknownElement openXmlUnknownElement23 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{4C010C04-55EC-89FE-7B95-CABF324B14FA}\" />");

            nonVisualDrawingPropertiesExtension23.Append(openXmlUnknownElement23);

            nonVisualDrawingPropertiesExtensionList23.Append(nonVisualDrawingPropertiesExtension23);

            nonVisualDrawingProperties29.Append(nonVisualDrawingPropertiesExtensionList23);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties23 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks23 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties23.Append(shapeLocks23);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties29 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape23 = new PlaceholderShape(){ Type = PlaceholderValues.Footer, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)11U };

            applicationNonVisualDrawingProperties29.Append(placeholderShape23);

            nonVisualShapeProperties23.Append(nonVisualDrawingProperties29);
            nonVisualShapeProperties23.Append(nonVisualShapeDrawingProperties23);
            nonVisualShapeProperties23.Append(applicationNonVisualDrawingProperties29);
            ShapeProperties shapeProperties23 = new ShapeProperties();

            TextBody textBody23 = new TextBody();
            A.BodyProperties bodyProperties23 = new A.BodyProperties();
            A.ListStyle listStyle23 = new A.ListStyle();

            A.Paragraph paragraph31 = new A.Paragraph();
            A.EndParagraphRunProperties endParagraphRunProperties21 = new A.EndParagraphRunProperties(){ Language = "pl-PL" };

            paragraph31.Append(endParagraphRunProperties21);

            textBody23.Append(bodyProperties23);
            textBody23.Append(listStyle23);
            textBody23.Append(paragraph31);

            shape23.Append(nonVisualShapeProperties23);
            shape23.Append(shapeProperties23);
            shape23.Append(textBody23);

            Shape shape24 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties24 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties30 = new NonVisualDrawingProperties(){ Id = (UInt32Value)4U, Name = "Slide Number Placeholder 3" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList24 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension24 = new A.NonVisualDrawingPropertiesExtension(){ Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            OpenXmlUnknownElement openXmlUnknownElement24 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{336B4F7D-A020-7502-696D-F5F9D466B74B}\" />");

            nonVisualDrawingPropertiesExtension24.Append(openXmlUnknownElement24);

            nonVisualDrawingPropertiesExtensionList24.Append(nonVisualDrawingPropertiesExtension24);

            nonVisualDrawingProperties30.Append(nonVisualDrawingPropertiesExtensionList24);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties24 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks24 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties24.Append(shapeLocks24);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties30 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape24 = new PlaceholderShape(){ Type = PlaceholderValues.SlideNumber, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)12U };

            applicationNonVisualDrawingProperties30.Append(placeholderShape24);

            nonVisualShapeProperties24.Append(nonVisualDrawingProperties30);
            nonVisualShapeProperties24.Append(nonVisualShapeDrawingProperties24);
            nonVisualShapeProperties24.Append(applicationNonVisualDrawingProperties30);
            ShapeProperties shapeProperties24 = new ShapeProperties();

            TextBody textBody24 = new TextBody();
            A.BodyProperties bodyProperties24 = new A.BodyProperties();
            A.ListStyle listStyle24 = new A.ListStyle();

            A.Paragraph paragraph32 = new A.Paragraph();

            A.Field field10 = new A.Field(){ Id = "{5FAB34AF-B09A-4E2A-BAD3-CCD6E0AD81E6}", Type = "slidenum" };

            A.RunProperties runProperties27 = new A.RunProperties(){ Language = "pl-PL" };
            runProperties27.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text27 = new A.Text();
            text27.Text = "‹#›";

            field10.Append(runProperties27);
            field10.Append(text27);
            A.EndParagraphRunProperties endParagraphRunProperties22 = new A.EndParagraphRunProperties(){ Language = "pl-PL" };

            paragraph32.Append(field10);
            paragraph32.Append(endParagraphRunProperties22);

            textBody24.Append(bodyProperties24);
            textBody24.Append(listStyle24);
            textBody24.Append(paragraph32);

            shape24.Append(nonVisualShapeProperties24);
            shape24.Append(shapeProperties24);
            shape24.Append(textBody24);

            shapeTree6.Append(nonVisualGroupShapeProperties6);
            shapeTree6.Append(groupShapeProperties6);
            shapeTree6.Append(shape22);
            shapeTree6.Append(shape23);
            shapeTree6.Append(shape24);

            CommonSlideDataExtensionList commonSlideDataExtensionList6 = new CommonSlideDataExtensionList();

            CommonSlideDataExtension commonSlideDataExtension6 = new CommonSlideDataExtension(){ Uri = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}" };

            P14.CreationId creationId6 = new P14.CreationId(){ Val = (UInt32Value)1065518391U };
            creationId6.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            commonSlideDataExtension6.Append(creationId6);

            commonSlideDataExtensionList6.Append(commonSlideDataExtension6);

            commonSlideData6.Append(shapeTree6);
            commonSlideData6.Append(commonSlideDataExtensionList6);

            ColorMapOverride colorMapOverride5 = new ColorMapOverride();
            A.MasterColorMapping masterColorMapping5 = new A.MasterColorMapping();

            colorMapOverride5.Append(masterColorMapping5);

            slideLayout4.Append(commonSlideData6);
            slideLayout4.Append(colorMapOverride5);

            slideLayoutPart4.SlideLayout = slideLayout4;
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
            A.LatinFont latinFont29 = new A.LatinFont(){ Typeface = "Aptos Display", Panose = "02110004020202020204" };
            A.EastAsianFont eastAsianFont29 = new A.EastAsianFont(){ Typeface = "" };
            A.ComplexScriptFont complexScriptFont29 = new A.ComplexScriptFont(){ Typeface = "" };
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

            majorFont1.Append(latinFont29);
            majorFont1.Append(eastAsianFont29);
            majorFont1.Append(complexScriptFont29);
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
            A.LatinFont latinFont30 = new A.LatinFont(){ Typeface = "Aptos", Panose = "02110004020202020204" };
            A.EastAsianFont eastAsianFont30 = new A.EastAsianFont(){ Typeface = "" };
            A.ComplexScriptFont complexScriptFont30 = new A.ComplexScriptFont(){ Typeface = "" };
            A.SupplementalFont supplementalFont48 = new A.SupplementalFont(){ Script = "Jpan", Typeface = "游ゴシック" };
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

            minorFont1.Append(latinFont30);
            minorFont1.Append(eastAsianFont30);
            minorFont1.Append(complexScriptFont30);
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

            A.SolidFill solidFill41 = new A.SolidFill();
            A.SchemeColor schemeColor42 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };

            solidFill41.Append(schemeColor42);

            A.GradientFill gradientFill1 = new A.GradientFill(){ RotateWithShape = true };

            A.GradientStopList gradientStopList1 = new A.GradientStopList();

            A.GradientStop gradientStop1 = new A.GradientStop(){ Position = 0 };

            A.SchemeColor schemeColor43 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };
            A.LuminanceModulation luminanceModulation1 = new A.LuminanceModulation(){ Val = 110000 };
            A.SaturationModulation saturationModulation1 = new A.SaturationModulation(){ Val = 105000 };
            A.Tint tint13 = new A.Tint(){ Val = 67000 };

            schemeColor43.Append(luminanceModulation1);
            schemeColor43.Append(saturationModulation1);
            schemeColor43.Append(tint13);

            gradientStop1.Append(schemeColor43);

            A.GradientStop gradientStop2 = new A.GradientStop(){ Position = 50000 };

            A.SchemeColor schemeColor44 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };
            A.LuminanceModulation luminanceModulation2 = new A.LuminanceModulation(){ Val = 105000 };
            A.SaturationModulation saturationModulation2 = new A.SaturationModulation(){ Val = 103000 };
            A.Tint tint14 = new A.Tint(){ Val = 73000 };

            schemeColor44.Append(luminanceModulation2);
            schemeColor44.Append(saturationModulation2);
            schemeColor44.Append(tint14);

            gradientStop2.Append(schemeColor44);

            A.GradientStop gradientStop3 = new A.GradientStop(){ Position = 100000 };

            A.SchemeColor schemeColor45 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };
            A.LuminanceModulation luminanceModulation3 = new A.LuminanceModulation(){ Val = 105000 };
            A.SaturationModulation saturationModulation3 = new A.SaturationModulation(){ Val = 109000 };
            A.Tint tint15 = new A.Tint(){ Val = 81000 };

            schemeColor45.Append(luminanceModulation3);
            schemeColor45.Append(saturationModulation3);
            schemeColor45.Append(tint15);

            gradientStop3.Append(schemeColor45);

            gradientStopList1.Append(gradientStop1);
            gradientStopList1.Append(gradientStop2);
            gradientStopList1.Append(gradientStop3);
            A.LinearGradientFill linearGradientFill1 = new A.LinearGradientFill(){ Angle = 5400000, Scaled = false };

            gradientFill1.Append(gradientStopList1);
            gradientFill1.Append(linearGradientFill1);

            A.GradientFill gradientFill2 = new A.GradientFill(){ RotateWithShape = true };

            A.GradientStopList gradientStopList2 = new A.GradientStopList();

            A.GradientStop gradientStop4 = new A.GradientStop(){ Position = 0 };

            A.SchemeColor schemeColor46 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };
            A.SaturationModulation saturationModulation4 = new A.SaturationModulation(){ Val = 103000 };
            A.LuminanceModulation luminanceModulation4 = new A.LuminanceModulation(){ Val = 102000 };
            A.Tint tint16 = new A.Tint(){ Val = 94000 };

            schemeColor46.Append(saturationModulation4);
            schemeColor46.Append(luminanceModulation4);
            schemeColor46.Append(tint16);

            gradientStop4.Append(schemeColor46);

            A.GradientStop gradientStop5 = new A.GradientStop(){ Position = 50000 };

            A.SchemeColor schemeColor47 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };
            A.SaturationModulation saturationModulation5 = new A.SaturationModulation(){ Val = 110000 };
            A.LuminanceModulation luminanceModulation5 = new A.LuminanceModulation(){ Val = 100000 };
            A.Shade shade1 = new A.Shade(){ Val = 100000 };

            schemeColor47.Append(saturationModulation5);
            schemeColor47.Append(luminanceModulation5);
            schemeColor47.Append(shade1);

            gradientStop5.Append(schemeColor47);

            A.GradientStop gradientStop6 = new A.GradientStop(){ Position = 100000 };

            A.SchemeColor schemeColor48 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };
            A.LuminanceModulation luminanceModulation6 = new A.LuminanceModulation(){ Val = 99000 };
            A.SaturationModulation saturationModulation6 = new A.SaturationModulation(){ Val = 120000 };
            A.Shade shade2 = new A.Shade(){ Val = 78000 };

            schemeColor48.Append(luminanceModulation6);
            schemeColor48.Append(saturationModulation6);
            schemeColor48.Append(shade2);

            gradientStop6.Append(schemeColor48);

            gradientStopList2.Append(gradientStop4);
            gradientStopList2.Append(gradientStop5);
            gradientStopList2.Append(gradientStop6);
            A.LinearGradientFill linearGradientFill2 = new A.LinearGradientFill(){ Angle = 5400000, Scaled = false };

            gradientFill2.Append(gradientStopList2);
            gradientFill2.Append(linearGradientFill2);

            fillStyleList1.Append(solidFill41);
            fillStyleList1.Append(gradientFill1);
            fillStyleList1.Append(gradientFill2);

            A.LineStyleList lineStyleList1 = new A.LineStyleList();

            A.Outline outline1 = new A.Outline(){ Width = 12700, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill42 = new A.SolidFill();
            A.SchemeColor schemeColor49 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };

            solidFill42.Append(schemeColor49);
            A.PresetDash presetDash1 = new A.PresetDash(){ Val = A.PresetLineDashValues.Solid };
            A.Miter miter1 = new A.Miter(){ Limit = 800000 };

            outline1.Append(solidFill42);
            outline1.Append(presetDash1);
            outline1.Append(miter1);

            A.Outline outline2 = new A.Outline(){ Width = 19050, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill43 = new A.SolidFill();
            A.SchemeColor schemeColor50 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };

            solidFill43.Append(schemeColor50);
            A.PresetDash presetDash2 = new A.PresetDash(){ Val = A.PresetLineDashValues.Solid };
            A.Miter miter2 = new A.Miter(){ Limit = 800000 };

            outline2.Append(solidFill43);
            outline2.Append(presetDash2);
            outline2.Append(miter2);

            A.Outline outline3 = new A.Outline(){ Width = 25400, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill44 = new A.SolidFill();
            A.SchemeColor schemeColor51 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };

            solidFill44.Append(schemeColor51);
            A.PresetDash presetDash3 = new A.PresetDash(){ Val = A.PresetLineDashValues.Solid };
            A.Miter miter3 = new A.Miter(){ Limit = 800000 };

            outline3.Append(solidFill44);
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

            A.SolidFill solidFill45 = new A.SolidFill();
            A.SchemeColor schemeColor52 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };

            solidFill45.Append(schemeColor52);

            A.SolidFill solidFill46 = new A.SolidFill();

            A.SchemeColor schemeColor53 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };
            A.Tint tint17 = new A.Tint(){ Val = 95000 };
            A.SaturationModulation saturationModulation7 = new A.SaturationModulation(){ Val = 170000 };

            schemeColor53.Append(tint17);
            schemeColor53.Append(saturationModulation7);

            solidFill46.Append(schemeColor53);

            A.GradientFill gradientFill3 = new A.GradientFill(){ RotateWithShape = true };

            A.GradientStopList gradientStopList3 = new A.GradientStopList();

            A.GradientStop gradientStop7 = new A.GradientStop(){ Position = 0 };

            A.SchemeColor schemeColor54 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };
            A.Tint tint18 = new A.Tint(){ Val = 93000 };
            A.SaturationModulation saturationModulation8 = new A.SaturationModulation(){ Val = 150000 };
            A.Shade shade3 = new A.Shade(){ Val = 98000 };
            A.LuminanceModulation luminanceModulation7 = new A.LuminanceModulation(){ Val = 102000 };

            schemeColor54.Append(tint18);
            schemeColor54.Append(saturationModulation8);
            schemeColor54.Append(shade3);
            schemeColor54.Append(luminanceModulation7);

            gradientStop7.Append(schemeColor54);

            A.GradientStop gradientStop8 = new A.GradientStop(){ Position = 50000 };

            A.SchemeColor schemeColor55 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };
            A.Tint tint19 = new A.Tint(){ Val = 98000 };
            A.SaturationModulation saturationModulation9 = new A.SaturationModulation(){ Val = 130000 };
            A.Shade shade4 = new A.Shade(){ Val = 90000 };
            A.LuminanceModulation luminanceModulation8 = new A.LuminanceModulation(){ Val = 103000 };

            schemeColor55.Append(tint19);
            schemeColor55.Append(saturationModulation9);
            schemeColor55.Append(shade4);
            schemeColor55.Append(luminanceModulation8);

            gradientStop8.Append(schemeColor55);

            A.GradientStop gradientStop9 = new A.GradientStop(){ Position = 100000 };

            A.SchemeColor schemeColor56 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };
            A.Shade shade5 = new A.Shade(){ Val = 63000 };
            A.SaturationModulation saturationModulation10 = new A.SaturationModulation(){ Val = 120000 };

            schemeColor56.Append(shade5);
            schemeColor56.Append(saturationModulation10);

            gradientStop9.Append(schemeColor56);

            gradientStopList3.Append(gradientStop7);
            gradientStopList3.Append(gradientStop8);
            gradientStopList3.Append(gradientStop9);
            A.LinearGradientFill linearGradientFill3 = new A.LinearGradientFill(){ Angle = 5400000, Scaled = false };

            gradientFill3.Append(gradientStopList3);
            gradientFill3.Append(linearGradientFill3);

            backgroundFillStyleList1.Append(solidFill45);
            backgroundFillStyleList1.Append(solidFill46);
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
            A.ShapeProperties shapeProperties25 = new A.ShapeProperties();
            A.BodyProperties bodyProperties25 = new A.BodyProperties();
            A.ListStyle listStyle25 = new A.ListStyle();

            A.ShapeStyle shapeStyle1 = new A.ShapeStyle();

            A.LineReference lineReference1 = new A.LineReference(){ Index = (UInt32Value)2U };
            A.SchemeColor schemeColor57 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent1 };

            lineReference1.Append(schemeColor57);

            A.FillReference fillReference1 = new A.FillReference(){ Index = (UInt32Value)0U };
            A.SchemeColor schemeColor58 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent1 };

            fillReference1.Append(schemeColor58);

            A.EffectReference effectReference1 = new A.EffectReference(){ Index = (UInt32Value)1U };
            A.SchemeColor schemeColor59 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent1 };

            effectReference1.Append(schemeColor59);

            A.FontReference fontReference1 = new A.FontReference(){ Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor60 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            fontReference1.Append(schemeColor60);

            shapeStyle1.Append(lineReference1);
            shapeStyle1.Append(fillReference1);
            shapeStyle1.Append(effectReference1);
            shapeStyle1.Append(fontReference1);

            lineDefault1.Append(shapeProperties25);
            lineDefault1.Append(bodyProperties25);
            lineDefault1.Append(listStyle25);
            lineDefault1.Append(shapeStyle1);

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

        // Generates content of slideLayoutPart5.
        private void GenerateSlideLayoutPart5Content(SlideLayoutPart slideLayoutPart5)
        {
            SlideLayout slideLayout5 = new SlideLayout(){ Type = SlideLayoutValues.Object, Preserve = true };
            slideLayout5.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            slideLayout5.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            slideLayout5.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            CommonSlideData commonSlideData7 = new CommonSlideData(){ Name = "Title and Content" };

            ShapeTree shapeTree7 = new ShapeTree();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties7 = new NonVisualGroupShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties31 = new NonVisualDrawingProperties(){ Id = (UInt32Value)1U, Name = "" };
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties7 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties31 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties7.Append(nonVisualDrawingProperties31);
            nonVisualGroupShapeProperties7.Append(nonVisualGroupShapeDrawingProperties7);
            nonVisualGroupShapeProperties7.Append(applicationNonVisualDrawingProperties31);

            GroupShapeProperties groupShapeProperties7 = new GroupShapeProperties();

            A.TransformGroup transformGroup7 = new A.TransformGroup();
            A.Offset offset19 = new A.Offset(){ X = 0L, Y = 0L };
            A.Extents extents19 = new A.Extents(){ Cx = 0L, Cy = 0L };
            A.ChildOffset childOffset7 = new A.ChildOffset(){ X = 0L, Y = 0L };
            A.ChildExtents childExtents7 = new A.ChildExtents(){ Cx = 0L, Cy = 0L };

            transformGroup7.Append(offset19);
            transformGroup7.Append(extents19);
            transformGroup7.Append(childOffset7);
            transformGroup7.Append(childExtents7);

            groupShapeProperties7.Append(transformGroup7);

            Shape shape25 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties25 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties32 = new NonVisualDrawingProperties(){ Id = (UInt32Value)2U, Name = "Title 1" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList25 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension25 = new A.NonVisualDrawingPropertiesExtension(){ Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            OpenXmlUnknownElement openXmlUnknownElement25 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{97C3C552-C929-FA6A-1A07-D563D3C96C31}\" />");

            nonVisualDrawingPropertiesExtension25.Append(openXmlUnknownElement25);

            nonVisualDrawingPropertiesExtensionList25.Append(nonVisualDrawingPropertiesExtension25);

            nonVisualDrawingProperties32.Append(nonVisualDrawingPropertiesExtensionList25);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties25 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks25 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties25.Append(shapeLocks25);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties32 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape25 = new PlaceholderShape(){ Type = PlaceholderValues.Title };

            applicationNonVisualDrawingProperties32.Append(placeholderShape25);

            nonVisualShapeProperties25.Append(nonVisualDrawingProperties32);
            nonVisualShapeProperties25.Append(nonVisualShapeDrawingProperties25);
            nonVisualShapeProperties25.Append(applicationNonVisualDrawingProperties32);
            ShapeProperties shapeProperties26 = new ShapeProperties();

            TextBody textBody25 = new TextBody();
            A.BodyProperties bodyProperties26 = new A.BodyProperties();
            A.ListStyle listStyle26 = new A.ListStyle();

            A.Paragraph paragraph33 = new A.Paragraph();

            A.Run run18 = new A.Run();
            A.RunProperties runProperties28 = new A.RunProperties(){ Language = "en-US" };
            A.Text text28 = new A.Text();
            text28.Text = "Click to edit Master title style";

            run18.Append(runProperties28);
            run18.Append(text28);
            A.EndParagraphRunProperties endParagraphRunProperties23 = new A.EndParagraphRunProperties(){ Language = "pl-PL" };

            paragraph33.Append(run18);
            paragraph33.Append(endParagraphRunProperties23);

            textBody25.Append(bodyProperties26);
            textBody25.Append(listStyle26);
            textBody25.Append(paragraph33);

            shape25.Append(nonVisualShapeProperties25);
            shape25.Append(shapeProperties26);
            shape25.Append(textBody25);

            Shape shape26 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties26 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties33 = new NonVisualDrawingProperties(){ Id = (UInt32Value)3U, Name = "Content Placeholder 2" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList26 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension26 = new A.NonVisualDrawingPropertiesExtension(){ Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            OpenXmlUnknownElement openXmlUnknownElement26 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{79AD22D5-35D7-F916-60EB-0E484F7F7AEC}\" />");

            nonVisualDrawingPropertiesExtension26.Append(openXmlUnknownElement26);

            nonVisualDrawingPropertiesExtensionList26.Append(nonVisualDrawingPropertiesExtension26);

            nonVisualDrawingProperties33.Append(nonVisualDrawingPropertiesExtensionList26);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties26 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks26 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties26.Append(shapeLocks26);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties33 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape26 = new PlaceholderShape(){ Index = (UInt32Value)1U };

            applicationNonVisualDrawingProperties33.Append(placeholderShape26);

            nonVisualShapeProperties26.Append(nonVisualDrawingProperties33);
            nonVisualShapeProperties26.Append(nonVisualShapeDrawingProperties26);
            nonVisualShapeProperties26.Append(applicationNonVisualDrawingProperties33);
            ShapeProperties shapeProperties27 = new ShapeProperties();

            TextBody textBody26 = new TextBody();
            A.BodyProperties bodyProperties27 = new A.BodyProperties();
            A.ListStyle listStyle27 = new A.ListStyle();

            A.Paragraph paragraph34 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties13 = new A.ParagraphProperties(){ Level = 0 };

            A.Run run19 = new A.Run();
            A.RunProperties runProperties29 = new A.RunProperties(){ Language = "en-US" };
            A.Text text29 = new A.Text();
            text29.Text = "Click to edit Master text styles";

            run19.Append(runProperties29);
            run19.Append(text29);

            paragraph34.Append(paragraphProperties13);
            paragraph34.Append(run19);

            A.Paragraph paragraph35 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties14 = new A.ParagraphProperties(){ Level = 1 };

            A.Run run20 = new A.Run();
            A.RunProperties runProperties30 = new A.RunProperties(){ Language = "en-US" };
            A.Text text30 = new A.Text();
            text30.Text = "Second level";

            run20.Append(runProperties30);
            run20.Append(text30);

            paragraph35.Append(paragraphProperties14);
            paragraph35.Append(run20);

            A.Paragraph paragraph36 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties15 = new A.ParagraphProperties(){ Level = 2 };

            A.Run run21 = new A.Run();
            A.RunProperties runProperties31 = new A.RunProperties(){ Language = "en-US" };
            A.Text text31 = new A.Text();
            text31.Text = "Third level";

            run21.Append(runProperties31);
            run21.Append(text31);

            paragraph36.Append(paragraphProperties15);
            paragraph36.Append(run21);

            A.Paragraph paragraph37 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties16 = new A.ParagraphProperties(){ Level = 3 };

            A.Run run22 = new A.Run();
            A.RunProperties runProperties32 = new A.RunProperties(){ Language = "en-US" };
            A.Text text32 = new A.Text();
            text32.Text = "Fourth level";

            run22.Append(runProperties32);
            run22.Append(text32);

            paragraph37.Append(paragraphProperties16);
            paragraph37.Append(run22);

            A.Paragraph paragraph38 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties17 = new A.ParagraphProperties(){ Level = 4 };

            A.Run run23 = new A.Run();
            A.RunProperties runProperties33 = new A.RunProperties(){ Language = "en-US" };
            A.Text text33 = new A.Text();
            text33.Text = "Fifth level";

            run23.Append(runProperties33);
            run23.Append(text33);
            A.EndParagraphRunProperties endParagraphRunProperties24 = new A.EndParagraphRunProperties(){ Language = "pl-PL" };

            paragraph38.Append(paragraphProperties17);
            paragraph38.Append(run23);
            paragraph38.Append(endParagraphRunProperties24);

            textBody26.Append(bodyProperties27);
            textBody26.Append(listStyle27);
            textBody26.Append(paragraph34);
            textBody26.Append(paragraph35);
            textBody26.Append(paragraph36);
            textBody26.Append(paragraph37);
            textBody26.Append(paragraph38);

            shape26.Append(nonVisualShapeProperties26);
            shape26.Append(shapeProperties27);
            shape26.Append(textBody26);

            Shape shape27 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties27 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties34 = new NonVisualDrawingProperties(){ Id = (UInt32Value)4U, Name = "Date Placeholder 3" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList27 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension27 = new A.NonVisualDrawingPropertiesExtension(){ Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            OpenXmlUnknownElement openXmlUnknownElement27 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{B87665A7-B3FD-88BE-92D6-5884B03B8408}\" />");

            nonVisualDrawingPropertiesExtension27.Append(openXmlUnknownElement27);

            nonVisualDrawingPropertiesExtensionList27.Append(nonVisualDrawingPropertiesExtension27);

            nonVisualDrawingProperties34.Append(nonVisualDrawingPropertiesExtensionList27);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties27 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks27 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties27.Append(shapeLocks27);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties34 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape27 = new PlaceholderShape(){ Type = PlaceholderValues.DateAndTime, Size = PlaceholderSizeValues.Half, Index = (UInt32Value)10U };

            applicationNonVisualDrawingProperties34.Append(placeholderShape27);

            nonVisualShapeProperties27.Append(nonVisualDrawingProperties34);
            nonVisualShapeProperties27.Append(nonVisualShapeDrawingProperties27);
            nonVisualShapeProperties27.Append(applicationNonVisualDrawingProperties34);
            ShapeProperties shapeProperties28 = new ShapeProperties();

            TextBody textBody27 = new TextBody();
            A.BodyProperties bodyProperties28 = new A.BodyProperties();
            A.ListStyle listStyle28 = new A.ListStyle();

            A.Paragraph paragraph39 = new A.Paragraph();

            A.Field field11 = new A.Field(){ Id = "{F550F806-175B-4EE1-ABEA-A5F04E4FD72B}", Type = "datetimeFigureOut" };

            A.RunProperties runProperties34 = new A.RunProperties(){ Language = "pl-PL" };
            runProperties34.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text34 = new A.Text();
            text34.Text = "20.08.2025";

            field11.Append(runProperties34);
            field11.Append(text34);
            A.EndParagraphRunProperties endParagraphRunProperties25 = new A.EndParagraphRunProperties(){ Language = "pl-PL" };

            paragraph39.Append(field11);
            paragraph39.Append(endParagraphRunProperties25);

            textBody27.Append(bodyProperties28);
            textBody27.Append(listStyle28);
            textBody27.Append(paragraph39);

            shape27.Append(nonVisualShapeProperties27);
            shape27.Append(shapeProperties28);
            shape27.Append(textBody27);

            Shape shape28 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties28 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties35 = new NonVisualDrawingProperties(){ Id = (UInt32Value)5U, Name = "Footer Placeholder 4" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList28 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension28 = new A.NonVisualDrawingPropertiesExtension(){ Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            OpenXmlUnknownElement openXmlUnknownElement28 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{10DB848D-1DBB-F7C5-72D0-DE776AA72E35}\" />");

            nonVisualDrawingPropertiesExtension28.Append(openXmlUnknownElement28);

            nonVisualDrawingPropertiesExtensionList28.Append(nonVisualDrawingPropertiesExtension28);

            nonVisualDrawingProperties35.Append(nonVisualDrawingPropertiesExtensionList28);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties28 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks28 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties28.Append(shapeLocks28);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties35 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape28 = new PlaceholderShape(){ Type = PlaceholderValues.Footer, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)11U };

            applicationNonVisualDrawingProperties35.Append(placeholderShape28);

            nonVisualShapeProperties28.Append(nonVisualDrawingProperties35);
            nonVisualShapeProperties28.Append(nonVisualShapeDrawingProperties28);
            nonVisualShapeProperties28.Append(applicationNonVisualDrawingProperties35);
            ShapeProperties shapeProperties29 = new ShapeProperties();

            TextBody textBody28 = new TextBody();
            A.BodyProperties bodyProperties29 = new A.BodyProperties();
            A.ListStyle listStyle29 = new A.ListStyle();

            A.Paragraph paragraph40 = new A.Paragraph();
            A.EndParagraphRunProperties endParagraphRunProperties26 = new A.EndParagraphRunProperties(){ Language = "pl-PL" };

            paragraph40.Append(endParagraphRunProperties26);

            textBody28.Append(bodyProperties29);
            textBody28.Append(listStyle29);
            textBody28.Append(paragraph40);

            shape28.Append(nonVisualShapeProperties28);
            shape28.Append(shapeProperties29);
            shape28.Append(textBody28);

            Shape shape29 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties29 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties36 = new NonVisualDrawingProperties(){ Id = (UInt32Value)6U, Name = "Slide Number Placeholder 5" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList29 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension29 = new A.NonVisualDrawingPropertiesExtension(){ Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            OpenXmlUnknownElement openXmlUnknownElement29 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{2D85A33A-A0AB-ED74-2677-03175D123739}\" />");

            nonVisualDrawingPropertiesExtension29.Append(openXmlUnknownElement29);

            nonVisualDrawingPropertiesExtensionList29.Append(nonVisualDrawingPropertiesExtension29);

            nonVisualDrawingProperties36.Append(nonVisualDrawingPropertiesExtensionList29);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties29 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks29 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties29.Append(shapeLocks29);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties36 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape29 = new PlaceholderShape(){ Type = PlaceholderValues.SlideNumber, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)12U };

            applicationNonVisualDrawingProperties36.Append(placeholderShape29);

            nonVisualShapeProperties29.Append(nonVisualDrawingProperties36);
            nonVisualShapeProperties29.Append(nonVisualShapeDrawingProperties29);
            nonVisualShapeProperties29.Append(applicationNonVisualDrawingProperties36);
            ShapeProperties shapeProperties30 = new ShapeProperties();

            TextBody textBody29 = new TextBody();
            A.BodyProperties bodyProperties30 = new A.BodyProperties();
            A.ListStyle listStyle30 = new A.ListStyle();

            A.Paragraph paragraph41 = new A.Paragraph();

            A.Field field12 = new A.Field(){ Id = "{5FAB34AF-B09A-4E2A-BAD3-CCD6E0AD81E6}", Type = "slidenum" };

            A.RunProperties runProperties35 = new A.RunProperties(){ Language = "pl-PL" };
            runProperties35.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text35 = new A.Text();
            text35.Text = "‹#›";

            field12.Append(runProperties35);
            field12.Append(text35);
            A.EndParagraphRunProperties endParagraphRunProperties27 = new A.EndParagraphRunProperties(){ Language = "pl-PL" };

            paragraph41.Append(field12);
            paragraph41.Append(endParagraphRunProperties27);

            textBody29.Append(bodyProperties30);
            textBody29.Append(listStyle30);
            textBody29.Append(paragraph41);

            shape29.Append(nonVisualShapeProperties29);
            shape29.Append(shapeProperties30);
            shape29.Append(textBody29);

            shapeTree7.Append(nonVisualGroupShapeProperties7);
            shapeTree7.Append(groupShapeProperties7);
            shapeTree7.Append(shape25);
            shapeTree7.Append(shape26);
            shapeTree7.Append(shape27);
            shapeTree7.Append(shape28);
            shapeTree7.Append(shape29);

            CommonSlideDataExtensionList commonSlideDataExtensionList7 = new CommonSlideDataExtensionList();

            CommonSlideDataExtension commonSlideDataExtension7 = new CommonSlideDataExtension(){ Uri = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}" };

            P14.CreationId creationId7 = new P14.CreationId(){ Val = (UInt32Value)3348220065U };
            creationId7.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            commonSlideDataExtension7.Append(creationId7);

            commonSlideDataExtensionList7.Append(commonSlideDataExtension7);

            commonSlideData7.Append(shapeTree7);
            commonSlideData7.Append(commonSlideDataExtensionList7);

            ColorMapOverride colorMapOverride6 = new ColorMapOverride();
            A.MasterColorMapping masterColorMapping6 = new A.MasterColorMapping();

            colorMapOverride6.Append(masterColorMapping6);

            slideLayout5.Append(commonSlideData7);
            slideLayout5.Append(colorMapOverride6);

            slideLayoutPart5.SlideLayout = slideLayout5;
        }

        // Generates content of slideLayoutPart6.
        private void GenerateSlideLayoutPart6Content(SlideLayoutPart slideLayoutPart6)
        {
            SlideLayout slideLayout6 = new SlideLayout(){ Type = SlideLayoutValues.TitleOnly, Preserve = true };
            slideLayout6.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            slideLayout6.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            slideLayout6.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            CommonSlideData commonSlideData8 = new CommonSlideData(){ Name = "Title Only" };

            ShapeTree shapeTree8 = new ShapeTree();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties8 = new NonVisualGroupShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties37 = new NonVisualDrawingProperties(){ Id = (UInt32Value)1U, Name = "" };
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties8 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties37 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties8.Append(nonVisualDrawingProperties37);
            nonVisualGroupShapeProperties8.Append(nonVisualGroupShapeDrawingProperties8);
            nonVisualGroupShapeProperties8.Append(applicationNonVisualDrawingProperties37);

            GroupShapeProperties groupShapeProperties8 = new GroupShapeProperties();

            A.TransformGroup transformGroup8 = new A.TransformGroup();
            A.Offset offset20 = new A.Offset(){ X = 0L, Y = 0L };
            A.Extents extents20 = new A.Extents(){ Cx = 0L, Cy = 0L };
            A.ChildOffset childOffset8 = new A.ChildOffset(){ X = 0L, Y = 0L };
            A.ChildExtents childExtents8 = new A.ChildExtents(){ Cx = 0L, Cy = 0L };

            transformGroup8.Append(offset20);
            transformGroup8.Append(extents20);
            transformGroup8.Append(childOffset8);
            transformGroup8.Append(childExtents8);

            groupShapeProperties8.Append(transformGroup8);

            Shape shape30 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties30 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties38 = new NonVisualDrawingProperties(){ Id = (UInt32Value)2U, Name = "Title 1" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList30 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension30 = new A.NonVisualDrawingPropertiesExtension(){ Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            OpenXmlUnknownElement openXmlUnknownElement30 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{FCA93AF8-F93E-99FA-0D8E-C9A46013178E}\" />");

            nonVisualDrawingPropertiesExtension30.Append(openXmlUnknownElement30);

            nonVisualDrawingPropertiesExtensionList30.Append(nonVisualDrawingPropertiesExtension30);

            nonVisualDrawingProperties38.Append(nonVisualDrawingPropertiesExtensionList30);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties30 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks30 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties30.Append(shapeLocks30);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties38 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape30 = new PlaceholderShape(){ Type = PlaceholderValues.Title };

            applicationNonVisualDrawingProperties38.Append(placeholderShape30);

            nonVisualShapeProperties30.Append(nonVisualDrawingProperties38);
            nonVisualShapeProperties30.Append(nonVisualShapeDrawingProperties30);
            nonVisualShapeProperties30.Append(applicationNonVisualDrawingProperties38);
            ShapeProperties shapeProperties31 = new ShapeProperties();

            TextBody textBody30 = new TextBody();
            A.BodyProperties bodyProperties31 = new A.BodyProperties();
            A.ListStyle listStyle31 = new A.ListStyle();

            A.Paragraph paragraph42 = new A.Paragraph();

            A.Run run24 = new A.Run();
            A.RunProperties runProperties36 = new A.RunProperties(){ Language = "en-US" };
            A.Text text36 = new A.Text();
            text36.Text = "Click to edit Master title style";

            run24.Append(runProperties36);
            run24.Append(text36);
            A.EndParagraphRunProperties endParagraphRunProperties28 = new A.EndParagraphRunProperties(){ Language = "pl-PL" };

            paragraph42.Append(run24);
            paragraph42.Append(endParagraphRunProperties28);

            textBody30.Append(bodyProperties31);
            textBody30.Append(listStyle31);
            textBody30.Append(paragraph42);

            shape30.Append(nonVisualShapeProperties30);
            shape30.Append(shapeProperties31);
            shape30.Append(textBody30);

            Shape shape31 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties31 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties39 = new NonVisualDrawingProperties(){ Id = (UInt32Value)3U, Name = "Date Placeholder 2" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList31 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension31 = new A.NonVisualDrawingPropertiesExtension(){ Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            OpenXmlUnknownElement openXmlUnknownElement31 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{A37AE221-4003-4E40-3C7E-FA7E63D4E399}\" />");

            nonVisualDrawingPropertiesExtension31.Append(openXmlUnknownElement31);

            nonVisualDrawingPropertiesExtensionList31.Append(nonVisualDrawingPropertiesExtension31);

            nonVisualDrawingProperties39.Append(nonVisualDrawingPropertiesExtensionList31);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties31 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks31 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties31.Append(shapeLocks31);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties39 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape31 = new PlaceholderShape(){ Type = PlaceholderValues.DateAndTime, Size = PlaceholderSizeValues.Half, Index = (UInt32Value)10U };

            applicationNonVisualDrawingProperties39.Append(placeholderShape31);

            nonVisualShapeProperties31.Append(nonVisualDrawingProperties39);
            nonVisualShapeProperties31.Append(nonVisualShapeDrawingProperties31);
            nonVisualShapeProperties31.Append(applicationNonVisualDrawingProperties39);
            ShapeProperties shapeProperties32 = new ShapeProperties();

            TextBody textBody31 = new TextBody();
            A.BodyProperties bodyProperties32 = new A.BodyProperties();
            A.ListStyle listStyle32 = new A.ListStyle();

            A.Paragraph paragraph43 = new A.Paragraph();

            A.Field field13 = new A.Field(){ Id = "{F550F806-175B-4EE1-ABEA-A5F04E4FD72B}", Type = "datetimeFigureOut" };

            A.RunProperties runProperties37 = new A.RunProperties(){ Language = "pl-PL" };
            runProperties37.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text37 = new A.Text();
            text37.Text = "20.08.2025";

            field13.Append(runProperties37);
            field13.Append(text37);
            A.EndParagraphRunProperties endParagraphRunProperties29 = new A.EndParagraphRunProperties(){ Language = "pl-PL" };

            paragraph43.Append(field13);
            paragraph43.Append(endParagraphRunProperties29);

            textBody31.Append(bodyProperties32);
            textBody31.Append(listStyle32);
            textBody31.Append(paragraph43);

            shape31.Append(nonVisualShapeProperties31);
            shape31.Append(shapeProperties32);
            shape31.Append(textBody31);

            Shape shape32 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties32 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties40 = new NonVisualDrawingProperties(){ Id = (UInt32Value)4U, Name = "Footer Placeholder 3" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList32 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension32 = new A.NonVisualDrawingPropertiesExtension(){ Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            OpenXmlUnknownElement openXmlUnknownElement32 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{AB5DA4FE-36DC-6B9B-2ACC-05D1ABFFBB53}\" />");

            nonVisualDrawingPropertiesExtension32.Append(openXmlUnknownElement32);

            nonVisualDrawingPropertiesExtensionList32.Append(nonVisualDrawingPropertiesExtension32);

            nonVisualDrawingProperties40.Append(nonVisualDrawingPropertiesExtensionList32);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties32 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks32 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties32.Append(shapeLocks32);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties40 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape32 = new PlaceholderShape(){ Type = PlaceholderValues.Footer, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)11U };

            applicationNonVisualDrawingProperties40.Append(placeholderShape32);

            nonVisualShapeProperties32.Append(nonVisualDrawingProperties40);
            nonVisualShapeProperties32.Append(nonVisualShapeDrawingProperties32);
            nonVisualShapeProperties32.Append(applicationNonVisualDrawingProperties40);
            ShapeProperties shapeProperties33 = new ShapeProperties();

            TextBody textBody32 = new TextBody();
            A.BodyProperties bodyProperties33 = new A.BodyProperties();
            A.ListStyle listStyle33 = new A.ListStyle();

            A.Paragraph paragraph44 = new A.Paragraph();
            A.EndParagraphRunProperties endParagraphRunProperties30 = new A.EndParagraphRunProperties(){ Language = "pl-PL" };

            paragraph44.Append(endParagraphRunProperties30);

            textBody32.Append(bodyProperties33);
            textBody32.Append(listStyle33);
            textBody32.Append(paragraph44);

            shape32.Append(nonVisualShapeProperties32);
            shape32.Append(shapeProperties33);
            shape32.Append(textBody32);

            Shape shape33 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties33 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties41 = new NonVisualDrawingProperties(){ Id = (UInt32Value)5U, Name = "Slide Number Placeholder 4" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList33 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension33 = new A.NonVisualDrawingPropertiesExtension(){ Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            OpenXmlUnknownElement openXmlUnknownElement33 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{32A33B40-1DBB-ECA9-F2C0-335BA10A18CA}\" />");

            nonVisualDrawingPropertiesExtension33.Append(openXmlUnknownElement33);

            nonVisualDrawingPropertiesExtensionList33.Append(nonVisualDrawingPropertiesExtension33);

            nonVisualDrawingProperties41.Append(nonVisualDrawingPropertiesExtensionList33);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties33 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks33 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties33.Append(shapeLocks33);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties41 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape33 = new PlaceholderShape(){ Type = PlaceholderValues.SlideNumber, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)12U };

            applicationNonVisualDrawingProperties41.Append(placeholderShape33);

            nonVisualShapeProperties33.Append(nonVisualDrawingProperties41);
            nonVisualShapeProperties33.Append(nonVisualShapeDrawingProperties33);
            nonVisualShapeProperties33.Append(applicationNonVisualDrawingProperties41);
            ShapeProperties shapeProperties34 = new ShapeProperties();

            TextBody textBody33 = new TextBody();
            A.BodyProperties bodyProperties34 = new A.BodyProperties();
            A.ListStyle listStyle34 = new A.ListStyle();

            A.Paragraph paragraph45 = new A.Paragraph();

            A.Field field14 = new A.Field(){ Id = "{5FAB34AF-B09A-4E2A-BAD3-CCD6E0AD81E6}", Type = "slidenum" };

            A.RunProperties runProperties38 = new A.RunProperties(){ Language = "pl-PL" };
            runProperties38.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text38 = new A.Text();
            text38.Text = "‹#›";

            field14.Append(runProperties38);
            field14.Append(text38);
            A.EndParagraphRunProperties endParagraphRunProperties31 = new A.EndParagraphRunProperties(){ Language = "pl-PL" };

            paragraph45.Append(field14);
            paragraph45.Append(endParagraphRunProperties31);

            textBody33.Append(bodyProperties34);
            textBody33.Append(listStyle34);
            textBody33.Append(paragraph45);

            shape33.Append(nonVisualShapeProperties33);
            shape33.Append(shapeProperties34);
            shape33.Append(textBody33);

            shapeTree8.Append(nonVisualGroupShapeProperties8);
            shapeTree8.Append(groupShapeProperties8);
            shapeTree8.Append(shape30);
            shapeTree8.Append(shape31);
            shapeTree8.Append(shape32);
            shapeTree8.Append(shape33);

            CommonSlideDataExtensionList commonSlideDataExtensionList8 = new CommonSlideDataExtensionList();

            CommonSlideDataExtension commonSlideDataExtension8 = new CommonSlideDataExtension(){ Uri = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}" };

            P14.CreationId creationId8 = new P14.CreationId(){ Val = (UInt32Value)246031094U };
            creationId8.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            commonSlideDataExtension8.Append(creationId8);

            commonSlideDataExtensionList8.Append(commonSlideDataExtension8);

            commonSlideData8.Append(shapeTree8);
            commonSlideData8.Append(commonSlideDataExtensionList8);

            ColorMapOverride colorMapOverride7 = new ColorMapOverride();
            A.MasterColorMapping masterColorMapping7 = new A.MasterColorMapping();

            colorMapOverride7.Append(masterColorMapping7);

            slideLayout6.Append(commonSlideData8);
            slideLayout6.Append(colorMapOverride7);

            slideLayoutPart6.SlideLayout = slideLayout6;
        }

        // Generates content of slideLayoutPart7.
        private void GenerateSlideLayoutPart7Content(SlideLayoutPart slideLayoutPart7)
        {
            SlideLayout slideLayout7 = new SlideLayout(){ Type = SlideLayoutValues.VerticalTitleAndText, Preserve = true };
            slideLayout7.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            slideLayout7.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            slideLayout7.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            CommonSlideData commonSlideData9 = new CommonSlideData(){ Name = "Vertical Title and Text" };

            ShapeTree shapeTree9 = new ShapeTree();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties9 = new NonVisualGroupShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties42 = new NonVisualDrawingProperties(){ Id = (UInt32Value)1U, Name = "" };
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties9 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties42 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties9.Append(nonVisualDrawingProperties42);
            nonVisualGroupShapeProperties9.Append(nonVisualGroupShapeDrawingProperties9);
            nonVisualGroupShapeProperties9.Append(applicationNonVisualDrawingProperties42);

            GroupShapeProperties groupShapeProperties9 = new GroupShapeProperties();

            A.TransformGroup transformGroup9 = new A.TransformGroup();
            A.Offset offset21 = new A.Offset(){ X = 0L, Y = 0L };
            A.Extents extents21 = new A.Extents(){ Cx = 0L, Cy = 0L };
            A.ChildOffset childOffset9 = new A.ChildOffset(){ X = 0L, Y = 0L };
            A.ChildExtents childExtents9 = new A.ChildExtents(){ Cx = 0L, Cy = 0L };

            transformGroup9.Append(offset21);
            transformGroup9.Append(extents21);
            transformGroup9.Append(childOffset9);
            transformGroup9.Append(childExtents9);

            groupShapeProperties9.Append(transformGroup9);

            Shape shape34 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties34 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties43 = new NonVisualDrawingProperties(){ Id = (UInt32Value)2U, Name = "Vertical Title 1" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList34 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension34 = new A.NonVisualDrawingPropertiesExtension(){ Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            OpenXmlUnknownElement openXmlUnknownElement34 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{CA4FEE10-459B-1FD5-3F1B-98B2BE2CE8CD}\" />");

            nonVisualDrawingPropertiesExtension34.Append(openXmlUnknownElement34);

            nonVisualDrawingPropertiesExtensionList34.Append(nonVisualDrawingPropertiesExtension34);

            nonVisualDrawingProperties43.Append(nonVisualDrawingPropertiesExtensionList34);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties34 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks34 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties34.Append(shapeLocks34);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties43 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape34 = new PlaceholderShape(){ Type = PlaceholderValues.Title, Orientation = DirectionValues.Vertical };

            applicationNonVisualDrawingProperties43.Append(placeholderShape34);

            nonVisualShapeProperties34.Append(nonVisualDrawingProperties43);
            nonVisualShapeProperties34.Append(nonVisualShapeDrawingProperties34);
            nonVisualShapeProperties34.Append(applicationNonVisualDrawingProperties43);

            ShapeProperties shapeProperties35 = new ShapeProperties();

            A.Transform2D transform2D13 = new A.Transform2D();
            A.Offset offset22 = new A.Offset(){ X = 8724900L, Y = 365125L };
            A.Extents extents22 = new A.Extents(){ Cx = 2628900L, Cy = 5811838L };

            transform2D13.Append(offset22);
            transform2D13.Append(extents22);

            shapeProperties35.Append(transform2D13);

            TextBody textBody34 = new TextBody();
            A.BodyProperties bodyProperties35 = new A.BodyProperties(){ Vertical = A.TextVerticalValues.EastAsianVetical };
            A.ListStyle listStyle35 = new A.ListStyle();

            A.Paragraph paragraph46 = new A.Paragraph();

            A.Run run25 = new A.Run();
            A.RunProperties runProperties39 = new A.RunProperties(){ Language = "en-US" };
            A.Text text39 = new A.Text();
            text39.Text = "Click to edit Master title style";

            run25.Append(runProperties39);
            run25.Append(text39);
            A.EndParagraphRunProperties endParagraphRunProperties32 = new A.EndParagraphRunProperties(){ Language = "pl-PL" };

            paragraph46.Append(run25);
            paragraph46.Append(endParagraphRunProperties32);

            textBody34.Append(bodyProperties35);
            textBody34.Append(listStyle35);
            textBody34.Append(paragraph46);

            shape34.Append(nonVisualShapeProperties34);
            shape34.Append(shapeProperties35);
            shape34.Append(textBody34);

            Shape shape35 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties35 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties44 = new NonVisualDrawingProperties(){ Id = (UInt32Value)3U, Name = "Vertical Text Placeholder 2" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList35 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension35 = new A.NonVisualDrawingPropertiesExtension(){ Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            OpenXmlUnknownElement openXmlUnknownElement35 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{8940CB41-24D0-61D1-D945-E91DC24F8D47}\" />");

            nonVisualDrawingPropertiesExtension35.Append(openXmlUnknownElement35);

            nonVisualDrawingPropertiesExtensionList35.Append(nonVisualDrawingPropertiesExtension35);

            nonVisualDrawingProperties44.Append(nonVisualDrawingPropertiesExtensionList35);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties35 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks35 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties35.Append(shapeLocks35);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties44 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape35 = new PlaceholderShape(){ Type = PlaceholderValues.Body, Orientation = DirectionValues.Vertical, Index = (UInt32Value)1U };

            applicationNonVisualDrawingProperties44.Append(placeholderShape35);

            nonVisualShapeProperties35.Append(nonVisualDrawingProperties44);
            nonVisualShapeProperties35.Append(nonVisualShapeDrawingProperties35);
            nonVisualShapeProperties35.Append(applicationNonVisualDrawingProperties44);

            ShapeProperties shapeProperties36 = new ShapeProperties();

            A.Transform2D transform2D14 = new A.Transform2D();
            A.Offset offset23 = new A.Offset(){ X = 838200L, Y = 365125L };
            A.Extents extents23 = new A.Extents(){ Cx = 7734300L, Cy = 5811838L };

            transform2D14.Append(offset23);
            transform2D14.Append(extents23);

            shapeProperties36.Append(transform2D14);

            TextBody textBody35 = new TextBody();
            A.BodyProperties bodyProperties36 = new A.BodyProperties(){ Vertical = A.TextVerticalValues.EastAsianVetical };
            A.ListStyle listStyle36 = new A.ListStyle();

            A.Paragraph paragraph47 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties18 = new A.ParagraphProperties(){ Level = 0 };

            A.Run run26 = new A.Run();
            A.RunProperties runProperties40 = new A.RunProperties(){ Language = "en-US" };
            A.Text text40 = new A.Text();
            text40.Text = "Click to edit Master text styles";

            run26.Append(runProperties40);
            run26.Append(text40);

            paragraph47.Append(paragraphProperties18);
            paragraph47.Append(run26);

            A.Paragraph paragraph48 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties19 = new A.ParagraphProperties(){ Level = 1 };

            A.Run run27 = new A.Run();
            A.RunProperties runProperties41 = new A.RunProperties(){ Language = "en-US" };
            A.Text text41 = new A.Text();
            text41.Text = "Second level";

            run27.Append(runProperties41);
            run27.Append(text41);

            paragraph48.Append(paragraphProperties19);
            paragraph48.Append(run27);

            A.Paragraph paragraph49 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties20 = new A.ParagraphProperties(){ Level = 2 };

            A.Run run28 = new A.Run();
            A.RunProperties runProperties42 = new A.RunProperties(){ Language = "en-US" };
            A.Text text42 = new A.Text();
            text42.Text = "Third level";

            run28.Append(runProperties42);
            run28.Append(text42);

            paragraph49.Append(paragraphProperties20);
            paragraph49.Append(run28);

            A.Paragraph paragraph50 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties21 = new A.ParagraphProperties(){ Level = 3 };

            A.Run run29 = new A.Run();
            A.RunProperties runProperties43 = new A.RunProperties(){ Language = "en-US" };
            A.Text text43 = new A.Text();
            text43.Text = "Fourth level";

            run29.Append(runProperties43);
            run29.Append(text43);

            paragraph50.Append(paragraphProperties21);
            paragraph50.Append(run29);

            A.Paragraph paragraph51 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties22 = new A.ParagraphProperties(){ Level = 4 };

            A.Run run30 = new A.Run();
            A.RunProperties runProperties44 = new A.RunProperties(){ Language = "en-US" };
            A.Text text44 = new A.Text();
            text44.Text = "Fifth level";

            run30.Append(runProperties44);
            run30.Append(text44);
            A.EndParagraphRunProperties endParagraphRunProperties33 = new A.EndParagraphRunProperties(){ Language = "pl-PL" };

            paragraph51.Append(paragraphProperties22);
            paragraph51.Append(run30);
            paragraph51.Append(endParagraphRunProperties33);

            textBody35.Append(bodyProperties36);
            textBody35.Append(listStyle36);
            textBody35.Append(paragraph47);
            textBody35.Append(paragraph48);
            textBody35.Append(paragraph49);
            textBody35.Append(paragraph50);
            textBody35.Append(paragraph51);

            shape35.Append(nonVisualShapeProperties35);
            shape35.Append(shapeProperties36);
            shape35.Append(textBody35);

            Shape shape36 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties36 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties45 = new NonVisualDrawingProperties(){ Id = (UInt32Value)4U, Name = "Date Placeholder 3" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList36 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension36 = new A.NonVisualDrawingPropertiesExtension(){ Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            OpenXmlUnknownElement openXmlUnknownElement36 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{E950EA4F-C90A-4241-C074-3EE9EC5D2BC6}\" />");

            nonVisualDrawingPropertiesExtension36.Append(openXmlUnknownElement36);

            nonVisualDrawingPropertiesExtensionList36.Append(nonVisualDrawingPropertiesExtension36);

            nonVisualDrawingProperties45.Append(nonVisualDrawingPropertiesExtensionList36);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties36 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks36 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties36.Append(shapeLocks36);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties45 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape36 = new PlaceholderShape(){ Type = PlaceholderValues.DateAndTime, Size = PlaceholderSizeValues.Half, Index = (UInt32Value)10U };

            applicationNonVisualDrawingProperties45.Append(placeholderShape36);

            nonVisualShapeProperties36.Append(nonVisualDrawingProperties45);
            nonVisualShapeProperties36.Append(nonVisualShapeDrawingProperties36);
            nonVisualShapeProperties36.Append(applicationNonVisualDrawingProperties45);
            ShapeProperties shapeProperties37 = new ShapeProperties();

            TextBody textBody36 = new TextBody();
            A.BodyProperties bodyProperties37 = new A.BodyProperties();
            A.ListStyle listStyle37 = new A.ListStyle();

            A.Paragraph paragraph52 = new A.Paragraph();

            A.Field field15 = new A.Field(){ Id = "{F550F806-175B-4EE1-ABEA-A5F04E4FD72B}", Type = "datetimeFigureOut" };

            A.RunProperties runProperties45 = new A.RunProperties(){ Language = "pl-PL" };
            runProperties45.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text45 = new A.Text();
            text45.Text = "20.08.2025";

            field15.Append(runProperties45);
            field15.Append(text45);
            A.EndParagraphRunProperties endParagraphRunProperties34 = new A.EndParagraphRunProperties(){ Language = "pl-PL" };

            paragraph52.Append(field15);
            paragraph52.Append(endParagraphRunProperties34);

            textBody36.Append(bodyProperties37);
            textBody36.Append(listStyle37);
            textBody36.Append(paragraph52);

            shape36.Append(nonVisualShapeProperties36);
            shape36.Append(shapeProperties37);
            shape36.Append(textBody36);

            Shape shape37 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties37 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties46 = new NonVisualDrawingProperties(){ Id = (UInt32Value)5U, Name = "Footer Placeholder 4" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList37 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension37 = new A.NonVisualDrawingPropertiesExtension(){ Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            OpenXmlUnknownElement openXmlUnknownElement37 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{1B7E3A5F-0975-F339-E34B-FE6C638870E9}\" />");

            nonVisualDrawingPropertiesExtension37.Append(openXmlUnknownElement37);

            nonVisualDrawingPropertiesExtensionList37.Append(nonVisualDrawingPropertiesExtension37);

            nonVisualDrawingProperties46.Append(nonVisualDrawingPropertiesExtensionList37);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties37 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks37 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties37.Append(shapeLocks37);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties46 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape37 = new PlaceholderShape(){ Type = PlaceholderValues.Footer, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)11U };

            applicationNonVisualDrawingProperties46.Append(placeholderShape37);

            nonVisualShapeProperties37.Append(nonVisualDrawingProperties46);
            nonVisualShapeProperties37.Append(nonVisualShapeDrawingProperties37);
            nonVisualShapeProperties37.Append(applicationNonVisualDrawingProperties46);
            ShapeProperties shapeProperties38 = new ShapeProperties();

            TextBody textBody37 = new TextBody();
            A.BodyProperties bodyProperties38 = new A.BodyProperties();
            A.ListStyle listStyle38 = new A.ListStyle();

            A.Paragraph paragraph53 = new A.Paragraph();
            A.EndParagraphRunProperties endParagraphRunProperties35 = new A.EndParagraphRunProperties(){ Language = "pl-PL" };

            paragraph53.Append(endParagraphRunProperties35);

            textBody37.Append(bodyProperties38);
            textBody37.Append(listStyle38);
            textBody37.Append(paragraph53);

            shape37.Append(nonVisualShapeProperties37);
            shape37.Append(shapeProperties38);
            shape37.Append(textBody37);

            Shape shape38 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties38 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties47 = new NonVisualDrawingProperties(){ Id = (UInt32Value)6U, Name = "Slide Number Placeholder 5" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList38 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension38 = new A.NonVisualDrawingPropertiesExtension(){ Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            OpenXmlUnknownElement openXmlUnknownElement38 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{C8E04515-967E-1BBE-E3A5-5B32FE5C5226}\" />");

            nonVisualDrawingPropertiesExtension38.Append(openXmlUnknownElement38);

            nonVisualDrawingPropertiesExtensionList38.Append(nonVisualDrawingPropertiesExtension38);

            nonVisualDrawingProperties47.Append(nonVisualDrawingPropertiesExtensionList38);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties38 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks38 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties38.Append(shapeLocks38);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties47 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape38 = new PlaceholderShape(){ Type = PlaceholderValues.SlideNumber, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)12U };

            applicationNonVisualDrawingProperties47.Append(placeholderShape38);

            nonVisualShapeProperties38.Append(nonVisualDrawingProperties47);
            nonVisualShapeProperties38.Append(nonVisualShapeDrawingProperties38);
            nonVisualShapeProperties38.Append(applicationNonVisualDrawingProperties47);
            ShapeProperties shapeProperties39 = new ShapeProperties();

            TextBody textBody38 = new TextBody();
            A.BodyProperties bodyProperties39 = new A.BodyProperties();
            A.ListStyle listStyle39 = new A.ListStyle();

            A.Paragraph paragraph54 = new A.Paragraph();

            A.Field field16 = new A.Field(){ Id = "{5FAB34AF-B09A-4E2A-BAD3-CCD6E0AD81E6}", Type = "slidenum" };

            A.RunProperties runProperties46 = new A.RunProperties(){ Language = "pl-PL" };
            runProperties46.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text46 = new A.Text();
            text46.Text = "‹#›";

            field16.Append(runProperties46);
            field16.Append(text46);
            A.EndParagraphRunProperties endParagraphRunProperties36 = new A.EndParagraphRunProperties(){ Language = "pl-PL" };

            paragraph54.Append(field16);
            paragraph54.Append(endParagraphRunProperties36);

            textBody38.Append(bodyProperties39);
            textBody38.Append(listStyle39);
            textBody38.Append(paragraph54);

            shape38.Append(nonVisualShapeProperties38);
            shape38.Append(shapeProperties39);
            shape38.Append(textBody38);

            shapeTree9.Append(nonVisualGroupShapeProperties9);
            shapeTree9.Append(groupShapeProperties9);
            shapeTree9.Append(shape34);
            shapeTree9.Append(shape35);
            shapeTree9.Append(shape36);
            shapeTree9.Append(shape37);
            shapeTree9.Append(shape38);

            CommonSlideDataExtensionList commonSlideDataExtensionList9 = new CommonSlideDataExtensionList();

            CommonSlideDataExtension commonSlideDataExtension9 = new CommonSlideDataExtension(){ Uri = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}" };

            P14.CreationId creationId9 = new P14.CreationId(){ Val = (UInt32Value)2514756079U };
            creationId9.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            commonSlideDataExtension9.Append(creationId9);

            commonSlideDataExtensionList9.Append(commonSlideDataExtension9);

            commonSlideData9.Append(shapeTree9);
            commonSlideData9.Append(commonSlideDataExtensionList9);

            ColorMapOverride colorMapOverride8 = new ColorMapOverride();
            A.MasterColorMapping masterColorMapping8 = new A.MasterColorMapping();

            colorMapOverride8.Append(masterColorMapping8);

            slideLayout7.Append(commonSlideData9);
            slideLayout7.Append(colorMapOverride8);

            slideLayoutPart7.SlideLayout = slideLayout7;
        }

        // Generates content of slideLayoutPart8.
        private void GenerateSlideLayoutPart8Content(SlideLayoutPart slideLayoutPart8)
        {
            SlideLayout slideLayout8 = new SlideLayout(){ Type = SlideLayoutValues.TwoTextAndTwoObjects, Preserve = true };
            slideLayout8.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            slideLayout8.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            slideLayout8.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            CommonSlideData commonSlideData10 = new CommonSlideData(){ Name = "Comparison" };

            ShapeTree shapeTree10 = new ShapeTree();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties10 = new NonVisualGroupShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties48 = new NonVisualDrawingProperties(){ Id = (UInt32Value)1U, Name = "" };
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties10 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties48 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties10.Append(nonVisualDrawingProperties48);
            nonVisualGroupShapeProperties10.Append(nonVisualGroupShapeDrawingProperties10);
            nonVisualGroupShapeProperties10.Append(applicationNonVisualDrawingProperties48);

            GroupShapeProperties groupShapeProperties10 = new GroupShapeProperties();

            A.TransformGroup transformGroup10 = new A.TransformGroup();
            A.Offset offset24 = new A.Offset(){ X = 0L, Y = 0L };
            A.Extents extents24 = new A.Extents(){ Cx = 0L, Cy = 0L };
            A.ChildOffset childOffset10 = new A.ChildOffset(){ X = 0L, Y = 0L };
            A.ChildExtents childExtents10 = new A.ChildExtents(){ Cx = 0L, Cy = 0L };

            transformGroup10.Append(offset24);
            transformGroup10.Append(extents24);
            transformGroup10.Append(childOffset10);
            transformGroup10.Append(childExtents10);

            groupShapeProperties10.Append(transformGroup10);

            Shape shape39 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties39 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties49 = new NonVisualDrawingProperties(){ Id = (UInt32Value)2U, Name = "Title 1" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList39 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension39 = new A.NonVisualDrawingPropertiesExtension(){ Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            OpenXmlUnknownElement openXmlUnknownElement39 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{CE6BB028-2054-749C-C942-F73C96BE24F4}\" />");

            nonVisualDrawingPropertiesExtension39.Append(openXmlUnknownElement39);

            nonVisualDrawingPropertiesExtensionList39.Append(nonVisualDrawingPropertiesExtension39);

            nonVisualDrawingProperties49.Append(nonVisualDrawingPropertiesExtensionList39);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties39 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks39 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties39.Append(shapeLocks39);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties49 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape39 = new PlaceholderShape(){ Type = PlaceholderValues.Title };

            applicationNonVisualDrawingProperties49.Append(placeholderShape39);

            nonVisualShapeProperties39.Append(nonVisualDrawingProperties49);
            nonVisualShapeProperties39.Append(nonVisualShapeDrawingProperties39);
            nonVisualShapeProperties39.Append(applicationNonVisualDrawingProperties49);

            ShapeProperties shapeProperties40 = new ShapeProperties();

            A.Transform2D transform2D15 = new A.Transform2D();
            A.Offset offset25 = new A.Offset(){ X = 839788L, Y = 365125L };
            A.Extents extents25 = new A.Extents(){ Cx = 10515600L, Cy = 1325563L };

            transform2D15.Append(offset25);
            transform2D15.Append(extents25);

            shapeProperties40.Append(transform2D15);

            TextBody textBody39 = new TextBody();
            A.BodyProperties bodyProperties40 = new A.BodyProperties();
            A.ListStyle listStyle40 = new A.ListStyle();

            A.Paragraph paragraph55 = new A.Paragraph();

            A.Run run31 = new A.Run();
            A.RunProperties runProperties47 = new A.RunProperties(){ Language = "en-US" };
            A.Text text47 = new A.Text();
            text47.Text = "Click to edit Master title style";

            run31.Append(runProperties47);
            run31.Append(text47);
            A.EndParagraphRunProperties endParagraphRunProperties37 = new A.EndParagraphRunProperties(){ Language = "pl-PL" };

            paragraph55.Append(run31);
            paragraph55.Append(endParagraphRunProperties37);

            textBody39.Append(bodyProperties40);
            textBody39.Append(listStyle40);
            textBody39.Append(paragraph55);

            shape39.Append(nonVisualShapeProperties39);
            shape39.Append(shapeProperties40);
            shape39.Append(textBody39);

            Shape shape40 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties40 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties50 = new NonVisualDrawingProperties(){ Id = (UInt32Value)3U, Name = "Text Placeholder 2" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList40 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension40 = new A.NonVisualDrawingPropertiesExtension(){ Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            OpenXmlUnknownElement openXmlUnknownElement40 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{74DEF3B3-0DC9-70DC-904F-4C1D2B2AB2A5}\" />");

            nonVisualDrawingPropertiesExtension40.Append(openXmlUnknownElement40);

            nonVisualDrawingPropertiesExtensionList40.Append(nonVisualDrawingPropertiesExtension40);

            nonVisualDrawingProperties50.Append(nonVisualDrawingPropertiesExtensionList40);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties40 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks40 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties40.Append(shapeLocks40);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties50 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape40 = new PlaceholderShape(){ Type = PlaceholderValues.Body, Index = (UInt32Value)1U };

            applicationNonVisualDrawingProperties50.Append(placeholderShape40);

            nonVisualShapeProperties40.Append(nonVisualDrawingProperties50);
            nonVisualShapeProperties40.Append(nonVisualShapeDrawingProperties40);
            nonVisualShapeProperties40.Append(applicationNonVisualDrawingProperties50);

            ShapeProperties shapeProperties41 = new ShapeProperties();

            A.Transform2D transform2D16 = new A.Transform2D();
            A.Offset offset26 = new A.Offset(){ X = 839788L, Y = 1681163L };
            A.Extents extents26 = new A.Extents(){ Cx = 5157787L, Cy = 823912L };

            transform2D16.Append(offset26);
            transform2D16.Append(extents26);

            shapeProperties41.Append(transform2D16);

            TextBody textBody40 = new TextBody();
            A.BodyProperties bodyProperties41 = new A.BodyProperties(){ Anchor = A.TextAnchoringTypeValues.Bottom };

            A.ListStyle listStyle41 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties15 = new A.Level1ParagraphProperties(){ LeftMargin = 0, Indent = 0 };
            A.NoBullet noBullet29 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties73 = new A.DefaultRunProperties(){ FontSize = 2400, Bold = true };

            level1ParagraphProperties15.Append(noBullet29);
            level1ParagraphProperties15.Append(defaultRunProperties73);

            A.Level2ParagraphProperties level2ParagraphProperties8 = new A.Level2ParagraphProperties(){ LeftMargin = 457200, Indent = 0 };
            A.NoBullet noBullet30 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties74 = new A.DefaultRunProperties(){ FontSize = 2000, Bold = true };

            level2ParagraphProperties8.Append(noBullet30);
            level2ParagraphProperties8.Append(defaultRunProperties74);

            A.Level3ParagraphProperties level3ParagraphProperties8 = new A.Level3ParagraphProperties(){ LeftMargin = 914400, Indent = 0 };
            A.NoBullet noBullet31 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties75 = new A.DefaultRunProperties(){ FontSize = 1800, Bold = true };

            level3ParagraphProperties8.Append(noBullet31);
            level3ParagraphProperties8.Append(defaultRunProperties75);

            A.Level4ParagraphProperties level4ParagraphProperties8 = new A.Level4ParagraphProperties(){ LeftMargin = 1371600, Indent = 0 };
            A.NoBullet noBullet32 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties76 = new A.DefaultRunProperties(){ FontSize = 1600, Bold = true };

            level4ParagraphProperties8.Append(noBullet32);
            level4ParagraphProperties8.Append(defaultRunProperties76);

            A.Level5ParagraphProperties level5ParagraphProperties8 = new A.Level5ParagraphProperties(){ LeftMargin = 1828800, Indent = 0 };
            A.NoBullet noBullet33 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties77 = new A.DefaultRunProperties(){ FontSize = 1600, Bold = true };

            level5ParagraphProperties8.Append(noBullet33);
            level5ParagraphProperties8.Append(defaultRunProperties77);

            A.Level6ParagraphProperties level6ParagraphProperties8 = new A.Level6ParagraphProperties(){ LeftMargin = 2286000, Indent = 0 };
            A.NoBullet noBullet34 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties78 = new A.DefaultRunProperties(){ FontSize = 1600, Bold = true };

            level6ParagraphProperties8.Append(noBullet34);
            level6ParagraphProperties8.Append(defaultRunProperties78);

            A.Level7ParagraphProperties level7ParagraphProperties8 = new A.Level7ParagraphProperties(){ LeftMargin = 2743200, Indent = 0 };
            A.NoBullet noBullet35 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties79 = new A.DefaultRunProperties(){ FontSize = 1600, Bold = true };

            level7ParagraphProperties8.Append(noBullet35);
            level7ParagraphProperties8.Append(defaultRunProperties79);

            A.Level8ParagraphProperties level8ParagraphProperties8 = new A.Level8ParagraphProperties(){ LeftMargin = 3200400, Indent = 0 };
            A.NoBullet noBullet36 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties80 = new A.DefaultRunProperties(){ FontSize = 1600, Bold = true };

            level8ParagraphProperties8.Append(noBullet36);
            level8ParagraphProperties8.Append(defaultRunProperties80);

            A.Level9ParagraphProperties level9ParagraphProperties8 = new A.Level9ParagraphProperties(){ LeftMargin = 3657600, Indent = 0 };
            A.NoBullet noBullet37 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties81 = new A.DefaultRunProperties(){ FontSize = 1600, Bold = true };

            level9ParagraphProperties8.Append(noBullet37);
            level9ParagraphProperties8.Append(defaultRunProperties81);

            listStyle41.Append(level1ParagraphProperties15);
            listStyle41.Append(level2ParagraphProperties8);
            listStyle41.Append(level3ParagraphProperties8);
            listStyle41.Append(level4ParagraphProperties8);
            listStyle41.Append(level5ParagraphProperties8);
            listStyle41.Append(level6ParagraphProperties8);
            listStyle41.Append(level7ParagraphProperties8);
            listStyle41.Append(level8ParagraphProperties8);
            listStyle41.Append(level9ParagraphProperties8);

            A.Paragraph paragraph56 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties23 = new A.ParagraphProperties(){ Level = 0 };

            A.Run run32 = new A.Run();
            A.RunProperties runProperties48 = new A.RunProperties(){ Language = "en-US" };
            A.Text text48 = new A.Text();
            text48.Text = "Click to edit Master text styles";

            run32.Append(runProperties48);
            run32.Append(text48);

            paragraph56.Append(paragraphProperties23);
            paragraph56.Append(run32);

            textBody40.Append(bodyProperties41);
            textBody40.Append(listStyle41);
            textBody40.Append(paragraph56);

            shape40.Append(nonVisualShapeProperties40);
            shape40.Append(shapeProperties41);
            shape40.Append(textBody40);

            Shape shape41 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties41 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties51 = new NonVisualDrawingProperties(){ Id = (UInt32Value)4U, Name = "Content Placeholder 3" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList41 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension41 = new A.NonVisualDrawingPropertiesExtension(){ Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            OpenXmlUnknownElement openXmlUnknownElement41 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{1D4F686E-CE94-9E71-7BF2-9C9A7E7C3127}\" />");

            nonVisualDrawingPropertiesExtension41.Append(openXmlUnknownElement41);

            nonVisualDrawingPropertiesExtensionList41.Append(nonVisualDrawingPropertiesExtension41);

            nonVisualDrawingProperties51.Append(nonVisualDrawingPropertiesExtensionList41);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties41 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks41 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties41.Append(shapeLocks41);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties51 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape41 = new PlaceholderShape(){ Size = PlaceholderSizeValues.Half, Index = (UInt32Value)2U };

            applicationNonVisualDrawingProperties51.Append(placeholderShape41);

            nonVisualShapeProperties41.Append(nonVisualDrawingProperties51);
            nonVisualShapeProperties41.Append(nonVisualShapeDrawingProperties41);
            nonVisualShapeProperties41.Append(applicationNonVisualDrawingProperties51);

            ShapeProperties shapeProperties42 = new ShapeProperties();

            A.Transform2D transform2D17 = new A.Transform2D();
            A.Offset offset27 = new A.Offset(){ X = 839788L, Y = 2505075L };
            A.Extents extents27 = new A.Extents(){ Cx = 5157787L, Cy = 3684588L };

            transform2D17.Append(offset27);
            transform2D17.Append(extents27);

            shapeProperties42.Append(transform2D17);

            TextBody textBody41 = new TextBody();
            A.BodyProperties bodyProperties42 = new A.BodyProperties();
            A.ListStyle listStyle42 = new A.ListStyle();

            A.Paragraph paragraph57 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties24 = new A.ParagraphProperties(){ Level = 0 };

            A.Run run33 = new A.Run();
            A.RunProperties runProperties49 = new A.RunProperties(){ Language = "en-US" };
            A.Text text49 = new A.Text();
            text49.Text = "Click to edit Master text styles";

            run33.Append(runProperties49);
            run33.Append(text49);

            paragraph57.Append(paragraphProperties24);
            paragraph57.Append(run33);

            A.Paragraph paragraph58 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties25 = new A.ParagraphProperties(){ Level = 1 };

            A.Run run34 = new A.Run();
            A.RunProperties runProperties50 = new A.RunProperties(){ Language = "en-US" };
            A.Text text50 = new A.Text();
            text50.Text = "Second level";

            run34.Append(runProperties50);
            run34.Append(text50);

            paragraph58.Append(paragraphProperties25);
            paragraph58.Append(run34);

            A.Paragraph paragraph59 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties26 = new A.ParagraphProperties(){ Level = 2 };

            A.Run run35 = new A.Run();
            A.RunProperties runProperties51 = new A.RunProperties(){ Language = "en-US" };
            A.Text text51 = new A.Text();
            text51.Text = "Third level";

            run35.Append(runProperties51);
            run35.Append(text51);

            paragraph59.Append(paragraphProperties26);
            paragraph59.Append(run35);

            A.Paragraph paragraph60 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties27 = new A.ParagraphProperties(){ Level = 3 };

            A.Run run36 = new A.Run();
            A.RunProperties runProperties52 = new A.RunProperties(){ Language = "en-US" };
            A.Text text52 = new A.Text();
            text52.Text = "Fourth level";

            run36.Append(runProperties52);
            run36.Append(text52);

            paragraph60.Append(paragraphProperties27);
            paragraph60.Append(run36);

            A.Paragraph paragraph61 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties28 = new A.ParagraphProperties(){ Level = 4 };

            A.Run run37 = new A.Run();
            A.RunProperties runProperties53 = new A.RunProperties(){ Language = "en-US" };
            A.Text text53 = new A.Text();
            text53.Text = "Fifth level";

            run37.Append(runProperties53);
            run37.Append(text53);
            A.EndParagraphRunProperties endParagraphRunProperties38 = new A.EndParagraphRunProperties(){ Language = "pl-PL" };

            paragraph61.Append(paragraphProperties28);
            paragraph61.Append(run37);
            paragraph61.Append(endParagraphRunProperties38);

            textBody41.Append(bodyProperties42);
            textBody41.Append(listStyle42);
            textBody41.Append(paragraph57);
            textBody41.Append(paragraph58);
            textBody41.Append(paragraph59);
            textBody41.Append(paragraph60);
            textBody41.Append(paragraph61);

            shape41.Append(nonVisualShapeProperties41);
            shape41.Append(shapeProperties42);
            shape41.Append(textBody41);

            Shape shape42 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties42 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties52 = new NonVisualDrawingProperties(){ Id = (UInt32Value)5U, Name = "Text Placeholder 4" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList42 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension42 = new A.NonVisualDrawingPropertiesExtension(){ Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            OpenXmlUnknownElement openXmlUnknownElement42 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{B619D884-F64E-9FFE-8448-FF3CBC5A74DF}\" />");

            nonVisualDrawingPropertiesExtension42.Append(openXmlUnknownElement42);

            nonVisualDrawingPropertiesExtensionList42.Append(nonVisualDrawingPropertiesExtension42);

            nonVisualDrawingProperties52.Append(nonVisualDrawingPropertiesExtensionList42);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties42 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks42 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties42.Append(shapeLocks42);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties52 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape42 = new PlaceholderShape(){ Type = PlaceholderValues.Body, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)3U };

            applicationNonVisualDrawingProperties52.Append(placeholderShape42);

            nonVisualShapeProperties42.Append(nonVisualDrawingProperties52);
            nonVisualShapeProperties42.Append(nonVisualShapeDrawingProperties42);
            nonVisualShapeProperties42.Append(applicationNonVisualDrawingProperties52);

            ShapeProperties shapeProperties43 = new ShapeProperties();

            A.Transform2D transform2D18 = new A.Transform2D();
            A.Offset offset28 = new A.Offset(){ X = 6172200L, Y = 1681163L };
            A.Extents extents28 = new A.Extents(){ Cx = 5183188L, Cy = 823912L };

            transform2D18.Append(offset28);
            transform2D18.Append(extents28);

            shapeProperties43.Append(transform2D18);

            TextBody textBody42 = new TextBody();
            A.BodyProperties bodyProperties43 = new A.BodyProperties(){ Anchor = A.TextAnchoringTypeValues.Bottom };

            A.ListStyle listStyle43 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties16 = new A.Level1ParagraphProperties(){ LeftMargin = 0, Indent = 0 };
            A.NoBullet noBullet38 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties82 = new A.DefaultRunProperties(){ FontSize = 2400, Bold = true };

            level1ParagraphProperties16.Append(noBullet38);
            level1ParagraphProperties16.Append(defaultRunProperties82);

            A.Level2ParagraphProperties level2ParagraphProperties9 = new A.Level2ParagraphProperties(){ LeftMargin = 457200, Indent = 0 };
            A.NoBullet noBullet39 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties83 = new A.DefaultRunProperties(){ FontSize = 2000, Bold = true };

            level2ParagraphProperties9.Append(noBullet39);
            level2ParagraphProperties9.Append(defaultRunProperties83);

            A.Level3ParagraphProperties level3ParagraphProperties9 = new A.Level3ParagraphProperties(){ LeftMargin = 914400, Indent = 0 };
            A.NoBullet noBullet40 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties84 = new A.DefaultRunProperties(){ FontSize = 1800, Bold = true };

            level3ParagraphProperties9.Append(noBullet40);
            level3ParagraphProperties9.Append(defaultRunProperties84);

            A.Level4ParagraphProperties level4ParagraphProperties9 = new A.Level4ParagraphProperties(){ LeftMargin = 1371600, Indent = 0 };
            A.NoBullet noBullet41 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties85 = new A.DefaultRunProperties(){ FontSize = 1600, Bold = true };

            level4ParagraphProperties9.Append(noBullet41);
            level4ParagraphProperties9.Append(defaultRunProperties85);

            A.Level5ParagraphProperties level5ParagraphProperties9 = new A.Level5ParagraphProperties(){ LeftMargin = 1828800, Indent = 0 };
            A.NoBullet noBullet42 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties86 = new A.DefaultRunProperties(){ FontSize = 1600, Bold = true };

            level5ParagraphProperties9.Append(noBullet42);
            level5ParagraphProperties9.Append(defaultRunProperties86);

            A.Level6ParagraphProperties level6ParagraphProperties9 = new A.Level6ParagraphProperties(){ LeftMargin = 2286000, Indent = 0 };
            A.NoBullet noBullet43 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties87 = new A.DefaultRunProperties(){ FontSize = 1600, Bold = true };

            level6ParagraphProperties9.Append(noBullet43);
            level6ParagraphProperties9.Append(defaultRunProperties87);

            A.Level7ParagraphProperties level7ParagraphProperties9 = new A.Level7ParagraphProperties(){ LeftMargin = 2743200, Indent = 0 };
            A.NoBullet noBullet44 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties88 = new A.DefaultRunProperties(){ FontSize = 1600, Bold = true };

            level7ParagraphProperties9.Append(noBullet44);
            level7ParagraphProperties9.Append(defaultRunProperties88);

            A.Level8ParagraphProperties level8ParagraphProperties9 = new A.Level8ParagraphProperties(){ LeftMargin = 3200400, Indent = 0 };
            A.NoBullet noBullet45 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties89 = new A.DefaultRunProperties(){ FontSize = 1600, Bold = true };

            level8ParagraphProperties9.Append(noBullet45);
            level8ParagraphProperties9.Append(defaultRunProperties89);

            A.Level9ParagraphProperties level9ParagraphProperties9 = new A.Level9ParagraphProperties(){ LeftMargin = 3657600, Indent = 0 };
            A.NoBullet noBullet46 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties90 = new A.DefaultRunProperties(){ FontSize = 1600, Bold = true };

            level9ParagraphProperties9.Append(noBullet46);
            level9ParagraphProperties9.Append(defaultRunProperties90);

            listStyle43.Append(level1ParagraphProperties16);
            listStyle43.Append(level2ParagraphProperties9);
            listStyle43.Append(level3ParagraphProperties9);
            listStyle43.Append(level4ParagraphProperties9);
            listStyle43.Append(level5ParagraphProperties9);
            listStyle43.Append(level6ParagraphProperties9);
            listStyle43.Append(level7ParagraphProperties9);
            listStyle43.Append(level8ParagraphProperties9);
            listStyle43.Append(level9ParagraphProperties9);

            A.Paragraph paragraph62 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties29 = new A.ParagraphProperties(){ Level = 0 };

            A.Run run38 = new A.Run();
            A.RunProperties runProperties54 = new A.RunProperties(){ Language = "en-US" };
            A.Text text54 = new A.Text();
            text54.Text = "Click to edit Master text styles";

            run38.Append(runProperties54);
            run38.Append(text54);

            paragraph62.Append(paragraphProperties29);
            paragraph62.Append(run38);

            textBody42.Append(bodyProperties43);
            textBody42.Append(listStyle43);
            textBody42.Append(paragraph62);

            shape42.Append(nonVisualShapeProperties42);
            shape42.Append(shapeProperties43);
            shape42.Append(textBody42);

            Shape shape43 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties43 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties53 = new NonVisualDrawingProperties(){ Id = (UInt32Value)6U, Name = "Content Placeholder 5" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList43 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension43 = new A.NonVisualDrawingPropertiesExtension(){ Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            OpenXmlUnknownElement openXmlUnknownElement43 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{106F62F3-7471-62F6-2CB2-66DEA739A900}\" />");

            nonVisualDrawingPropertiesExtension43.Append(openXmlUnknownElement43);

            nonVisualDrawingPropertiesExtensionList43.Append(nonVisualDrawingPropertiesExtension43);

            nonVisualDrawingProperties53.Append(nonVisualDrawingPropertiesExtensionList43);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties43 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks43 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties43.Append(shapeLocks43);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties53 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape43 = new PlaceholderShape(){ Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)4U };

            applicationNonVisualDrawingProperties53.Append(placeholderShape43);

            nonVisualShapeProperties43.Append(nonVisualDrawingProperties53);
            nonVisualShapeProperties43.Append(nonVisualShapeDrawingProperties43);
            nonVisualShapeProperties43.Append(applicationNonVisualDrawingProperties53);

            ShapeProperties shapeProperties44 = new ShapeProperties();

            A.Transform2D transform2D19 = new A.Transform2D();
            A.Offset offset29 = new A.Offset(){ X = 6172200L, Y = 2505075L };
            A.Extents extents29 = new A.Extents(){ Cx = 5183188L, Cy = 3684588L };

            transform2D19.Append(offset29);
            transform2D19.Append(extents29);

            shapeProperties44.Append(transform2D19);

            TextBody textBody43 = new TextBody();
            A.BodyProperties bodyProperties44 = new A.BodyProperties();
            A.ListStyle listStyle44 = new A.ListStyle();

            A.Paragraph paragraph63 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties30 = new A.ParagraphProperties(){ Level = 0 };

            A.Run run39 = new A.Run();
            A.RunProperties runProperties55 = new A.RunProperties(){ Language = "en-US" };
            A.Text text55 = new A.Text();
            text55.Text = "Click to edit Master text styles";

            run39.Append(runProperties55);
            run39.Append(text55);

            paragraph63.Append(paragraphProperties30);
            paragraph63.Append(run39);

            A.Paragraph paragraph64 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties31 = new A.ParagraphProperties(){ Level = 1 };

            A.Run run40 = new A.Run();
            A.RunProperties runProperties56 = new A.RunProperties(){ Language = "en-US" };
            A.Text text56 = new A.Text();
            text56.Text = "Second level";

            run40.Append(runProperties56);
            run40.Append(text56);

            paragraph64.Append(paragraphProperties31);
            paragraph64.Append(run40);

            A.Paragraph paragraph65 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties32 = new A.ParagraphProperties(){ Level = 2 };

            A.Run run41 = new A.Run();
            A.RunProperties runProperties57 = new A.RunProperties(){ Language = "en-US" };
            A.Text text57 = new A.Text();
            text57.Text = "Third level";

            run41.Append(runProperties57);
            run41.Append(text57);

            paragraph65.Append(paragraphProperties32);
            paragraph65.Append(run41);

            A.Paragraph paragraph66 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties33 = new A.ParagraphProperties(){ Level = 3 };

            A.Run run42 = new A.Run();
            A.RunProperties runProperties58 = new A.RunProperties(){ Language = "en-US" };
            A.Text text58 = new A.Text();
            text58.Text = "Fourth level";

            run42.Append(runProperties58);
            run42.Append(text58);

            paragraph66.Append(paragraphProperties33);
            paragraph66.Append(run42);

            A.Paragraph paragraph67 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties34 = new A.ParagraphProperties(){ Level = 4 };

            A.Run run43 = new A.Run();
            A.RunProperties runProperties59 = new A.RunProperties(){ Language = "en-US" };
            A.Text text59 = new A.Text();
            text59.Text = "Fifth level";

            run43.Append(runProperties59);
            run43.Append(text59);
            A.EndParagraphRunProperties endParagraphRunProperties39 = new A.EndParagraphRunProperties(){ Language = "pl-PL" };

            paragraph67.Append(paragraphProperties34);
            paragraph67.Append(run43);
            paragraph67.Append(endParagraphRunProperties39);

            textBody43.Append(bodyProperties44);
            textBody43.Append(listStyle44);
            textBody43.Append(paragraph63);
            textBody43.Append(paragraph64);
            textBody43.Append(paragraph65);
            textBody43.Append(paragraph66);
            textBody43.Append(paragraph67);

            shape43.Append(nonVisualShapeProperties43);
            shape43.Append(shapeProperties44);
            shape43.Append(textBody43);

            Shape shape44 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties44 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties54 = new NonVisualDrawingProperties(){ Id = (UInt32Value)7U, Name = "Date Placeholder 6" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList44 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension44 = new A.NonVisualDrawingPropertiesExtension(){ Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            OpenXmlUnknownElement openXmlUnknownElement44 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{1019A1BA-DAE9-001F-4C1C-ED08C6A81764}\" />");

            nonVisualDrawingPropertiesExtension44.Append(openXmlUnknownElement44);

            nonVisualDrawingPropertiesExtensionList44.Append(nonVisualDrawingPropertiesExtension44);

            nonVisualDrawingProperties54.Append(nonVisualDrawingPropertiesExtensionList44);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties44 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks44 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties44.Append(shapeLocks44);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties54 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape44 = new PlaceholderShape(){ Type = PlaceholderValues.DateAndTime, Size = PlaceholderSizeValues.Half, Index = (UInt32Value)10U };

            applicationNonVisualDrawingProperties54.Append(placeholderShape44);

            nonVisualShapeProperties44.Append(nonVisualDrawingProperties54);
            nonVisualShapeProperties44.Append(nonVisualShapeDrawingProperties44);
            nonVisualShapeProperties44.Append(applicationNonVisualDrawingProperties54);
            ShapeProperties shapeProperties45 = new ShapeProperties();

            TextBody textBody44 = new TextBody();
            A.BodyProperties bodyProperties45 = new A.BodyProperties();
            A.ListStyle listStyle45 = new A.ListStyle();

            A.Paragraph paragraph68 = new A.Paragraph();

            A.Field field17 = new A.Field(){ Id = "{F550F806-175B-4EE1-ABEA-A5F04E4FD72B}", Type = "datetimeFigureOut" };

            A.RunProperties runProperties60 = new A.RunProperties(){ Language = "pl-PL" };
            runProperties60.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text60 = new A.Text();
            text60.Text = "20.08.2025";

            field17.Append(runProperties60);
            field17.Append(text60);
            A.EndParagraphRunProperties endParagraphRunProperties40 = new A.EndParagraphRunProperties(){ Language = "pl-PL" };

            paragraph68.Append(field17);
            paragraph68.Append(endParagraphRunProperties40);

            textBody44.Append(bodyProperties45);
            textBody44.Append(listStyle45);
            textBody44.Append(paragraph68);

            shape44.Append(nonVisualShapeProperties44);
            shape44.Append(shapeProperties45);
            shape44.Append(textBody44);

            Shape shape45 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties45 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties55 = new NonVisualDrawingProperties(){ Id = (UInt32Value)8U, Name = "Footer Placeholder 7" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList45 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension45 = new A.NonVisualDrawingPropertiesExtension(){ Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            OpenXmlUnknownElement openXmlUnknownElement45 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{BF929D73-CB7B-6744-136E-F33DB5F0904C}\" />");

            nonVisualDrawingPropertiesExtension45.Append(openXmlUnknownElement45);

            nonVisualDrawingPropertiesExtensionList45.Append(nonVisualDrawingPropertiesExtension45);

            nonVisualDrawingProperties55.Append(nonVisualDrawingPropertiesExtensionList45);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties45 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks45 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties45.Append(shapeLocks45);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties55 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape45 = new PlaceholderShape(){ Type = PlaceholderValues.Footer, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)11U };

            applicationNonVisualDrawingProperties55.Append(placeholderShape45);

            nonVisualShapeProperties45.Append(nonVisualDrawingProperties55);
            nonVisualShapeProperties45.Append(nonVisualShapeDrawingProperties45);
            nonVisualShapeProperties45.Append(applicationNonVisualDrawingProperties55);
            ShapeProperties shapeProperties46 = new ShapeProperties();

            TextBody textBody45 = new TextBody();
            A.BodyProperties bodyProperties46 = new A.BodyProperties();
            A.ListStyle listStyle46 = new A.ListStyle();

            A.Paragraph paragraph69 = new A.Paragraph();
            A.EndParagraphRunProperties endParagraphRunProperties41 = new A.EndParagraphRunProperties(){ Language = "pl-PL" };

            paragraph69.Append(endParagraphRunProperties41);

            textBody45.Append(bodyProperties46);
            textBody45.Append(listStyle46);
            textBody45.Append(paragraph69);

            shape45.Append(nonVisualShapeProperties45);
            shape45.Append(shapeProperties46);
            shape45.Append(textBody45);

            Shape shape46 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties46 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties56 = new NonVisualDrawingProperties(){ Id = (UInt32Value)9U, Name = "Slide Number Placeholder 8" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList46 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension46 = new A.NonVisualDrawingPropertiesExtension(){ Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            OpenXmlUnknownElement openXmlUnknownElement46 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{33497FF4-3C24-DC06-5760-1ECE11A8B3CB}\" />");

            nonVisualDrawingPropertiesExtension46.Append(openXmlUnknownElement46);

            nonVisualDrawingPropertiesExtensionList46.Append(nonVisualDrawingPropertiesExtension46);

            nonVisualDrawingProperties56.Append(nonVisualDrawingPropertiesExtensionList46);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties46 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks46 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties46.Append(shapeLocks46);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties56 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape46 = new PlaceholderShape(){ Type = PlaceholderValues.SlideNumber, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)12U };

            applicationNonVisualDrawingProperties56.Append(placeholderShape46);

            nonVisualShapeProperties46.Append(nonVisualDrawingProperties56);
            nonVisualShapeProperties46.Append(nonVisualShapeDrawingProperties46);
            nonVisualShapeProperties46.Append(applicationNonVisualDrawingProperties56);
            ShapeProperties shapeProperties47 = new ShapeProperties();

            TextBody textBody46 = new TextBody();
            A.BodyProperties bodyProperties47 = new A.BodyProperties();
            A.ListStyle listStyle47 = new A.ListStyle();

            A.Paragraph paragraph70 = new A.Paragraph();

            A.Field field18 = new A.Field(){ Id = "{5FAB34AF-B09A-4E2A-BAD3-CCD6E0AD81E6}", Type = "slidenum" };

            A.RunProperties runProperties61 = new A.RunProperties(){ Language = "pl-PL" };
            runProperties61.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text61 = new A.Text();
            text61.Text = "‹#›";

            field18.Append(runProperties61);
            field18.Append(text61);
            A.EndParagraphRunProperties endParagraphRunProperties42 = new A.EndParagraphRunProperties(){ Language = "pl-PL" };

            paragraph70.Append(field18);
            paragraph70.Append(endParagraphRunProperties42);

            textBody46.Append(bodyProperties47);
            textBody46.Append(listStyle47);
            textBody46.Append(paragraph70);

            shape46.Append(nonVisualShapeProperties46);
            shape46.Append(shapeProperties47);
            shape46.Append(textBody46);

            shapeTree10.Append(nonVisualGroupShapeProperties10);
            shapeTree10.Append(groupShapeProperties10);
            shapeTree10.Append(shape39);
            shapeTree10.Append(shape40);
            shapeTree10.Append(shape41);
            shapeTree10.Append(shape42);
            shapeTree10.Append(shape43);
            shapeTree10.Append(shape44);
            shapeTree10.Append(shape45);
            shapeTree10.Append(shape46);

            CommonSlideDataExtensionList commonSlideDataExtensionList10 = new CommonSlideDataExtensionList();

            CommonSlideDataExtension commonSlideDataExtension10 = new CommonSlideDataExtension(){ Uri = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}" };

            P14.CreationId creationId10 = new P14.CreationId(){ Val = (UInt32Value)1479905673U };
            creationId10.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            commonSlideDataExtension10.Append(creationId10);

            commonSlideDataExtensionList10.Append(commonSlideDataExtension10);

            commonSlideData10.Append(shapeTree10);
            commonSlideData10.Append(commonSlideDataExtensionList10);

            ColorMapOverride colorMapOverride9 = new ColorMapOverride();
            A.MasterColorMapping masterColorMapping9 = new A.MasterColorMapping();

            colorMapOverride9.Append(masterColorMapping9);

            slideLayout8.Append(commonSlideData10);
            slideLayout8.Append(colorMapOverride9);

            slideLayoutPart8.SlideLayout = slideLayout8;
        }

        // Generates content of slideLayoutPart9.
        private void GenerateSlideLayoutPart9Content(SlideLayoutPart slideLayoutPart9)
        {
            SlideLayout slideLayout9 = new SlideLayout(){ Type = SlideLayoutValues.VerticalText, Preserve = true };
            slideLayout9.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            slideLayout9.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            slideLayout9.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            CommonSlideData commonSlideData11 = new CommonSlideData(){ Name = "Title and Vertical Text" };

            ShapeTree shapeTree11 = new ShapeTree();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties11 = new NonVisualGroupShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties57 = new NonVisualDrawingProperties(){ Id = (UInt32Value)1U, Name = "" };
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties11 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties57 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties11.Append(nonVisualDrawingProperties57);
            nonVisualGroupShapeProperties11.Append(nonVisualGroupShapeDrawingProperties11);
            nonVisualGroupShapeProperties11.Append(applicationNonVisualDrawingProperties57);

            GroupShapeProperties groupShapeProperties11 = new GroupShapeProperties();

            A.TransformGroup transformGroup11 = new A.TransformGroup();
            A.Offset offset30 = new A.Offset(){ X = 0L, Y = 0L };
            A.Extents extents30 = new A.Extents(){ Cx = 0L, Cy = 0L };
            A.ChildOffset childOffset11 = new A.ChildOffset(){ X = 0L, Y = 0L };
            A.ChildExtents childExtents11 = new A.ChildExtents(){ Cx = 0L, Cy = 0L };

            transformGroup11.Append(offset30);
            transformGroup11.Append(extents30);
            transformGroup11.Append(childOffset11);
            transformGroup11.Append(childExtents11);

            groupShapeProperties11.Append(transformGroup11);

            Shape shape47 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties47 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties58 = new NonVisualDrawingProperties(){ Id = (UInt32Value)2U, Name = "Title 1" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList47 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension47 = new A.NonVisualDrawingPropertiesExtension(){ Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            OpenXmlUnknownElement openXmlUnknownElement47 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{60577B76-5D41-DF0C-5067-982852358610}\" />");

            nonVisualDrawingPropertiesExtension47.Append(openXmlUnknownElement47);

            nonVisualDrawingPropertiesExtensionList47.Append(nonVisualDrawingPropertiesExtension47);

            nonVisualDrawingProperties58.Append(nonVisualDrawingPropertiesExtensionList47);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties47 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks47 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties47.Append(shapeLocks47);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties58 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape47 = new PlaceholderShape(){ Type = PlaceholderValues.Title };

            applicationNonVisualDrawingProperties58.Append(placeholderShape47);

            nonVisualShapeProperties47.Append(nonVisualDrawingProperties58);
            nonVisualShapeProperties47.Append(nonVisualShapeDrawingProperties47);
            nonVisualShapeProperties47.Append(applicationNonVisualDrawingProperties58);
            ShapeProperties shapeProperties48 = new ShapeProperties();

            TextBody textBody47 = new TextBody();
            A.BodyProperties bodyProperties48 = new A.BodyProperties();
            A.ListStyle listStyle48 = new A.ListStyle();

            A.Paragraph paragraph71 = new A.Paragraph();

            A.Run run44 = new A.Run();
            A.RunProperties runProperties62 = new A.RunProperties(){ Language = "en-US" };
            A.Text text62 = new A.Text();
            text62.Text = "Click to edit Master title style";

            run44.Append(runProperties62);
            run44.Append(text62);
            A.EndParagraphRunProperties endParagraphRunProperties43 = new A.EndParagraphRunProperties(){ Language = "pl-PL" };

            paragraph71.Append(run44);
            paragraph71.Append(endParagraphRunProperties43);

            textBody47.Append(bodyProperties48);
            textBody47.Append(listStyle48);
            textBody47.Append(paragraph71);

            shape47.Append(nonVisualShapeProperties47);
            shape47.Append(shapeProperties48);
            shape47.Append(textBody47);

            Shape shape48 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties48 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties59 = new NonVisualDrawingProperties(){ Id = (UInt32Value)3U, Name = "Vertical Text Placeholder 2" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList48 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension48 = new A.NonVisualDrawingPropertiesExtension(){ Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            OpenXmlUnknownElement openXmlUnknownElement48 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{CB94687B-2111-11BC-7FB0-B651D2E43603}\" />");

            nonVisualDrawingPropertiesExtension48.Append(openXmlUnknownElement48);

            nonVisualDrawingPropertiesExtensionList48.Append(nonVisualDrawingPropertiesExtension48);

            nonVisualDrawingProperties59.Append(nonVisualDrawingPropertiesExtensionList48);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties48 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks48 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties48.Append(shapeLocks48);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties59 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape48 = new PlaceholderShape(){ Type = PlaceholderValues.Body, Orientation = DirectionValues.Vertical, Index = (UInt32Value)1U };

            applicationNonVisualDrawingProperties59.Append(placeholderShape48);

            nonVisualShapeProperties48.Append(nonVisualDrawingProperties59);
            nonVisualShapeProperties48.Append(nonVisualShapeDrawingProperties48);
            nonVisualShapeProperties48.Append(applicationNonVisualDrawingProperties59);
            ShapeProperties shapeProperties49 = new ShapeProperties();

            TextBody textBody48 = new TextBody();
            A.BodyProperties bodyProperties49 = new A.BodyProperties(){ Vertical = A.TextVerticalValues.EastAsianVetical };
            A.ListStyle listStyle49 = new A.ListStyle();

            A.Paragraph paragraph72 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties35 = new A.ParagraphProperties(){ Level = 0 };

            A.Run run45 = new A.Run();
            A.RunProperties runProperties63 = new A.RunProperties(){ Language = "en-US" };
            A.Text text63 = new A.Text();
            text63.Text = "Click to edit Master text styles";

            run45.Append(runProperties63);
            run45.Append(text63);

            paragraph72.Append(paragraphProperties35);
            paragraph72.Append(run45);

            A.Paragraph paragraph73 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties36 = new A.ParagraphProperties(){ Level = 1 };

            A.Run run46 = new A.Run();
            A.RunProperties runProperties64 = new A.RunProperties(){ Language = "en-US" };
            A.Text text64 = new A.Text();
            text64.Text = "Second level";

            run46.Append(runProperties64);
            run46.Append(text64);

            paragraph73.Append(paragraphProperties36);
            paragraph73.Append(run46);

            A.Paragraph paragraph74 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties37 = new A.ParagraphProperties(){ Level = 2 };

            A.Run run47 = new A.Run();
            A.RunProperties runProperties65 = new A.RunProperties(){ Language = "en-US" };
            A.Text text65 = new A.Text();
            text65.Text = "Third level";

            run47.Append(runProperties65);
            run47.Append(text65);

            paragraph74.Append(paragraphProperties37);
            paragraph74.Append(run47);

            A.Paragraph paragraph75 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties38 = new A.ParagraphProperties(){ Level = 3 };

            A.Run run48 = new A.Run();
            A.RunProperties runProperties66 = new A.RunProperties(){ Language = "en-US" };
            A.Text text66 = new A.Text();
            text66.Text = "Fourth level";

            run48.Append(runProperties66);
            run48.Append(text66);

            paragraph75.Append(paragraphProperties38);
            paragraph75.Append(run48);

            A.Paragraph paragraph76 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties39 = new A.ParagraphProperties(){ Level = 4 };

            A.Run run49 = new A.Run();
            A.RunProperties runProperties67 = new A.RunProperties(){ Language = "en-US" };
            A.Text text67 = new A.Text();
            text67.Text = "Fifth level";

            run49.Append(runProperties67);
            run49.Append(text67);
            A.EndParagraphRunProperties endParagraphRunProperties44 = new A.EndParagraphRunProperties(){ Language = "pl-PL" };

            paragraph76.Append(paragraphProperties39);
            paragraph76.Append(run49);
            paragraph76.Append(endParagraphRunProperties44);

            textBody48.Append(bodyProperties49);
            textBody48.Append(listStyle49);
            textBody48.Append(paragraph72);
            textBody48.Append(paragraph73);
            textBody48.Append(paragraph74);
            textBody48.Append(paragraph75);
            textBody48.Append(paragraph76);

            shape48.Append(nonVisualShapeProperties48);
            shape48.Append(shapeProperties49);
            shape48.Append(textBody48);

            Shape shape49 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties49 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties60 = new NonVisualDrawingProperties(){ Id = (UInt32Value)4U, Name = "Date Placeholder 3" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList49 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension49 = new A.NonVisualDrawingPropertiesExtension(){ Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            OpenXmlUnknownElement openXmlUnknownElement49 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{D64F779F-1CA1-3EB0-C3B2-388D36152299}\" />");

            nonVisualDrawingPropertiesExtension49.Append(openXmlUnknownElement49);

            nonVisualDrawingPropertiesExtensionList49.Append(nonVisualDrawingPropertiesExtension49);

            nonVisualDrawingProperties60.Append(nonVisualDrawingPropertiesExtensionList49);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties49 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks49 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties49.Append(shapeLocks49);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties60 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape49 = new PlaceholderShape(){ Type = PlaceholderValues.DateAndTime, Size = PlaceholderSizeValues.Half, Index = (UInt32Value)10U };

            applicationNonVisualDrawingProperties60.Append(placeholderShape49);

            nonVisualShapeProperties49.Append(nonVisualDrawingProperties60);
            nonVisualShapeProperties49.Append(nonVisualShapeDrawingProperties49);
            nonVisualShapeProperties49.Append(applicationNonVisualDrawingProperties60);
            ShapeProperties shapeProperties50 = new ShapeProperties();

            TextBody textBody49 = new TextBody();
            A.BodyProperties bodyProperties50 = new A.BodyProperties();
            A.ListStyle listStyle50 = new A.ListStyle();

            A.Paragraph paragraph77 = new A.Paragraph();

            A.Field field19 = new A.Field(){ Id = "{F550F806-175B-4EE1-ABEA-A5F04E4FD72B}", Type = "datetimeFigureOut" };

            A.RunProperties runProperties68 = new A.RunProperties(){ Language = "pl-PL" };
            runProperties68.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text68 = new A.Text();
            text68.Text = "20.08.2025";

            field19.Append(runProperties68);
            field19.Append(text68);
            A.EndParagraphRunProperties endParagraphRunProperties45 = new A.EndParagraphRunProperties(){ Language = "pl-PL" };

            paragraph77.Append(field19);
            paragraph77.Append(endParagraphRunProperties45);

            textBody49.Append(bodyProperties50);
            textBody49.Append(listStyle50);
            textBody49.Append(paragraph77);

            shape49.Append(nonVisualShapeProperties49);
            shape49.Append(shapeProperties50);
            shape49.Append(textBody49);

            Shape shape50 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties50 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties61 = new NonVisualDrawingProperties(){ Id = (UInt32Value)5U, Name = "Footer Placeholder 4" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList50 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension50 = new A.NonVisualDrawingPropertiesExtension(){ Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            OpenXmlUnknownElement openXmlUnknownElement50 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{943C26C4-AA11-A748-56F1-72D504AE6CC2}\" />");

            nonVisualDrawingPropertiesExtension50.Append(openXmlUnknownElement50);

            nonVisualDrawingPropertiesExtensionList50.Append(nonVisualDrawingPropertiesExtension50);

            nonVisualDrawingProperties61.Append(nonVisualDrawingPropertiesExtensionList50);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties50 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks50 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties50.Append(shapeLocks50);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties61 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape50 = new PlaceholderShape(){ Type = PlaceholderValues.Footer, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)11U };

            applicationNonVisualDrawingProperties61.Append(placeholderShape50);

            nonVisualShapeProperties50.Append(nonVisualDrawingProperties61);
            nonVisualShapeProperties50.Append(nonVisualShapeDrawingProperties50);
            nonVisualShapeProperties50.Append(applicationNonVisualDrawingProperties61);
            ShapeProperties shapeProperties51 = new ShapeProperties();

            TextBody textBody50 = new TextBody();
            A.BodyProperties bodyProperties51 = new A.BodyProperties();
            A.ListStyle listStyle51 = new A.ListStyle();

            A.Paragraph paragraph78 = new A.Paragraph();
            A.EndParagraphRunProperties endParagraphRunProperties46 = new A.EndParagraphRunProperties(){ Language = "pl-PL" };

            paragraph78.Append(endParagraphRunProperties46);

            textBody50.Append(bodyProperties51);
            textBody50.Append(listStyle51);
            textBody50.Append(paragraph78);

            shape50.Append(nonVisualShapeProperties50);
            shape50.Append(shapeProperties51);
            shape50.Append(textBody50);

            Shape shape51 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties51 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties62 = new NonVisualDrawingProperties(){ Id = (UInt32Value)6U, Name = "Slide Number Placeholder 5" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList51 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension51 = new A.NonVisualDrawingPropertiesExtension(){ Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            OpenXmlUnknownElement openXmlUnknownElement51 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{0E75C361-56F3-5B56-BECA-E4C4A127A0A5}\" />");

            nonVisualDrawingPropertiesExtension51.Append(openXmlUnknownElement51);

            nonVisualDrawingPropertiesExtensionList51.Append(nonVisualDrawingPropertiesExtension51);

            nonVisualDrawingProperties62.Append(nonVisualDrawingPropertiesExtensionList51);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties51 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks51 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties51.Append(shapeLocks51);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties62 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape51 = new PlaceholderShape(){ Type = PlaceholderValues.SlideNumber, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)12U };

            applicationNonVisualDrawingProperties62.Append(placeholderShape51);

            nonVisualShapeProperties51.Append(nonVisualDrawingProperties62);
            nonVisualShapeProperties51.Append(nonVisualShapeDrawingProperties51);
            nonVisualShapeProperties51.Append(applicationNonVisualDrawingProperties62);
            ShapeProperties shapeProperties52 = new ShapeProperties();

            TextBody textBody51 = new TextBody();
            A.BodyProperties bodyProperties52 = new A.BodyProperties();
            A.ListStyle listStyle52 = new A.ListStyle();

            A.Paragraph paragraph79 = new A.Paragraph();

            A.Field field20 = new A.Field(){ Id = "{5FAB34AF-B09A-4E2A-BAD3-CCD6E0AD81E6}", Type = "slidenum" };

            A.RunProperties runProperties69 = new A.RunProperties(){ Language = "pl-PL" };
            runProperties69.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text69 = new A.Text();
            text69.Text = "‹#›";

            field20.Append(runProperties69);
            field20.Append(text69);
            A.EndParagraphRunProperties endParagraphRunProperties47 = new A.EndParagraphRunProperties(){ Language = "pl-PL" };

            paragraph79.Append(field20);
            paragraph79.Append(endParagraphRunProperties47);

            textBody51.Append(bodyProperties52);
            textBody51.Append(listStyle52);
            textBody51.Append(paragraph79);

            shape51.Append(nonVisualShapeProperties51);
            shape51.Append(shapeProperties52);
            shape51.Append(textBody51);

            shapeTree11.Append(nonVisualGroupShapeProperties11);
            shapeTree11.Append(groupShapeProperties11);
            shapeTree11.Append(shape47);
            shapeTree11.Append(shape48);
            shapeTree11.Append(shape49);
            shapeTree11.Append(shape50);
            shapeTree11.Append(shape51);

            CommonSlideDataExtensionList commonSlideDataExtensionList11 = new CommonSlideDataExtensionList();

            CommonSlideDataExtension commonSlideDataExtension11 = new CommonSlideDataExtension(){ Uri = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}" };

            P14.CreationId creationId11 = new P14.CreationId(){ Val = (UInt32Value)1153641642U };
            creationId11.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            commonSlideDataExtension11.Append(creationId11);

            commonSlideDataExtensionList11.Append(commonSlideDataExtension11);

            commonSlideData11.Append(shapeTree11);
            commonSlideData11.Append(commonSlideDataExtensionList11);

            ColorMapOverride colorMapOverride10 = new ColorMapOverride();
            A.MasterColorMapping masterColorMapping10 = new A.MasterColorMapping();

            colorMapOverride10.Append(masterColorMapping10);

            slideLayout9.Append(commonSlideData11);
            slideLayout9.Append(colorMapOverride10);

            slideLayoutPart9.SlideLayout = slideLayout9;
        }

        // Generates content of slideLayoutPart10.
        private void GenerateSlideLayoutPart10Content(SlideLayoutPart slideLayoutPart10)
        {
            SlideLayout slideLayout10 = new SlideLayout(){ Type = SlideLayoutValues.TwoObjects, Preserve = true };
            slideLayout10.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            slideLayout10.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            slideLayout10.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            CommonSlideData commonSlideData12 = new CommonSlideData(){ Name = "Two Content" };

            ShapeTree shapeTree12 = new ShapeTree();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties12 = new NonVisualGroupShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties63 = new NonVisualDrawingProperties(){ Id = (UInt32Value)1U, Name = "" };
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties12 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties63 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties12.Append(nonVisualDrawingProperties63);
            nonVisualGroupShapeProperties12.Append(nonVisualGroupShapeDrawingProperties12);
            nonVisualGroupShapeProperties12.Append(applicationNonVisualDrawingProperties63);

            GroupShapeProperties groupShapeProperties12 = new GroupShapeProperties();

            A.TransformGroup transformGroup12 = new A.TransformGroup();
            A.Offset offset31 = new A.Offset(){ X = 0L, Y = 0L };
            A.Extents extents31 = new A.Extents(){ Cx = 0L, Cy = 0L };
            A.ChildOffset childOffset12 = new A.ChildOffset(){ X = 0L, Y = 0L };
            A.ChildExtents childExtents12 = new A.ChildExtents(){ Cx = 0L, Cy = 0L };

            transformGroup12.Append(offset31);
            transformGroup12.Append(extents31);
            transformGroup12.Append(childOffset12);
            transformGroup12.Append(childExtents12);

            groupShapeProperties12.Append(transformGroup12);

            Shape shape52 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties52 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties64 = new NonVisualDrawingProperties(){ Id = (UInt32Value)2U, Name = "Title 1" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList52 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension52 = new A.NonVisualDrawingPropertiesExtension(){ Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            OpenXmlUnknownElement openXmlUnknownElement52 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{9A9CDBE5-DE92-8A26-67BE-A69E03B1489B}\" />");

            nonVisualDrawingPropertiesExtension52.Append(openXmlUnknownElement52);

            nonVisualDrawingPropertiesExtensionList52.Append(nonVisualDrawingPropertiesExtension52);

            nonVisualDrawingProperties64.Append(nonVisualDrawingPropertiesExtensionList52);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties52 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks52 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties52.Append(shapeLocks52);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties64 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape52 = new PlaceholderShape(){ Type = PlaceholderValues.Title };

            applicationNonVisualDrawingProperties64.Append(placeholderShape52);

            nonVisualShapeProperties52.Append(nonVisualDrawingProperties64);
            nonVisualShapeProperties52.Append(nonVisualShapeDrawingProperties52);
            nonVisualShapeProperties52.Append(applicationNonVisualDrawingProperties64);
            ShapeProperties shapeProperties53 = new ShapeProperties();

            TextBody textBody52 = new TextBody();
            A.BodyProperties bodyProperties53 = new A.BodyProperties();
            A.ListStyle listStyle53 = new A.ListStyle();

            A.Paragraph paragraph80 = new A.Paragraph();

            A.Run run50 = new A.Run();
            A.RunProperties runProperties70 = new A.RunProperties(){ Language = "en-US" };
            A.Text text70 = new A.Text();
            text70.Text = "Click to edit Master title style";

            run50.Append(runProperties70);
            run50.Append(text70);
            A.EndParagraphRunProperties endParagraphRunProperties48 = new A.EndParagraphRunProperties(){ Language = "pl-PL" };

            paragraph80.Append(run50);
            paragraph80.Append(endParagraphRunProperties48);

            textBody52.Append(bodyProperties53);
            textBody52.Append(listStyle53);
            textBody52.Append(paragraph80);

            shape52.Append(nonVisualShapeProperties52);
            shape52.Append(shapeProperties53);
            shape52.Append(textBody52);

            Shape shape53 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties53 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties65 = new NonVisualDrawingProperties(){ Id = (UInt32Value)3U, Name = "Content Placeholder 2" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList53 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension53 = new A.NonVisualDrawingPropertiesExtension(){ Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            OpenXmlUnknownElement openXmlUnknownElement53 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{06033BBF-FAA9-D126-F725-8C12C150DF8B}\" />");

            nonVisualDrawingPropertiesExtension53.Append(openXmlUnknownElement53);

            nonVisualDrawingPropertiesExtensionList53.Append(nonVisualDrawingPropertiesExtension53);

            nonVisualDrawingProperties65.Append(nonVisualDrawingPropertiesExtensionList53);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties53 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks53 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties53.Append(shapeLocks53);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties65 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape53 = new PlaceholderShape(){ Size = PlaceholderSizeValues.Half, Index = (UInt32Value)1U };

            applicationNonVisualDrawingProperties65.Append(placeholderShape53);

            nonVisualShapeProperties53.Append(nonVisualDrawingProperties65);
            nonVisualShapeProperties53.Append(nonVisualShapeDrawingProperties53);
            nonVisualShapeProperties53.Append(applicationNonVisualDrawingProperties65);

            ShapeProperties shapeProperties54 = new ShapeProperties();

            A.Transform2D transform2D20 = new A.Transform2D();
            A.Offset offset32 = new A.Offset(){ X = 838200L, Y = 1825625L };
            A.Extents extents32 = new A.Extents(){ Cx = 5181600L, Cy = 4351338L };

            transform2D20.Append(offset32);
            transform2D20.Append(extents32);

            shapeProperties54.Append(transform2D20);

            TextBody textBody53 = new TextBody();
            A.BodyProperties bodyProperties54 = new A.BodyProperties();
            A.ListStyle listStyle54 = new A.ListStyle();

            A.Paragraph paragraph81 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties40 = new A.ParagraphProperties(){ Level = 0 };

            A.Run run51 = new A.Run();
            A.RunProperties runProperties71 = new A.RunProperties(){ Language = "en-US" };
            A.Text text71 = new A.Text();
            text71.Text = "Click to edit Master text styles";

            run51.Append(runProperties71);
            run51.Append(text71);

            paragraph81.Append(paragraphProperties40);
            paragraph81.Append(run51);

            A.Paragraph paragraph82 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties41 = new A.ParagraphProperties(){ Level = 1 };

            A.Run run52 = new A.Run();
            A.RunProperties runProperties72 = new A.RunProperties(){ Language = "en-US" };
            A.Text text72 = new A.Text();
            text72.Text = "Second level";

            run52.Append(runProperties72);
            run52.Append(text72);

            paragraph82.Append(paragraphProperties41);
            paragraph82.Append(run52);

            A.Paragraph paragraph83 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties42 = new A.ParagraphProperties(){ Level = 2 };

            A.Run run53 = new A.Run();
            A.RunProperties runProperties73 = new A.RunProperties(){ Language = "en-US" };
            A.Text text73 = new A.Text();
            text73.Text = "Third level";

            run53.Append(runProperties73);
            run53.Append(text73);

            paragraph83.Append(paragraphProperties42);
            paragraph83.Append(run53);

            A.Paragraph paragraph84 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties43 = new A.ParagraphProperties(){ Level = 3 };

            A.Run run54 = new A.Run();
            A.RunProperties runProperties74 = new A.RunProperties(){ Language = "en-US" };
            A.Text text74 = new A.Text();
            text74.Text = "Fourth level";

            run54.Append(runProperties74);
            run54.Append(text74);

            paragraph84.Append(paragraphProperties43);
            paragraph84.Append(run54);

            A.Paragraph paragraph85 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties44 = new A.ParagraphProperties(){ Level = 4 };

            A.Run run55 = new A.Run();
            A.RunProperties runProperties75 = new A.RunProperties(){ Language = "en-US" };
            A.Text text75 = new A.Text();
            text75.Text = "Fifth level";

            run55.Append(runProperties75);
            run55.Append(text75);
            A.EndParagraphRunProperties endParagraphRunProperties49 = new A.EndParagraphRunProperties(){ Language = "pl-PL" };

            paragraph85.Append(paragraphProperties44);
            paragraph85.Append(run55);
            paragraph85.Append(endParagraphRunProperties49);

            textBody53.Append(bodyProperties54);
            textBody53.Append(listStyle54);
            textBody53.Append(paragraph81);
            textBody53.Append(paragraph82);
            textBody53.Append(paragraph83);
            textBody53.Append(paragraph84);
            textBody53.Append(paragraph85);

            shape53.Append(nonVisualShapeProperties53);
            shape53.Append(shapeProperties54);
            shape53.Append(textBody53);

            Shape shape54 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties54 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties66 = new NonVisualDrawingProperties(){ Id = (UInt32Value)4U, Name = "Content Placeholder 3" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList54 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension54 = new A.NonVisualDrawingPropertiesExtension(){ Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            OpenXmlUnknownElement openXmlUnknownElement54 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{7E2C3E91-5DA8-469B-3502-73BAEF25F66B}\" />");

            nonVisualDrawingPropertiesExtension54.Append(openXmlUnknownElement54);

            nonVisualDrawingPropertiesExtensionList54.Append(nonVisualDrawingPropertiesExtension54);

            nonVisualDrawingProperties66.Append(nonVisualDrawingPropertiesExtensionList54);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties54 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks54 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties54.Append(shapeLocks54);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties66 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape54 = new PlaceholderShape(){ Size = PlaceholderSizeValues.Half, Index = (UInt32Value)2U };

            applicationNonVisualDrawingProperties66.Append(placeholderShape54);

            nonVisualShapeProperties54.Append(nonVisualDrawingProperties66);
            nonVisualShapeProperties54.Append(nonVisualShapeDrawingProperties54);
            nonVisualShapeProperties54.Append(applicationNonVisualDrawingProperties66);

            ShapeProperties shapeProperties55 = new ShapeProperties();

            A.Transform2D transform2D21 = new A.Transform2D();
            A.Offset offset33 = new A.Offset(){ X = 6172200L, Y = 1825625L };
            A.Extents extents33 = new A.Extents(){ Cx = 5181600L, Cy = 4351338L };

            transform2D21.Append(offset33);
            transform2D21.Append(extents33);

            shapeProperties55.Append(transform2D21);

            TextBody textBody54 = new TextBody();
            A.BodyProperties bodyProperties55 = new A.BodyProperties();
            A.ListStyle listStyle55 = new A.ListStyle();

            A.Paragraph paragraph86 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties45 = new A.ParagraphProperties(){ Level = 0 };

            A.Run run56 = new A.Run();
            A.RunProperties runProperties76 = new A.RunProperties(){ Language = "en-US" };
            A.Text text76 = new A.Text();
            text76.Text = "Click to edit Master text styles";

            run56.Append(runProperties76);
            run56.Append(text76);

            paragraph86.Append(paragraphProperties45);
            paragraph86.Append(run56);

            A.Paragraph paragraph87 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties46 = new A.ParagraphProperties(){ Level = 1 };

            A.Run run57 = new A.Run();
            A.RunProperties runProperties77 = new A.RunProperties(){ Language = "en-US" };
            A.Text text77 = new A.Text();
            text77.Text = "Second level";

            run57.Append(runProperties77);
            run57.Append(text77);

            paragraph87.Append(paragraphProperties46);
            paragraph87.Append(run57);

            A.Paragraph paragraph88 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties47 = new A.ParagraphProperties(){ Level = 2 };

            A.Run run58 = new A.Run();
            A.RunProperties runProperties78 = new A.RunProperties(){ Language = "en-US" };
            A.Text text78 = new A.Text();
            text78.Text = "Third level";

            run58.Append(runProperties78);
            run58.Append(text78);

            paragraph88.Append(paragraphProperties47);
            paragraph88.Append(run58);

            A.Paragraph paragraph89 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties48 = new A.ParagraphProperties(){ Level = 3 };

            A.Run run59 = new A.Run();
            A.RunProperties runProperties79 = new A.RunProperties(){ Language = "en-US" };
            A.Text text79 = new A.Text();
            text79.Text = "Fourth level";

            run59.Append(runProperties79);
            run59.Append(text79);

            paragraph89.Append(paragraphProperties48);
            paragraph89.Append(run59);

            A.Paragraph paragraph90 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties49 = new A.ParagraphProperties(){ Level = 4 };

            A.Run run60 = new A.Run();
            A.RunProperties runProperties80 = new A.RunProperties(){ Language = "en-US" };
            A.Text text80 = new A.Text();
            text80.Text = "Fifth level";

            run60.Append(runProperties80);
            run60.Append(text80);
            A.EndParagraphRunProperties endParagraphRunProperties50 = new A.EndParagraphRunProperties(){ Language = "pl-PL" };

            paragraph90.Append(paragraphProperties49);
            paragraph90.Append(run60);
            paragraph90.Append(endParagraphRunProperties50);

            textBody54.Append(bodyProperties55);
            textBody54.Append(listStyle55);
            textBody54.Append(paragraph86);
            textBody54.Append(paragraph87);
            textBody54.Append(paragraph88);
            textBody54.Append(paragraph89);
            textBody54.Append(paragraph90);

            shape54.Append(nonVisualShapeProperties54);
            shape54.Append(shapeProperties55);
            shape54.Append(textBody54);

            Shape shape55 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties55 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties67 = new NonVisualDrawingProperties(){ Id = (UInt32Value)5U, Name = "Date Placeholder 4" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList55 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension55 = new A.NonVisualDrawingPropertiesExtension(){ Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            OpenXmlUnknownElement openXmlUnknownElement55 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{0C843B07-892D-AEFC-B9EA-F29956C9BA74}\" />");

            nonVisualDrawingPropertiesExtension55.Append(openXmlUnknownElement55);

            nonVisualDrawingPropertiesExtensionList55.Append(nonVisualDrawingPropertiesExtension55);

            nonVisualDrawingProperties67.Append(nonVisualDrawingPropertiesExtensionList55);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties55 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks55 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties55.Append(shapeLocks55);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties67 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape55 = new PlaceholderShape(){ Type = PlaceholderValues.DateAndTime, Size = PlaceholderSizeValues.Half, Index = (UInt32Value)10U };

            applicationNonVisualDrawingProperties67.Append(placeholderShape55);

            nonVisualShapeProperties55.Append(nonVisualDrawingProperties67);
            nonVisualShapeProperties55.Append(nonVisualShapeDrawingProperties55);
            nonVisualShapeProperties55.Append(applicationNonVisualDrawingProperties67);
            ShapeProperties shapeProperties56 = new ShapeProperties();

            TextBody textBody55 = new TextBody();
            A.BodyProperties bodyProperties56 = new A.BodyProperties();
            A.ListStyle listStyle56 = new A.ListStyle();

            A.Paragraph paragraph91 = new A.Paragraph();

            A.Field field21 = new A.Field(){ Id = "{F550F806-175B-4EE1-ABEA-A5F04E4FD72B}", Type = "datetimeFigureOut" };

            A.RunProperties runProperties81 = new A.RunProperties(){ Language = "pl-PL" };
            runProperties81.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text81 = new A.Text();
            text81.Text = "20.08.2025";

            field21.Append(runProperties81);
            field21.Append(text81);
            A.EndParagraphRunProperties endParagraphRunProperties51 = new A.EndParagraphRunProperties(){ Language = "pl-PL" };

            paragraph91.Append(field21);
            paragraph91.Append(endParagraphRunProperties51);

            textBody55.Append(bodyProperties56);
            textBody55.Append(listStyle56);
            textBody55.Append(paragraph91);

            shape55.Append(nonVisualShapeProperties55);
            shape55.Append(shapeProperties56);
            shape55.Append(textBody55);

            Shape shape56 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties56 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties68 = new NonVisualDrawingProperties(){ Id = (UInt32Value)6U, Name = "Footer Placeholder 5" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList56 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension56 = new A.NonVisualDrawingPropertiesExtension(){ Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            OpenXmlUnknownElement openXmlUnknownElement56 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{5467C875-9421-F4B6-3E20-46E856A91BE5}\" />");

            nonVisualDrawingPropertiesExtension56.Append(openXmlUnknownElement56);

            nonVisualDrawingPropertiesExtensionList56.Append(nonVisualDrawingPropertiesExtension56);

            nonVisualDrawingProperties68.Append(nonVisualDrawingPropertiesExtensionList56);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties56 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks56 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties56.Append(shapeLocks56);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties68 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape56 = new PlaceholderShape(){ Type = PlaceholderValues.Footer, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)11U };

            applicationNonVisualDrawingProperties68.Append(placeholderShape56);

            nonVisualShapeProperties56.Append(nonVisualDrawingProperties68);
            nonVisualShapeProperties56.Append(nonVisualShapeDrawingProperties56);
            nonVisualShapeProperties56.Append(applicationNonVisualDrawingProperties68);
            ShapeProperties shapeProperties57 = new ShapeProperties();

            TextBody textBody56 = new TextBody();
            A.BodyProperties bodyProperties57 = new A.BodyProperties();
            A.ListStyle listStyle57 = new A.ListStyle();

            A.Paragraph paragraph92 = new A.Paragraph();
            A.EndParagraphRunProperties endParagraphRunProperties52 = new A.EndParagraphRunProperties(){ Language = "pl-PL" };

            paragraph92.Append(endParagraphRunProperties52);

            textBody56.Append(bodyProperties57);
            textBody56.Append(listStyle57);
            textBody56.Append(paragraph92);

            shape56.Append(nonVisualShapeProperties56);
            shape56.Append(shapeProperties57);
            shape56.Append(textBody56);

            Shape shape57 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties57 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties69 = new NonVisualDrawingProperties(){ Id = (UInt32Value)7U, Name = "Slide Number Placeholder 6" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList57 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension57 = new A.NonVisualDrawingPropertiesExtension(){ Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            OpenXmlUnknownElement openXmlUnknownElement57 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{81FA18BB-203B-F167-F27A-FE902142A696}\" />");

            nonVisualDrawingPropertiesExtension57.Append(openXmlUnknownElement57);

            nonVisualDrawingPropertiesExtensionList57.Append(nonVisualDrawingPropertiesExtension57);

            nonVisualDrawingProperties69.Append(nonVisualDrawingPropertiesExtensionList57);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties57 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks57 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties57.Append(shapeLocks57);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties69 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape57 = new PlaceholderShape(){ Type = PlaceholderValues.SlideNumber, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)12U };

            applicationNonVisualDrawingProperties69.Append(placeholderShape57);

            nonVisualShapeProperties57.Append(nonVisualDrawingProperties69);
            nonVisualShapeProperties57.Append(nonVisualShapeDrawingProperties57);
            nonVisualShapeProperties57.Append(applicationNonVisualDrawingProperties69);
            ShapeProperties shapeProperties58 = new ShapeProperties();

            TextBody textBody57 = new TextBody();
            A.BodyProperties bodyProperties58 = new A.BodyProperties();
            A.ListStyle listStyle58 = new A.ListStyle();

            A.Paragraph paragraph93 = new A.Paragraph();

            A.Field field22 = new A.Field(){ Id = "{5FAB34AF-B09A-4E2A-BAD3-CCD6E0AD81E6}", Type = "slidenum" };

            A.RunProperties runProperties82 = new A.RunProperties(){ Language = "pl-PL" };
            runProperties82.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text82 = new A.Text();
            text82.Text = "‹#›";

            field22.Append(runProperties82);
            field22.Append(text82);
            A.EndParagraphRunProperties endParagraphRunProperties53 = new A.EndParagraphRunProperties(){ Language = "pl-PL" };

            paragraph93.Append(field22);
            paragraph93.Append(endParagraphRunProperties53);

            textBody57.Append(bodyProperties58);
            textBody57.Append(listStyle58);
            textBody57.Append(paragraph93);

            shape57.Append(nonVisualShapeProperties57);
            shape57.Append(shapeProperties58);
            shape57.Append(textBody57);

            shapeTree12.Append(nonVisualGroupShapeProperties12);
            shapeTree12.Append(groupShapeProperties12);
            shapeTree12.Append(shape52);
            shapeTree12.Append(shape53);
            shapeTree12.Append(shape54);
            shapeTree12.Append(shape55);
            shapeTree12.Append(shape56);
            shapeTree12.Append(shape57);

            CommonSlideDataExtensionList commonSlideDataExtensionList12 = new CommonSlideDataExtensionList();

            CommonSlideDataExtension commonSlideDataExtension12 = new CommonSlideDataExtension(){ Uri = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}" };

            P14.CreationId creationId12 = new P14.CreationId(){ Val = (UInt32Value)3105625174U };
            creationId12.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            commonSlideDataExtension12.Append(creationId12);

            commonSlideDataExtensionList12.Append(commonSlideDataExtension12);

            commonSlideData12.Append(shapeTree12);
            commonSlideData12.Append(commonSlideDataExtensionList12);

            ColorMapOverride colorMapOverride11 = new ColorMapOverride();
            A.MasterColorMapping masterColorMapping11 = new A.MasterColorMapping();

            colorMapOverride11.Append(masterColorMapping11);

            slideLayout10.Append(commonSlideData12);
            slideLayout10.Append(colorMapOverride11);

            slideLayoutPart10.SlideLayout = slideLayout10;
        }

        // Generates content of slideLayoutPart11.
        private void GenerateSlideLayoutPart11Content(SlideLayoutPart slideLayoutPart11)
        {
            SlideLayout slideLayout11 = new SlideLayout(){ Type = SlideLayoutValues.PictureText, Preserve = true };
            slideLayout11.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            slideLayout11.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            slideLayout11.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            CommonSlideData commonSlideData13 = new CommonSlideData(){ Name = "Picture with Caption" };

            ShapeTree shapeTree13 = new ShapeTree();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties13 = new NonVisualGroupShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties70 = new NonVisualDrawingProperties(){ Id = (UInt32Value)1U, Name = "" };
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties13 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties70 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties13.Append(nonVisualDrawingProperties70);
            nonVisualGroupShapeProperties13.Append(nonVisualGroupShapeDrawingProperties13);
            nonVisualGroupShapeProperties13.Append(applicationNonVisualDrawingProperties70);

            GroupShapeProperties groupShapeProperties13 = new GroupShapeProperties();

            A.TransformGroup transformGroup13 = new A.TransformGroup();
            A.Offset offset34 = new A.Offset(){ X = 0L, Y = 0L };
            A.Extents extents34 = new A.Extents(){ Cx = 0L, Cy = 0L };
            A.ChildOffset childOffset13 = new A.ChildOffset(){ X = 0L, Y = 0L };
            A.ChildExtents childExtents13 = new A.ChildExtents(){ Cx = 0L, Cy = 0L };

            transformGroup13.Append(offset34);
            transformGroup13.Append(extents34);
            transformGroup13.Append(childOffset13);
            transformGroup13.Append(childExtents13);

            groupShapeProperties13.Append(transformGroup13);

            Shape shape58 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties58 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties71 = new NonVisualDrawingProperties(){ Id = (UInt32Value)2U, Name = "Title 1" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList58 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension58 = new A.NonVisualDrawingPropertiesExtension(){ Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            OpenXmlUnknownElement openXmlUnknownElement58 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{E1CF4EB0-B301-DC9E-B2C5-E1905AB20F7B}\" />");

            nonVisualDrawingPropertiesExtension58.Append(openXmlUnknownElement58);

            nonVisualDrawingPropertiesExtensionList58.Append(nonVisualDrawingPropertiesExtension58);

            nonVisualDrawingProperties71.Append(nonVisualDrawingPropertiesExtensionList58);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties58 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks58 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties58.Append(shapeLocks58);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties71 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape58 = new PlaceholderShape(){ Type = PlaceholderValues.Title };

            applicationNonVisualDrawingProperties71.Append(placeholderShape58);

            nonVisualShapeProperties58.Append(nonVisualDrawingProperties71);
            nonVisualShapeProperties58.Append(nonVisualShapeDrawingProperties58);
            nonVisualShapeProperties58.Append(applicationNonVisualDrawingProperties71);

            ShapeProperties shapeProperties59 = new ShapeProperties();

            A.Transform2D transform2D22 = new A.Transform2D();
            A.Offset offset35 = new A.Offset(){ X = 839788L, Y = 457200L };
            A.Extents extents35 = new A.Extents(){ Cx = 3932237L, Cy = 1600200L };

            transform2D22.Append(offset35);
            transform2D22.Append(extents35);

            shapeProperties59.Append(transform2D22);

            TextBody textBody58 = new TextBody();
            A.BodyProperties bodyProperties59 = new A.BodyProperties(){ Anchor = A.TextAnchoringTypeValues.Bottom };

            A.ListStyle listStyle59 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties17 = new A.Level1ParagraphProperties();
            A.DefaultRunProperties defaultRunProperties91 = new A.DefaultRunProperties(){ FontSize = 3200 };

            level1ParagraphProperties17.Append(defaultRunProperties91);

            listStyle59.Append(level1ParagraphProperties17);

            A.Paragraph paragraph94 = new A.Paragraph();

            A.Run run61 = new A.Run();
            A.RunProperties runProperties83 = new A.RunProperties(){ Language = "en-US" };
            A.Text text83 = new A.Text();
            text83.Text = "Click to edit Master title style";

            run61.Append(runProperties83);
            run61.Append(text83);
            A.EndParagraphRunProperties endParagraphRunProperties54 = new A.EndParagraphRunProperties(){ Language = "pl-PL" };

            paragraph94.Append(run61);
            paragraph94.Append(endParagraphRunProperties54);

            textBody58.Append(bodyProperties59);
            textBody58.Append(listStyle59);
            textBody58.Append(paragraph94);

            shape58.Append(nonVisualShapeProperties58);
            shape58.Append(shapeProperties59);
            shape58.Append(textBody58);

            Shape shape59 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties59 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties72 = new NonVisualDrawingProperties(){ Id = (UInt32Value)3U, Name = "Picture Placeholder 2" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList59 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension59 = new A.NonVisualDrawingPropertiesExtension(){ Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            OpenXmlUnknownElement openXmlUnknownElement59 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{B1B062AF-5584-FBDA-0968-08234852218C}\" />");

            nonVisualDrawingPropertiesExtension59.Append(openXmlUnknownElement59);

            nonVisualDrawingPropertiesExtensionList59.Append(nonVisualDrawingPropertiesExtension59);

            nonVisualDrawingProperties72.Append(nonVisualDrawingPropertiesExtensionList59);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties59 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks59 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties59.Append(shapeLocks59);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties72 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape59 = new PlaceholderShape(){ Type = PlaceholderValues.Picture, Index = (UInt32Value)1U };

            applicationNonVisualDrawingProperties72.Append(placeholderShape59);

            nonVisualShapeProperties59.Append(nonVisualDrawingProperties72);
            nonVisualShapeProperties59.Append(nonVisualShapeDrawingProperties59);
            nonVisualShapeProperties59.Append(applicationNonVisualDrawingProperties72);

            ShapeProperties shapeProperties60 = new ShapeProperties();

            A.Transform2D transform2D23 = new A.Transform2D();
            A.Offset offset36 = new A.Offset(){ X = 5183188L, Y = 987425L };
            A.Extents extents36 = new A.Extents(){ Cx = 6172200L, Cy = 4873625L };

            transform2D23.Append(offset36);
            transform2D23.Append(extents36);

            shapeProperties60.Append(transform2D23);

            TextBody textBody59 = new TextBody();
            A.BodyProperties bodyProperties60 = new A.BodyProperties();

            A.ListStyle listStyle60 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties18 = new A.Level1ParagraphProperties(){ LeftMargin = 0, Indent = 0 };
            A.NoBullet noBullet47 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties92 = new A.DefaultRunProperties(){ FontSize = 3200 };

            level1ParagraphProperties18.Append(noBullet47);
            level1ParagraphProperties18.Append(defaultRunProperties92);

            A.Level2ParagraphProperties level2ParagraphProperties10 = new A.Level2ParagraphProperties(){ LeftMargin = 457200, Indent = 0 };
            A.NoBullet noBullet48 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties93 = new A.DefaultRunProperties(){ FontSize = 2800 };

            level2ParagraphProperties10.Append(noBullet48);
            level2ParagraphProperties10.Append(defaultRunProperties93);

            A.Level3ParagraphProperties level3ParagraphProperties10 = new A.Level3ParagraphProperties(){ LeftMargin = 914400, Indent = 0 };
            A.NoBullet noBullet49 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties94 = new A.DefaultRunProperties(){ FontSize = 2400 };

            level3ParagraphProperties10.Append(noBullet49);
            level3ParagraphProperties10.Append(defaultRunProperties94);

            A.Level4ParagraphProperties level4ParagraphProperties10 = new A.Level4ParagraphProperties(){ LeftMargin = 1371600, Indent = 0 };
            A.NoBullet noBullet50 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties95 = new A.DefaultRunProperties(){ FontSize = 2000 };

            level4ParagraphProperties10.Append(noBullet50);
            level4ParagraphProperties10.Append(defaultRunProperties95);

            A.Level5ParagraphProperties level5ParagraphProperties10 = new A.Level5ParagraphProperties(){ LeftMargin = 1828800, Indent = 0 };
            A.NoBullet noBullet51 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties96 = new A.DefaultRunProperties(){ FontSize = 2000 };

            level5ParagraphProperties10.Append(noBullet51);
            level5ParagraphProperties10.Append(defaultRunProperties96);

            A.Level6ParagraphProperties level6ParagraphProperties10 = new A.Level6ParagraphProperties(){ LeftMargin = 2286000, Indent = 0 };
            A.NoBullet noBullet52 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties97 = new A.DefaultRunProperties(){ FontSize = 2000 };

            level6ParagraphProperties10.Append(noBullet52);
            level6ParagraphProperties10.Append(defaultRunProperties97);

            A.Level7ParagraphProperties level7ParagraphProperties10 = new A.Level7ParagraphProperties(){ LeftMargin = 2743200, Indent = 0 };
            A.NoBullet noBullet53 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties98 = new A.DefaultRunProperties(){ FontSize = 2000 };

            level7ParagraphProperties10.Append(noBullet53);
            level7ParagraphProperties10.Append(defaultRunProperties98);

            A.Level8ParagraphProperties level8ParagraphProperties10 = new A.Level8ParagraphProperties(){ LeftMargin = 3200400, Indent = 0 };
            A.NoBullet noBullet54 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties99 = new A.DefaultRunProperties(){ FontSize = 2000 };

            level8ParagraphProperties10.Append(noBullet54);
            level8ParagraphProperties10.Append(defaultRunProperties99);

            A.Level9ParagraphProperties level9ParagraphProperties10 = new A.Level9ParagraphProperties(){ LeftMargin = 3657600, Indent = 0 };
            A.NoBullet noBullet55 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties100 = new A.DefaultRunProperties(){ FontSize = 2000 };

            level9ParagraphProperties10.Append(noBullet55);
            level9ParagraphProperties10.Append(defaultRunProperties100);

            listStyle60.Append(level1ParagraphProperties18);
            listStyle60.Append(level2ParagraphProperties10);
            listStyle60.Append(level3ParagraphProperties10);
            listStyle60.Append(level4ParagraphProperties10);
            listStyle60.Append(level5ParagraphProperties10);
            listStyle60.Append(level6ParagraphProperties10);
            listStyle60.Append(level7ParagraphProperties10);
            listStyle60.Append(level8ParagraphProperties10);
            listStyle60.Append(level9ParagraphProperties10);

            A.Paragraph paragraph95 = new A.Paragraph();
            A.EndParagraphRunProperties endParagraphRunProperties55 = new A.EndParagraphRunProperties(){ Language = "pl-PL" };

            paragraph95.Append(endParagraphRunProperties55);

            textBody59.Append(bodyProperties60);
            textBody59.Append(listStyle60);
            textBody59.Append(paragraph95);

            shape59.Append(nonVisualShapeProperties59);
            shape59.Append(shapeProperties60);
            shape59.Append(textBody59);

            Shape shape60 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties60 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties73 = new NonVisualDrawingProperties(){ Id = (UInt32Value)4U, Name = "Text Placeholder 3" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList60 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension60 = new A.NonVisualDrawingPropertiesExtension(){ Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            OpenXmlUnknownElement openXmlUnknownElement60 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{FF457D08-F650-24DF-05DF-841A9C5BBCB5}\" />");

            nonVisualDrawingPropertiesExtension60.Append(openXmlUnknownElement60);

            nonVisualDrawingPropertiesExtensionList60.Append(nonVisualDrawingPropertiesExtension60);

            nonVisualDrawingProperties73.Append(nonVisualDrawingPropertiesExtensionList60);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties60 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks60 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties60.Append(shapeLocks60);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties73 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape60 = new PlaceholderShape(){ Type = PlaceholderValues.Body, Size = PlaceholderSizeValues.Half, Index = (UInt32Value)2U };

            applicationNonVisualDrawingProperties73.Append(placeholderShape60);

            nonVisualShapeProperties60.Append(nonVisualDrawingProperties73);
            nonVisualShapeProperties60.Append(nonVisualShapeDrawingProperties60);
            nonVisualShapeProperties60.Append(applicationNonVisualDrawingProperties73);

            ShapeProperties shapeProperties61 = new ShapeProperties();

            A.Transform2D transform2D24 = new A.Transform2D();
            A.Offset offset37 = new A.Offset(){ X = 839788L, Y = 2057400L };
            A.Extents extents37 = new A.Extents(){ Cx = 3932237L, Cy = 3811588L };

            transform2D24.Append(offset37);
            transform2D24.Append(extents37);

            shapeProperties61.Append(transform2D24);

            TextBody textBody60 = new TextBody();
            A.BodyProperties bodyProperties61 = new A.BodyProperties();

            A.ListStyle listStyle61 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties19 = new A.Level1ParagraphProperties(){ LeftMargin = 0, Indent = 0 };
            A.NoBullet noBullet56 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties101 = new A.DefaultRunProperties(){ FontSize = 1600 };

            level1ParagraphProperties19.Append(noBullet56);
            level1ParagraphProperties19.Append(defaultRunProperties101);

            A.Level2ParagraphProperties level2ParagraphProperties11 = new A.Level2ParagraphProperties(){ LeftMargin = 457200, Indent = 0 };
            A.NoBullet noBullet57 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties102 = new A.DefaultRunProperties(){ FontSize = 1400 };

            level2ParagraphProperties11.Append(noBullet57);
            level2ParagraphProperties11.Append(defaultRunProperties102);

            A.Level3ParagraphProperties level3ParagraphProperties11 = new A.Level3ParagraphProperties(){ LeftMargin = 914400, Indent = 0 };
            A.NoBullet noBullet58 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties103 = new A.DefaultRunProperties(){ FontSize = 1200 };

            level3ParagraphProperties11.Append(noBullet58);
            level3ParagraphProperties11.Append(defaultRunProperties103);

            A.Level4ParagraphProperties level4ParagraphProperties11 = new A.Level4ParagraphProperties(){ LeftMargin = 1371600, Indent = 0 };
            A.NoBullet noBullet59 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties104 = new A.DefaultRunProperties(){ FontSize = 1000 };

            level4ParagraphProperties11.Append(noBullet59);
            level4ParagraphProperties11.Append(defaultRunProperties104);

            A.Level5ParagraphProperties level5ParagraphProperties11 = new A.Level5ParagraphProperties(){ LeftMargin = 1828800, Indent = 0 };
            A.NoBullet noBullet60 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties105 = new A.DefaultRunProperties(){ FontSize = 1000 };

            level5ParagraphProperties11.Append(noBullet60);
            level5ParagraphProperties11.Append(defaultRunProperties105);

            A.Level6ParagraphProperties level6ParagraphProperties11 = new A.Level6ParagraphProperties(){ LeftMargin = 2286000, Indent = 0 };
            A.NoBullet noBullet61 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties106 = new A.DefaultRunProperties(){ FontSize = 1000 };

            level6ParagraphProperties11.Append(noBullet61);
            level6ParagraphProperties11.Append(defaultRunProperties106);

            A.Level7ParagraphProperties level7ParagraphProperties11 = new A.Level7ParagraphProperties(){ LeftMargin = 2743200, Indent = 0 };
            A.NoBullet noBullet62 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties107 = new A.DefaultRunProperties(){ FontSize = 1000 };

            level7ParagraphProperties11.Append(noBullet62);
            level7ParagraphProperties11.Append(defaultRunProperties107);

            A.Level8ParagraphProperties level8ParagraphProperties11 = new A.Level8ParagraphProperties(){ LeftMargin = 3200400, Indent = 0 };
            A.NoBullet noBullet63 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties108 = new A.DefaultRunProperties(){ FontSize = 1000 };

            level8ParagraphProperties11.Append(noBullet63);
            level8ParagraphProperties11.Append(defaultRunProperties108);

            A.Level9ParagraphProperties level9ParagraphProperties11 = new A.Level9ParagraphProperties(){ LeftMargin = 3657600, Indent = 0 };
            A.NoBullet noBullet64 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties109 = new A.DefaultRunProperties(){ FontSize = 1000 };

            level9ParagraphProperties11.Append(noBullet64);
            level9ParagraphProperties11.Append(defaultRunProperties109);

            listStyle61.Append(level1ParagraphProperties19);
            listStyle61.Append(level2ParagraphProperties11);
            listStyle61.Append(level3ParagraphProperties11);
            listStyle61.Append(level4ParagraphProperties11);
            listStyle61.Append(level5ParagraphProperties11);
            listStyle61.Append(level6ParagraphProperties11);
            listStyle61.Append(level7ParagraphProperties11);
            listStyle61.Append(level8ParagraphProperties11);
            listStyle61.Append(level9ParagraphProperties11);

            A.Paragraph paragraph96 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties50 = new A.ParagraphProperties(){ Level = 0 };

            A.Run run62 = new A.Run();
            A.RunProperties runProperties84 = new A.RunProperties(){ Language = "en-US" };
            A.Text text84 = new A.Text();
            text84.Text = "Click to edit Master text styles";

            run62.Append(runProperties84);
            run62.Append(text84);

            paragraph96.Append(paragraphProperties50);
            paragraph96.Append(run62);

            textBody60.Append(bodyProperties61);
            textBody60.Append(listStyle61);
            textBody60.Append(paragraph96);

            shape60.Append(nonVisualShapeProperties60);
            shape60.Append(shapeProperties61);
            shape60.Append(textBody60);

            Shape shape61 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties61 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties74 = new NonVisualDrawingProperties(){ Id = (UInt32Value)5U, Name = "Date Placeholder 4" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList61 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension61 = new A.NonVisualDrawingPropertiesExtension(){ Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            OpenXmlUnknownElement openXmlUnknownElement61 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{F7656F4A-A937-7B2E-F7A2-3B946B3073B4}\" />");

            nonVisualDrawingPropertiesExtension61.Append(openXmlUnknownElement61);

            nonVisualDrawingPropertiesExtensionList61.Append(nonVisualDrawingPropertiesExtension61);

            nonVisualDrawingProperties74.Append(nonVisualDrawingPropertiesExtensionList61);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties61 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks61 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties61.Append(shapeLocks61);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties74 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape61 = new PlaceholderShape(){ Type = PlaceholderValues.DateAndTime, Size = PlaceholderSizeValues.Half, Index = (UInt32Value)10U };

            applicationNonVisualDrawingProperties74.Append(placeholderShape61);

            nonVisualShapeProperties61.Append(nonVisualDrawingProperties74);
            nonVisualShapeProperties61.Append(nonVisualShapeDrawingProperties61);
            nonVisualShapeProperties61.Append(applicationNonVisualDrawingProperties74);
            ShapeProperties shapeProperties62 = new ShapeProperties();

            TextBody textBody61 = new TextBody();
            A.BodyProperties bodyProperties62 = new A.BodyProperties();
            A.ListStyle listStyle62 = new A.ListStyle();

            A.Paragraph paragraph97 = new A.Paragraph();

            A.Field field23 = new A.Field(){ Id = "{F550F806-175B-4EE1-ABEA-A5F04E4FD72B}", Type = "datetimeFigureOut" };

            A.RunProperties runProperties85 = new A.RunProperties(){ Language = "pl-PL" };
            runProperties85.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text85 = new A.Text();
            text85.Text = "20.08.2025";

            field23.Append(runProperties85);
            field23.Append(text85);
            A.EndParagraphRunProperties endParagraphRunProperties56 = new A.EndParagraphRunProperties(){ Language = "pl-PL" };

            paragraph97.Append(field23);
            paragraph97.Append(endParagraphRunProperties56);

            textBody61.Append(bodyProperties62);
            textBody61.Append(listStyle62);
            textBody61.Append(paragraph97);

            shape61.Append(nonVisualShapeProperties61);
            shape61.Append(shapeProperties62);
            shape61.Append(textBody61);

            Shape shape62 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties62 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties75 = new NonVisualDrawingProperties(){ Id = (UInt32Value)6U, Name = "Footer Placeholder 5" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList62 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension62 = new A.NonVisualDrawingPropertiesExtension(){ Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            OpenXmlUnknownElement openXmlUnknownElement62 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{09B9912F-47D6-CEA3-3B4E-F02FD11E3781}\" />");

            nonVisualDrawingPropertiesExtension62.Append(openXmlUnknownElement62);

            nonVisualDrawingPropertiesExtensionList62.Append(nonVisualDrawingPropertiesExtension62);

            nonVisualDrawingProperties75.Append(nonVisualDrawingPropertiesExtensionList62);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties62 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks62 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties62.Append(shapeLocks62);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties75 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape62 = new PlaceholderShape(){ Type = PlaceholderValues.Footer, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)11U };

            applicationNonVisualDrawingProperties75.Append(placeholderShape62);

            nonVisualShapeProperties62.Append(nonVisualDrawingProperties75);
            nonVisualShapeProperties62.Append(nonVisualShapeDrawingProperties62);
            nonVisualShapeProperties62.Append(applicationNonVisualDrawingProperties75);
            ShapeProperties shapeProperties63 = new ShapeProperties();

            TextBody textBody62 = new TextBody();
            A.BodyProperties bodyProperties63 = new A.BodyProperties();
            A.ListStyle listStyle63 = new A.ListStyle();

            A.Paragraph paragraph98 = new A.Paragraph();
            A.EndParagraphRunProperties endParagraphRunProperties57 = new A.EndParagraphRunProperties(){ Language = "pl-PL" };

            paragraph98.Append(endParagraphRunProperties57);

            textBody62.Append(bodyProperties63);
            textBody62.Append(listStyle63);
            textBody62.Append(paragraph98);

            shape62.Append(nonVisualShapeProperties62);
            shape62.Append(shapeProperties63);
            shape62.Append(textBody62);

            Shape shape63 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties63 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties76 = new NonVisualDrawingProperties(){ Id = (UInt32Value)7U, Name = "Slide Number Placeholder 6" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList63 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension63 = new A.NonVisualDrawingPropertiesExtension(){ Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            OpenXmlUnknownElement openXmlUnknownElement63 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{610ED5CD-E0F0-7DAC-3A3D-6C8EE773E7EC}\" />");

            nonVisualDrawingPropertiesExtension63.Append(openXmlUnknownElement63);

            nonVisualDrawingPropertiesExtensionList63.Append(nonVisualDrawingPropertiesExtension63);

            nonVisualDrawingProperties76.Append(nonVisualDrawingPropertiesExtensionList63);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties63 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks63 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties63.Append(shapeLocks63);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties76 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape63 = new PlaceholderShape(){ Type = PlaceholderValues.SlideNumber, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)12U };

            applicationNonVisualDrawingProperties76.Append(placeholderShape63);

            nonVisualShapeProperties63.Append(nonVisualDrawingProperties76);
            nonVisualShapeProperties63.Append(nonVisualShapeDrawingProperties63);
            nonVisualShapeProperties63.Append(applicationNonVisualDrawingProperties76);
            ShapeProperties shapeProperties64 = new ShapeProperties();

            TextBody textBody63 = new TextBody();
            A.BodyProperties bodyProperties64 = new A.BodyProperties();
            A.ListStyle listStyle64 = new A.ListStyle();

            A.Paragraph paragraph99 = new A.Paragraph();

            A.Field field24 = new A.Field(){ Id = "{5FAB34AF-B09A-4E2A-BAD3-CCD6E0AD81E6}", Type = "slidenum" };

            A.RunProperties runProperties86 = new A.RunProperties(){ Language = "pl-PL" };
            runProperties86.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text86 = new A.Text();
            text86.Text = "‹#›";

            field24.Append(runProperties86);
            field24.Append(text86);
            A.EndParagraphRunProperties endParagraphRunProperties58 = new A.EndParagraphRunProperties(){ Language = "pl-PL" };

            paragraph99.Append(field24);
            paragraph99.Append(endParagraphRunProperties58);

            textBody63.Append(bodyProperties64);
            textBody63.Append(listStyle64);
            textBody63.Append(paragraph99);

            shape63.Append(nonVisualShapeProperties63);
            shape63.Append(shapeProperties64);
            shape63.Append(textBody63);

            shapeTree13.Append(nonVisualGroupShapeProperties13);
            shapeTree13.Append(groupShapeProperties13);
            shapeTree13.Append(shape58);
            shapeTree13.Append(shape59);
            shapeTree13.Append(shape60);
            shapeTree13.Append(shape61);
            shapeTree13.Append(shape62);
            shapeTree13.Append(shape63);

            CommonSlideDataExtensionList commonSlideDataExtensionList13 = new CommonSlideDataExtensionList();

            CommonSlideDataExtension commonSlideDataExtension13 = new CommonSlideDataExtension(){ Uri = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}" };

            P14.CreationId creationId13 = new P14.CreationId(){ Val = (UInt32Value)4258389826U };
            creationId13.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            commonSlideDataExtension13.Append(creationId13);

            commonSlideDataExtensionList13.Append(commonSlideDataExtension13);

            commonSlideData13.Append(shapeTree13);
            commonSlideData13.Append(commonSlideDataExtensionList13);

            ColorMapOverride colorMapOverride12 = new ColorMapOverride();
            A.MasterColorMapping masterColorMapping12 = new A.MasterColorMapping();

            colorMapOverride12.Append(masterColorMapping12);

            slideLayout11.Append(commonSlideData13);
            slideLayout11.Append(colorMapOverride12);

            slideLayoutPart11.SlideLayout = slideLayout11;
        }

        // Generates content of tableStylesPart1.
        private void GenerateTableStylesPart1Content(TableStylesPart tableStylesPart1)
        {
            A.TableStyleList tableStyleList1 = new A.TableStyleList(){ Default = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}" };
            tableStyleList1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            tableStylesPart1.TableStyleList = tableStyleList1;
        }

        // Generates content of viewPropertiesPart1.
        private void GenerateViewPropertiesPart1Content(ViewPropertiesPart viewPropertiesPart1)
        {
            ViewProperties viewProperties1 = new ViewProperties();
            viewProperties1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            viewProperties1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            viewProperties1.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            NormalViewProperties normalViewProperties1 = new NormalViewProperties();
            RestoredLeft restoredLeft1 = new RestoredLeft(){ Size = 15989, AutoAdjust = false };
            RestoredTop restoredTop1 = new RestoredTop(){ Size = 94660 };

            normalViewProperties1.Append(restoredLeft1);
            normalViewProperties1.Append(restoredTop1);

            SlideViewProperties slideViewProperties1 = new SlideViewProperties();

            CommonSlideViewProperties commonSlideViewProperties1 = new CommonSlideViewProperties(){ SnapToGrid = false };

            CommonViewProperties commonViewProperties1 = new CommonViewProperties(){ VariableScale = true };

            ScaleFactor scaleFactor1 = new ScaleFactor();
            A.ScaleX scaleX1 = new A.ScaleX(){ Numerator = 142, Denominator = 100 };
            A.ScaleY scaleY1 = new A.ScaleY(){ Numerator = 142, Denominator = 100 };

            scaleFactor1.Append(scaleX1);
            scaleFactor1.Append(scaleY1);
            Origin origin1 = new Origin(){ X = 84L, Y = 316L };

            commonViewProperties1.Append(scaleFactor1);
            commonViewProperties1.Append(origin1);
            GuideList guideList1 = new GuideList();

            commonSlideViewProperties1.Append(commonViewProperties1);
            commonSlideViewProperties1.Append(guideList1);

            slideViewProperties1.Append(commonSlideViewProperties1);

            NotesTextViewProperties notesTextViewProperties1 = new NotesTextViewProperties();

            CommonViewProperties commonViewProperties2 = new CommonViewProperties();

            ScaleFactor scaleFactor2 = new ScaleFactor();
            A.ScaleX scaleX2 = new A.ScaleX(){ Numerator = 1, Denominator = 1 };
            A.ScaleY scaleY2 = new A.ScaleY(){ Numerator = 1, Denominator = 1 };

            scaleFactor2.Append(scaleX2);
            scaleFactor2.Append(scaleY2);
            Origin origin2 = new Origin(){ X = 0L, Y = 0L };

            commonViewProperties2.Append(scaleFactor2);
            commonViewProperties2.Append(origin2);

            notesTextViewProperties1.Append(commonViewProperties2);
            GridSpacing gridSpacing1 = new GridSpacing(){ Cx = 72008L, Cy = 72008L };

            viewProperties1.Append(normalViewProperties1);
            viewProperties1.Append(slideViewProperties1);
            viewProperties1.Append(notesTextViewProperties1);
            viewProperties1.Append(gridSpacing1);

            viewPropertiesPart1.ViewProperties = viewProperties1;
        }

        // Generates content of extendedFilePropertiesPart1.
        private void GenerateExtendedFilePropertiesPart1Content(ExtendedFilePropertiesPart extendedFilePropertiesPart1)
        {
            Ap.Properties properties1 = new Ap.Properties();
            properties1.AddNamespaceDeclaration("vt", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes");
            Ap.TotalTime totalTime1 = new Ap.TotalTime();
            totalTime1.Text = "0";
            Ap.Words words1 = new Ap.Words();
            words1.Text = "0";
            Ap.Application application1 = new Ap.Application();
            application1.Text = "Microsoft Office PowerPoint";
            Ap.PresentationFormat presentationFormat1 = new Ap.PresentationFormat();
            presentationFormat1.Text = "Widescreen";
            Ap.Paragraphs paragraphs1 = new Ap.Paragraphs();
            paragraphs1.Text = "0";
            Ap.Slides slides1 = new Ap.Slides();
            slides1.Text = "1";
            Ap.Notes notes1 = new Ap.Notes();
            notes1.Text = "0";
            Ap.HiddenSlides hiddenSlides1 = new Ap.HiddenSlides();
            hiddenSlides1.Text = "0";
            Ap.MultimediaClips multimediaClips1 = new Ap.MultimediaClips();
            multimediaClips1.Text = "0";
            Ap.ScaleCrop scaleCrop1 = new Ap.ScaleCrop();
            scaleCrop1.Text = "false";

            Ap.HeadingPairs headingPairs1 = new Ap.HeadingPairs();

            Vt.VTVector vTVector1 = new Vt.VTVector(){ BaseType = Vt.VectorBaseValues.Variant, Size = (UInt32Value)6U };

            Vt.Variant variant1 = new Vt.Variant();
            Vt.VTLPSTR vTLPSTR1 = new Vt.VTLPSTR();
            vTLPSTR1.Text = "Fonts Used";

            variant1.Append(vTLPSTR1);

            Vt.Variant variant2 = new Vt.Variant();
            Vt.VTInt32 vTInt321 = new Vt.VTInt32();
            vTInt321.Text = "3";

            variant2.Append(vTInt321);

            Vt.Variant variant3 = new Vt.Variant();
            Vt.VTLPSTR vTLPSTR2 = new Vt.VTLPSTR();
            vTLPSTR2.Text = "Theme";

            variant3.Append(vTLPSTR2);

            Vt.Variant variant4 = new Vt.Variant();
            Vt.VTInt32 vTInt322 = new Vt.VTInt32();
            vTInt322.Text = "1";

            variant4.Append(vTInt322);

            Vt.Variant variant5 = new Vt.Variant();
            Vt.VTLPSTR vTLPSTR3 = new Vt.VTLPSTR();
            vTLPSTR3.Text = "Slide Titles";

            variant5.Append(vTLPSTR3);

            Vt.Variant variant6 = new Vt.Variant();
            Vt.VTInt32 vTInt323 = new Vt.VTInt32();
            vTInt323.Text = "1";

            variant6.Append(vTInt323);

            vTVector1.Append(variant1);
            vTVector1.Append(variant2);
            vTVector1.Append(variant3);
            vTVector1.Append(variant4);
            vTVector1.Append(variant5);
            vTVector1.Append(variant6);

            headingPairs1.Append(vTVector1);

            Ap.TitlesOfParts titlesOfParts1 = new Ap.TitlesOfParts();

            Vt.VTVector vTVector2 = new Vt.VTVector(){ BaseType = Vt.VectorBaseValues.Lpstr, Size = (UInt32Value)5U };
            Vt.VTLPSTR vTLPSTR4 = new Vt.VTLPSTR();
            vTLPSTR4.Text = "Aptos";
            Vt.VTLPSTR vTLPSTR5 = new Vt.VTLPSTR();
            vTLPSTR5.Text = "Aptos Display";
            Vt.VTLPSTR vTLPSTR6 = new Vt.VTLPSTR();
            vTLPSTR6.Text = "Arial";
            Vt.VTLPSTR vTLPSTR7 = new Vt.VTLPSTR();
            vTLPSTR7.Text = "Office Theme";
            Vt.VTLPSTR vTLPSTR8 = new Vt.VTLPSTR();
            vTLPSTR8.Text = "PowerPoint Presentation";

            vTVector2.Append(vTLPSTR4);
            vTVector2.Append(vTLPSTR5);
            vTVector2.Append(vTLPSTR6);
            vTVector2.Append(vTLPSTR7);
            vTVector2.Append(vTLPSTR8);

            titlesOfParts1.Append(vTVector2);
            Ap.Company company1 = new Ap.Company();
            company1.Text = "";
            Ap.LinksUpToDate linksUpToDate1 = new Ap.LinksUpToDate();
            linksUpToDate1.Text = "false";
            Ap.SharedDocument sharedDocument1 = new Ap.SharedDocument();
            sharedDocument1.Text = "false";
            Ap.HyperlinksChanged hyperlinksChanged1 = new Ap.HyperlinksChanged();
            hyperlinksChanged1.Text = "false";
            Ap.ApplicationVersion applicationVersion1 = new Ap.ApplicationVersion();
            applicationVersion1.Text = "16.0000";

            properties1.Append(totalTime1);
            properties1.Append(words1);
            properties1.Append(application1);
            properties1.Append(presentationFormat1);
            properties1.Append(paragraphs1);
            properties1.Append(slides1);
            properties1.Append(notes1);
            properties1.Append(hiddenSlides1);
            properties1.Append(multimediaClips1);
            properties1.Append(scaleCrop1);
            properties1.Append(headingPairs1);
            properties1.Append(titlesOfParts1);
            properties1.Append(company1);
            properties1.Append(linksUpToDate1);
            properties1.Append(sharedDocument1);
            properties1.Append(hyperlinksChanged1);
            properties1.Append(applicationVersion1);

            extendedFilePropertiesPart1.Properties = properties1;
        }

        private void SetPackageProperties(OpenXmlPackage document)
        {
            document.PackageProperties.Creator = "Przemysław Kłys";
            document.PackageProperties.Title = "";
            document.PackageProperties.Revision = "1";
            document.PackageProperties.Created = System.Xml.XmlConvert.ToDateTime("2025-08-20T16:14:14Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
            document.PackageProperties.Modified = System.Xml.XmlConvert.ToDateTime("2025-08-20T16:14:52Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
            document.PackageProperties.LastModifiedBy = "Przemysław Kłys";
        }

        #region Binary Data
        private string thumbnailPart1Data = "/9j/4AAQSkZJRgABAQEAYABgAAD/2wBDAAMCAgMCAgMDAwMEAwMEBQgFBQQEBQoHBwYIDAoMDAsKCwsNDhIQDQ4RDgsLEBYQERMUFRUVDA8XGBYUGBIUFRT/2wBDAQMEBAUEBQkFBQkUDQsNFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBT/wAARCACQAQADASIAAhEBAxEB/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAAAgEDAwIEAwUFBAQAAAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJicoKSo0NTY3ODk6Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uHi4+Tl5ufo6erx8vP09fb3+Pn6/8QAHwEAAwEBAQEBAQEBAQAAAAAAAAECAwQFBgcICQoL/8QAtREAAgECBAQDBAcFBAQAAQJ3AAECAxEEBSExBhJBUQdhcRMiMoEIFEKRobHBCSMzUvAVYnLRChYkNOEl8RcYGRomJygpKjU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6goOEhYaHiImKkpOUlZaXmJmaoqOkpaanqKmqsrO0tba3uLm6wsPExcbHyMnK0tPU1dbX2Nna4uPk5ebn6Onq8vP09fb3+Pn6/9oADAMBAAIRAxEAPwD9U6KKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooA//Z";

        private System.IO.Stream GetBinaryDataStream(string base64String)
        {
            return new System.IO.MemoryStream(System.Convert.FromBase64String(base64String));
        }

        #endregion

    }
}
