using System;
using System.IO;
using System.Linq;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf {
    public partial class PdfComposePageOptionsTests {
        [Fact]
        public void StandardFontNames_RejectInvalidValuesInsteadOfFallingBack() {
            Assert.Equal("Helvetica", PdfStandardFont.Helvetica.ToBaseFontName());
            Assert.Equal("Times-Roman", PdfStandardFont.TimesRoman.ToBaseFontName());
            Assert.Equal("Courier-BoldOblique", PdfStandardFont.CourierBoldOblique.ToBaseFontName());

            var exception = Assert.Throws<ArgumentOutOfRangeException>(() =>
                ((PdfStandardFont)99).ToBaseFontName());

            Assert.Equal("font", exception.ParamName);
            Assert.Contains("PDF font must be one of the supported standard PDF fonts.", exception.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void StandardFontDictionaryBuilder_EmitsType1WinAnsiFontObjects() {
            Assert.Equal(
                "<< /Type /Font /Subtype /Type1 /BaseFont /Times-BoldItalic /Encoding /WinAnsiEncoding >>\n",
                PdfStandardFontDictionaryBuilder.BuildStandardType1FontObject(PdfStandardFont.TimesBoldItalic));
            Assert.Equal(
                "<< /Type /Font /Subtype /Type1 /BaseFont /Times-BoldItalic /Encoding /WinAnsiEncoding /ToUnicode 7 0 R >>\n",
                PdfStandardFontDictionaryBuilder.BuildStandardType1FontObject(PdfStandardFont.TimesBoldItalic, 7));

            PdfDictionary dictionary = PdfStandardFontDictionaryBuilder.BuildStandardType1FontDictionary(PdfStandardFont.CourierOblique);

            Assert.Equal("Font", Assert.IsType<PdfName>(dictionary.Items["Type"]).Name);
            Assert.Equal("Type1", Assert.IsType<PdfName>(dictionary.Items["Subtype"]).Name);
            Assert.Equal("Courier-Oblique", Assert.IsType<PdfName>(dictionary.Items["BaseFont"]).Name);
            Assert.Equal("WinAnsiEncoding", Assert.IsType<PdfName>(dictionary.Items["Encoding"]).Name);

            Assert.Throws<ArgumentOutOfRangeException>(() =>
                PdfStandardFontDictionaryBuilder.BuildStandardType1FontObject((PdfStandardFont)99));
            Assert.Throws<ArgumentOutOfRangeException>(() =>
                PdfStandardFontDictionaryBuilder.BuildStandardType1FontObject(PdfStandardFont.Helvetica, -1));
        }

        [Fact]
        public void CatalogDictionaryBuilder_EmitsGeneratedCatalogsAndSharedEntries() {
            Assert.Equal(
                "<< /Type /Catalog /Pages 2 0 R >>\n",
                PdfCatalogDictionaryBuilder.BuildGeneratedCatalogDictionary(2, 0));

            Assert.Equal(
                "<< /Type /Catalog /Pages 2 0 R /Outlines 5 0 R /PageMode /UseOutlines >>\n",
                PdfCatalogDictionaryBuilder.BuildGeneratedCatalogDictionary(2, 5));

            Assert.Equal(
                "<< /Type /Catalog /Pages 2 0 R /Names << /Dests 7 0 R >> >>\n",
                PdfCatalogDictionaryBuilder.BuildGeneratedCatalogDictionary(2, 0, 7));

            Assert.Equal(
                "<< /Type /Catalog /Pages 2 0 R /Outlines 5 0 R /PageMode /UseOutlines /Names << /Dests 7 0 R >> >>\n",
                PdfCatalogDictionaryBuilder.BuildGeneratedCatalogDictionary(2, 5, 7));

            Assert.Equal(
                "<< /Type /Catalog /Pages 2 0 R /AcroForm 8 0 R >>\n",
                PdfCatalogDictionaryBuilder.BuildGeneratedCatalogDictionary(2, 0, 0, 8));

            Assert.Equal(
                "<< /Type /Catalog /Pages 2 0 R /Metadata 9 0 R >>\n",
                PdfCatalogDictionaryBuilder.BuildGeneratedCatalogDictionary(2, 0, 0, 0, 9));

            Assert.Equal(
                "<< /Type /Catalog /Pages 2 0 R /OutputIntents [10 0 R] >>\n",
                PdfCatalogDictionaryBuilder.BuildGeneratedCatalogDictionary(2, 0, 0, 0, 0, 10));

            Assert.Equal(
                "<< /Type /Catalog /Pages 2 0 R /Metadata 9 0 R /OutputIntents [10 0 R] >>\n",
                PdfCatalogDictionaryBuilder.BuildGeneratedCatalogDictionary(2, 0, 0, 0, 9, 10));

            Assert.Equal(
                "<< /Type /Catalog /Pages 2 0 R /Lang <656E2D5553> >>\n",
                PdfCatalogDictionaryBuilder.BuildGeneratedCatalogDictionary(2, 0, 0, 0, 0, 0, "en-US"));

            Assert.Equal(
                "<< /Type /Catalog /Pages 2 0 R /Lang <656E2D5553> /Metadata 9 0 R /OutputIntents [10 0 R] >>\n",
                PdfCatalogDictionaryBuilder.BuildGeneratedCatalogDictionary(2, 0, 0, 0, 9, 10, "en-US"));

            Assert.Equal(
                "<< /Type /Catalog /Pages 2 0 R /PageLabels 14 0 R >>\n",
                PdfCatalogDictionaryBuilder.BuildGeneratedCatalogDictionary(2, 0, pageLabelsId: 14));

            Assert.Equal(
                "<< /Type /Catalog /Pages 2 0 R /Lang <656E2D5553> /PageLabels 14 0 R /Metadata 9 0 R /OutputIntents [10 0 R] >>\n",
                PdfCatalogDictionaryBuilder.BuildGeneratedCatalogDictionary(2, 0, 0, 0, 9, 10, "en-US", pageLabelsId: 14));

            Assert.Equal(
                "<< /Type /Catalog /Pages 2 0 R /ViewerPreferences 15 0 R >>\n",
                PdfCatalogDictionaryBuilder.BuildGeneratedCatalogDictionary(2, 0, viewerPreferencesId: 15));

            Assert.Equal(
                "<< /Type /Catalog /Pages 2 0 R /OpenAction [3 0 R /XYZ 0 700 0] >>\n",
                PdfCatalogDictionaryBuilder.BuildGeneratedCatalogDictionary(2, 0, openAction: "[3 0 R /XYZ 0 700 0]"));

            Assert.Equal(
                "<< /Type /Catalog /Pages 2 0 R /URI << /Base (https://evotec.xyz/docs\\(pdf\\)/) >> >>\n",
                PdfCatalogDictionaryBuilder.BuildGeneratedCatalogDictionary(2, 0, catalogUriBase: "https://evotec.xyz/docs(pdf)/"));

            Assert.Equal(
                "<< /Type /Catalog /Pages 2 0 R /PageMode /FullScreen /PageLayout /TwoColumnLeft >>\n",
                PdfCatalogDictionaryBuilder.BuildGeneratedCatalogDictionary(2, 0, pageMode: "FullScreen", pageLayout: "TwoColumnLeft"));

            Assert.Equal(
                "<< /Type /Catalog /Pages 2 0 R /Outlines 5 0 R /PageMode /UseThumbs /PageLayout /SinglePage >>\n",
                PdfCatalogDictionaryBuilder.BuildGeneratedCatalogDictionary(2, 5, pageMode: "UseThumbs", pageLayout: "SinglePage"));

            Assert.Equal(
                "<< /Type /Catalog /Pages 2 0 R /Lang <656E2D5553> /PageLabels 14 0 R /ViewerPreferences 15 0 R /Metadata 9 0 R /OutputIntents [10 0 R] >>\n",
                PdfCatalogDictionaryBuilder.BuildGeneratedCatalogDictionary(2, 0, 0, 0, 9, 10, "en-US", pageLabelsId: 14, viewerPreferencesId: 15));

            Assert.Equal(
                "<< /Type /Catalog /Pages 2 0 R /Names << /EmbeddedFiles 11 0 R >> /AF [12 0 R] >>\n",
                PdfCatalogDictionaryBuilder.BuildGeneratedCatalogDictionary(2, 0, embeddedFilesNameTreeId: 11, associatedFileIds: new[] { 12 }));

            Assert.Equal(
                "<< /Type /Catalog /Pages 2 0 R /Names << /Dests 7 0 R /EmbeddedFiles 11 0 R >> /AF [12 0 R 13 0 R] >>\n",
                PdfCatalogDictionaryBuilder.BuildGeneratedCatalogDictionary(2, 0, 7, embeddedFilesNameTreeId: 11, associatedFileIds: new[] { 12, 13 }));

            var sb = new StringBuilder();
            PdfCatalogDictionaryBuilder.AppendCatalogStart(sb, 3);
            PdfCatalogDictionaryBuilder.AppendNameEntry(sb, "PageLayout", "TwoColumnLeft");
            PdfCatalogDictionaryBuilder.AppendTextStringEntry(sb, "Lang", "pl-PL");
            PdfCatalogDictionaryBuilder.AppendReferenceEntry(sb, "Outlines", 9);
            sb.Append(" >>\n");

            Assert.Equal("<< /Type /Catalog /Pages 3 0 R /PageLayout /TwoColumnLeft /Lang <706C2D504C> /Outlines 9 0 R >>\n", sb.ToString());
            Assert.Throws<ArgumentOutOfRangeException>(() => PdfCatalogDictionaryBuilder.BuildGeneratedCatalogDictionary(0, 0));
            Assert.Throws<ArgumentOutOfRangeException>(() => PdfCatalogDictionaryBuilder.BuildGeneratedCatalogDictionary(2, -1));
            Assert.Throws<ArgumentOutOfRangeException>(() => PdfCatalogDictionaryBuilder.BuildGeneratedCatalogDictionary(2, 0, -1));
            Assert.Throws<ArgumentOutOfRangeException>(() => PdfCatalogDictionaryBuilder.BuildGeneratedCatalogDictionary(2, 0, 0, 0, -1));
            Assert.Throws<ArgumentOutOfRangeException>(() => PdfCatalogDictionaryBuilder.BuildGeneratedCatalogDictionary(2, 0, 0, -1));
            Assert.Throws<ArgumentOutOfRangeException>(() => PdfCatalogDictionaryBuilder.BuildGeneratedCatalogDictionary(2, 0, 0, 0, 0, -1));
            Assert.Throws<ArgumentOutOfRangeException>(() => PdfCatalogDictionaryBuilder.BuildGeneratedCatalogDictionary(2, 0, embeddedFilesNameTreeId: -1));
            Assert.Throws<ArgumentOutOfRangeException>(() => PdfCatalogDictionaryBuilder.BuildGeneratedCatalogDictionary(2, 0, associatedFileIds: new[] { 0 }));
            Assert.Throws<ArgumentOutOfRangeException>(() => PdfCatalogDictionaryBuilder.BuildGeneratedCatalogDictionary(2, 0, pageLabelsId: -1));
            Assert.Throws<ArgumentOutOfRangeException>(() => PdfCatalogDictionaryBuilder.BuildGeneratedCatalogDictionary(2, 0, viewerPreferencesId: -1));
            Assert.Throws<ArgumentException>(() => PdfCatalogDictionaryBuilder.BuildGeneratedCatalogDictionary(2, 0, openAction: ""));
            Assert.Throws<ArgumentException>(() => PdfCatalogDictionaryBuilder.BuildGeneratedCatalogDictionary(2, 0, pageMode: ""));
            Assert.Throws<ArgumentException>(() => PdfCatalogDictionaryBuilder.BuildGeneratedCatalogDictionary(2, 0, pageLayout: ""));
            Assert.Throws<ArgumentException>(() => PdfCatalogDictionaryBuilder.BuildGeneratedCatalogDictionary(2, 0, catalogUriBase: ""));
            Assert.Throws<ArgumentException>(() => PdfCatalogDictionaryBuilder.BuildGeneratedCatalogDictionary(2, 0, 0, 0, 0, 0, ""));
        }

        [Fact]
        public void PageLabelDictionaryBuilder_EmitsSimpleNumberTree() {
            Assert.Equal(
                "<< /Nums [0 << /S /D /St 1 >>] >>\n",
                PdfPageLabelDictionaryBuilder.BuildGeneratedPageLabelsDictionary(PdfPageNumberStyle.Arabic, 1));

            Assert.Equal(
                "<< /Nums [0 << /S /R /St 5 /P <412D> >>] >>\n",
                PdfPageLabelDictionaryBuilder.BuildGeneratedPageLabelsDictionary(PdfPageNumberStyle.UpperRoman, 5, "A-"));

            Assert.Equal("r", PdfPageLabelDictionaryBuilder.GetStyleName(PdfPageNumberStyle.LowerRoman));
            Assert.Equal("a", PdfPageLabelDictionaryBuilder.GetStyleName(PdfPageNumberStyle.LowerLetter));
            Assert.Equal("A", PdfPageLabelDictionaryBuilder.GetStyleName(PdfPageNumberStyle.UpperLetter));
            Assert.Throws<ArgumentOutOfRangeException>(() =>
                PdfPageLabelDictionaryBuilder.BuildGeneratedPageLabelsDictionary(PdfPageNumberStyle.Arabic, 0));
            Assert.Throws<ArgumentException>(() =>
                PdfPageLabelDictionaryBuilder.BuildGeneratedPageLabelsDictionary(PdfPageNumberStyle.Arabic, 1, ""));
            Assert.Throws<ArgumentException>(() =>
                PdfPageLabelDictionaryBuilder.BuildGeneratedPageLabelsDictionary((PdfPageNumberStyle)99, 1));
        }

        [Fact]
        public void PageLabelDictionaryBuilder_EmitsMultipleNumberTreeRanges() {
            string dictionary = PdfPageLabelDictionaryBuilder.BuildGeneratedPageLabelsDictionary(new[] {
                new PdfPageLabelRange(3, PdfPageNumberStyle.Arabic, 1),
                new PdfPageLabelRange(1, PdfPageNumberStyle.LowerRoman, 1, "front-"),
                new PdfPageLabelRange(7, PdfPageNumberStyle.UpperLetter, 2, "A-")
            });

            Assert.Equal("<< /Nums [0 << /S /r /St 1 /P <66726F6E742D> >> 2 << /S /D /St 1 >> 6 << /S /A /St 2 /P <412D> >>] >>\n", dictionary);
            Assert.Throws<ArgumentException>(() => PdfPageLabelDictionaryBuilder.BuildGeneratedPageLabelsDictionary(Array.Empty<PdfPageLabelRange>()));
            Assert.Throws<ArgumentException>(() => PdfPageLabelDictionaryBuilder.BuildGeneratedPageLabelsDictionary(new[] {
                new PdfPageLabelRange(1, PdfPageNumberStyle.Arabic),
                new PdfPageLabelRange(1, PdfPageNumberStyle.LowerRoman)
            }));
        }

        [Fact]
        public void ViewerPreferenceDictionaryBuilder_EmitsConfiguredEntries() {
            var preferences = new PdfViewerPreferencesOptions {
                HideToolbar = true,
                FitWindow = false,
                DisplayDocTitle = true,
                NonFullScreenPageMode = PdfNonFullScreenPageMode.UseThumbs,
                Direction = PdfViewerDirection.RightToLeft,
                PrintScaling = PdfPrintScaling.None,
                Duplex = PdfDuplexMode.DuplexFlipShortEdge,
                ViewArea = PdfPageBoundaryBox.CropBox,
                ViewClip = PdfPageBoundaryBox.BleedBox,
                PrintArea = PdfPageBoundaryBox.TrimBox,
                PrintClip = PdfPageBoundaryBox.ArtBox,
                PickTrayByPdfSize = true,
                NumCopies = 3
            }.AddPrintPageRange(1, 1).AddPrintPageRange(3, 5);

            Assert.Equal(
                "<< /HideToolbar true /FitWindow false /DisplayDocTitle true /PickTrayByPDFSize true /NonFullScreenPageMode /UseThumbs /Direction /R2L /PrintScaling /None /Duplex /DuplexFlipShortEdge /ViewArea /CropBox /ViewClip /BleedBox /PrintArea /TrimBox /PrintClip /ArtBox /NumCopies 3 /PrintPageRange [1 1 3 5] >>\n",
                PdfViewerPreferenceDictionaryBuilder.BuildGeneratedViewerPreferencesDictionary(preferences));

            Assert.Throws<ArgumentNullException>(() =>
                PdfViewerPreferenceDictionaryBuilder.BuildGeneratedViewerPreferencesDictionary(null!));
            Assert.Throws<ArgumentException>(() =>
                PdfViewerPreferenceDictionaryBuilder.BuildGeneratedViewerPreferencesDictionary(new PdfViewerPreferencesOptions()));
        }

        [Fact]
        public void OpenActionBuilder_EmitsDestinationArray() {
            Assert.Equal(
                "[8 0 R /XYZ 0 712.25 0]",
                PdfCatalogDictionaryBuilder.BuildGeneratedOpenActionDestination(8, 712.25));

            Assert.Equal(
                "[8 0 R /Fit]",
                PdfCatalogDictionaryBuilder.BuildGeneratedOpenActionDestination(8, 712.25, PdfOpenActionDestinationMode.Fit));

            Assert.Equal(
                "[8 0 R /FitH 712.25]",
                PdfCatalogDictionaryBuilder.BuildGeneratedOpenActionDestination(8, 712.25, PdfOpenActionDestinationMode.FitHorizontal));

            Assert.Equal(
                "[8 0 R /FitV 24.5]",
                PdfCatalogDictionaryBuilder.BuildGeneratedOpenActionDestination(8, 712.25, PdfOpenActionDestinationMode.FitVertical, 24.5));

            Assert.Equal(
                "[8 0 R /FitR 24.5 12.25 300.75 712.25]",
                PdfCatalogDictionaryBuilder.BuildGeneratedOpenActionDestination(8, 712.25, PdfOpenActionDestinationMode.FitRectangle, 24.5, 12.25, 300.75));

            Assert.Equal(
                "[8 0 R /FitB]",
                PdfCatalogDictionaryBuilder.BuildGeneratedOpenActionDestination(8, 712.25, PdfOpenActionDestinationMode.FitBoundingBox));

            Assert.Equal(
                "[8 0 R /FitBH 712.25]",
                PdfCatalogDictionaryBuilder.BuildGeneratedOpenActionDestination(8, 712.25, PdfOpenActionDestinationMode.FitBoundingBoxHorizontal));

            Assert.Equal(
                "[8 0 R /FitBV 24.5]",
                PdfCatalogDictionaryBuilder.BuildGeneratedOpenActionDestination(8, 712.25, PdfOpenActionDestinationMode.FitBoundingBoxVertical, 24.5));

            Assert.Throws<ArgumentOutOfRangeException>(() =>
                PdfCatalogDictionaryBuilder.BuildGeneratedOpenActionDestination(0, 712.25));
            Assert.Throws<ArgumentOutOfRangeException>(() =>
                PdfCatalogDictionaryBuilder.BuildGeneratedOpenActionDestination(8, double.NaN));
            Assert.Throws<ArgumentOutOfRangeException>(() =>
                PdfCatalogDictionaryBuilder.BuildGeneratedOpenActionDestination(8, 712.25, PdfOpenActionDestinationMode.FitRectangle, 24.5, 12.25, 20));
            Assert.Throws<ArgumentOutOfRangeException>(() =>
                PdfCatalogDictionaryBuilder.BuildGeneratedOpenActionDestination(8, 712.25, (PdfOpenActionDestinationMode)99));
        }

        [Fact]
        public void CatalogViewBuilders_MapTypedModeAndLayoutNames() {
            Assert.Equal("UseNone", PdfCatalogDictionaryBuilder.GetPageModeName(PdfCatalogPageMode.UseNone));
            Assert.Equal("UseOutlines", PdfCatalogDictionaryBuilder.GetPageModeName(PdfCatalogPageMode.UseOutlines));
            Assert.Equal("UseThumbs", PdfCatalogDictionaryBuilder.GetPageModeName(PdfCatalogPageMode.UseThumbs));
            Assert.Equal("FullScreen", PdfCatalogDictionaryBuilder.GetPageModeName(PdfCatalogPageMode.FullScreen));
            Assert.Equal("UseOC", PdfCatalogDictionaryBuilder.GetPageModeName(PdfCatalogPageMode.UseOC));
            Assert.Equal("UseAttachments", PdfCatalogDictionaryBuilder.GetPageModeName(PdfCatalogPageMode.UseAttachments));
            Assert.Equal("SinglePage", PdfCatalogDictionaryBuilder.GetPageLayoutName(PdfCatalogPageLayout.SinglePage));
            Assert.Equal("OneColumn", PdfCatalogDictionaryBuilder.GetPageLayoutName(PdfCatalogPageLayout.OneColumn));
            Assert.Equal("TwoColumnLeft", PdfCatalogDictionaryBuilder.GetPageLayoutName(PdfCatalogPageLayout.TwoColumnLeft));
            Assert.Equal("TwoColumnRight", PdfCatalogDictionaryBuilder.GetPageLayoutName(PdfCatalogPageLayout.TwoColumnRight));
            Assert.Equal("TwoPageLeft", PdfCatalogDictionaryBuilder.GetPageLayoutName(PdfCatalogPageLayout.TwoPageLeft));
            Assert.Equal("TwoPageRight", PdfCatalogDictionaryBuilder.GetPageLayoutName(PdfCatalogPageLayout.TwoPageRight));
            Assert.Throws<ArgumentOutOfRangeException>(() => PdfCatalogDictionaryBuilder.GetPageModeName((PdfCatalogPageMode)99));
            Assert.Throws<ArgumentOutOfRangeException>(() => PdfCatalogDictionaryBuilder.GetPageLayoutName((PdfCatalogPageLayout)99));
        }

        [Fact]
        public void ViewerPreferenceBuilders_MapTypedNameValues() {
            Assert.Equal("UseNone", PdfViewerPreferenceDictionaryBuilder.GetNonFullScreenPageModeName(PdfNonFullScreenPageMode.UseNone));
            Assert.Equal("UseOutlines", PdfViewerPreferenceDictionaryBuilder.GetNonFullScreenPageModeName(PdfNonFullScreenPageMode.UseOutlines));
            Assert.Equal("UseThumbs", PdfViewerPreferenceDictionaryBuilder.GetNonFullScreenPageModeName(PdfNonFullScreenPageMode.UseThumbs));
            Assert.Equal("UseOC", PdfViewerPreferenceDictionaryBuilder.GetNonFullScreenPageModeName(PdfNonFullScreenPageMode.UseOC));
            Assert.Equal("L2R", PdfViewerPreferenceDictionaryBuilder.GetDirectionName(PdfViewerDirection.LeftToRight));
            Assert.Equal("R2L", PdfViewerPreferenceDictionaryBuilder.GetDirectionName(PdfViewerDirection.RightToLeft));
            Assert.Equal("AppDefault", PdfViewerPreferenceDictionaryBuilder.GetPrintScalingName(PdfPrintScaling.AppDefault));
            Assert.Equal("None", PdfViewerPreferenceDictionaryBuilder.GetPrintScalingName(PdfPrintScaling.None));
            Assert.Equal("Simplex", PdfViewerPreferenceDictionaryBuilder.GetDuplexName(PdfDuplexMode.Simplex));
            Assert.Equal("DuplexFlipShortEdge", PdfViewerPreferenceDictionaryBuilder.GetDuplexName(PdfDuplexMode.DuplexFlipShortEdge));
            Assert.Equal("DuplexFlipLongEdge", PdfViewerPreferenceDictionaryBuilder.GetDuplexName(PdfDuplexMode.DuplexFlipLongEdge));
            Assert.Equal("MediaBox", PdfViewerPreferenceDictionaryBuilder.GetPageBoundaryBoxName(PdfPageBoundaryBox.MediaBox));
            Assert.Equal("CropBox", PdfViewerPreferenceDictionaryBuilder.GetPageBoundaryBoxName(PdfPageBoundaryBox.CropBox));
            Assert.Equal("BleedBox", PdfViewerPreferenceDictionaryBuilder.GetPageBoundaryBoxName(PdfPageBoundaryBox.BleedBox));
            Assert.Equal("TrimBox", PdfViewerPreferenceDictionaryBuilder.GetPageBoundaryBoxName(PdfPageBoundaryBox.TrimBox));
            Assert.Equal("ArtBox", PdfViewerPreferenceDictionaryBuilder.GetPageBoundaryBoxName(PdfPageBoundaryBox.ArtBox));
            Assert.Throws<ArgumentOutOfRangeException>(() => PdfViewerPreferenceDictionaryBuilder.GetNonFullScreenPageModeName((PdfNonFullScreenPageMode)99));
            Assert.Throws<ArgumentOutOfRangeException>(() => PdfViewerPreferenceDictionaryBuilder.GetDirectionName((PdfViewerDirection)99));
            Assert.Throws<ArgumentOutOfRangeException>(() => PdfViewerPreferenceDictionaryBuilder.GetPrintScalingName((PdfPrintScaling)99));
            Assert.Throws<ArgumentOutOfRangeException>(() => PdfViewerPreferenceDictionaryBuilder.GetDuplexName((PdfDuplexMode)99));
            Assert.Throws<ArgumentOutOfRangeException>(() => PdfViewerPreferenceDictionaryBuilder.GetPageBoundaryBoxName((PdfPageBoundaryBox)99));
        }

        [Fact]
        public void PageDictionaryBuilder_EmitsGeneratedPageWithResourcesAndAnnotations() {
            var fonts = new[] { ("F1", 4) };
            var xobjects = new[] { ("/Im1", 5) };
            var graphicsStates = new[] { ("/GS1", 6) };
            var shadings = new[] { ("/Sh1", 7) };
            var annotations = new[] { 8, 9 };

            string page = PdfPageDictionaryBuilder.BuildGeneratedPageDictionary(
                2,
                612,
                792,
                10,
                fonts,
                xobjects,
                graphicsStates,
                shadings,
                annotations);

            Assert.Equal(
                "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 4 0 R >> /XObject << /Im1 5 0 R >> /ExtGState << /GS1 6 0 R >> /Shading << /Sh1 7 0 R >> >> /Contents 10 0 R /Annots [ 8 0 R 9 0 R ] >>\n",
                page);

            Assert.Equal(
                " /XObject << /Image#201 3 0 R >>",
                PdfPageDictionaryBuilder.BuildResourcePart("XObject", new[] { ("Image 1", 3) }));

            Assert.Throws<ArgumentOutOfRangeException>(() =>
                PdfPageDictionaryBuilder.BuildGeneratedPageDictionary(2, 0, 792, 10, fonts, xobjects, graphicsStates, shadings, annotations));
        }

        [Fact]
        public void PageDictionaryBuilder_EmitsTaggedStructureTabOrder() {
            string page = PdfPageDictionaryBuilder.BuildGeneratedPageDictionary(
                2,
                612,
                792,
                10,
                Array.Empty<(string Name, int Id)>(),
                Array.Empty<(string Name, int Id)>(),
                Array.Empty<(string Name, int Id)>(),
                Array.Empty<(string Name, int Id)>(),
                Array.Empty<int>(),
                structParents: 0,
                useStructureTabOrder: true);

            Assert.Equal(
                "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << >> /Contents 10 0 R /StructParents 0 /Tabs /S >>\n",
                page);
        }

        [Fact]
        public void StructTreeRootDictionaryBuilder_EmitsParentTreeNextKey() {
            var parentTreeEntries = new[] {
                PdfStructTreeRootDictionaryBuilder.ParentTreeEntry.ForMarkedContentPage(0, new[] { 6 }),
                PdfStructTreeRootDictionaryBuilder.ParentTreeEntry.ForObjectReference(1, 7)
            };

            string parentTree = PdfStructTreeRootDictionaryBuilder.BuildParentTree(parentTreeEntries);
            string document = PdfStructTreeRootDictionaryBuilder.BuildDocumentStructElement(3, new[] { 6, 7 });
            string root = PdfStructTreeRootDictionaryBuilder.BuildStructTreeRootDictionary(new[] { 6, 7 }, parentTreeId: 8, parentTreeNextKey: 2);

            Assert.Equal("<< /Nums [0 [6 0 R] 1 7 0 R] >>\n", parentTree);
            Assert.Equal("<< /Type /StructElem /S /Document /P 3 0 R /K [6 0 R 7 0 R] >>\n", document);
            Assert.Equal(
                "<< /Type /StructElem /S /Document /P 3 0 R /K [6 0 R 7 0 R] /Lang <656E2D5553> >>\n",
                PdfStructTreeRootDictionaryBuilder.BuildDocumentStructElement(3, new[] { 6, 7 }, "en-US"));
            Assert.Equal(
                "<< /Type /StructElem /S /Link /P 3 0 R /Pg 4 0 R /K [<< /Type /MCR /Pg 4 0 R /MCID 5 >> << /Type /OBJR /Obj 9 0 R >>] >>\n",
                PdfStructTreeRootDictionaryBuilder.BuildAnnotationStructElement(3, 4, 9, 5));
            Assert.Equal(
                "<< /Type /StructElem /S /Link /P 3 0 R /Pg 4 0 R /K [<< /Type /MCR /Pg 4 0 R /MCID 5 >> << /Type /MCR /Pg 4 0 R /MCID 6 >> << /Type /OBJR /Obj 9 0 R >>] >>\n",
                PdfStructTreeRootDictionaryBuilder.BuildAnnotationStructElement(3, 4, 9, 5, new[] { 6 }));
            Assert.Equal("<< /Type /StructTreeRoot /K [6 0 R 7 0 R] /ParentTree 8 0 R /ParentTreeNextKey 2 /RoleMap << >> >>\n", root);
            Assert.Throws<ArgumentOutOfRangeException>(() =>
                PdfStructTreeRootDictionaryBuilder.BuildStructTreeRootDictionary(new[] { 6 }, parentTreeId: 8, parentTreeNextKey: -1));
        }

        [Fact]
        public void EmbeddedFileDictionaryBuilder_EmitsDeterministicStreamParameters() {
            var file = new PdfEmbeddedFile(
                "invoice.xml",
                new byte[] { 1, 2, 3 },
                "application/xml",
                PdfAssociatedFileRelationship.Data,
                "Invoice XML");

            string dictionary = PdfEmbeddedFileDictionaryBuilder.BuildEmbeddedFileStreamDictionary(file, new byte[] { 1, 2, 3 });

            Assert.Equal(
                "<< /Type /EmbeddedFile /Subtype /application#2Fxml /Length 3 /Params << /Size 3 /CheckSum <5289DF737DF57326FCDD22597AFB1FAC> >> >>",
                dictionary);
            Assert.Throws<ArgumentNullException>(() =>
                PdfEmbeddedFileDictionaryBuilder.BuildEmbeddedFileStreamDictionary(file, null!));
            Assert.Throws<ArgumentException>(() =>
                PdfEmbeddedFileDictionaryBuilder.BuildEmbeddedFileStreamDictionary(file, Array.Empty<byte>()));
        }

        [Fact]
        public void AnnotationDictionaryBuilder_EmitsUriLinkAnnotationsWithEscapedUri() {
            Assert.Equal(
                "<< /Type /Annot /Subtype /Link /Border [0 0 0] /Rect [10 20.5 110 44.25] /A << /S /URI /URI (https://evotec.xyz/docs\\(pdf\\)) >> >>\n",
                PdfAnnotationDictionaryBuilder.BuildUriLinkAnnotation(10, 20.5, 110, 44.25, "https://evotec.xyz/docs(pdf)"));

            Assert.Equal(
                "<< /Type /Annot /Subtype /Link /Border [0 0 0] /Contents (Jump metadata) /Rect [10 20.5 110 44.25] /A << /S /GoTo /D (Intro\\(A\\)) >> >>\n",
                PdfAnnotationDictionaryBuilder.BuildGoToNamedDestinationLinkAnnotation(10, 20.5, 110, 44.25, "Intro(A)", "Jump metadata"));

            Assert.Equal(
                "<< /Type /Annot /Subtype /Link /Border [0 0 0] /Contents (Jump metadata) /Rect [10 20.5 110 44.25] /A << /S /URI /URI (https://evotec.xyz/docs) >> /StructParent 7 >>\n",
                PdfAnnotationDictionaryBuilder.BuildUriLinkAnnotation(10, 20.5, 110, 44.25, "https://evotec.xyz/docs", "Jump metadata", 7));

            Assert.Equal(
                "<< /Type /Annot /Subtype /Link /Border [0 0 0] /Rect [10 20.5 110 44.25] /A << /S /URI /URI (assets/report.html#summary) >> >>\n",
                PdfAnnotationDictionaryBuilder.BuildUriLinkAnnotation(10, 20.5, 110, 44.25, "assets/report.html#summary"));

            Assert.Equal(
                "<< /Type /Annot /Subtype /Link /Border [0 0 0] /Rect [10 20.5 110 44.25] /A << /S /GoTo /D (Intro) >> /StructParent 8 >>\n",
                PdfAnnotationDictionaryBuilder.BuildGoToNamedDestinationLinkAnnotation(10, 20.5, 110, 44.25, "Intro", structParentIndex: 8));

            Assert.Equal(
                "<< /Type /Annot /Subtype /Text /Rect [10 20.5 26 36.5] /Contents (Review \\(layout\\)) /Name /Key /C [1 0 0] /Open true >>\n",
                PdfAnnotationDictionaryBuilder.BuildTextAnnotation(10, 20.5, 26, 36.5, "Review (layout)", PdfTextAnnotationIcon.Key, new PdfColor(1D, 0D, 0D), open: true));

            Assert.Equal(
                "<< /Type /Annot /Subtype /FreeText /Rect [10 20.5 110 44.25] /Contents (Reviewer note) /DA (/Helv 9.5 Tf 0 0 1 rg) /Border [0 0 0.75] /C [1 0 0] /IC [1 1 0.8] >>\n",
                PdfAnnotationDictionaryBuilder.BuildFreeTextAnnotation(10, 20.5, 110, 44.25, "Reviewer note", 9.5, new PdfColor(0D, 0D, 1D), new PdfColor(1D, 0D, 0D), 0.75, new PdfColor(1D, 1D, 0.8D)));
            string freeTextAppearance = PdfAnnotationDictionaryBuilder.BuildFreeTextAppearanceContent(
                70,
                48,
                "Reviewer note wraps across words",
                10,
                new PdfColor(0D, 0D, 1D),
                new PdfColor(1D, 0D, 0D),
                0.75,
                new PdfColor(1D, 1D, 0.8D),
                PdfAlign.Center,
                4D,
                11D);
            Assert.True(CountOccurrences(freeTextAppearance, " Tj ET") >= 2);
            Assert.Contains("BT /Helv 10 Tf 0 0 1 rg", freeTextAppearance, StringComparison.Ordinal);
            Assert.Contains("1 1 0.8 rg 0 0 70 48 re f", freeTextAppearance, StringComparison.Ordinal);

            Assert.Equal(
                "<< /Type /Annot /Subtype /Highlight /Rect [10 20.5 110 44.25] /Contents (Important) /C [1 0.9 0.1] /QuadPoints [10 44.25 110 44.25 10 20.5 110 20.5] >>\n",
                PdfAnnotationDictionaryBuilder.BuildHighlightAnnotation(10, 20.5, 110, 44.25, "Important", new PdfColor(1D, 0.9D, 0.1D)));

            Assert.Equal(
                "<< /Type /Annot /Subtype /Widget /FT /Tx /T <506572736F6E2E4E616D65> /V <416461> /DV <416461> /Rect [10 20.5 110 44.25] /F 4 /DA (/Helv 10 Tf 0 0 0 rg) /MK << /BC [0.75 0.75 0.75] /BG [1 1 1] >> /AP << /N 12 0 R >> >>\n",
                PdfAnnotationDictionaryBuilder.BuildTextFieldWidgetAnnotation(10, 20.5, 110, 44.25, "Person.Name", "Ada", 10, 12));
            Assert.Contains(
                "/TU <506572736F6E206E616D65> /TM <706572736F6E2E6E616D65>",
                PdfAnnotationDictionaryBuilder.BuildTextFieldWidgetAnnotation(
                    10,
                    20.5,
                    110,
                    44.25,
                    "Person.Name",
                    "Ada",
                    10,
                    12,
                    new PdfFormFieldStyle {
                        AlternateName = "Person name",
                        MappingName = "person.name"
                    }),
                StringComparison.Ordinal);
            var flaggedStyle = new PdfFormFieldStyle {
                IsReadOnly = true,
                IsRequired = true,
                IsNoExport = true
            };
            Assert.Contains(
                "/FT /Tx /T <506572736F6E2E4E616D65> /Ff 7 /V <416461>",
                PdfAnnotationDictionaryBuilder.BuildTextFieldWidgetAnnotation(10, 20.5, 110, 44.25, "Person.Name", "Ada", 10, 12, flaggedStyle),
                StringComparison.Ordinal);
            Assert.Contains(
                "/FT /Tx /T <506572736F6E2E4E616D65> /MaxLen 64 /V <416461>",
                PdfAnnotationDictionaryBuilder.BuildTextFieldWidgetAnnotation(
                    10,
                    20.5,
                    110,
                    44.25,
                    "Person.Name",
                    "Ada",
                    10,
                    12,
                    new PdfFormFieldStyle {
                        MaxLength = 64
                    }),
                StringComparison.Ordinal);
            var textFlagStyle = new PdfFormFieldStyle {
                IsMultiline = true,
                IsPassword = true,
                DoesNotSpellCheck = true,
                DoesNotScroll = true
            };
            Assert.Contains(
                "/FT /Tx /T <506572736F6E2E4E616D65> /Ff 12595200 /V <416461>",
                PdfAnnotationDictionaryBuilder.BuildTextFieldWidgetAnnotation(10, 20.5, 110, 44.25, "Person.Name", "Ada", 10, 12, textFlagStyle),
                StringComparison.Ordinal);
            Assert.Contains(
                "/FT /Tx /T <506572736F6E2E4E616D65> /Ff 1048576 /V <416461>",
                PdfAnnotationDictionaryBuilder.BuildTextFieldWidgetAnnotation(
                    10,
                    20.5,
                    110,
                    44.25,
                    "Person.Name",
                    "Ada",
                    10,
                    12,
                    new PdfFormFieldStyle {
                        IsFileSelect = true
                    }),
                StringComparison.Ordinal);
            Assert.Contains(
                "/FT /Tx /T <506572736F6E2E4E616D65> /Ff 16777216 /MaxLen 4 /V <416461>",
                PdfAnnotationDictionaryBuilder.BuildTextFieldWidgetAnnotation(
                    10,
                    20.5,
                    110,
                    44.25,
                    "Person.Name",
                    "Ada",
                    10,
                    12,
                    new PdfFormFieldStyle {
                        IsComb = true,
                        MaxLength = 4
                    }),
                StringComparison.Ordinal);
            Assert.Throws<ArgumentException>(() =>
                PdfAnnotationDictionaryBuilder.BuildTextFieldWidgetAnnotation(
                    10,
                    20.5,
                    110,
                    44.25,
                    "Person.Name",
                    "Ada",
                    10,
                    12,
                    new PdfFormFieldStyle {
                        IsComb = true
                    }));

            Assert.Equal(
                "<< /Type /Annot /Subtype /Widget /FT /Btn /T <4163636570745465726D73> /V /Yes /DV /Yes /Rect [10 20.5 26 36.5] /F 4 /AS /Yes /MK << /BC [0.75 0.75 0.75] /BG [1 1 1] >> /AP << /N << /Off 12 0 R /Yes 13 0 R >> >> >>\n",
                PdfAnnotationDictionaryBuilder.BuildCheckBoxWidgetAnnotation(10, 20.5, 26, 36.5, "AcceptTerms", true, "Yes", 12, 13));
            Assert.Contains(
                "/FT /Btn /T <4163636570745465726D73> /Ff 7 /V /Yes",
                PdfAnnotationDictionaryBuilder.BuildCheckBoxWidgetAnnotation(10, 20.5, 26, 36.5, "AcceptTerms", true, "Yes", 12, 13, flaggedStyle),
                StringComparison.Ordinal);

            Assert.Equal(
                "<< /Type /Annot /Subtype /Widget /FT /Ch /T <436F756E747279> /V <506F6C616E64> /DV <506F6C616E64> /Opt [ <506F6C616E64> <556E6974656420537461746573> ] /Ff 131072 /Rect [10 20.5 110 44.25] /F 4 /DA (/Helv 10 Tf 0 0 0 rg) /MK << /BC [0.75 0.75 0.75] /BG [1 1 1] >> /AP << /N 12 0 R >> >>\n",
                PdfAnnotationDictionaryBuilder.BuildChoiceFieldWidgetAnnotation(10, 20.5, 110, 44.25, "Country", new[] { "Poland", "United States" }, "Poland", 10, 12, isComboBox: true));
            Assert.Contains(
                "/Ff 131079",
                PdfAnnotationDictionaryBuilder.BuildChoiceFieldWidgetAnnotation(10, 20.5, 110, 44.25, "Country", new[] { "Poland", "United States" }, "Poland", 10, 12, isComboBox: true, style: flaggedStyle),
                StringComparison.Ordinal);
            Assert.Contains(
                "/Ff 4325376",
                PdfAnnotationDictionaryBuilder.BuildChoiceFieldWidgetAnnotation(10, 20.5, 110, 44.25, "Country", new[] { "Poland", "United States" }, "Poland", 10, 12, isComboBox: true, style: textFlagStyle),
                StringComparison.Ordinal);
            var choiceFlagStyle = new PdfFormFieldStyle {
                IsEditableChoice = true,
                IsSortedChoice = true,
                CommitsOnSelectionChange = true
            };
            Assert.Contains(
                "/Ff 68026368",
                PdfAnnotationDictionaryBuilder.BuildChoiceFieldWidgetAnnotation(10, 20.5, 110, 44.25, "Country", new[] { "Poland", "United States" }, "Poland", 10, 12, isComboBox: true, style: choiceFlagStyle),
                StringComparison.Ordinal);

            Assert.Equal(
                "<< /Type /Annot /Subtype /Widget /FT /Ch /T <436F756E7472696573> /V [<506F6C616E64> <556E6974656420537461746573>] /DV [<506F6C616E64> <556E6974656420537461746573>] /Opt [ <506F6C616E64> <4765726D616E79> <556E6974656420537461746573> ] /Ff 2097152 /Rect [10 20.5 110 70] /F 4 /DA (/Helv 10 Tf 0 0 0 rg) /MK << /BC [0.75 0.75 0.75] /BG [1 1 1] >> /AP << /N 12 0 R >> >>\n",
                PdfAnnotationDictionaryBuilder.BuildChoiceFieldWidgetAnnotation(10, 20.5, 110, 70, "Countries", new[] { "Poland", "Germany", "United States" }, new[] { "Poland", "United States" }, 10, 12, isComboBox: false, allowsMultipleSelection: true));
            Assert.Contains(
                "/Ff 69730304",
                PdfAnnotationDictionaryBuilder.BuildChoiceFieldWidgetAnnotation(10, 20.5, 110, 70, "Countries", new[] { "Poland", "Germany", "United States" }, new[] { "Poland", "United States" }, 10, 12, isComboBox: false, allowsMultipleSelection: true, style: choiceFlagStyle),
                StringComparison.Ordinal);
            Assert.Contains(
                "/FT /Btn /T <436F6E74616374> /Ff 49159 /V /Email",
                PdfAnnotationDictionaryBuilder.BuildRadioButtonFieldDictionary("Contact", new[] { "Email", "Phone" }, "Email", new[] { 12, 13 }, flaggedStyle),
                StringComparison.Ordinal);

            Assert.Contains("/T <FEFF540D>", PdfAnnotationDictionaryBuilder.BuildTextFieldWidgetAnnotation(10, 20, 110, 44, "名", "Ada", 10, 12), StringComparison.Ordinal);

            Assert.Throws<ArgumentException>(() =>
                PdfAnnotationDictionaryBuilder.BuildUriLinkAnnotation(10, 20, 110, 44, "bad\u0001uri"));
            Assert.Throws<ArgumentOutOfRangeException>(() =>
                PdfAnnotationDictionaryBuilder.BuildUriLinkAnnotation(10, 20, 10, 44, "https://evotec.xyz"));
            Assert.Throws<ArgumentOutOfRangeException>(() =>
                PdfAnnotationDictionaryBuilder.BuildUriLinkAnnotation(10, 20, 110, double.NaN, "https://evotec.xyz"));
            Assert.Throws<ArgumentException>(() =>
                PdfAnnotationDictionaryBuilder.BuildGoToNamedDestinationLinkAnnotation(10, 20, 110, 44, " "));
            Assert.Throws<ArgumentException>(() =>
                PdfAnnotationDictionaryBuilder.BuildTextAnnotation(10, 20, 110, 44, " "));
            Assert.Throws<ArgumentException>(() =>
                PdfAnnotationDictionaryBuilder.BuildFreeTextAnnotation(10, 20, 110, 44, " "));
            Assert.Throws<ArgumentOutOfRangeException>(() =>
                PdfAnnotationDictionaryBuilder.BuildFreeTextAnnotation(10, 20, 110, 44, "note", 0));
            Assert.Throws<ArgumentException>(() =>
                PdfAnnotationDictionaryBuilder.BuildFreeTextAppearanceContent(100, 40, "note", textAlign: PdfAlign.Justify));
            Assert.Throws<ArgumentOutOfRangeException>(() =>
                PdfAnnotationDictionaryBuilder.BuildFreeTextAppearanceContent(100, 40, "note", padding: -1));
            Assert.Throws<ArgumentOutOfRangeException>(() =>
                PdfAnnotationDictionaryBuilder.BuildFreeTextAppearanceContent(100, 40, "note", lineHeight: 0));
            Assert.Throws<ArgumentException>(() =>
                PdfAnnotationDictionaryBuilder.BuildHighlightAnnotation(10, 20, 110, 44, " "));
            Assert.Throws<ArgumentException>(() =>
                PdfAnnotationDictionaryBuilder.BuildCheckBoxWidgetAnnotation(10, 20, 26, 36, "AcceptTerms", true, "Off", 12, 13));
            Assert.Throws<ArgumentException>(() =>
                PdfAnnotationDictionaryBuilder.BuildCheckBoxWidgetAnnotation(10, 20, 26, 36, "AcceptTerms", true, "Y\u2713", 12, 13));
            Assert.Throws<ArgumentException>(() =>
                PdfAnnotationDictionaryBuilder.BuildChoiceFieldWidgetAnnotation(10, 20, 110, 44, "Country", Array.Empty<string>(), "Poland", 10, 12, isComboBox: true));
            Assert.Throws<ArgumentException>(() =>
                PdfAnnotationDictionaryBuilder.BuildChoiceFieldWidgetAnnotation(10, 20, 110, 44, "Country", new[] { "Poland" }, "Germany", 10, 12, isComboBox: true));
            Assert.Throws<ArgumentException>(() =>
                PdfAnnotationDictionaryBuilder.BuildChoiceFieldWidgetAnnotation(10, 20, 110, 44, "Countries", new[] { "Poland" }, Array.Empty<string>(), 10, 12, isComboBox: false, allowsMultipleSelection: true));
            Assert.Throws<ArgumentException>(() =>
                PdfAnnotationDictionaryBuilder.BuildChoiceFieldWidgetAnnotation(10, 20, 110, 44, "Countries", new[] { "Poland" }, new[] { "Poland", "Poland" }, 10, 12, isComboBox: false, allowsMultipleSelection: true));
            Assert.Throws<ArgumentException>(() =>
                PdfAnnotationDictionaryBuilder.BuildChoiceFieldWidgetAnnotation(10, 20, 110, 44, "Countries", new[] { "Poland", "Germany" }, new[] { "Poland", "Germany" }, 10, 12, isComboBox: true, allowsMultipleSelection: true));
        }

        [Fact]
        public void AcroFormDictionaryBuilder_EmitsFieldsAndTextAppearance() {
            Assert.Equal(
                "<< /Fields [ 4 0 R 5 0 R ] /NeedAppearances false /DR << /Font << /Helv 3 0 R >> >> /DA (/Helv 10 Tf 0 g) >>\n",
                PdfAcroFormDictionaryBuilder.BuildAcroFormDictionary(new[] { 4, 5 }, 3));

            Assert.Equal(
                "<< /Fields [ 4 0 R ] /NeedAppearances false /DR << /Font << /Helv 3 0 R >> >> /DA (/Helv 10 Tf 0 g) /Q 1 >>\n",
                PdfAcroFormDictionaryBuilder.BuildAcroFormDictionary(new[] { 4 }, 3, PdfFormFieldTextAlignment.Center));

            string content = PdfAcroFormDictionaryBuilder.BuildTextFieldAppearanceContent(120, 20, "Ada", 10);
            Assert.Contains("<416461> Tj", content);

            string passwordContent = PdfAcroFormDictionaryBuilder.BuildTextFieldAppearanceContent(120, 20, "Secret", 10, new PdfFormFieldStyle { IsPassword = true });
            Assert.Contains("<2A2A2A2A2A2A> Tj", passwordContent);
            Assert.DoesNotContain("<536563726574> Tj", passwordContent);

            string centeredContent = PdfAcroFormDictionaryBuilder.BuildTextFieldAppearanceContent(120, 20, "Ada", 10, textAlignment: PdfFormFieldTextAlignment.Center, textWidth: 30);
            Assert.Contains("45 12.2 Td <416461> Tj", centeredContent);

            string rightContent = PdfAcroFormDictionaryBuilder.BuildTextFieldAppearanceContent(120, 20, "Ada", 10, textAlignment: PdfFormFieldTextAlignment.Right, textWidth: 30);
            Assert.Contains("87 12.2 Td <416461> Tj", rightContent);

            string dictionary = PdfAcroFormDictionaryBuilder.BuildTextFieldAppearanceStreamDictionary(120, 20, 3, Encoding.ASCII.GetByteCount(content));
            Assert.Equal(
                "<< /Type /XObject /Subtype /Form /BBox [0 0 120 20] /Resources << /Font << /Helv 3 0 R >> >> /Length " + Encoding.ASCII.GetByteCount(content) + " >>",
                dictionary);

            string checkedContent = PdfAcroFormDictionaryBuilder.BuildCheckBoxAppearanceContent(16, 16, selected: true);
            Assert.Contains(" l S", checkedContent);

            string checkBoxDictionary = PdfAcroFormDictionaryBuilder.BuildCheckBoxAppearanceStreamDictionary(16, 16, Encoding.ASCII.GetByteCount(checkedContent));
            Assert.Equal(
                "<< /Type /XObject /Subtype /Form /BBox [0 0 16 16] /Length " + Encoding.ASCII.GetByteCount(checkedContent) + " >>",
                checkBoxDictionary);

            Assert.Throws<ArgumentException>(() => PdfAcroFormDictionaryBuilder.BuildAcroFormDictionary(Array.Empty<int>(), 3));
            Assert.Throws<ArgumentOutOfRangeException>(() => PdfAcroFormDictionaryBuilder.BuildAcroFormDictionary(new[] { 4 }, 3, (PdfFormFieldTextAlignment)999));
            Assert.Throws<ArgumentOutOfRangeException>(() => PdfAcroFormDictionaryBuilder.BuildTextFieldAppearanceContent(120, 20, "Ada", 10, textAlignment: PdfFormFieldTextAlignment.Unknown));
            Assert.Throws<ArgumentOutOfRangeException>(() => PdfAcroFormDictionaryBuilder.BuildTextFieldAppearanceContent(120, 20, "Ada", 10, textWidth: -1));
            Assert.Throws<ArgumentOutOfRangeException>(() => PdfAcroFormDictionaryBuilder.BuildTextFieldAppearanceStreamDictionary(120, 20, 3, -1));
            Assert.Throws<ArgumentOutOfRangeException>(() => PdfAcroFormDictionaryBuilder.BuildCheckBoxAppearanceStreamDictionary(16, 16, -1));
        }

        [Fact]
        public void OutlineDictionaryBuilder_EmitsRootAndNestedItems() {
            Assert.Equal(
                "<< /Type /Outlines /First 6 0 R /Last 9 0 R /Count 3 >>\n",
                PdfOutlineDictionaryBuilder.BuildOutlineRoot(6, 9, 3));

            Assert.Equal(
                "<< /Title (Chapter \\(A\\)) /Parent 5 0 R /Prev 6 0 R /Next 8 0 R /First 10 0 R /Last 11 0 R /Count 2 /Dest [3 0 R /XYZ 0 712.25 0] >>\n",
                PdfOutlineDictionaryBuilder.BuildOutlineItem("Chapter (A)", 5, 6, 8, 10, 11, 2, 3, 712.25));

            Assert.Equal(
                "<< /Title (Closed) /Parent 5 0 R /First 10 0 R /Last 11 0 R /Count -2 /Dest [3 0 R /XYZ 0 712.25 0] >>\n",
                PdfOutlineDictionaryBuilder.BuildOutlineItem("Closed", 5, 0, 0, 10, 11, -2, 3, 712.25));

            Assert.Equal(
                "<< /Title (Leaf) /Parent 5 0 R /Dest [3 0 R /XYZ 0 0 0] >>\n",
                PdfOutlineDictionaryBuilder.BuildOutlineItem("Leaf", 5, 0, 0, 0, 0, 0, 3, 0));

            Assert.Throws<ArgumentOutOfRangeException>(() =>
                PdfOutlineDictionaryBuilder.BuildOutlineRoot(6, 9, -1));
            Assert.Throws<ArgumentOutOfRangeException>(() =>
                PdfOutlineDictionaryBuilder.BuildOutlineItem("Bad", 5, 0, 0, 0, 0, 0, 3, double.PositiveInfinity));
        }

        [Fact]
        public void VisualResourceDictionaryBuilder_EmitsOpacityAndAxialShadingResources() {
            Assert.Equal(
                "<< /Type /ExtGState /ca 0.35 /CA 0.75 >>\n",
                PdfVisualResourceDictionaryBuilder.BuildExtGStateObject(0.35, 0.75));

            Assert.Equal(
                "<< /ShadingType 2 /ColorSpace /DeviceRGB /Coords [30 118 120 118] /Function << /FunctionType 2 /Domain [0 1] /C0 [0.039 0.078 0.118] /C1 [1 0.502 0] /N 1 >> /Extend [true true] >>\n",
                PdfVisualResourceDictionaryBuilder.BuildAxialShadingObject(
                    30,
                    118,
                    120,
                    118,
                    OfficeColor.FromRgb(10, 20, 30),
                    OfficeColor.FromRgb(255, 128, 0)));

            Assert.Throws<ArgumentOutOfRangeException>(() =>
                PdfVisualResourceDictionaryBuilder.BuildExtGStateObject(-0.1, 1));
            Assert.Throws<ArgumentOutOfRangeException>(() =>
                PdfVisualResourceDictionaryBuilder.BuildExtGStateObject(1, 1.1));
            Assert.Throws<ArgumentOutOfRangeException>(() =>
                PdfVisualResourceDictionaryBuilder.BuildAxialShadingObject(double.NaN, 0, 1, 1, OfficeColor.Black, OfficeColor.White));
        }

        private static int CountOccurrences(string text, string value) {
            int count = 0;
            int startIndex = 0;
            while (true) {
                int index = text.IndexOf(value, startIndex, StringComparison.Ordinal);
                if (index < 0) {
                    return count;
                }

                count++;
                startIndex = index + value.Length;
            }
        }

    }
}
