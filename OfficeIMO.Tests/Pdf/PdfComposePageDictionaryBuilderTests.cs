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
        public void ViewerPreferenceDictionaryBuilder_EmitsConfiguredBooleans() {
            var preferences = new PdfViewerPreferencesOptions {
                HideToolbar = true,
                FitWindow = false,
                DisplayDocTitle = true
            };

            Assert.Equal(
                "<< /HideToolbar true /FitWindow false /DisplayDocTitle true >>\n",
                PdfViewerPreferenceDictionaryBuilder.BuildGeneratedViewerPreferencesDictionary(preferences));

            Assert.Throws<ArgumentNullException>(() =>
                PdfViewerPreferenceDictionaryBuilder.BuildGeneratedViewerPreferencesDictionary(null!));
            Assert.Throws<ArgumentException>(() =>
                PdfViewerPreferenceDictionaryBuilder.BuildGeneratedViewerPreferencesDictionary(new PdfViewerPreferencesOptions()));
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
                "<< /Type /Annot /Subtype /Link /Border [0 0 0] /Rect [10 20.5 110 44.25] /A << /S /GoTo /D (Intro) >> /StructParent 8 >>\n",
                PdfAnnotationDictionaryBuilder.BuildGoToNamedDestinationLinkAnnotation(10, 20.5, 110, 44.25, "Intro", structParentIndex: 8));

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

            Assert.Equal(
                "<< /Type /Annot /Subtype /Widget /FT /Btn /T <4163636570745465726D73> /V /Yes /DV /Yes /Rect [10 20.5 26 36.5] /F 4 /AS /Yes /MK << /BC [0.75 0.75 0.75] /BG [1 1 1] >> /AP << /N << /Off 12 0 R /Yes 13 0 R >> >> >>\n",
                PdfAnnotationDictionaryBuilder.BuildCheckBoxWidgetAnnotation(10, 20.5, 26, 36.5, "AcceptTerms", true, "Yes", 12, 13));

            Assert.Equal(
                "<< /Type /Annot /Subtype /Widget /FT /Ch /T <436F756E747279> /V <506F6C616E64> /DV <506F6C616E64> /Opt [ <506F6C616E64> <556E6974656420537461746573> ] /Ff 131072 /Rect [10 20.5 110 44.25] /F 4 /DA (/Helv 10 Tf 0 0 0 rg) /MK << /BC [0.75 0.75 0.75] /BG [1 1 1] >> /AP << /N 12 0 R >> >>\n",
                PdfAnnotationDictionaryBuilder.BuildChoiceFieldWidgetAnnotation(10, 20.5, 110, 44.25, "Country", new[] { "Poland", "United States" }, "Poland", 10, 12, isComboBox: true));

            Assert.Equal(
                "<< /Type /Annot /Subtype /Widget /FT /Ch /T <436F756E7472696573> /V [<506F6C616E64> <556E6974656420537461746573>] /DV [<506F6C616E64> <556E6974656420537461746573>] /Opt [ <506F6C616E64> <4765726D616E79> <556E6974656420537461746573> ] /Ff 2097152 /Rect [10 20.5 110 70] /F 4 /DA (/Helv 10 Tf 0 0 0 rg) /MK << /BC [0.75 0.75 0.75] /BG [1 1 1] >> /AP << /N 12 0 R >> >>\n",
                PdfAnnotationDictionaryBuilder.BuildChoiceFieldWidgetAnnotation(10, 20.5, 110, 70, "Countries", new[] { "Poland", "Germany", "United States" }, new[] { "Poland", "United States" }, 10, 12, isComboBox: false, allowsMultipleSelection: true));

            Assert.Contains("/T <FEFF540D>", PdfAnnotationDictionaryBuilder.BuildTextFieldWidgetAnnotation(10, 20, 110, 44, "名", "Ada", 10, 12), StringComparison.Ordinal);

            Assert.Throws<ArgumentException>(() =>
                PdfAnnotationDictionaryBuilder.BuildUriLinkAnnotation(10, 20, 110, 44, "not-a-uri"));
            Assert.Throws<ArgumentOutOfRangeException>(() =>
                PdfAnnotationDictionaryBuilder.BuildUriLinkAnnotation(10, 20, 10, 44, "https://evotec.xyz"));
            Assert.Throws<ArgumentOutOfRangeException>(() =>
                PdfAnnotationDictionaryBuilder.BuildUriLinkAnnotation(10, 20, 110, double.NaN, "https://evotec.xyz"));
            Assert.Throws<ArgumentException>(() =>
                PdfAnnotationDictionaryBuilder.BuildGoToNamedDestinationLinkAnnotation(10, 20, 110, 44, " "));
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

            string content = PdfAcroFormDictionaryBuilder.BuildTextFieldAppearanceContent(120, 20, "Ada", 10);
            Assert.Contains("<416461> Tj", content);

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
                "<< /Title (Leaf) /Parent 5 0 R /Dest [3 0 R /XYZ 0 0 0] >>\n",
                PdfOutlineDictionaryBuilder.BuildOutlineItem("Leaf", 5, 0, 0, 0, 0, 0, 3, 0));

            Assert.Throws<ArgumentOutOfRangeException>(() =>
                PdfOutlineDictionaryBuilder.BuildOutlineRoot(6, 9, -1));
            Assert.Throws<ArgumentOutOfRangeException>(() =>
                PdfOutlineDictionaryBuilder.BuildOutlineItem("Bad", 5, 0, 0, 0, 0, -1, 3, 10));
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

    }
}
