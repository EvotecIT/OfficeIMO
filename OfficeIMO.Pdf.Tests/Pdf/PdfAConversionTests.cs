using System.Globalization;
using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfAConversionTests {
    [Theory]
    [InlineData(PdfComplianceProfile.PdfA2B, "2b")]
    [InlineData(PdfComplianceProfile.PdfA2U, "2u")]
    [InlineData(PdfComplianceProfile.PdfA3B, "3b")]
    [InlineData(PdfComplianceProfile.PdfA3U, "3u")]
    public void ConvertToPdfA_RemovesActiveContentAndFixesStructureWithoutMovingContent(PdfComplianceProfile profile, string flavour) {
        byte[] source = RawPdf.Build(
            content: RawPdf.Lzw("1 0 0 rg 20 20 80 80 re f\n"),
            contentFilter: "/Filter /LZWDecode",
            pageEntries: "/Annots [5 0 R 6 0 R] /AA << /O 7 0 R >>",
            catalogEntries: "/OpenAction 7 0 R /AA << /WC 7 0 R >> /NeedsRendering true /Names << /JavaScript << /Names [(init) 7 0 R] >> >>",
            extraObjects: new[] {
                "<< /Type /Annot /Subtype /Link /Rect [20 20 100 100] /Border [0 0 0] /A 7 0 R >>",
                "<< /Type /Annot /Subtype /Text /Rect [10 150 30 170] /F 2 /Contents (hidden note) >>",
                "<< /S /JavaScript /JS (app.alert\\('x'\\)) >>"
            });

        PdfAConversionResult result = PdfDocument.Load(new MemoryStream(source)).ConvertToPdfA(new PdfAConversionOptions { Profile = profile });

        Assert.True(result.Succeeded, string.Join(Environment.NewLine, result.Report.Violations));
        IReadOnlyList<PdfAConversionChange> changes = result.Report.Changes;
        Assert.Contains(changes, change => change.Kind == PdfAConversionChangeKind.ActionRemoved && change.Location == "catalog open action");
        Assert.Contains(changes, change => change.Kind == PdfAConversionChangeKind.ActionRemoved && change.Location == "catalog names");
        Assert.Contains(changes, change => change.Kind == PdfAConversionChangeKind.AnnotationRemoved && change.ObjectNumber == 6);
        Assert.Contains(changes, change => change.Kind == PdfAConversionChangeKind.AnnotationFlagsUpdated && change.ObjectNumber == 5);
        Assert.Contains(changes, change => change.Kind == PdfAConversionChangeKind.StreamReencoded && change.ObjectNumber == 4);
        Assert.Contains(changes, change => change.Kind == PdfAConversionChangeKind.OutputIntentAdded);
        Assert.Contains(changes, change => change.Kind == PdfAConversionChangeKind.MetadataReplaced);
        Assert.True(result.Report.HasLoss);

        byte[] converted = result.ToBytes();
        PdfReadDocument output = PdfReadDocument.Open(converted);
        PdfDictionary catalog = output.CatalogDictionary!;
        Assert.False(catalog.Items.ContainsKey("OpenAction"));
        Assert.False(catalog.Items.ContainsKey("AA"));
        Assert.False(catalog.Items.ContainsKey("NeedsRendering"));
        Assert.StartsWith("%PDF-1.7", Encoding.ASCII.GetString(converted, 0, 8), StringComparison.Ordinal);
        PdfDocumentInfo info = PdfInspector.Inspect(converted);
        Assert.True(info.HasOutputIntents);
        Assert.Equal(profile is PdfComplianceProfile.PdfA2B or PdfComplianceProfile.PdfA2U ? 2 : 3, info.XmpMetadata!.PdfAPart);
        PdfDictionary page = output.Pages[0].PageDictionary;
        Assert.False(page.Items.ContainsKey("AA"));
        var annotations = (PdfArray)PdfObjectLookup.Resolve(output.Objects, page.Items["Annots"])!;
        PdfDictionary link = Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(output.Objects, Assert.Single(annotations.Items)));
        Assert.Equal(4D, ((PdfNumber)link.Items["F"]).Value);
        Assert.False(link.Items.ContainsKey("A"));
        Assert.DoesNotContain("LZWDecode", Encoding.ASCII.GetString(converted), StringComparison.Ordinal);

        AssertRendersSame(source, converted);
        AssertVeraPdfPasses(converted, flavour);
    }

    [Theory]
    [InlineData(PdfComplianceProfile.PdfA2B, "2b")]
    [InlineData(PdfComplianceProfile.PdfA3U, "3u")]
    public void ConvertToPdfA_EmbedsAnInstalledSubstituteForAStandardFont(PdfComplianceProfile profile, string flavour) {
        byte[] source = PdfDocument.Create()
            .Meta(title: "Quarterly archive", author: "Records team")
            .H1("Quarterly archive")
            .Paragraph(paragraph => paragraph.AddText("Retained as PDF/A; every glyph keeps its width."))
            .ToBytes();

        PdfAConversionResult result = PdfDocument.Load(new MemoryStream(source)).ConvertToPdfA(new PdfAConversionOptions { Profile = profile });

        if (!HasHelveticaSubstitute()) {
            Assert.False(result.Succeeded);
            Assert.All(result.Report.Violations, violation => Assert.Equal(PdfAConversionRefusalReason.FontNotEmbeddable, violation.Reason));
            return;
        }
        Assert.True(result.Succeeded, string.Join(Environment.NewLine, result.Report.Violations));
        Assert.Contains(result.Report.Changes, change => change.Kind == PdfAConversionChangeKind.FontEmbedded);
        if (profile == PdfComplianceProfile.PdfA3U) {
            Assert.Contains(result.Report.Changes, change => change.Kind == PdfAConversionChangeKind.ToUnicodeAdded);
        }
        byte[] converted = result.ToBytes();
        PdfReadDocument output = PdfReadDocument.Open(converted);
        foreach (PdfIndirectObject indirect in output.Objects.Values) {
            if (indirect.Value is PdfDictionary dictionary && dictionary.Get<PdfName>("Type")?.Name == "Font") {
                PdfDictionary descriptor = Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(output.Objects, dictionary.Items["FontDescriptor"]));
                Assert.True(descriptor.Items.ContainsKey("FontFile2"));
            }
        }
        Assert.Equal("Quarterly archive", result.RequireValue().Properties.Title);

        AssertRendersSame(source, converted);
        AssertVeraPdfPasses(converted, flavour);
    }

    [Fact]
    public void ConvertToPdfA_StoresTheAppearanceViewersDrawForAnnotationsWithoutOne() {
        byte[] source = RawPdf.Build("0 g 0 0 10 10 re f\n", pageEntries: "/Annots [5 0 R]",
            extraObjects: new[] { "<< /Type /Annot /Subtype /Square /Rect [40 40 160 120] /C [0 0 1] /Border [0 0 3] >>" });

        PdfAConversionResult result = PdfDocument.Load(new MemoryStream(source)).ConvertToPdfA();

        Assert.True(result.Succeeded, string.Join(Environment.NewLine, result.Report.Violations));
        Assert.Contains(result.Report.Changes, change => change.Kind == PdfAConversionChangeKind.AnnotationAppearanceUpdated && change.ObjectNumber == 5);
        Assert.Contains(result.Report.Changes, change => change.Kind == PdfAConversionChangeKind.AnnotationFlagsUpdated && change.ObjectNumber == 5);
        AssertRendersSame(source, result.ToBytes());
        AssertVeraPdfPasses(result.ToBytes(), "2b");
    }

    [Fact]
    public void ConvertToPdfA_RemovesEncryptionOnlyWithTheOwnerPassword() {
        byte[] plain = RawPdf.Build("0 g 20 20 50 50 re f\n");
        byte[] encrypted = PdfDocument.Load(new MemoryStream(plain)).Security
            .Encrypt(PdfStandardEncryptionOptions.ForPermissions("owner", PdfStandardPermissions.Print)).Pdf;

        PdfAConversionResult refused = PdfDocument.Load(new MemoryStream(encrypted)).ConvertToPdfA();
        PdfAConversionResult converted = PdfDocument.Load(new MemoryStream(encrypted), new PdfLoadOptions { Password = "owner" }).ConvertToPdfA();

        Assert.Equal(PdfAConversionRefusalReason.EncryptionNotAuthorized, Assert.Single(refused.Report.Violations).Reason);
        Assert.True(converted.Succeeded, string.Join(Environment.NewLine, converted.Report.Violations));
        Assert.Equal(PdfAConversionChangeKind.EncryptionRemoved, converted.Report.Changes[0].Kind);
        Assert.False(PdfInspector.Probe(converted.ToBytes()).HasEncryption);
        AssertVeraPdfPasses(converted.ToBytes(), "2b");
    }

    [Fact]
    public void ConvertToPdfA_KeepsAttachmentsForPdfA3AndRemovesThemForPdfA2() {
        byte[] source = RawPdf.Build(
            "0 g 20 20 50 50 re f\n",
            catalogEntries: "/Names << /EmbeddedFiles << /Names [(data.xml) 5 0 R] >> >>",
            extraObjects: new[] {
                "<< /Type /Filespec /F (data.xml) /EF << /F 6 0 R >> >>",
                RawPdf.Stream("/Type /EmbeddedFile", "<data/>")
            });

        PdfAConversionResult archive3 = PdfDocument.Load(new MemoryStream(source)).ConvertToPdfA(new PdfAConversionOptions { Profile = PdfComplianceProfile.PdfA3B });
        PdfAConversionResult archive2 = PdfDocument.Load(new MemoryStream(source)).ConvertToPdfA(new PdfAConversionOptions { Profile = PdfComplianceProfile.PdfA2B });

        Assert.True(archive3.Succeeded, string.Join(Environment.NewLine, archive3.Report.Violations));
        Assert.Single(archive3.RequireValue().Attachments.Extract());
        Assert.Contains(archive3.Report.Changes, change => change.Kind == PdfAConversionChangeKind.EmbeddedFileUpdated && change.ObjectNumber == 5);
        Assert.True(archive2.Succeeded, string.Join(Environment.NewLine, archive2.Report.Violations));
        Assert.Empty(archive2.RequireValue().Attachments.Extract());
        Assert.Contains(archive2.Report.Changes, change => change.Kind == PdfAConversionChangeKind.EmbeddedFileRemoved && change.LossKind == OfficeIMO.OfficeConversionLossKind.Omission);
        AssertVeraPdfPasses(archive3.ToBytes(), "3b");
        AssertVeraPdfPasses(archive2.ToBytes(), "2b");
    }

    [Theory]
    [InlineData(PdfComplianceProfile.PdfA4, "4")]
    [InlineData(PdfComplianceProfile.PdfA4F, "4f")]
    public void ConvertToPdfA_WritesPdfA4AsPdf20WithXmpAuthoritativeMetadataAndItsActionRules(PdfComplianceProfile profile, string flavour) {
        byte[] source = RawPdf.Build(
            content: "0 g 20 20 80 80 re f\n",
            pageEntries: "/Annots [5 0 R] /AA << /O 7 0 R >>",
            catalogEntries: "/Version /1.7 /OpenAction 6 0 R /AA << /WC 6 0 R >> /Requirements [<< /Type /Requirement /S /EnableJavaScripts >>] " +
                "/Names << /EmbeddedFiles << /Names [(data.xml) 9 0 R] >> >>",
            trailerEntries: "/Info 8 0 R",
            extraObjects: new[] {
                "<< /Type /Annot /Subtype /Link /Rect [20 20 100 100] /Border [0 0 0] /A 7 0 R >>",
                "<< /S /JavaScript /JS (app.alert\\('x'\\)) >>",
                "<< /S /Launch /F (calc.exe) >>",
                "<< /Title (Quarterly archive) /Author (Records team) /CreationDate (D:20240102030405Z) >>",
                "<< /Type /Filespec /F (data.xml) /EF << /F 10 0 R >> >>",
                RawPdf.Stream("/Type /EmbeddedFile", "<data/>")
            });

        PdfAConversionResult result = PdfDocument.Load(new MemoryStream(source)).ConvertToPdfA(new PdfAConversionOptions { Profile = profile });

        Assert.True(result.Succeeded, string.Join(Environment.NewLine, result.Report.Violations));
        Assert.Equal(profile == PdfComplianceProfile.PdfA4 ? "PDF/A-4" : "PDF/A-4f", result.Report.Target);
        Assert.Contains(result.Report.Changes, change => change.Kind == PdfAConversionChangeKind.DocumentInfoUpdated && change.ObjectNumber == 8);
        Assert.Contains(result.Report.Changes, change => change.Kind == PdfAConversionChangeKind.ActionRemoved && change.ObjectNumber == 7);
        byte[] converted = result.ToBytes();
        Assert.StartsWith("%PDF-2.0", Encoding.ASCII.GetString(converted, 0, 8), StringComparison.Ordinal);
        PdfReadDocument output = PdfReadDocument.Open(converted);
        Assert.Null(output.Security.InfoObjectNumber);
        PdfDictionary catalog = output.CatalogDictionary!;
        Assert.False(catalog.Items.ContainsKey("Version"));
        Assert.False(catalog.Items.ContainsKey("Requirements"));
        Assert.False(catalog.Items.ContainsKey("AA"));
        // PDF/A-4 permits JavaScript actions; only the triggers and action types it forbids are removed.
        Assert.True(catalog.Items.ContainsKey("OpenAction"));
        Assert.False(output.Pages[0].PageDictionary.Items.ContainsKey("AA"));
        var link = (PdfDictionary)PdfObjectLookup.Resolve(output.Objects, ((PdfArray)PdfObjectLookup.Resolve(output.Objects, output.Pages[0].PageDictionary.Items["Annots"])!).Items[0])!;
        Assert.False(link.Items.ContainsKey("A"));
        string xmp = Encoding.UTF8.GetString(((PdfStream)PdfObjectLookup.Resolve(output.Objects, catalog.Items["Metadata"])!).Data);
        Assert.Contains("<pdfaid:part>4</pdfaid:part>", xmp, StringComparison.Ordinal);
        Assert.Contains("<pdfaid:rev>2020</pdfaid:rev>", xmp, StringComparison.Ordinal);
        Assert.Contains("Quarterly archive", xmp, StringComparison.Ordinal);
        if (profile == PdfComplianceProfile.PdfA4F) {
            Assert.Contains("<pdfaid:conformance>F</pdfaid:conformance>", xmp, StringComparison.Ordinal);
            Assert.Single(result.RequireValue().Attachments.Extract());
        } else {
            Assert.DoesNotContain("pdfaid:conformance", xmp, StringComparison.Ordinal);
            Assert.Empty(result.RequireValue().Attachments.Extract());
            Assert.Contains(result.Report.Changes, change => change.Kind == PdfAConversionChangeKind.EmbeddedFileRemoved);
        }

        AssertRendersSame(source, converted);
        AssertVeraPdfPasses(converted, flavour);
    }

    [Fact]
    public void ConvertToPdfA_KeepsOnlyModDateInTheInfoDictionaryWhenPieceInfoNeedsItForPdfA4() {
        byte[] source = RawPdf.Build("0 g 20 20 50 50 re f\n",
            catalogEntries: "/PieceInfo << /Producer.App << /LastModified (D:20240102030405Z) /Private << /Data 1 >> >> >>",
            trailerEntries: "/Info 5 0 R",
            extraObjects: new[] { "<< /Title (Kept in XMP) /ModDate (D:20240102030405Z) >>" });

        PdfAConversionResult result = PdfDocument.Load(new MemoryStream(source)).ConvertToPdfA(new PdfAConversionOptions { Profile = PdfComplianceProfile.PdfA4 });

        Assert.True(result.Succeeded, string.Join(Environment.NewLine, result.Report.Violations));
        PdfReadDocument output = PdfReadDocument.Open(result.ToBytes());
        PdfDictionary info = Assert.IsType<PdfDictionary>(output.Objects[output.Security.InfoObjectNumber!.Value].Value);
        Assert.Equal("ModDate", Assert.Single(info.Items.Keys));
        AssertVeraPdfPasses(result.ToBytes(), "4");
    }

    [Fact]
    public void ConvertToPdfA_ListsAnnotationOnlyAttachmentsInTheEmbeddedFilesTreeForPdfA4F() {
        byte[] source = RawPdf.Build("0 g 20 20 50 50 re f\n",
            pageEntries: "/Annots [5 0 R]",
            extraObjects: new[] {
                "<< /Type /Annot /Subtype /FileAttachment /Rect [120 120 140 140] /F 4 /FS 6 0 R /AP << /N 8 0 R >> >>",
                "<< /Type /Filespec /F (notes.txt) /UF (notes.txt) /EF << /F 7 0 R >> >>",
                RawPdf.Stream("/Type /EmbeddedFile /Subtype /text#2Fplain", "archived notes"),
                RawPdf.Stream("/Type /XObject /Subtype /Form /BBox [0 0 20 20]", "0 0 1 rg 0 0 20 20 re f")
            });

        PdfAConversionResult archive4f = PdfDocument.Load(new MemoryStream(source)).ConvertToPdfA(new PdfAConversionOptions { Profile = PdfComplianceProfile.PdfA4F });
        PdfAConversionResult archive4 = PdfDocument.Load(new MemoryStream(source)).ConvertToPdfA(new PdfAConversionOptions { Profile = PdfComplianceProfile.PdfA4 });

        Assert.True(archive4f.Succeeded, string.Join(Environment.NewLine, archive4f.Report.Violations));
        Assert.Contains(archive4f.Report.Changes, change => change.Kind == PdfAConversionChangeKind.EmbeddedFileUpdated && change.Location == "catalog names");
        Assert.Single(archive4f.RequireValue().Attachments.Extract());
        AssertRendersSame(source, archive4f.ToBytes());
        AssertVeraPdfPasses(archive4f.ToBytes(), "4f");
        // PDF/A-4 forbids the FileAttachment annotation, and this one draws an icon, so removing it would change the page.
        Assert.Equal(PdfAConversionRefusalReason.ForbiddenAnnotation, Assert.Single(archive4.Report.Violations).Reason);
    }

    [Fact]
    public void ConvertToPdfA_ChecksPageOutputIntentsForPdfA4() {
        byte[] source = RawPdf.Build("0 g 20 20 50 50 re f\n",
            pageEntries: "/OutputIntents [5 0 R]",
            extraObjects: new[] {
                "<< /Type /OutputIntent /S /GTS_PDFA1 /OutputConditionIdentifier (sRGB) /DestOutputProfile 6 0 R /DestOutputProfileRef << /URLs [(http://example.com/srgb.icc)] >> >>",
                RawPdf.Stream("/N 3", RawPdf.Latin1(PdfIccProfiles.SrgbIec6196621))
            });

        PdfAConversionResult result = PdfDocument.Load(new MemoryStream(source)).ConvertToPdfA(new PdfAConversionOptions { Profile = PdfComplianceProfile.PdfA4 });

        Assert.True(result.Succeeded, string.Join(Environment.NewLine, result.Report.Violations));
        Assert.Contains(result.Report.Changes, change => change.ObjectNumber == 5 && change.Message.Contains("DestOutputProfileRef", StringComparison.Ordinal));
        AssertVeraPdfPasses(result.ToBytes(), "4");
    }

    [Fact]
    public void ConvertToPdfA_FixesTheEmbeddedFontRulesPdfA4AddsWithoutChangingGlyphs() {
        // The program's Unicode cmap is listed only under the Unicode platform, and ToUnicode maps code 65 to U+0000, which
        // PDF/A-4 forbids, and code 66 to U+FFFD, which it allows.
        byte[] program = TrueTypeWithUnicodePlatformCmapOnly();
        const string toUnicode = "/CIDInit /ProcSet findresource begin 12 dict begin begincmap 1 begincodespacerange <00> <FF> endcodespacerange " +
            "2 beginbfchar <41> <0000> <42> <FFFD> endbfchar endcmap CMapName currentdict /CMap defineresource pop end end";
        byte[] source = RawPdf.Build("BT /F1 24 Tf 20 100 Td (A) Tj ET\n", resources: "/Font << /F1 5 0 R >>",
            extraObjects: new[] {
                "<< /Type /Font /Subtype /TrueType /BaseFont /Carlito /FirstChar 65 /LastChar 65 /Widths [578.613] /Encoding /WinAnsiEncoding " +
                    "/FontDescriptor 6 0 R /ToUnicode 8 0 R >>",
                "<< /Type /FontDescriptor /FontName /Carlito /Flags 32 /FontBBox [-500 -300 1300 1000] /ItalicAngle 0 /Ascent 950 /Descent -250 /CapHeight 640 /StemV 80 /FontFile2 7 0 R >>",
                RawPdf.Stream("/Length1 " + program.Length.ToString(CultureInfo.InvariantCulture), RawPdf.Latin1(program)),
                RawPdf.Stream(string.Empty, toUnicode)
            });

        PdfAConversionResult result = PdfDocument.Load(new MemoryStream(source)).ConvertToPdfA(new PdfAConversionOptions { Profile = PdfComplianceProfile.PdfA4 });

        Assert.True(result.Succeeded, string.Join(Environment.NewLine, result.Report.Violations));
        Assert.Contains(result.Report.Changes, change => change.Kind == PdfAConversionChangeKind.FontDescriptorUpdated && change.ObjectNumber == 7);
        Assert.Contains(result.Report.Changes, change => change.Kind == PdfAConversionChangeKind.ToUnicodeAdded && change.LossKind == OfficeIMO.OfficeConversionLossKind.Omission);
        PdfReadDocument output = PdfReadDocument.Open(result.ToBytes());
        var fonts = (PdfDictionary)PdfObjectLookup.Resolve(output.Objects, output.Pages[0].GetFontInspectionResources()!.Items["Font"])!;
        var font = (PdfDictionary)PdfObjectLookup.Resolve(output.Objects, fonts.Items["F1"])!;
        Dictionary<string, string> mappings = ResourceResolver.CreateFontResource("F1", font, output.Objects).CMap!.Mappings
            .ToDictionary(static entry => entry.Key, static entry => entry.Value, StringComparer.OrdinalIgnoreCase);
        Assert.False(mappings.ContainsKey("41"));
        Assert.Equal("\uFFFD", mappings["42"]);
        AssertRendersSame(source, result.ToBytes());
        AssertVeraPdfPasses(result.ToBytes(), "4");
    }

    [Fact]
    public void ConvertToPdfA_ShortensStructureTypeNamesBeyondTheNameLimit() {
        string longType = "Caption" + string.Concat(Enumerable.Repeat(",Caption#20Char", 12));
        byte[] source = RawPdf.Build("0 g 20 20 50 50 re f\n",
            catalogEntries: "/StructTreeRoot 5 0 R /MarkInfo << /Marked true >>",
            extraObjects: new[] {
                "<< /Type /StructTreeRoot /K [6 0 R] /RoleMap << /" + longType + " /P >> >>",
                "<< /Type /StructElem /S /" + longType + " /P 5 0 R >>"
            });

        PdfAConversionResult result = PdfDocument.Load(new MemoryStream(source)).ConvertToPdfA();

        Assert.True(result.Succeeded, string.Join(Environment.NewLine, result.Report.Violations));
        Assert.Contains(result.Report.Changes, change => change.Kind == PdfAConversionChangeKind.StructureUpdated && change.ObjectNumber == 5);
        PdfReadDocument output = PdfReadDocument.Open(result.ToBytes());
        var root = (PdfDictionary)PdfObjectLookup.Resolve(output.Objects, output.CatalogDictionary!.Items["StructTreeRoot"])!;
        var roleMap = (PdfDictionary)PdfObjectLookup.Resolve(output.Objects, root.Items["RoleMap"])!;
        string renamed = Assert.Single(roleMap.Items.Keys);
        Assert.True(Encoding.UTF8.GetByteCount(renamed) <= 127);
        Assert.Equal("P", ((PdfName)roleMap.Items[renamed]).Name);
        var element = (PdfDictionary)PdfObjectLookup.Resolve(output.Objects, ((PdfArray)PdfObjectLookup.Resolve(output.Objects, root.Items["K"])!).Items[0])!;
        Assert.Equal(renamed, ((PdfName)element.Items["S"]).Name);
        AssertVeraPdfPasses(result.ToBytes(), "2b");
    }

    [Fact]
    public void ConvertToPdfA_RejectsPdfA4E() {
        PdfDocument document = PdfDocument.Load(new MemoryStream(RawPdf.Build("0 g 0 0 10 10 re f\n")));

        ArgumentException exception = Assert.Throws<ArgumentException>(() => document.ConvertToPdfA(new PdfAConversionOptions { Profile = PdfComplianceProfile.PdfA4E }));

        Assert.Contains("PDF/A-4e", exception.Message, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(PdfComplianceProfile.PdfA2B, "2b")]
    [InlineData(PdfComplianceProfile.PdfA3U, "3u")]
    [InlineData(PdfComplianceProfile.PdfA4, "4")]
    public void ConvertToPdfA_EmbedsSymbolAndZapfDingbatsSubstitutesByGlyphName(PdfComplianceProfile profile, string flavour) {
        // Symbol draws alpha, beta and, through Differences, summation; ZapfDingbats draws a check mark and a star.
        byte[] source = RawPdf.Build(
            "BT /F1 24 Tf 20 120 Td (ab) Tj (A) Tj ET BT /F2 24 Tf 20 60 Td (34H) Tj ET\n",
            resources: "/Font << /F1 5 0 R /F2 6 0 R >>",
            extraObjects: new[] {
                "<< /Type /Font /Subtype /Type1 /BaseFont /Symbol /Encoding << /Differences [65 /summation] >> >>",
                "<< /Type /Font /Subtype /Type1 /BaseFont /ZapfDingbats >>"
            });

        PdfAConversionResult result = PdfDocument.Load(new MemoryStream(source)).ConvertToPdfA(new PdfAConversionOptions { Profile = profile });

        if (!HasSymbolicSubstitutes()) {
            Assert.False(result.Succeeded);
            Assert.All(result.Report.Violations, violation => Assert.Equal(PdfAConversionRefusalReason.FontNotEmbeddable, violation.Reason));
            return;
        }
        Assert.True(result.Succeeded, string.Join(Environment.NewLine, result.Report.Violations));
        Assert.Equal(2, result.Report.Changes.Count(change => change.Kind == PdfAConversionChangeKind.FontEmbedded));
        byte[] converted = result.ToBytes();
        PdfReadDocument output = PdfReadDocument.Open(converted);
        foreach (string name in new[] { "F1", "F2" }) {
            var pageFonts = (PdfDictionary)PdfObjectLookup.Resolve(output.Objects, output.Pages[0].GetFontInspectionResources()!.Items["Font"])!;
            var font = (PdfDictionary)PdfObjectLookup.Resolve(output.Objects, pageFonts.Items[name])!;
            Assert.Equal("TrueType", font.Get<PdfName>("Subtype")!.Name);
            Assert.False(font.Items.ContainsKey("Encoding"));
            PdfDictionary descriptor = Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(output.Objects, font.Items["FontDescriptor"]));
            Assert.Equal(4, (int)((PdfNumber)descriptor.Items["Flags"]).Value & 36);
        }
        Assert.Contains("\u03B1\u03B2\u2211", PdfDocument.Load(new MemoryStream(converted)).ExtractText(), StringComparison.Ordinal);
        // Viewers position non-embedded standard fonts by the Adobe metrics, so the embedded fonts keep those advances.
        var fonts = (PdfDictionary)PdfObjectLookup.Resolve(output.Objects, output.Pages[0].GetFontInspectionResources()!.Items["Font"])!;
        var symbol = (PdfDictionary)PdfObjectLookup.Resolve(output.Objects, fonts.Items["F1"])!;
        double[] widths = ((PdfArray)PdfObjectLookup.Resolve(output.Objects, symbol.Items["Widths"])!).Items.Select(item => ((PdfNumber)item).Value).ToArray();
        int firstChar = (int)((PdfNumber)symbol.Items["FirstChar"]).Value;
        Assert.Equal(713D, widths[65 - firstChar]);
        Assert.Equal(631D, widths[97 - firstChar]);
        Assert.Equal(549D, widths[98 - firstChar]);

        AssertVeraPdfPasses(converted, flavour);
    }

    [Fact]
    public void ConvertToPdfA_RefusesASymbolSubstituteWhoseWidthsDifferFromThePage() {
        byte[] source = RawPdf.Build("BT /F1 24 Tf 20 120 Td (a) Tj ET\n", resources: "/Font << /F1 5 0 R >>",
            extraObjects: new[] { "<< /Type /Font /Subtype /Type1 /BaseFont /Symbol /FirstChar 97 /LastChar 97 /Widths [100] >>" });

        PdfAConversionResult result = PdfDocument.Load(new MemoryStream(source)).ConvertToPdfA();

        Assert.False(result.Succeeded);
        Assert.Equal(PdfEmbeddedFontFamily.TryFromSystem("Symbol", out _) ? PdfAConversionRefusalReason.FontMetricsMismatch : PdfAConversionRefusalReason.FontNotEmbeddable,
            Assert.Single(result.Report.Violations).Reason);
    }

    [Theory]
    [MemberData(nameof(RefusalCases))]
    public void ConvertToPdfA_RefusesWithEveryReasonAndWritesNothing(string name, PdfAConversionRefusalReason reason, PdfComplianceProfile profile) {
        PdfAConversionOptions options = new() {
            Profile = profile,
            Attachments = name == "embedded file" ? PdfAAttachmentPolicy.Refuse : PdfAAttachmentPolicy.Remove
        };
        byte[] source = CreateRefusalSource(name);
        PdfDocument document = PdfDocument.Load(new MemoryStream(source));

        PdfAConversionResult result = document.ConvertToPdfA(options);

        Assert.False(result.Succeeded);
        Assert.Null(result.Value);
        Assert.Empty(result.Report.Changes);
        Assert.Contains(result.Report.Violations, violation => violation.Reason == reason);
        Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == "PdfA." + reason);
        Assert.Throws<PdfAConversionException>(() => result.ToBytes());
        string path = Path.Combine(Path.GetTempPath(), "officeimo-pdfa-refused-" + Guid.NewGuid().ToString("N") + ".pdf");
        PdfAConversionException exception = Assert.Throws<PdfAConversionException>(() => document.SaveAsPdfA(path, options));
        Assert.False(File.Exists(path));
        Assert.Contains(exception.Report.Violations, violation => violation.Reason == reason);
        using var stream = new MemoryStream();
        Assert.Throws<PdfAConversionException>(() => document.SaveAsPdfA(stream, options));
        Assert.Equal(0, stream.Length);
    }

    [Fact]
    public void SaveAsPdfA_WritesTheConvertedFileToAStream() {
        byte[] source = RawPdf.Build("0 g 20 20 50 50 re f\n");
        using var stream = new MemoryStream();

        PdfAConversionReport report = PdfDocument.Load(new MemoryStream(source)).SaveAsPdfA(stream,
            new PdfAConversionOptions { Profile = PdfComplianceProfile.PdfA3B });

        Assert.Empty(report.Violations);
        Assert.Equal(PdfComplianceProfile.PdfA3B, report.Profile);
        Assert.Contains(report.Changes, change => change.Kind == PdfAConversionChangeKind.OutputIntentAdded);
        AssertVeraPdfPasses(stream.ToArray(), "3b");
    }

    [Fact]
    public void ConvertToPdfA_EmbedsSuppliedFontsWithTheSameChecksAsInstalledOnes() {
        string? fontPath = PdfComplianceTestFonts.FindBundledTrueTypeFont();
        Assert.NotNull(fontPath);
        byte[] font = File.ReadAllBytes(fontPath!);
        PdfTrueTypeFontProgram program = PdfTrueTypeFontProgram.Parse(font);
        string widths = string.Join(" ", "AB".Select(character => {
            Assert.True(program.TryGetGlyphId(character, out int glyph));
            return program.GetExactGlyphWidth1000(glyph).ToString("0.###", CultureInfo.InvariantCulture);
        }));
        // A family no machine installs, so only the supplied face can stand in for it.
        byte[] source = RawPdf.Build("BT /F1 24 Tf 20 120 Td (AB) Tj ET\n", resources: "/Font << /F1 5 0 R >>",
            extraObjects: new[] { "<< /Type /Font /Subtype /TrueType /BaseFont /OfficeIMOArchiveSans /FirstChar 65 /LastChar 66 /Widths [" + widths + "] /Encoding /WinAnsiEncoding >>" });
        PdfAConversionResult Convert(PdfAFontSubstitution substitution) => PdfDocument.Load(new MemoryStream(source))
            .ConvertToPdfA(new PdfAConversionOptions { FontSubstitution = substitution });

        PdfAConversionResult supplied = Convert(PdfAFontSubstitution.Supplied(new PdfEmbeddedFontFamily("OfficeIMOArchiveSans", font)));
        PdfAConversionResult installedOnly = Convert(PdfAFontSubstitution.InstalledFonts);
        byte[] restricted = (byte[])font.Clone();
        WriteFsType(restricted, 0x0002);
        PdfAConversionResult licensed = Convert(PdfAFontSubstitution.Supplied(new PdfEmbeddedFontFamily("OfficeIMOArchiveSans", restricted)).ThenInstalledFonts());

        Assert.True(supplied.Succeeded, string.Join(Environment.NewLine, supplied.Report.Violations));
        Assert.Contains(supplied.Report.Changes, change => change.Kind == PdfAConversionChangeKind.FontEmbedded &&
            change.Message.Contains("supplied OfficeIMOArchiveSans face", StringComparison.Ordinal));
        AssertVeraPdfPasses(supplied.ToBytes(), "2b");
        Assert.Equal(PdfAConversionRefusalReason.FontNotEmbeddable, Assert.Single(installedOnly.Report.Violations).Reason);
        PdfAConversionViolation refusal = Assert.Single(licensed.Report.Violations);
        Assert.Equal(PdfAConversionRefusalReason.FontNotEmbeddable, refusal.Reason);
        Assert.Contains("fsType", refusal.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void ConvertToPdfA_ChecksSharedContentAgainstEachPdfA4PageOutputIntent() {
        // Both pages draw one RGB form; the second page declares a CMYK output intent of its own.
        byte[] cmyk = new byte[128];
        cmyk[3] = 128;
        cmyk[8] = 4;
        Encoding.ASCII.GetBytes("prtrCMYKLab ", 0, 12, cmyk, 12);
        Encoding.ASCII.GetBytes("acsp", 0, 4, cmyk, 36);
        byte[] source = RawPdf.Build("/X1 Do\n", resources: "/XObject << /X1 6 0 R >>", kids: "3 0 R 5 0 R",
            extraObjects: new[] {
                "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Resources << /XObject << /X1 6 0 R >> >> /Contents 7 0 R /OutputIntents [8 0 R] >>",
                RawPdf.Stream("/Type /XObject /Subtype /Form /BBox [0 0 20 20] /Resources << >>", "1 0 0 rg 0 0 10 10 re f\n"),
                RawPdf.Stream(string.Empty, "/X1 Do\n"),
                "<< /Type /OutputIntent /S /GTS_PDFA1 /OutputConditionIdentifier (Press) /DestOutputProfile 9 0 R >>",
                RawPdf.Stream("/N 4", RawPdf.Latin1(cmyk))
            });

        PdfAConversionResult result = PdfDocument.Load(new MemoryStream(source)).ConvertToPdfA(new PdfAConversionOptions { Profile = PdfComplianceProfile.PdfA4 });

        PdfAConversionViolation violation = Assert.Single(result.Report.Violations, violation => violation.Reason == PdfAConversionRefusalReason.DeviceColorNotCharacterized);
        Assert.StartsWith("page 2", violation.Location, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("Dingbats")]
    [InlineData("ITCZapfDingbats")]
    public void ConvertToPdfA_ReadsDingbatsAliasesInTheZapfDingbatsEncoding(string baseFont) {
        // 0x80 is a ZapfDingbats glyph but has no StandardEncoding name; only the missing substitute should stop conversion.
        byte[] source = RawPdf.Build("BT /F1 24 Tf 20 120 Td <80> Tj ET\n", resources: "/Font << /F1 5 0 R >>",
            extraObjects: new[] { "<< /Type /Font /Subtype /Type1 /BaseFont /" + baseFont + " >>" });

        PdfAConversionResult result = PdfDocument.Load(new MemoryStream(source))
            .ConvertToPdfA(new PdfAConversionOptions { FontSubstitution = PdfAFontSubstitution.None });

        Assert.Equal(PdfAConversionRefusalReason.FontNotEmbeddable, Assert.Single(result.Report.Violations).Reason);
    }

    private static void WriteFsType(byte[] font, ushort fsType) {
        int count = (font[4] << 8) | font[5];
        for (int index = 0; index < count; index++) {
            int record = 12 + index * 16;
            if (Encoding.ASCII.GetString(font, record, 4) != "OS/2") continue;
            int offset = (font[record + 8] << 24) | (font[record + 9] << 16) | (font[record + 10] << 8) | font[record + 11];
            font[offset + 8] = (byte)(fsType >> 8);
            font[offset + 9] = (byte)fsType;
            return;
        }
        throw new InvalidOperationException("The test font has no OS/2 table.");
    }

    public static TheoryData<string, PdfAConversionRefusalReason, PdfComplianceProfile> RefusalCases() => new() {
        { "signature", PdfAConversionRefusalReason.SignedDocument, PdfComplianceProfile.PdfA2B },
        { "composite font", PdfAConversionRefusalReason.FontNotEmbeddable, PdfComplianceProfile.PdfA2B },
        { "unnamed glyph", PdfAConversionRefusalReason.FontGlyphMissing, PdfComplianceProfile.PdfA2B },
        { "invisible notdef", PdfAConversionRefusalReason.FontGlyphMissing, PdfComplianceProfile.PdfA2B },
        { "truetype differences", PdfAConversionRefusalReason.FontEncodingNotCompliant, PdfComplianceProfile.PdfA2B },
        { "widths", HasHelveticaSubstitute() ? PdfAConversionRefusalReason.FontMetricsMismatch : PdfAConversionRefusalReason.FontNotEmbeddable, PdfComplianceProfile.PdfA2B },
        { "type3 unicode", PdfAConversionRefusalReason.UnicodeMappingUnavailable, PdfComplianceProfile.PdfA2U },
        { "cmyk", PdfAConversionRefusalReason.DeviceColorNotCharacterized, PdfComplianceProfile.PdfA2B },
        { "icc", PdfAConversionRefusalReason.InvalidIccProfile, PdfComplianceProfile.PdfA2B },
        { "output intents", PdfAConversionRefusalReason.ConflictingOutputIntents, PdfComplianceProfile.PdfA2B },
        { "no appearance", PdfAConversionRefusalReason.AnnotationAppearanceMissing, PdfComplianceProfile.PdfA2B },
        { "screen", PdfAConversionRefusalReason.ForbiddenAnnotation, PdfComplianceProfile.PdfA2B },
        { "embedded file", PdfAConversionRefusalReason.EmbeddedFileNotAllowed, PdfComplianceProfile.PdfA2B },
        { "postscript", PdfAConversionRefusalReason.ForbiddenXObject, PdfComplianceProfile.PdfA2B },
        { "blend mode", PdfAConversionRefusalReason.ForbiddenGraphicsState, PdfComplianceProfile.PdfA2B },
        { "nesting", PdfAConversionRefusalReason.ImplementationLimitExceeded, PdfComplianceProfile.PdfA2B },
        { "external stream", PdfAConversionRefusalReason.UnsupportedStream, PdfComplianceProfile.PdfA2B },
        { "inline image", PdfAConversionRefusalReason.InlineImageNotCompliant, PdfComplianceProfile.PdfA2B },
        { "operator", PdfAConversionRefusalReason.ContentNotCompliant, PdfComplianceProfile.PdfA2B },
        { "xfa", PdfAConversionRefusalReason.DynamicXfaForm, PdfComplianceProfile.PdfA2B },
        { "no attachment", PdfAConversionRefusalReason.EmbeddedFileRequired, PdfComplianceProfile.PdfA4F },
        { "undefined resource", PdfAConversionRefusalReason.ContentNotCompliant, PdfComplianceProfile.PdfA4 },
        { "page output intents", PdfAConversionRefusalReason.ConflictingOutputIntents, PdfComplianceProfile.PdfA4 }
    };

    [Fact]
    public void ConvertToPdfA_ListsEveryViolationByObject() {
        byte[] source = RawPdf.Build(
            "0 0 0 1 k 0 0 10 10 re f\n/G1 gs\n",
            resources: "/ExtGState << /G1 5 0 R >>",
            pageEntries: "/Annots [6 0 R]",
            extraObjects: new[] {
                "<< /Type /ExtGState /BM /Weird >>",
                "<< /Type /Annot /Subtype /Watermark /Rect [10 10 50 50] /F 4 >>"
            });

        PdfAConversionResult result = PdfDocument.Load(new MemoryStream(source)).ConvertToPdfA();

        Assert.Contains(result.Report.Violations, violation => violation.Reason == PdfAConversionRefusalReason.DeviceColorNotCharacterized && violation.ObjectNumber == 3);
        Assert.Contains(result.Report.Violations, violation => violation.Reason == PdfAConversionRefusalReason.ForbiddenGraphicsState && violation.ObjectNumber == 5);
        Assert.Contains(result.Report.Violations, violation => violation.Reason == PdfAConversionRefusalReason.AnnotationAppearanceMissing && violation.ObjectNumber == 6);
    }

    private static byte[] CreateRefusalSource(string name) => name switch {
        "signature" => RawPdf.Build("0 g 0 0 10 10 re f\n",
            catalogEntries: "/AcroForm << /Fields [5 0 R] /SigFlags 3 >>",
            extraObjects: new[] {
                "<< /FT /Sig /T (Approval) /V 6 0 R >>",
                "<< /Type /Sig /Filter /Adobe.PPKLite /SubFilter /adbe.pkcs7.detached /ByteRange [0 10 20 10] /Contents <00000000> >>"
            }),
        "composite font" => RawPdf.Build("BT /F1 12 Tf 10 10 Td <0041> Tj ET\n", resources: "/Font << /F1 5 0 R >>",
            extraObjects: new[] {
                "<< /Type /Font /Subtype /Type0 /BaseFont /MSMincho /Encoding /Identity-H /DescendantFonts [6 0 R] >>",
                "<< /Type /Font /Subtype /CIDFontType2 /BaseFont /MSMincho /CIDSystemInfo << /Registry (Adobe) /Ordering (Identity) /Supplement 0 >> /FontDescriptor 7 0 R >>",
                "<< /Type /FontDescriptor /FontName /MSMincho /Flags 4 /FontBBox [0 0 1000 1000] /ItalicAngle 0 /Ascent 880 /Descent -120 /CapHeight 700 /StemV 80 >>"
            }),
        "unnamed glyph" => RawPdf.Build("BT /F1 12 Tf 10 10 Td (A) Tj ET\n", resources: "/Font << /F1 5 0 R >>",
            extraObjects: new[] { "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding << /Differences [65 /g123] >> >>" }),
        "invisible notdef" => RawPdf.Build("BT 3 Tr /F1 12 Tf 10 10 Td (\\004) Tj ET\n", resources: "/Font << /F1 5 0 R >>",
            extraObjects: new[] { "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>" }),
        "truetype differences" => RawPdf.Build("BT /F1 12 Tf 10 10 Td (A) Tj ET\n", resources: "/Font << /F1 5 0 R >>",
            extraObjects: new[] {
                "<< /Type /Font /Subtype /TrueType /BaseFont /Carlito /FirstChar 65 /LastChar 65 /Widths [606] " +
                    "/Encoding << /BaseEncoding /WinAnsiEncoding /Differences [65 /B] >> /FontDescriptor 6 0 R >>",
                "<< /Type /FontDescriptor /FontName /Carlito /Flags 32 /FontBBox [0 -250 1000 950] /ItalicAngle 0 /Ascent 950 /Descent -250 /CapHeight 640 /StemV 80 /FontFile2 7 0 R >>",
                RawPdf.Stream("/Length1 " + TrueTypeWithoutUnicodeCmap().Length.ToString(CultureInfo.InvariantCulture), RawPdf.Latin1(TrueTypeWithoutUnicodeCmap()))
            }),
        "widths" => RawPdf.Build("BT /F1 12 Tf 10 10 Td (A) Tj ET\n", resources: "/Font << /F1 5 0 R >>",
            extraObjects: new[] { "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding /FirstChar 65 /LastChar 65 /Widths [100] >>" }),
        "type3 unicode" => RawPdf.Build("BT /F1 12 Tf 10 10 Td (A) Tj ET\n", resources: "/Font << /F1 5 0 R >>",
            extraObjects: new[] {
                "<< /Type /Font /Subtype /Type3 /FontBBox [0 0 1000 1000] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /g1 6 0 R >> " +
                    "/Encoding << /Type /Encoding /Differences [65 /g1] >> /FirstChar 65 /LastChar 65 /Widths [1000] /Resources << >> >>",
                RawPdf.Stream(string.Empty, "1000 0 0 0 1000 1000 d1 0 0 1000 1000 re f")
            }),
        "cmyk" => RawPdf.Build("0 0 0 1 k 0 0 10 10 re f\n"),
        "icc" => RawPdf.Build("/CS0 cs 0.5 0.5 0.5 sc 0 0 10 10 re f\n", resources: "/ColorSpace << /CS0 [/ICCBased 5 0 R] >>",
            extraObjects: new[] { RawPdf.Stream("/N 3", "not an ICC profile") }),
        "output intents" => RawPdf.Build("0 g 0 0 10 10 re f\n",
            catalogEntries: "/OutputIntents [5 0 R 6 0 R]",
            extraObjects: new[] {
                "<< /Type /OutputIntent /S /GTS_PDFA1 /OutputConditionIdentifier (sRGB) /DestOutputProfile 7 0 R >>",
                "<< /Type /OutputIntent /S /GTS_PDFX /OutputConditionIdentifier (sRGB) /DestOutputProfile 8 0 R >>",
                RawPdf.Stream("/N 3", RawPdf.Latin1(PdfIccProfiles.SrgbIec6196621)),
                RawPdf.Stream("/N 3", RawPdf.Latin1(PdfIccProfiles.SrgbIec6196621))
            }),
        "no appearance" => RawPdf.Build("0 g 0 0 10 10 re f\n", pageEntries: "/Annots [5 0 R]",
            extraObjects: new[] { "<< /Type /Annot /Subtype /Watermark /Rect [10 10 50 50] /F 4 >>" }),
        "screen" => RawPdf.Build("0 g 0 0 10 10 re f\n", pageEntries: "/Annots [5 0 R]",
            extraObjects: new[] {
                "<< /Type /Annot /Subtype /Screen /Rect [10 10 50 50] /F 4 /AP << /N 6 0 R >> >>",
                RawPdf.Stream("/Type /XObject /Subtype /Form /BBox [0 0 40 40]", "0 g 0 0 40 40 re f")
            }),
        "embedded file" => RawPdf.Build("0 g 0 0 10 10 re f\n",
            catalogEntries: "/Names << /EmbeddedFiles << /Names [(a.txt) 5 0 R] >> >>",
            extraObjects: new[] {
                "<< /Type /Filespec /F (a.txt) /EF << /F 6 0 R >> >>",
                RawPdf.Stream("/Type /EmbeddedFile", "hello")
            }),
        "postscript" => RawPdf.Build("/P1 Do\n", resources: "/XObject << /P1 5 0 R >>",
            extraObjects: new[] { RawPdf.Stream("/Type /XObject /Subtype /PS", "showpage") }),
        "blend mode" => RawPdf.Build("/G1 gs 0 g 0 0 10 10 re f\n", resources: "/ExtGState << /G1 5 0 R >>",
            extraObjects: new[] { "<< /Type /ExtGState /BM /Weird >>" }),
        "nesting" => RawPdf.Build(string.Concat(Enumerable.Repeat("q ", 30)) + "0 g 0 0 10 10 re f " + string.Concat(Enumerable.Repeat("Q ", 30)) + "\n"),
        "external stream" => RawPdf.Build("/X1 Do\n", resources: "/XObject << /X1 5 0 R >>",
            extraObjects: new[] { RawPdf.Stream("/Type /XObject /Subtype /Form /BBox [0 0 1 1] /F (external.dat)", string.Empty) }),
        "inline image" => RawPdf.Build("q 10 0 0 10 0 0 cm BI /W 1 /H 1 /CS /G /BPC 8 /I true ID \u0080 EI Q\n"),
        "operator" => RawPdf.Build("0 0 m xyz\n"),
        "xfa" => RawPdf.Build("0 g 0 0 10 10 re f\n",
            catalogEntries: "/NeedsRendering true /AcroForm << /Fields [] /XFA 5 0 R >>",
            extraObjects: new[] { RawPdf.Stream(string.Empty, "<xdp:xdp xmlns:xdp=\"http://ns.adobe.com/xdp/\"/>") }),
        "no attachment" => RawPdf.Build("0 g 0 0 10 10 re f\n"),
        "undefined resource" => RawPdf.Build("/Missing gs 0 g 0 0 10 10 re f\n"),
        "page output intents" => RawPdf.Build("0 g 0 0 10 10 re f\n",
            pageEntries: "/OutputIntents [5 0 R 6 0 R]",
            extraObjects: new[] {
                "<< /Type /OutputIntent /S /GTS_PDFA1 /OutputConditionIdentifier (sRGB) /DestOutputProfile 7 0 R >>",
                "<< /Type /OutputIntent /S /GTS_PDFX /OutputConditionIdentifier (sRGB) /DestOutputProfile 8 0 R >>",
                RawPdf.Stream("/N 3", RawPdf.Latin1(PdfIccProfiles.SrgbIec6196621)),
                RawPdf.Stream("/N 3", RawPdf.Latin1(PdfIccProfiles.SrgbIec6196621))
            }),
        _ => throw new ArgumentOutOfRangeException(nameof(name))
    };

    /// <summary>The bundled Carlito face with its Windows Unicode and Macintosh Roman cmaps relabelled to encodings viewers do not use.</summary>
    private static byte[] TrueTypeWithUnicodePlatformCmapOnly() {
        string? path = PdfComplianceTestFonts.FindBundledTrueTypeFont();
        Assert.NotNull(path);
        byte[] font = File.ReadAllBytes(path!);
        int tables = (font[4] << 8) | font[5];
        for (int table = 0; table < tables; table++) {
            int record = 12 + table * 16;
            if (Encoding.ASCII.GetString(font, record, 4) != "cmap") continue;
            int offset = (font[record + 8] << 24) | (font[record + 9] << 16) | (font[record + 10] << 8) | font[record + 11];
            int count = (font[offset + 2] << 8) | font[offset + 3];
            for (int entry = 0; entry < count; entry++) {
                int at = offset + 4 + entry * 8;
                if (font[at + 1] == 3 && font[at + 3] == 1) {
                    font[at + 1] = 0;
                    font[at + 3] = 4;
                } else if (font[at + 1] == 1 && font[at + 3] == 0) {
                    font[at + 3] = 1;
                }
            }
        }
        return font;
    }

    /// <summary>The bundled Carlito face with its Windows Unicode cmap relabelled as a symbol cmap.</summary>
    private static byte[] TrueTypeWithoutUnicodeCmap() {
        string? path = PdfComplianceTestFonts.FindBundledTrueTypeFont();
        Assert.NotNull(path);
        byte[] font = File.ReadAllBytes(path!);
        int tables = (font[4] << 8) | font[5];
        for (int table = 0; table < tables; table++) {
            int record = 12 + table * 16;
            if (Encoding.ASCII.GetString(font, record, 4) != "cmap") continue;
            int offset = (font[record + 8] << 24) | (font[record + 9] << 16) | (font[record + 10] << 8) | font[record + 11];
            int count = (font[offset + 2] << 8) | font[offset + 3];
            for (int entry = 0; entry < count; entry++) {
                int at = offset + 4 + entry * 8;
                if (font[at + 1] == 3 && font[at + 3] == 1) font[at + 3] = 0;
            }
        }
        return font;
    }

    private static bool HasSymbolicSubstitutes() =>
        PdfEmbeddedFontFamily.TryFromSystem("Symbol", out _) && PdfEmbeddedFontFamily.TryFromSystem("Zapf Dingbats", out _);

    private static bool HasHelveticaSubstitute() =>
        new[] { "Arial", "Liberation Sans", "Arimo", "Helvetica" }.Any(family => PdfEmbeddedFontFamily.TryFromSystem(family, out _));

    /// <summary>Compares pixels with the snapshot tolerances; a source drawn from a synthesized appearance still compares by pixels.</summary>
    private static void AssertRendersSame(byte[] source, byte[] converted) {
        var options = new PdfVisualComparisonOptions {
            Scale = 2D,
            ChannelTolerance = 48,
            AllowedDifferenceRatio = 0.000001D,
            IgnoreAntialiasing = true
        };
        PdfVisualComparisonReport report = PdfVisualComparer.Compare(source, converted, options: options);
        Assert.Empty(report.StructuralDifferences);
        foreach (PdfVisualPageComparison page in report.Pages) {
            Assert.False(page.HasSizeDifference, "page " + page.PageNumber + " changed size");
            Assert.True(page.DifferenceRatio <= options.AllowedDifferenceRatio,
                "page " + page.PageNumber + " differs in " + page.DifferentPixels.ToString(CultureInfo.InvariantCulture) + " pixels");
        }
    }

    private static void AssertVeraPdfPasses(byte[] pdf, string flavour) {
        PdfExternalValidator validator = PdfExternalValidator.VeraPdf(flavour);
        if (!validator.IsAvailable) {
            PdfExternalValidator.SkipUnlessRequired(validator);
            return;
        }
        PdfExternalProcessResult result = validator.Run(pdf, "converted-pdfa-" + flavour + ".pdf");
        Assert.True(result.ExitCode == 0, result.GetDiagnosticText());
    }

    /// <summary>Builds small single-page PDFs whose objects the tests control exactly.</summary>
    private static class RawPdf {
        private static readonly Encoding Iso88591 = Encoding.GetEncoding(28591);

        internal static byte[] Build(string content, string resources = "", string pageEntries = "", string catalogEntries = "",
            string contentFilter = "", string trailerEntries = "", string kids = "3 0 R", params string[] extraObjects) {
            var objects = new List<string> {
                "<< /Type /Catalog /Pages 2 0 R " + catalogEntries + " >>",
                "<< /Type /Pages /Kids [" + kids + "] /Count " + kids.Split(new[] { " R" }, StringSplitOptions.RemoveEmptyEntries).Count(part => part.Trim().Length > 0).ToString(CultureInfo.InvariantCulture) + " >>",
                "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Resources << " + resources + " >> /Contents 4 0 R " + pageEntries + " >>",
                Stream(contentFilter, content)
            };
            objects.AddRange(extraObjects);
            var pdf = new StringBuilder("%PDF-1.4\n%âãÏÓ\n");
            var offsets = new List<int>();
            for (int index = 0; index < objects.Count; index++) {
                offsets.Add(pdf.Length);
                pdf.Append(index + 1).Append(" 0 obj\n").Append(objects[index]).Append("\nendobj\n");
            }
            int xref = pdf.Length;
            pdf.Append("xref\n0 ").Append(objects.Count + 1).Append("\n0000000000 65535 f \n");
            foreach (int offset in offsets) pdf.Append(offset.ToString("0000000000", CultureInfo.InvariantCulture)).Append(" 00000 n \n");
            pdf.Append("trailer\n<< /Size ").Append(objects.Count + 1).Append(" /Root 1 0 R ").Append(trailerEntries).Append(" >>\nstartxref\n").Append(xref).Append("\n%%EOF\n");
            return Iso88591.GetBytes(pdf.ToString());
        }

        internal static string Stream(string dictionary, string data) =>
            "<< " + dictionary + " /Length " + data.Length.ToString(CultureInfo.InvariantCulture) + " >>\nstream\n" + data + "\nendstream";

        internal static string Latin1(byte[] data) => Iso88591.GetString(data);

        /// <summary>LZW-encodes text with the PDF defaults (early change 1), returning the bytes as Latin-1 text.</summary>
        internal static string Lzw(string text) {
            byte[] input = Iso88591.GetBytes(text);
            var output = new List<byte>();
            int buffer = 0, bits = 0, width = 9;
            void Write(int code) {
                buffer = (buffer << width) | code;
                bits += width;
                while (bits >= 8) {
                    output.Add((byte)(buffer >> (bits - 8)));
                    bits -= 8;
                    buffer &= (1 << bits) - 1;
                }
            }
            var table = new Dictionary<string, int>();
            for (int code = 0; code < 256; code++) table[((char)code).ToString()] = code;
            int next = 258;
            Write(256);
            string current = string.Empty;
            foreach (byte value in input) {
                string candidate = current + (char)value;
                if (table.ContainsKey(candidate)) {
                    current = candidate;
                    continue;
                }
                Write(table[current]);
                table[candidate] = next++;
                if (next >= (1 << width) - 1 && width < 12) width++;
                current = ((char)value).ToString();
            }
            if (current.Length > 0) Write(table[current]);
            Write(257);
            if (bits > 0) output.Add((byte)(buffer << (8 - bits)));
            return Iso88591.GetString(output.ToArray());
        }
    }
}
