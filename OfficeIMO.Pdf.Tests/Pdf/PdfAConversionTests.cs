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
            catalogEntries: "/OpenAction 7 0 R /AA << /WC 7 0 R >> /Names << /JavaScript << /Names [(init) 7 0 R] >> >>",
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
        PdfAConversionResult converted = PdfDocument.Load(new MemoryStream(encrypted)).ConvertToPdfA(new PdfAConversionOptions { Password = "owner" });

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
    [MemberData(nameof(RefusalCases))]
    public void ConvertToPdfA_RefusesWithEveryReasonAndWritesNothing(string name, PdfAConversionRefusalReason reason, PdfComplianceProfile profile) {
        PdfAConversionOptions options = new() { Profile = profile, RemoveEmbeddedFiles = name != "embedded file" };
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
        { "xfa", PdfAConversionRefusalReason.DynamicXfaForm, PdfComplianceProfile.PdfA2B }
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
        _ => throw new ArgumentOutOfRangeException(nameof(name))
    };

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
            string contentFilter = "", params string[] extraObjects) {
            var objects = new List<string> {
                "<< /Type /Catalog /Pages 2 0 R " + catalogEntries + " >>",
                "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
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
            pdf.Append("trailer\n<< /Size ").Append(objects.Count + 1).Append(" /Root 1 0 R >>\nstartxref\n").Append(xref).Append("\n%%EOF\n");
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
