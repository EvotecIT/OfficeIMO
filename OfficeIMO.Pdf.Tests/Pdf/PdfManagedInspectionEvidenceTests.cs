using System.Globalization;
using System.Text;
using System.Threading;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfManagedInspectionEvidenceTests {
    [Fact]
    public void InspectionRetainsSuccessfulOccurrencesSourceLinksAndExplicitMappingFailures() {
        PdfDocument pdf = PdfDocument.Load(BuildMixedDocument());
        Assert.Throws<PdfUnsupportedTextMappingException>(() => pdf.Read());
        PdfInspectionReadResult inspection = pdf.ReadForInspection(new PdfReadOptions {
            PageSelection = PdfPageSelection.Parse("3,2,1,3")
        });
        Assert.False(inspection.CompletedAllSelectedPages);
        Assert.Equal(3, inspection.SourceInfo.PageCount);
        Assert.Equal("Inspection fixture", inspection.SourceInfo.Metadata.Title);
        Assert.Equal(new[] { 3, 2, 1, 3 }, inspection.PageOutcomes.Select(outcome => outcome.PageNumber));
        Assert.Equal(new[] { 3, 1, 3 }, inspection.Document.Pages.Select(page => page.PageNumber));
        Assert.Equal(new[] { 0, 1, 2, 3 }, inspection.PageOutcomes.Select(outcome => outcome.SelectionIndex));
        PdfPageReadOutcome failure = inspection.PageOutcomes[1];
        Assert.Equal(PdfPageReadStatus.UnmappableText, failure.Status);
        Assert.Null(failure.Page);
        Assert.Equal("Mystery-H", failure.MappingFailure!.EncodingName);
        Assert.Equal("F1", failure.MappingFailure.FontResource);
        Assert.False(failure.MappingFailure.HasDeclaredToUnicode);
        Assert.Equal("https://example.test/source-link", Assert.Single(inspection.SourceInfo.Pages[1].LinkAnnotations).Uri);
        Assert.Contains("First readable page", inspection.Document.Text);
        Assert.Contains("Last readable page", inspection.Document.Text);
        Assert.Throws<PdfUnsupportedTextMappingException>(() => pdf.Read());
    }

    [Fact]
    public void InspectionReportsUnsupportedFilterWithoutReturningAnEmptySuccessfulPage() {
        byte[] bytes = BuildMixedDocument(unsupportedFilter: true);
        PdfInspectionReadResult inspection = PdfDocument.Load(bytes).ReadForInspection();
        PdfPageReadOutcome failed = inspection.PageOutcomes[1];
        Assert.Equal(PdfPageReadStatus.UnsupportedContentStreamFilter, failed.Status);
        Assert.Contains("UnsupportedDecode", failed.UnsupportedFilters);
        Assert.Null(failed.Page);
        Assert.Equal(2, inspection.Document.Pages.Count);
    }

    [Theory]
    [InlineData("Type1", "WinAnsiEncoding")]
    [InlineData("Type0", "Identity-H")]
    public void DecodingEvidenceDistinguishesToUnicodeRawFallbackEncodingAndActualText(string fontSubtype, string encodingName) {
        byte[] mapped = BuildMappedDocument(encoding: encodingName, fontSubtype: fontSubtype);
        PdfTextSpan span = Assert.Single(PdfReadDocument.Open(mapped).Pages[0].GetTextSpans());
        Assert.Equal("\u03A9B", span.Text);
        Assert.Equal(PdfTextDecodingSource.ToUnicode, span.DecodingEvidence.Source);
        Assert.True(span.DecodingEvidence.UsedFallback);
        Assert.Equal(1, span.DecodingEvidence.UnmappedCodeBytes);
        PdfTextSpan encoding = Assert.Single(PdfReadDocument.Open(mapped,
            new PdfLoadOptions { PreferToUnicode = false }).Pages[0].GetTextSpans());
        Assert.Equal("AB", encoding.Text);
        Assert.Equal(fontSubtype == "Type1" ? PdfTextDecodingSource.FontEncoding : PdfTextDecodingSource.WinAnsiFallback,
            encoding.DecodingEvidence.Source);
        Assert.Equal(fontSubtype == "Type0", encoding.DecodingEvidence.UsedFallback);
        Assert.Throws<PdfUnsupportedTextMappingException>(() => PdfDocument.Load(mapped,
            new PdfLoadOptions { UseWinAnsiFallback = false }).Read());

        byte[] actualText = BuildMappedDocument(actualText: true, encoding: encodingName, fontSubtype: fontSubtype);
        PdfTextSpan replacement = Assert.Single(PdfReadDocument.Open(actualText).Pages[0].GetTextSpans());
        Assert.Equal("Replacement", replacement.Text);
        Assert.True(replacement.HasActualText);
        Assert.Equal(PdfTextDecodingSource.ActualText, replacement.DecodingEvidence.Source);
        Assert.Equal(PdfTextDecodingSource.ToUnicode, replacement.DecodingEvidence.UnderlyingSource);
        Assert.False(replacement.DecodingEvidence.UsedFallback);
        Assert.Equal(0, replacement.DecodingEvidence.UnmappedCodeBytes);
        PdfInspectionReadResult actualInspection = PdfDocument.Load(actualText).ReadForInspection();
        Assert.True(actualInspection.CompletedAllSelectedPages);
        Assert.Contains("Replacement", actualInspection.Document.Text);
        PdfTextSpan strictReplacement = Assert.Single(PdfReadDocument.Open(actualText,
            new PdfLoadOptions { UseWinAnsiFallback = false }).Pages[0].GetTextSpans());
        Assert.Equal(PdfTextDecodingSource.ActualText, strictReplacement.DecodingEvidence.Source);
        Assert.Equal(PdfTextDecodingSource.Unmappable, strictReplacement.DecodingEvidence.UnderlyingSource);
        Assert.Equal(0, strictReplacement.DecodingEvidence.UnmappedCodeBytes);
        byte[] fallback = BuildMappedDocument(includeToUnicode: false, encoding: null, baseFont: "UnknownFont");
        PdfTextSpan fallbackSpan = Assert.Single(PdfReadDocument.Open(fallback).Pages[0].GetTextSpans());
        Assert.Equal(PdfTextDecodingSource.WinAnsiFallback, fallbackSpan.DecodingEvidence.Source);
        Assert.True(fallbackSpan.DecodingEvidence.UsedFallback);
        Assert.Throws<PdfUnsupportedTextMappingException>(() => PdfDocument.Load(fallback,
            new PdfLoadOptions { UseWinAnsiFallback = false }).Read());
        byte[] differences = BuildMappedDocument(includeToUnicode: false, encoding: null, baseFont: "UnknownFont",
            contentOverride: "BT /F1 12 Tf 20 100 Td <41> Tj ET", fontEncodingDictionary: "<< /Differences [65 /uni03B1] >>");
        PdfTextSpan difference = Assert.Single(PdfReadDocument.Open(differences,
            new PdfLoadOptions { UseWinAnsiFallback = false }).Pages[0].GetTextSpans());
        Assert.Equal("\u03B1", difference.Text);
        Assert.Equal(PdfTextDecodingSource.FontEncoding, difference.DecodingEvidence.Source);
        Assert.True(difference.DecodingEvidence.UsesDifferences);
        Assert.False(difference.DecodingEvidence.UsedFallback);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void ActualTextEvidenceRetainsTheCompositeRunOutcomeRatherThanMapPresence(bool includeToUnicode) {
        byte[] incomplete = BuildCompositeEvidenceDocument(includeToUnicode, complete: false, actualText: false);
        Assert.Throws<PdfUnsupportedTextMappingException>(() => PdfDocument.Load(incomplete).Read());
        PdfPageReadOutcome failed = Assert.Single(PdfDocument.Load(incomplete).ReadForInspection().PageOutcomes);
        Assert.Equal(PdfPageReadStatus.UnmappableText, failed.Status);
        Assert.Equal(includeToUnicode, failed.MappingFailure!.HasDeclaredToUnicode);

        byte[] replaced = BuildCompositeEvidenceDocument(includeToUnicode, complete: false, actualText: true);
        PdfTextSpan replacement = Assert.Single(PdfReadDocument.Open(replaced).Pages[0].GetTextSpans());
        Assert.Equal("Replacement", replacement.Text);
        Assert.Equal(PdfTextDecodingSource.ActualText, replacement.DecodingEvidence.Source);
        Assert.Equal(PdfTextDecodingSource.Unmappable, replacement.DecodingEvidence.UnderlyingSource);
        Assert.Equal("UniJIS-UCS2-H", replacement.DecodingEvidence.EncodingName);
        Assert.Equal(includeToUnicode, replacement.DecodingEvidence.HasDeclaredToUnicode);
        Assert.Equal(0, replacement.DecodingEvidence.UnmappedCodeBytes);
        Assert.False(replacement.DecodingEvidence.UsedFallback);
        PdfInspectionReadResult inspection = PdfDocument.Load(replaced).ReadForInspection();
        Assert.True(inspection.CompletedAllSelectedPages);
        PdfTextSpan logical = Assert.Single(Assert.Single(inspection.Document.Pages).TextBlocks.SelectMany(block => block.Spans));
        Assert.Equal(PdfTextDecodingSource.Unmappable, logical.DecodingEvidence.UnderlyingSource);

        PdfTextDecodingSource successfulSource = includeToUnicode
            ? PdfTextDecodingSource.ToUnicode : PdfTextDecodingSource.PredefinedCMap;
        byte[] complete = BuildCompositeEvidenceDocument(includeToUnicode, complete: true, actualText: false);
        PdfTextSpan mapped = Assert.Single(PdfReadDocument.Open(complete).Pages[0].GetTextSpans());
        Assert.Equal("\u3042\u3044", mapped.Text);
        Assert.Equal(successfulSource, mapped.DecodingEvidence.Source);
        Assert.Equal(0, mapped.DecodingEvidence.UnmappedCodeBytes);
        Assert.False(mapped.DecodingEvidence.UsedFallback);
        byte[] completeReplacement = BuildCompositeEvidenceDocument(includeToUnicode, complete: true, actualText: true);
        PdfTextSpan replacedMapped = Assert.Single(PdfReadDocument.Open(completeReplacement).Pages[0].GetTextSpans());
        Assert.Equal(successfulSource, replacedMapped.DecodingEvidence.UnderlyingSource);
        Assert.Equal(0, replacedMapped.DecodingEvidence.UnmappedCodeBytes);
    }

    [Fact]
    public void InspectionDoesNotConvertResourceCancellationParsingOrPermissionFailuresToPartialSuccess() {
        byte[] bytes = BuildMixedDocument();
        Assert.Throws<PdfReadLimitException>(() => PdfDocument.Load(bytes,
            new PdfLoadOptions { Limits = new PdfReadLimits { MaxPages = 2 } }).ReadForInspection());
        Assert.Throws<PdfReadLimitException>(() => PdfDocument.Load(bytes).ReadForInspection(
            new PdfReadOptions { Pipeline = new PdfUnderstandingPipelineOptions { MaxPages = 2 } }));
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        Assert.ThrowsAny<OperationCanceledException>(() => PdfDocument.Load(bytes).ReadForInspection(cancellationToken: cancellation.Token));
        byte[] noXref = BuildMixedDocument(includeXref: false);
        Assert.Throws<PdfParseException>(() => PdfDocument.Load(noXref,
            new PdfLoadOptions { ParsingMode = PdfParsingMode.Strict }).ReadForInspection());
        Assert.True(PdfDocument.Load(noXref).ReadForInspection().RepairReport.HasRepairs);

        PdfDocument privateDocument = PdfDocument.Create();
        privateDocument.Content.Text("Private");
        byte[] encrypted = privateDocument.Security.Encrypt(new PdfStandardEncryptionOptions("reader") {
                OwnerPassword = "owner", AllowedPermissions = PdfStandardPermissions.None
            }).Pdf;
        Assert.Throws<PdfPermissionDeniedException>(() => PdfDocument.Load(encrypted,
            new PdfLoadOptions { Password = "reader" }).ReadForInspection());
        Assert.Throws<PdfInvalidPasswordException>(() => PdfDocument.Load(encrypted,
            new PdfLoadOptions { Password = "wrong" }).ReadForInspection());
    }

    [Fact]
    public void ExplicitKerningPolicyChangesPositionedAdvancesAndClonesRetainEveryBudget() {
        byte[] bytes = BuildMappedDocument(contentOverride: "BT /F1 12 Tf 1 0 0 1 20 100 Tm [(A) 300 (B)] TJ ET");
        double adjusted = PdfReadDocument.Open(bytes).Pages[0].GetTextSpans().Last().X;
        double unadjusted = PdfReadDocument.Open(bytes, new PdfLoadOptions { AdjustKerningFromTJ = false })
            .Pages[0].GetTextSpans().Last().X;
        Assert.NotEqual(adjusted, unadjusted);
        var limits = new PdfReadLimits { MaxInputBytes = 321, MaxWidgetActions = 87, MaxTextSearchMatches = 23 };
        PdfLoadOptions clone = new PdfLoadOptions { Limits = limits, IncludeArtifactText = true, ParsingMode = PdfParsingMode.Strict }.Clone();
        Assert.NotSame(limits, clone.Limits);
        Assert.Equal(321, clone.Limits.MaxInputBytes);
        Assert.Equal(87, clone.Limits.MaxWidgetActions);
        Assert.Equal(23, clone.Limits.MaxTextSearchMatches);
        Assert.True(clone.IncludeArtifactText);
        Assert.Equal(PdfParsingMode.Strict, clone.ParsingMode);
    }

    private static byte[] BuildMixedDocument(bool unsupportedFilter = false, bool includeXref = true) => BuildPdf(new[] {
        "<< /Type /Catalog /Pages 2 0 R >>",
        "<< /Type /Pages /Count 3 /Kids [3 0 R 4 0 R 5 0 R] >>",
        "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 200] /Resources << /Font << /F1 6 0 R >> >> /Contents 8 0 R >>",
        "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 200] /Resources << /Font << /F1 7 0 R >> >> /Contents 9 0 R /Annots [11 0 R] >>",
        "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 200] /Resources << /Font << /F1 6 0 R >> >> /Contents 10 0 R >>",
        "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>",
        "<< /Type /Font /Subtype /Type0 /BaseFont /Unmapped /Encoding /Mystery-H >>",
        Stream("BT /F1 12 Tf 20 100 Td (First readable page) Tj ET"),
        Stream("BT /F1 12 Tf 20 100 Td <0001> Tj ET", unsupportedFilter ? "/Filter /UnsupportedDecode" : ""),
        Stream("BT /F1 12 Tf 20 100 Td (Last readable page) Tj ET"),
        "<< /Type /Annot /Subtype /Link /Rect [20 30 90 50] /A << /S /URI /URI (https://example.test/source-link) >> >>",
        "<< /Title (Inspection fixture) >>"
    }, 12, includeXref);

    private static byte[] BuildMappedDocument(bool actualText = false, string? contentOverride = null,
        bool includeToUnicode = true, string? encoding = "WinAnsiEncoding", string baseFont = "Helvetica",
        string? fontEncodingDictionary = null, string fontSubtype = "Type1") {
        string content = contentOverride ?? "BT /F1 12 Tf 20 100 Td <4142> Tj ET";
        if (actualText) content = "/Span << /ActualText (Replacement) >> BDC " + content + " EMC";
        return BuildPdf(new[] {
            "<< /Type /Catalog /Pages 2 0 R >>", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 200] /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>",
            "<< /Type /Font /Subtype /" + fontSubtype + " /BaseFont /" + baseFont +
                (fontEncodingDictionary is not null ? " /Encoding " + fontEncodingDictionary : encoding is null ? "" : " /Encoding /" + encoding) +
                (includeToUnicode ? " /ToUnicode 6 0 R" : "") + " >>",
            Stream(content), Stream("1 beginbfchar\n<41> <03A9>\nendbfchar")
        });
    }

    private static byte[] BuildCompositeEvidenceDocument(bool includeToUnicode, bool complete, bool actualText) {
        // The named composite decoder requires complete coverage. FFFF has no predefined
        // Unicode mapping, but remains a valid painted CID 0 for ActualText geometry.
        string codes = includeToUnicode || complete ? "30423044" : "3042FFFF";
        string content = "BT /F1 12 Tf 20 100 Td <" + codes + "> Tj ET";
        if (actualText) content = "/Span << /ActualText (Replacement) >> BDC " + content + " EMC";
        string unicode = complete ? "2 beginbfchar\n<3042> <3042>\n<3044> <3044>\nendbfchar"
            : "1 beginbfchar\n<3042> <3042>\nendbfchar";
        return BuildPdf(new[] {
            "<< /Type /Catalog /Pages 2 0 R >>", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 200] /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>",
            "<< /Type /Font /Subtype /Type0 /BaseFont /HeiseiMin-W3 /Encoding /UniJIS-UCS2-H /DescendantFonts [7 0 R]" +
                (includeToUnicode ? " /ToUnicode 6 0 R" : "") + " >>",
            Stream(content), Stream(unicode),
            "<< /Type /Font /Subtype /CIDFontType0 /BaseFont /HeiseiMin-W3 /CIDSystemInfo << /Registry (Adobe) /Ordering (Japan1) /Supplement 0 >> /DW 1000 >>"
        });
    }

    private static string Stream(string content, string entries = "") => "<< /Length " + content.Length.ToString(CultureInfo.InvariantCulture) +
        " " + entries + " >>\nstream\n" + content + "\nendstream";

    private static byte[] BuildPdf(string[] objects, int? infoObject = null, bool includeXref = true) {
        var text = new StringBuilder("%PDF-1.7\n");
        var offsets = new List<int>();
        for (int index = 0; index < objects.Length; index++) {
            offsets.Add(text.Length);
            text.Append(index + 1).Append(" 0 obj\n").Append(objects[index]).Append("\nendobj\n");
        }
        int xref = text.Length;
        if (includeXref) {
            text.Append("xref\n0 ").Append(objects.Length + 1).Append("\n0000000000 65535 f \n");
            foreach (int offset in offsets) text.Append(offset.ToString("D10", CultureInfo.InvariantCulture)).Append(" 00000 n \n");
        }
        text.Append("trailer\n<< /Root 1 0 R /Size ").Append(objects.Length + 1);
        if (infoObject.HasValue) text.Append(" /Info ").Append(infoObject.Value).Append(" 0 R");
        text.Append(" >>\n");
        if (includeXref) text.Append("startxref\n").Append(xref).Append('\n');
        text.Append("%%EOF\n");
        return Encoding.ASCII.GetBytes(text.ToString());
    }
}
