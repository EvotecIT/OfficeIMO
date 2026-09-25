using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfProductionPreflightTests {
    [Fact]
    public void ProposedPageBoxesRequireSelectionAndAreReinspectedAfterRewrite() {
        PdfDocument source = PdfDocument.Load(PdfDocument.Create(new PdfOptions { PageWidth = 300D, PageHeight = 200D })
            .Paragraph(paragraph => paragraph.Text("Print proof"))
            .ToBytes());
        PdfProductionPreflightReport report = source.Proof.PreflightProduction(
            new PdfProductionPreflightOptions { Profile = PdfProductionPreflightProfile.PdfX4Candidate });

        Assert.Contains(report.Findings, finding => finding.Kind == PdfProductionFindingKind.MissingOutputIntent && finding.PageNumber is null);
        Assert.Contains(report.Findings, finding => finding.Kind == PdfProductionFindingKind.InvalidPageBoxes && finding.PageNumber == 1);
        Assert.Equal(new[] { PdfPageBoundaryBox.TrimBox, PdfPageBoundaryBox.BleedBox },
            report.FixupProposals.Select(static proposal => proposal.Box));
        Assert.Null(source.Inspect().Pages[0].TrimBox);

        PdfProductionFixupResult fixedBoxes = report.ApplySelected(report.FixupProposals.Select(static proposal => proposal.Index).ToArray());

        Assert.NotNull(fixedBoxes.Document.Inspect().Pages[0].TrimBox);
        Assert.DoesNotContain(fixedBoxes.After.Findings, finding => finding.Kind == PdfProductionFindingKind.InvalidPageBoxes);
        Assert.Contains(fixedBoxes.After.Findings, finding => finding.Kind == PdfProductionFindingKind.MissingOutputIntent);
        Assert.Equal(PdfArtifactFingerprint.ComputeSha256(fixedBoxes.Document.ToBytes()), fixedBoxes.After.SourceSha256);
    }

    [Fact]
    public void SelectedPageReportsLowImageResolutionAndDeviceColorWithPageGeometry() {
        byte[] image = PdfPngTestImages.CreateRgbPng(30, 90, 180);
        PdfDocument source = PdfDocument.Load(PdfDocument.Create(new PdfOptions { PageWidth = 240D, PageHeight = 180D })
            .Paragraph(paragraph => paragraph.Text("First page"))
            .PageBreak()
            .Canvas(canvas => canvas.Image(image, 20D, 20D, 100D, 100D))
            .ToBytes());

        PdfProductionPreflightReport report = source.Proof.PreflightProduction(new PdfProductionPreflightOptions {
            Profile = PdfProductionPreflightProfile.PdfX1aCandidate,
            PageSelection = PdfPageSelection.From(2),
            MaxPages = 1
        });

        Assert.Equal(new[] { 2 }, report.InspectedPages);
        PdfProductionFinding resolution = Assert.Single(report.Findings,
            finding => finding.Kind == PdfProductionFindingKind.LowImageResolution);
        Assert.Equal(2, resolution.PageNumber);
        Assert.NotNull(resolution.VisualBounds);
        Assert.True(resolution.ObservedImagePpi < 300D);
        Assert.Contains(report.Findings, finding => finding.Kind == PdfProductionFindingKind.DeviceRgbColor && finding.PageNumber == 2);
        Assert.DoesNotContain(report.Findings, finding => finding.PageNumber == 1);
    }

    [Fact]
    public void EffectiveImageResolutionAccountsForPageUserUnit() {
        byte[] image = PdfPngTestImages.CreateRgbPng(30, 30);
        byte[] source = PdfDocument.Create(new PdfOptions { PageWidth = 240D, PageHeight = 180D })
            .Canvas(canvas => canvas.Image(image, 20D, 20D, 72D, 72D)).ToBytes();
        byte[] scaled = PdfDocumentObjectGraphRewriter.Rewrite(source, null, null, (objects, security) => {
            PdfIndirectObject page = Assert.Single(objects.Values, static item =>
                item.Value is PdfDictionary dictionary && dictionary.Get<PdfName>("Type")?.Name == "Page");
            Assert.IsType<PdfDictionary>(page.Value).Items["UserUnit"] = new PdfNumber(2D);
            return security.InfoObjectNumber;
        });

        PdfProductionPreflightReport report = PdfDocument.Load(scaled).Proof.PreflightProduction(
            new PdfProductionPreflightOptions { MinimumImagePpi = 20D });

        PdfProductionFinding resolution = Assert.Single(report.Findings,
            static finding => finding.Kind == PdfProductionFindingKind.LowImageResolution);
        Assert.InRange(resolution.ObservedImagePpi!.Value, 14.9D, 15.1D);
    }

    [Fact]
    public void DisplayProfileDoesNotSatisfyPrintOutputIntent() {
        byte[] source = PdfDocument.Create(new PdfOptions().SetSrgbOutputIntent())
            .Paragraph(paragraph => paragraph.Text("Screen profile"))
            .ToBytes();

        PdfProductionPreflightReport report = PdfDocument.Load(source).Proof.PreflightProduction();

        Assert.Contains(report.Findings, static finding =>
            finding.Kind == PdfProductionFindingKind.InvalidOutputIntent);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PrintableAnnotationMarksColorAndResolutionEvidenceIncomplete(bool generateAppearance) {
        byte[] source = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Page")).ToBytes();
        PdfDocument annotated = PdfDocument.Load(source).Annotations.Add(new PdfAnnotationCreateOptions {
            Subtype = "Text", Contents = "Print note", GenerateAppearance = generateAppearance, Flags = 4
        }).ToDocument();

        PdfProductionPreflightReport report = annotated.Proof.PreflightProduction();

        Assert.Contains(report.Findings, static finding =>
            finding.Kind == PdfProductionFindingKind.UninspectableColor &&
            finding.Severity == PdfProductionFindingSeverity.Indeterminate);
        Assert.Contains(report.Findings, static finding =>
            finding.Kind == PdfProductionFindingKind.UninspectableImageResolution &&
            finding.Severity == PdfProductionFindingSeverity.Indeterminate);
    }

    [Fact]
    public void ViewHiddenImageWithPrintUsageKeepsResolutionIndeterminate() {
        const string content = "q 80 0 0 80 20 20 cm /Im0 Do Q\n";
        byte[] source = System.Text.Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj", "<< /Type /Catalog /Pages 2 0 R /OCProperties << /OCGs [6 0 R] /D << /BaseState /ON /OFF [6 0 R] /AS [<< /Event /Print /Category [/Print] /OCGs [6 0 R] >>] >> >> >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 120] /Resources << /XObject << /Im0 5 0 R >> >> /Contents 4 0 R >>", "endobj",
            "4 0 obj", "<< /Length " + System.Text.Encoding.ASCII.GetByteCount(content) + " >>", "stream", content.TrimEnd('\n'), "endstream", "endobj",
            "5 0 obj", "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8 /OC 6 0 R /Length 3 >>", "stream", "abc", "endstream", "endobj",
            "6 0 obj", "<< /Type /OCG /Name (Print image) /Usage << /Print << /PrintState /ON >> >> >>", "endobj",
            "trailer", "<< /Root 1 0 R /Size 7 >>", "%%EOF", string.Empty
        }));

        PdfProductionPreflightReport report = PdfDocument.Load(source).Proof.PreflightProduction();

        Assert.Contains(report.Findings, static finding =>
            finding.Kind == PdfProductionFindingKind.UninspectableImageResolution &&
            finding.Severity == PdfProductionFindingSeverity.Indeterminate &&
            finding.Message.Contains("optional-content image", StringComparison.Ordinal));
    }
}
