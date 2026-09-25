using OfficeIMO.Pdf;
using OfficeIMO.Drawing;
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
    public void SelectedPageBoxFixupsApplyAcrossPagesInOneResult() {
        PdfDocument source = PdfDocument.Load(PdfDocument.Create(new PdfOptions { PageWidth = 300D, PageHeight = 200D })
            .Paragraph(paragraph => paragraph.Text("First"))
            .PageBreak()
            .Paragraph(paragraph => paragraph.Text("Second"))
            .ToBytes());
        PdfProductionPreflightReport report = source.Proof.PreflightProduction();

        PdfProductionFixupResult result = report.ApplySelected(report.FixupProposals.Select(static proposal => proposal.Index).ToArray());

        Assert.Equal(2, result.Document.Inspect().Pages.Count);
        Assert.All(result.Document.Inspect().Pages, static page => {
            Assert.NotNull(page.TrimBox);
            Assert.NotNull(page.BleedBox);
        });
        Assert.DoesNotContain(result.After.Findings, static finding => finding.Kind == PdfProductionFindingKind.InvalidPageBoxes);
    }

    [Fact]
    public void PageBoxFixupReopensGrowthBeyondSourceInputLimit() {
        byte[] bytes = PdfDocument.Create(new PdfOptions { PageWidth = 300D, PageHeight = 200D })
            .Paragraph(paragraph => paragraph.Text("Trim proof")).ToBytes();
        var limits = new PdfLoadOptions { Limits = new PdfReadLimits { MaxInputBytes = bytes.LongLength } };
        PdfProductionPreflightReport report = PdfDocument.Load(bytes, limits).Proof.PreflightProduction();

        PdfProductionFixupResult result = report.ApplySelected(
            report.FixupProposals.Select(static proposal => proposal.Index).ToArray());

        Assert.True(result.Document.ToBytes().LongLength > bytes.LongLength);
        Assert.DoesNotContain(result.After.Findings,
            static finding => finding.Kind == PdfProductionFindingKind.InvalidPageBoxes);
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
    public void ExtractedImageVariantMatchesOnlyItsPaintIntentPlacement() {
        var relative = new PdfImagePlacement(1, "Im0", 5, 0,
            10D, 0D, 0D, 10D, 0D, 0D, 0D, 0D, 10D, 10D,
            renderingIntent: OfficeIccRenderingIntent.RelativeColorimetric);
        var perceptual = new PdfImagePlacement(1, "Im0", 5, 0,
            10D, 0D, 0D, 10D, 20D, 0D, 20D, 0D, 10D, 10D,
            renderingIntent: OfficeIccRenderingIntent.Perceptual);
        var extracted = new PdfExtractedImage(1, "Im0", 5, 10, 10, 8,
            "DeviceRGB", "", new byte[] { 1, 2, 3 }, null, null, false,
            renderingIntent: OfficeIccRenderingIntent.Perceptual);

        Assert.Same(perceptual, Assert.Single(PdfLogicalPage.MatchImagePlacements(extracted,
            new[] { relative, perceptual })));
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

    [Fact]
    public void PdfX4CandidateRejectsRgbPrintOutputIntent() {
        byte[] bytes = PdfDocument.Create(new PdfOptions().SetSrgbOutputIntent())
            .Paragraph(paragraph => paragraph.Text("Output profile")).ToBytes();
        byte[] rgbPrintProfile = IccMabTestProfiles.CreateRgbXyz16OutputDeviceWithDistinctOutputIntents();
        byte[] rewritten = PdfDocumentObjectGraphRewriter.Rewrite(bytes, null, null, (objects, security) => {
            PdfDictionary catalog = Assert.IsType<PdfDictionary>(Assert.Single(objects.Values,
                static item => item.Value is PdfDictionary dictionary &&
                    dictionary.Get<PdfName>("Type")?.Name == "Catalog").Value);
            PdfArray intents = Assert.IsType<PdfArray>(PdfObjectLookup.ResolveChain(objects, catalog.Items["OutputIntents"]));
            PdfDictionary intent = Assert.IsType<PdfDictionary>(PdfObjectLookup.ResolveChain(objects, Assert.Single(intents.Items)));
            intent.Items["S"] = new PdfName("GTS_PDFX");
            int profileNumber = objects.Keys.Max() + 1;
            objects[profileNumber] = new PdfIndirectObject(profileNumber, 0, new PdfStream(new PdfDictionary {
                Items = { ["N"] = new PdfNumber(3) }
            }, rgbPrintProfile));
            intent.Items["DestOutputProfile"] = new PdfReference(profileNumber, 0);
            return security.InfoObjectNumber;
        });

        PdfProductionPreflightReport report = PdfDocument.Load(rewritten).Proof.PreflightProduction(
            new PdfProductionPreflightOptions { Profile = PdfProductionPreflightProfile.PdfX4Candidate });

        Assert.Contains(report.Findings, static finding =>
            finding.Kind == PdfProductionFindingKind.InvalidOutputIntent);
    }

    [Fact]
    public void PrintableFreeTextWithoutNormalAppearanceMarksFontEvidenceIncomplete() {
        byte[] source = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Page")).ToBytes();
        PdfDocument annotated = PdfDocument.Load(source).Annotations.Add(new PdfAnnotationCreateOptions {
            Subtype = "FreeText", Contents = "Printable text", GenerateAppearance = false, Flags = 4
        }).ToDocument();

        PdfProductionPreflightReport report = annotated.Proof.PreflightProduction();

        Assert.Contains(report.Findings, static finding =>
            finding.Kind == PdfProductionFindingKind.UninspectableFont &&
            finding.Severity == PdfProductionFindingSeverity.Indeterminate);
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

    [Theory]
    [InlineData(true, "ON", true)]
    [InlineData(false, "OFF", false)]
    public void ImageResolutionUsesPrintLayerState(bool hiddenInView, string printState, bool prints) {
        const string content = "q 80 0 0 80 20 20 cm /Im0 Do Q\n";
        byte[] source = System.Text.Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj", $"<< /Type /Catalog /Pages 2 0 R /OCProperties << /OCGs [6 0 R] /D << /BaseState /ON {(hiddenInView ? "/OFF [6 0 R]" : string.Empty)} /AS [<< /Event /Print /Category [/Print] /OCGs [6 0 R] >>] >> >> >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 120] /Resources << /XObject << /Im0 5 0 R >> >> /Contents 4 0 R >>", "endobj",
            "4 0 obj", "<< /Length " + System.Text.Encoding.ASCII.GetByteCount(content) + " >>", "stream", content.TrimEnd('\n'), "endstream", "endobj",
            "5 0 obj", "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8 /OC 6 0 R /Length 3 >>", "stream", "abc", "endstream", "endobj",
            "6 0 obj", $"<< /Type /OCG /Name (Print image) /Usage << /Print << /PrintState /{printState} >> >> >>", "endobj",
            "trailer", "<< /Root 1 0 R /Size 7 >>", "%%EOF", string.Empty
        }));

        PdfProductionPreflightReport report = PdfDocument.Load(source).Proof.PreflightProduction();

        Assert.Equal(prints, report.Findings.Any(static finding =>
            finding.Kind == PdfProductionFindingKind.LowImageResolution));
        Assert.DoesNotContain(report.Findings, static finding =>
            finding.Kind == PdfProductionFindingKind.UninspectableImageResolution);
        if (prints) {
            Assert.Contains(report.Findings, static finding => finding.Kind == PdfProductionFindingKind.DeviceRgbColor);
        } else {
            Assert.DoesNotContain(report.Findings, static finding => finding.Kind == PdfProductionFindingKind.DeviceRgbColor);
            Assert.Contains(report.Findings, static finding => finding.Kind == PdfProductionFindingKind.UninspectableColor);
            Assert.Contains(report.Findings, static finding => finding.Kind == PdfProductionFindingKind.UninspectableFont);
        }
    }

    [Fact]
    public void HiddenPrintLayerDoesNotSuppressDefiniteUnlayeredColorFindings() {
        const string content = "0 1 0 rg 10 10 20 20 re f /GS0 gs 40 10 20 20 re f q 80 0 0 80 20 20 cm /Im0 Do Q\n";
        byte[] source = System.Text.Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj", "<< /Type /Catalog /Pages 2 0 R /OCProperties << /OCGs [6 0 R] /D << /BaseState /ON /AS [<< /Event /Print /Category [/Print] /OCGs [6 0 R] >>] >> >> >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 120] /Resources << /XObject << /Im0 5 0 R >> /ExtGState << /GS0 7 0 R >> >> /Contents 4 0 R >>", "endobj",
            "4 0 obj", "<< /Length " + System.Text.Encoding.ASCII.GetByteCount(content) + " >>", "stream", content.TrimEnd('\n'), "endstream", "endobj",
            "5 0 obj", "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8 /OC 6 0 R /Length 3 >>", "stream", "abc", "endstream", "endobj",
            "6 0 obj", "<< /Type /OCG /Name (Hidden print image) /Usage << /Print << /PrintState /OFF >> >> >>", "endobj",
            "7 0 obj", "<< /Type /ExtGState /ca 0.5 >>", "endobj",
            "trailer", "<< /Root 1 0 R /Size 8 >>", "%%EOF", string.Empty
        }));

        PdfProductionPreflightReport report = PdfDocument.Load(source).Proof.PreflightProduction(
            new PdfProductionPreflightOptions { Profile = PdfProductionPreflightProfile.PdfX1aCandidate });

        Assert.Contains(report.Findings, static finding => finding.Kind == PdfProductionFindingKind.DeviceRgbColor);
        Assert.Contains(report.Findings, static finding => finding.Kind == PdfProductionFindingKind.Transparency);
        Assert.Contains(report.Findings, static finding => finding.Kind == PdfProductionFindingKind.UninspectableColor);
        Assert.DoesNotContain(report.Findings, static finding => finding.Kind == PdfProductionFindingKind.LowImageResolution);
    }
}
