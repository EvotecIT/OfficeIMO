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

    [Fact]
    public void UnusedHiddenPrintGroupDoesNotSuppressVisibleLayerFontFinding() {
        const string content = "/OC /Visible BDC BT /F1 12 Tf 10 10 Td (Print text) Tj ET EMC\n";
        byte[] source = RawPrintLayerPdf(content,
            "/Font << /F1 5 0 R >> /Properties << /Visible 6 0 R >>",
            "5 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>\nendobj\n" +
            "6 0 obj\n<< /Type /OCG /Name (Visible) /Usage << /Print << /PrintState /ON >> >> >>\nendobj\n" +
            "7 0 obj\n<< /Type /OCG /Name (Unused) /Usage << /Print << /PrintState /OFF >> >> >>\nendobj",
            "[6 0 R 7 0 R]", "[6 0 R 7 0 R]");

        PdfProductionPreflightReport report = PdfDocument.Load(source).Proof.PreflightProduction();

        Assert.Contains(report.Findings, static finding => finding.Kind == PdfProductionFindingKind.UnembeddedFont);
        Assert.DoesNotContain(report.Findings, static finding => finding.Kind == PdfProductionFindingKind.UninspectableFont);
    }

    [Theory]
    [InlineData("BT /F1 12 Tf 10 10 Td (Print text) Tj ET", "/Font << /F1 5 0 R >>")]
    [InlineData("q 72 0 0 72 10 10 cm /Im0 Do Q", "/XObject << /Im0 5 0 R >>")]
    public void UnlayeredTextOrImageStillReportsRgbBesideHiddenPrintContent(string painted, string resources) {
        string content = "1 0 0 rg " + painted + " /OC /Hidden BDC 0 1 0 rg 10 10 10 10 re f EMC\n";
        string objectFive = resources.IndexOf("/Font", StringComparison.Ordinal) >= 0
            ? "5 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>\nendobj"
            : "5 0 obj\n<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Length 3 >>\nstream\nabc\nendstream\nendobj";
        byte[] source = RawPrintLayerPdf(content, resources + " /Properties << /Hidden 6 0 R >>",
            objectFive + "\n6 0 obj\n<< /Type /OCG /Name (Hidden) /Usage << /Print << /PrintState /OFF >> >> >>\nendobj",
            "[6 0 R]", "[6 0 R]");

        PdfProductionPreflightReport report = PdfDocument.Load(source).Proof.PreflightProduction();

        Assert.Contains(report.Findings, static finding => finding.Kind == PdfProductionFindingKind.DeviceRgbColor);
        Assert.Contains(report.Findings, static finding => finding.Kind == PdfProductionFindingKind.UninspectableColor);
    }

    [Fact]
    public void InvisibleUnlayeredTextDoesNotCreateDefinitePrintColorFinding() {
        const string content = "1 0 0 rg BT /F1 12 Tf 3 Tr 10 10 Td (Screen text) Tj ET /OC /Hidden BDC 10 10 10 10 re f EMC\n";
        byte[] source = RawPrintLayerPdf(content, "/Font << /F1 5 0 R >> /Properties << /Hidden 6 0 R >>",
            "5 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>\nendobj\n" +
            "6 0 obj\n<< /Type /OCG /Name (Hidden) /Usage << /Print << /PrintState /OFF >> >> >>\nendobj",
            "[6 0 R]", "[6 0 R]");

        PdfProductionPreflightReport report = PdfDocument.Load(source).Proof.PreflightProduction();

        Assert.DoesNotContain(report.Findings, static finding => finding.Kind == PdfProductionFindingKind.DeviceRgbColor);
        Assert.Contains(report.Findings, static finding => finding.Kind == PdfProductionFindingKind.UninspectableColor);
    }

    [Fact]
    public void UnlayeredFormPaintRetainsDefiniteRgbFindingBesideHiddenPrintContent() {
        const string formContent = "1 0 0 rg 10 10 20 20 re f\n";
        const string content = "q /Form0 Do Q /OC /Hidden BDC 10 10 10 10 re f EMC\n";
        byte[] source = RawPrintLayerPdf(content, "/XObject << /Form0 5 0 R >> /Properties << /Hidden 6 0 R >>",
            "5 0 obj\n<< /Type /XObject /Subtype /Form /BBox [0 0 100 100] /Length " +
            System.Text.Encoding.ASCII.GetByteCount(formContent) + " >>\nstream\n" + formContent.TrimEnd('\n') +
            "\nendstream\nendobj\n" +
            "6 0 obj\n<< /Type /OCG /Name (Hidden) /Usage << /Print << /PrintState /OFF >> >> >>\nendobj",
            "[6 0 R]", "[6 0 R]");

        PdfProductionPreflightReport report = PdfDocument.Load(source).Proof.PreflightProduction();

        Assert.Contains(report.Findings, static finding => finding.Kind == PdfProductionFindingKind.DeviceRgbColor);
        Assert.Contains(report.Findings, static finding => finding.Kind == PdfProductionFindingKind.UninspectableColor);
    }

    [Fact]
    public void PrintHiddenMalformedImageIsNotDecodedThroughViewState() {
        const string content = "q 72 0 0 72 10 10 cm /Im0 Do Q\n";
        byte[] source = RawPrintLayerPdf(content, "/XObject << /Im0 5 0 R >>",
            "5 0 obj\n<< /Type /XObject /Subtype /Image /Width 10000 /Height 10000 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /FlateDecode /OC 6 0 R /Length 3 >>\nstream\nabc\nendstream\nendobj\n" +
            "6 0 obj\n<< /Type /OCG /Name (Hidden) /Usage << /Print << /PrintState /OFF >> >> >>\nendobj",
            "[6 0 R]", "[6 0 R]");

        PdfProductionPreflightReport report = PdfDocument.Load(source).Proof.PreflightProduction();

        Assert.DoesNotContain(report.Findings, static finding => finding.Kind == PdfProductionFindingKind.LowImageResolution);
    }

    [Fact]
    public void UnsupportedPrintLayerRuleRetainsUnlayeredLowResolutionImage() {
        const string content = "q 72 0 0 72 10 10 cm /Im0 Do Q /OC /Layer BDC 10 10 10 10 re f EMC\n";
        byte[] source = RawPrintLayerPdf(content,
            "/XObject << /Im0 5 0 R >> /Properties << /Layer 6 0 R >>",
            "5 0 obj\n<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Length 3 >>\nstream\nabc\nendstream\nendobj\n" +
            "6 0 obj\n<< /Type /OCG /Name (Layer) >>\nendobj",
            "[6 0 R]", "[6 0 R]", "/Print /View");

        PdfProductionPreflightReport report = PdfDocument.Load(source).Proof.PreflightProduction();

        Assert.Contains(report.Findings, static finding => finding.Kind == PdfProductionFindingKind.LowImageResolution);
        Assert.Contains(report.Findings, static finding => finding.Kind == PdfProductionFindingKind.UninspectableImageResolution);
    }

    [Fact]
    public void ColorSelectedBeforeHiddenPrintBlockStillPaintsAfterIt() {
        const string content = "1 0 0 rg /OC /Hidden BDC 0 1 0 rg 10 10 10 10 re f EMC 40 10 20 20 re f\n";
        byte[] source = RawPrintLayerPdf(content, "/Properties << /Hidden 6 0 R >>",
            "6 0 obj\n<< /Type /OCG /Name (Hidden) /Usage << /Print << /PrintState /OFF >> >> >>\nendobj",
            "[6 0 R]", "[6 0 R]");

        PdfProductionPreflightReport report = PdfDocument.Load(source).Proof.PreflightProduction();

        Assert.Contains(report.Findings, static finding => finding.Kind == PdfProductionFindingKind.DeviceRgbColor);
    }

    [Fact]
    public void GraphicsStateSelectedInsideHiddenPrintBlockAffectsLaterPaint() {
        const string content = "/OC /Hidden BDC 1 0 0 rg /GS0 gs EMC 40 10 20 20 re f\n";
        byte[] source = RawPrintLayerPdf(content,
            "/Properties << /Hidden 6 0 R >> /ExtGState << /GS0 7 0 R >>",
            "6 0 obj\n<< /Type /OCG /Name (Hidden) /Usage << /Print << /PrintState /OFF >> >> >>\nendobj\n" +
            "7 0 obj\n<< /Type /ExtGState /ca 0.5 >>\nendobj",
            "[6 0 R]", "[6 0 R]");

        PdfProductionPreflightReport report = PdfDocument.Load(source).Proof.PreflightProduction(
            new PdfProductionPreflightOptions { Profile = PdfProductionPreflightProfile.PdfX1aCandidate });

        Assert.Contains(report.Findings, static finding => finding.Kind == PdfProductionFindingKind.DeviceRgbColor);
        Assert.Contains(report.Findings, static finding => finding.Kind == PdfProductionFindingKind.Transparency);
    }

    [Theory]
    [InlineData("/ca 0.5", "/ca 1")]
    [InlineData("/BM /Multiply", "/BM /Normal")]
    public void GraphicsStateResetInsideHiddenPrintBlockDoesNotMarkLaterPaintTransparent(
        string firstState, string resetState) {
        const string content = "/OC /Hidden BDC /GS1 gs /GS2 gs EMC 40 10 20 20 re f\n";
        byte[] source = RawPrintLayerPdf(content,
            "/Properties << /Hidden 6 0 R >> /ExtGState << /GS1 << " + firstState + " >> /GS2 << " + resetState + " >> >>",
            "6 0 obj\n<< /Type /OCG /Name (Hidden) /Usage << /Print << /PrintState /OFF >> >> >>\nendobj",
            "[6 0 R]", "[6 0 R]");

        PdfProductionPreflightReport report = PdfDocument.Load(source).Proof.PreflightProduction(
            new PdfProductionPreflightOptions { Profile = PdfProductionPreflightProfile.PdfX1aCandidate });

        Assert.DoesNotContain(report.Findings, static finding => finding.Kind == PdfProductionFindingKind.Transparency);
        Assert.Contains(report.Findings, static finding => finding.Kind == PdfProductionFindingKind.UninspectableColor);
    }

    [Fact]
    public void UnsupportedPrintRuleRetainsUnlayeredImageInsideForm() {
        const string formContent = "q 72 0 0 72 10 10 cm /Im0 Do Q\n";
        const string content = "q /Fm0 Do Q /OC /Layer BDC 10 10 10 10 re f EMC\n";
        byte[] source = RawPrintLayerPdf(content,
            "/XObject << /Fm0 5 0 R >> /Properties << /Layer 6 0 R >>",
            "5 0 obj\n<< /Type /XObject /Subtype /Form /BBox [0 0 100 100] /Resources << /XObject << /Im0 7 0 R >> >> /Length " +
            System.Text.Encoding.ASCII.GetByteCount(formContent) + " >>\nstream\n" + formContent.TrimEnd('\n') + "\nendstream\nendobj\n" +
            "6 0 obj\n<< /Type /OCG /Name (Layer) >>\nendobj\n" +
            "7 0 obj\n<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Length 3 >>\nstream\nabc\nendstream\nendobj",
            "[6 0 R]", "[6 0 R]", "/Print /View");

        PdfProductionPreflightReport report = PdfDocument.Load(source).Proof.PreflightProduction();

        Assert.Contains(report.Findings, static finding => finding.Kind == PdfProductionFindingKind.LowImageResolution);
        Assert.Contains(report.Findings, static finding => finding.Kind == PdfProductionFindingKind.UninspectableImageResolution);
    }

    [Fact]
    public void UnsupportedPrintRuleStillReportsDefinitelyUnembeddedUnlayeredFont() {
        const string content = "BT /F1 12 Tf 10 80 Td (Visible) Tj ET /OC /Layer BDC BT /F1 12 Tf 10 40 Td (Hidden) Tj ET EMC\n";
        byte[] source = RawPrintLayerPdf(content,
            "/Font << /F1 7 0 R >> /Properties << /Layer 6 0 R >>",
            "6 0 obj\n<< /Type /OCG /Name (Layer) >>\nendobj\n" +
            "7 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>\nendobj",
            "[6 0 R]", "[6 0 R]", "/Print /View");

        PdfProductionPreflightReport report = PdfDocument.Load(source).Proof.PreflightProduction(
            new PdfProductionPreflightOptions { Profile = PdfProductionPreflightProfile.PdfX1aCandidate });

        Assert.Contains(report.Findings, static finding => finding.Kind == PdfProductionFindingKind.UnembeddedFont);
        Assert.Contains(report.Findings, static finding => finding.Kind == PdfProductionFindingKind.UninspectableFont);
    }

    [Fact]
    public void UnsupportedPrintRuleRetainsFontInPrintableAnnotationAppearance() {
        const string content = "/OC /Layer BDC 10 10 20 20 re f EMC\n";
        const string appearance = "BT /F1 12 Tf 10 10 Td (Print) Tj ET";
        byte[] source = RawPrintLayerPdf(content,
            "/Properties << /Layer 6 0 R >>",
            "5 0 obj\n<< /Type /Annot /Subtype /Stamp /F 4 /Rect [10 10 90 30] /AP << /N 8 0 R >> >>\nendobj\n" +
            "6 0 obj\n<< /Type /OCG /Name (Layer) >>\nendobj\n" +
            "7 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>\nendobj\n" +
            "8 0 obj\n<< /Type /XObject /Subtype /Form /BBox [0 0 80 20] /Resources << /Font << /F1 7 0 R >> >> /Length " + appearance.Length + " >>\nstream\n" + appearance + "\nendstream\nendobj",
            "[6 0 R]", "[6 0 R]", "/Print /View", "/Annots [5 0 R]", size: 9);

        PdfProductionPreflightReport report = PdfDocument.Load(source).Proof.PreflightProduction(
            new PdfProductionPreflightOptions { Profile = PdfProductionPreflightProfile.PdfX1aCandidate });

        Assert.Contains(report.Findings, static finding => finding.Kind == PdfProductionFindingKind.UnembeddedFont);
    }

    [Fact]
    public void InvalidOptionalContentPropertyMakesPrintEvidenceIndeterminate() {
        const string content = "/OC /Missing BDC BT /F1 12 Tf 10 80 Td (Unknown) Tj ET EMC\n";
        byte[] source = RawPrintLayerPdf(content,
            "/Font << /F1 7 0 R >> /Properties << /Layer 6 0 R >>",
            "6 0 obj\n<< /Type /OCG /Name (Layer) >>\nendobj\n" +
            "7 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>\nendobj",
            "[6 0 R]", "[6 0 R]", "/Print /View");

        PdfProductionPreflightReport report = PdfDocument.Load(source).Proof.PreflightProduction(
            new PdfProductionPreflightOptions { Profile = PdfProductionPreflightProfile.PdfX1aCandidate });

        Assert.Contains(report.Findings, static finding => finding.Kind == PdfProductionFindingKind.UninspectableFont);
        Assert.DoesNotContain(report.Findings, static finding => finding.Kind == PdfProductionFindingKind.UnembeddedFont);
    }

    [Fact]
    public void DeviceRgbImageAliasRemainsDefiniteBesideUnresolvedPrintLayer() {
        const string content = "q 72 0 0 72 10 10 cm /Im0 Do Q /OC /Layer BDC 10 10 20 20 re f EMC\n";
        byte[] source = RawPrintLayerPdf(content,
            "/XObject << /Im0 5 0 R >> /ColorSpace << /CS1 /DeviceRGB >> /Properties << /Layer 6 0 R >>",
            "5 0 obj\n<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /CS1 /BitsPerComponent 8 /Length 3 >>\nstream\nabc\nendstream\nendobj\n" +
            "6 0 obj\n<< /Type /OCG /Name (Layer) >>\nendobj",
            "[6 0 R]", "[6 0 R]", "/Print /View");

        PdfProductionPreflightReport report = PdfDocument.Load(source).Proof.PreflightProduction(
            new PdfProductionPreflightOptions { Profile = PdfProductionPreflightProfile.PdfX1aCandidate });

        Assert.Contains(report.Findings, static finding => finding.Kind == PdfProductionFindingKind.DeviceRgbColor);
    }

    [Fact]
    public void UnsupportedPrintRuleStillRejectsMalformedAlwaysVisibleType3Glyph() {
        const string content = "BT /F1 12 Tf 10 80 Td (A) Tj ET /OC /Layer BDC 10 10 20 20 re f EMC\n";
        const string glyph = "0 0 m 10 10 l S";
        byte[] source = RawPrintLayerPdf(content,
            "/Font << /F1 7 0 R >> /Properties << /Layer 6 0 R >>",
            "6 0 obj\n<< /Type /OCG /Name (Layer) >>\nendobj\n" +
            "7 0 obj\n<< /Type /Font /Subtype /Type3 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] " +
            "/CharProcs << /A 8 0 R >> /Encoding << /Type /Encoding /Differences [65 /A] >> " +
            "/FirstChar 65 /LastChar 65 /Widths [500] /Resources << >> >>\nendobj\n" +
            "8 0 obj\n<< /Length " + glyph.Length + " >>\nstream\n" + glyph + "\nendstream\nendobj",
            "[6 0 R]", "[6 0 R]", "/Print /View");

        PdfProductionPreflightReport report = PdfDocument.Load(source).Proof.PreflightProduction(
            new PdfProductionPreflightOptions { Profile = PdfProductionPreflightProfile.PdfX1aCandidate });

        Assert.Contains(report.Findings, static finding => finding.Kind == PdfProductionFindingKind.UnembeddedFont);
    }

    [Fact]
    public void UnsupportedPrintRuleStillReportsDefinitelyUnlayeredCalRgbPaint() {
        const string content = "/CS1 cs 0.2 0.3 0.4 sc 10 10 20 20 re f /OC /Layer BDC 40 10 20 20 re f EMC\n";
        byte[] source = RawPrintLayerPdf(content,
            "/ColorSpace << /CS1 [/CalRGB << /WhitePoint [1 1 1] >>] >> /Properties << /Layer 6 0 R >>",
            "6 0 obj\n<< /Type /OCG /Name (Layer) >>\nendobj",
            "[6 0 R]", "[6 0 R]", "/Print /View");

        PdfProductionPreflightReport report = PdfDocument.Load(source).Proof.PreflightProduction(
            new PdfProductionPreflightOptions { Profile = PdfProductionPreflightProfile.PdfX1aCandidate });

        Assert.Contains(report.Findings, static finding => finding.Kind == PdfProductionFindingKind.DeviceIndependentColor);
        Assert.Contains(report.Findings, static finding => finding.Kind == PdfProductionFindingKind.UninspectableColor);
    }

    [Fact]
    public void UnsupportedPrintRuleStillReportsIndependentIndexedColorPaint() {
        const string content = "/CS1 cs 0 sc 10 10 20 20 re f /OC /Layer BDC 40 10 20 20 re f EMC\n";
        byte[] source = RawPrintLayerPdf(content,
            "/ColorSpace << /CS1 [/Indexed [/CalRGB << /WhitePoint [1 1 1] >>] 0 <FF0000>] >> /Properties << /Layer 6 0 R >>",
            "6 0 obj\n<< /Type /OCG /Name (Layer) >>\nendobj",
            "[6 0 R]", "[6 0 R]", "/Print /View");

        PdfProductionPreflightReport report = PdfDocument.Load(source).Proof.PreflightProduction(
            new PdfProductionPreflightOptions { Profile = PdfProductionPreflightProfile.PdfX1aCandidate });

        Assert.Contains(report.Findings, static finding => finding.Kind == PdfProductionFindingKind.DeviceIndependentColor);
    }

    [Theory]
    [InlineData("/DeviceRGB", "")]
    [InlineData("/CS1", "/ColorSpace << /CS1 /DeviceRGB >>")]
    public void UnsupportedPrintRuleStillReportsRgbSelectedByGenericColorOperators(string colorSpace, string colorResources) {
        string content = colorSpace + " cs 1 0 0 sc 10 10 20 20 re f " +
            colorSpace + " CS 0 1 0 SC 40 10 20 20 re S /OC /Layer BDC 80 10 20 20 re f EMC\n";
        byte[] source = RawPrintLayerPdf(content,
            colorResources + " /Properties << /Layer 6 0 R >>",
            "6 0 obj\n<< /Type /OCG /Name (Layer) >>\nendobj",
            "[6 0 R]", "[6 0 R]", "/Print /View");

        PdfProductionPreflightReport report = PdfDocument.Load(source).Proof.PreflightProduction(
            new PdfProductionPreflightOptions { Profile = PdfProductionPreflightProfile.PdfX1aCandidate });

        Assert.Contains(report.Findings, static finding => finding.Kind == PdfProductionFindingKind.DeviceRgbColor);
    }

    [Fact]
    public void MalformedCatalogOutputIntentIsInvalidRatherThanMissing() {
        byte[] source = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Print source")).ToBytes();
        byte[] malformed = PdfDocumentObjectGraphRewriter.Rewrite(source, null, null, (objects, security) => {
            PdfDictionary catalog = Assert.IsType<PdfDictionary>(Assert.Single(objects.Values,
                static item => item.Value is PdfDictionary dictionary &&
                    dictionary.Get<PdfName>("Type")?.Name == "Catalog").Value);
            catalog.Items["OutputIntents"] = new PdfName("Broken");
            return security.InfoObjectNumber;
        });

        PdfProductionPreflightReport report = PdfDocument.Load(malformed).Proof.PreflightProduction();

        Assert.Contains(report.Findings, static finding => finding.Kind == PdfProductionFindingKind.InvalidOutputIntent);
        Assert.DoesNotContain(report.Findings, static finding => finding.Kind == PdfProductionFindingKind.MissingOutputIntent);
    }

    [Fact]
    public void UnsupportedPrintLayerAndPrintableAnnotationShareOneResolutionUnknown() {
        const string content = "/OC /Layer BDC 10 10 20 20 re f EMC\n";
        byte[] source = RawPrintLayerPdf(content, "/Properties << /Layer 6 0 R >>",
            "6 0 obj\n<< /Type /OCG /Name (Layer) >>\nendobj\n" +
            "7 0 obj\n<< /Type /Annot /Subtype /Text /Rect [10 10 30 30] /F 4 >>\nendobj",
            "[6 0 R]", "[6 0 R]", "/Print /View", "/Annots [7 0 R]");

        PdfProductionPreflightReport report = PdfDocument.Load(source).Proof.PreflightProduction();

        Assert.Single(report.Findings, static finding =>
            finding.Kind == PdfProductionFindingKind.UninspectableImageResolution);
    }

    [Fact]
    public void UnusedGroupWithUnsupportedPrintStateDoesNotDegradePageEvidence() {
        const string content = "/OC /Visible BDC q 72 0 0 72 10 10 cm /Im0 Do Q EMC\n";
        byte[] source = RawPrintLayerPdf(content, "/XObject << /Im0 5 0 R >> /Properties << /Visible 7 0 R >>",
            "5 0 obj\n<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Length 3 >>\nstream\nabc\nendstream\nendobj\n" +
            "6 0 obj\n<< /Type /OCG /Name (Unused) /Usage << /Print << /PrintState /Maybe >> >> >>\nendobj\n" +
            "7 0 obj\n<< /Type /OCG /Name (Visible) /Usage << /Print << /PrintState /ON >> >> >>\nendobj",
            "[6 0 R 7 0 R]", "[6 0 R 7 0 R]");

        PdfProductionPreflightReport report = PdfDocument.Load(source).Proof.PreflightProduction();

        Assert.Contains(report.Findings, static finding => finding.Kind == PdfProductionFindingKind.LowImageResolution);
        Assert.DoesNotContain(report.Findings, static finding => finding.Kind == PdfProductionFindingKind.UninspectableImageResolution);
    }

    [Fact]
    public void InactiveAnnotationAppearanceDoesNotContributeFontFinding() {
        const string active = "";
        const string inactive = "BT /F1 12 Tf 10 10 Td (Inactive) Tj ET";
        byte[] source = System.Text.Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.7", "1 0 obj", "<< /Type /Catalog /Pages 2 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 120] /Resources << >> /Contents 4 0 R /Annots [5 0 R] >>", "endobj",
            "4 0 obj", "<< /Length 0 >>", "stream", active, "endstream", "endobj",
            "5 0 obj", "<< /Type /Annot /Subtype /Widget /Rect [10 10 100 30] /F 4 /AS /Off /AP << /N << /Off 6 0 R /On 7 0 R >> >> >>", "endobj",
            "6 0 obj", "<< /Type /XObject /Subtype /Form /BBox [0 0 90 20] /Length 0 >>", "stream", active, "endstream", "endobj",
            "7 0 obj", "<< /Type /XObject /Subtype /Form /BBox [0 0 90 20] /Resources << /Font << /F1 8 0 R >> >> /Length " + inactive.Length + " >>", "stream", inactive, "endstream", "endobj",
            "8 0 obj", "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>", "endobj",
            "trailer", "<< /Root 1 0 R /Size 9 >>", "%%EOF", string.Empty
        }));

        PdfProductionPreflightReport report = PdfDocument.Load(source).Proof.PreflightProduction();

        Assert.DoesNotContain(report.Findings, static finding => finding.Kind == PdfProductionFindingKind.UnembeddedFont);
    }

    [Fact]
    public void ScreenOnlyAnnotationLayerDoesNotDegradePrintedImageEvidence() {
        const string content = "q 72 0 0 72 10 10 cm /Im0 Do Q\n";
        byte[] source = RawPrintLayerPdf(content, "/XObject << /Im0 5 0 R >>",
            "5 0 obj\n<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Length 3 >>\nstream\nabc\nendstream\nendobj\n" +
            "6 0 obj\n<< /Type /OCG /Name (Screen only) >>\nendobj\n" +
            "7 0 obj\n<< /Type /Annot /Subtype /Text /Rect [0 0 10 10] /F 0 /OC 6 0 R >>\nendobj",
            "[6 0 R]", "[6 0 R]", "/Print /View", "/Annots [7 0 R]");

        PdfProductionPreflightReport report = PdfDocument.Load(source).Proof.PreflightProduction();

        Assert.Contains(report.Findings, static finding => finding.Kind == PdfProductionFindingKind.LowImageResolution);
        Assert.DoesNotContain(report.Findings, static finding => finding.Kind == PdfProductionFindingKind.UninspectableImageResolution);
    }

    private static byte[] RawPrintLayerPdf(string content, string resources, string extraObjects,
        string groups, string printGroups, string categories = "/Print", string pageEntries = "", int size = 8) => System.Text.Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj", $"<< /Type /Catalog /Pages 2 0 R /OCProperties << /OCGs {groups} /D << /BaseState /ON /AS [<< /Event /Print /Category [{categories}] /OCGs {printGroups} >>] >> >> >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj", $"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 120] /Resources << {resources} >> /Contents 4 0 R {pageEntries} >>", "endobj",
            "4 0 obj", "<< /Length " + System.Text.Encoding.ASCII.GetByteCount(content) + " >>", "stream", content.TrimEnd('\n'), "endstream", "endobj",
            extraObjects, "trailer", $"<< /Root 1 0 R /Size {size} >>", "%%EOF", string.Empty
        }));
}
