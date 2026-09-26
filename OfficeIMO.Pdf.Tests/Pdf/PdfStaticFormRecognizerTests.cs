using System;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfStaticFormRecognizerTests {
    [Fact]
    public void ProposalsRequireSelectionAndCreateReopenableFields() {
        byte[] sourceBytes = CreateStaticForm();
        PdfDocument source = PdfDocument.Load(sourceBytes);

        PdfStaticFormRecognitionReport report = source.Forms.RecognizeStaticLayout();

        Assert.True(report.Proposals.Count == 2,
            "Expected two proposals; diagnostics: " + string.Join(", ", report.Diagnostics.Select(static diagnostic => diagnostic.Code)));
        Assert.Equal(new[] { PdfFormFieldCreationKind.Text, PdfFormFieldCreationKind.CheckBox },
            report.Proposals.Select(static proposal => proposal.Kind));
        Assert.Empty(PdfInspector.Inspect(sourceBytes).FormFields);

        PdfAcroFormEditResult edited = report.ApplySelected(report.Proposals.Select(static proposal => proposal.Index).ToArray());
        PdfFormField[] fields = PdfInspector.Inspect(edited.ToBytes()).FormFields.ToArray();
        Assert.Equal(2, fields.Length);
        Assert.All(report.Proposals, proposal => Assert.Contains(fields, field => field.Name == proposal.SuggestedName));
        Assert.All(fields, static field => Assert.Single(field.Widgets));
        PdfVisualComparisonReport visual = source.Proof.CompareVisual(edited.ToDocument());
        OfficeRasterImage rendered = VisualBaselineTestSupport.DecodePng(visual.Pages[0].ActualPng, "Edited form page must render as PNG.");
        Assert.True(rendered.GetPixel(110, 27).R < 200, "The static text-field outline must remain visible below its widget.");
        Assert.True(rendered.GetPixel(110, 35).R > 245, "The proposed widget must leave the writing area unobscured.");
    }

    [Fact]
    public void OcrLabelCanSupportAProposalFromAnExactSourceSnapshot() {
        byte[] sourceBytes = PdfDocument.Create(new PdfOptions { PageWidth = 400, PageHeight = 300 })
            .Canvas(canvas => canvas.Shape(Box(140D, 20D), 100D, 30D))
            .ToBytes();
        var ocr = new[] { new PdfStaticFormTextEvidence(1, "Customer ID", 20D, 30D, 85D, 49D, 0.95D) };
        PdfStaticFormRecognitionReport report = PdfDocument.Load(sourceBytes).Forms.RecognizeStaticLayout(ocrText: ocr);

        PdfStaticFormFieldProposal proposal = Assert.Single(report.Proposals);
        Assert.True(proposal.UsedOcrLabel);
        Assert.Equal("Customer ID", proposal.Label);

        Assert.Equal(PdfArtifactFingerprint.ComputeSha256(sourceBytes), report.SourceSha256);
        Assert.Contains(report.ApplySelected(new[] { proposal.Index }).Fields, field => field.Name == proposal.SuggestedName);
    }

    [Fact]
    public void OcrEvidenceBelowHalfConfidenceUsesTheConfiguredProposalThreshold() {
        byte[] source = PdfDocument.Create(new PdfOptions { PageWidth = 400, PageHeight = 300 })
            .Canvas(canvas => canvas.Shape(Box(140D, 20D), 100D, 30D)).ToBytes();
        var ocr = new[] { new PdfStaticFormTextEvidence(1, "Customer ID", 20D, 30D, 85D, 49D, 0.49D) };

        PdfStaticFormRecognitionReport accepted = PdfDocument.Load(source).Forms.RecognizeStaticLayout(
            new PdfStaticFormRecognitionOptions { MinimumConfidence = 0.3D }, ocr);
        PdfStaticFormRecognitionReport rejected = PdfDocument.Load(source).Forms.RecognizeStaticLayout(
            new PdfStaticFormRecognitionOptions { MinimumConfidence = 0.4D }, ocr);

        Assert.Single(accepted.Proposals);
        Assert.Empty(rejected.Proposals);
        Assert.Contains(rejected.Diagnostics, static diagnostic => diagnostic.Code == "low-confidence");
    }

    [Fact]
    public void GradientStrokeIsAVisibleOutlineCandidate() {
        const string content = "/Pattern CS /P1 SCN 1 w 80 80 120 20 re S";
        byte[] source = System.Text.Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.4", "1 0 obj", "<< /Type /Catalog /Pages 2 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 240 200] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /Resources << /Pattern << /P1 5 0 R >> >> /Contents 4 0 R >>", "endobj",
            "4 0 obj", "<< /Length " + content.Length + " >>", "stream", content, "endstream", "endobj",
            "5 0 obj", "<< /Type /Pattern /PatternType 2 /Shading << /ShadingType 2 /ColorSpace /DeviceRGB /Coords [80 80 200 80] /Function << /FunctionType 2 /Domain [0 1] /C0 [1 0 0] /C1 [0 0 1] /N 1 >> /Extend [true true] >> >>", "endobj",
            "trailer", "<< /Root 1 0 R >>", "%%EOF", ""
        }));

        IReadOnlyList<PdfPageVisualPrimitive> primitives = PdfReadDocument.Open(source).Pages[0].GetIdentityVisualPrimitives();
        Assert.Contains(primitives, primitive => primitive.Kind == PdfPageVisualPrimitiveKind.Rectangle && primitive.StrokeGradient != null);
        var ocr = new[] { new PdfStaticFormTextEvidence(1, "Name", 10D, 100D, 70D, 120D, 1D) };
        PdfStaticFormRecognitionReport report = PdfDocument.Load(source).Forms.RecognizeStaticLayout(ocrText: ocr);
        Assert.Single(report.Proposals);
    }

    [Fact]
    public void PartialWhiteRepaintDoesNotClearAnOccupiedField() {
        OfficeShape painted = OfficeShape.Rectangle(140D, 20D);
        painted.FillColor = OfficeColor.Red;
        painted.StrokeColor = null;
        OfficeShape white = OfficeShape.Rectangle(126D, 20D);
        white.FillColor = OfficeColor.White;
        white.StrokeColor = null;
        OfficeShape outline = Box(140D, 20D);
        outline.FillColor = null;
        byte[] source = PdfDocument.Create(new PdfOptions { PageWidth = 400, PageHeight = 300 })
            .Canvas(canvas => canvas.Shape(painted, 100D, 28D).Shape(white, 100D, 28D)
                .Shape(outline, 100D, 28D).Text("Name:", 20D, 28D, 70D, 20D)).ToBytes();

        Assert.Empty(PdfDocument.Load(source).Forms.RecognizeStaticLayout().Proposals);
    }

    [Fact]
    public void LargerOpaqueWhiteRepaintClearsAnOccupiedField() {
        OfficeShape painted = OfficeShape.Rectangle(140D, 20D);
        painted.FillColor = OfficeColor.Red;
        painted.StrokeColor = null;
        OfficeShape white = OfficeShape.Rectangle(400D, 300D);
        white.FillColor = OfficeColor.White;
        white.StrokeColor = null;
        OfficeShape outline = Box(140D, 20D);
        outline.FillColor = null;
        byte[] source = PdfDocument.Create(new PdfOptions { PageWidth = 400, PageHeight = 300 })
            .Canvas(canvas => canvas.Shape(painted, 100D, 28D).Shape(white, 0D, 0D)
                .Shape(outline, 100D, 28D).Text("Name:", 20D, 28D, 70D, 20D)).ToBytes();

        Assert.Single(PdfDocument.Load(source).Forms.RecognizeStaticLayout().Proposals);
    }

    [Fact]
    public void WhiteRepaintCanClearOnlyTheImageAreaInsideAField() {
        byte[] image = PdfPngTestImages.CreateRgbPng(180, 60);
        OfficeShape white = OfficeShape.Rectangle(140D, 20D);
        white.FillColor = OfficeColor.White;
        white.StrokeColor = null;
        OfficeShape outline = Box(140D, 20D);
        outline.FillColor = null;
        byte[] source = PdfDocument.Create(new PdfOptions { PageWidth = 400, PageHeight = 300 })
            .Canvas(canvas => canvas.Image(image, 80D, 20D, 180D, 60D)
                .Shape(white, 100D, 30D).Shape(outline, 100D, 30D))
            .ToBytes();
        var ocr = new[] { new PdfStaticFormTextEvidence(1, "Name", 20D, 30D, 70D, 50D, 1D) };

        Assert.Single(PdfDocument.Load(source).Forms.RecognizeStaticLayout(ocrText: ocr).Proposals);
    }

    [Fact]
    public void DifferenceBlendedOutlineIsNotAnEmptyField() {
        const string content = "/GS1 gs 1 1 1 rg 1 w 80 80 120 20 re B";
        byte[] source = System.Text.Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.4", "1 0 obj", "<< /Type /Catalog /Pages 2 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 240 200] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /Resources << /ExtGState << /GS1 << /BM /Difference >> >> >> /Contents 4 0 R >>", "endobj",
            "4 0 obj", "<< /Length " + content.Length + " >>", "stream", content, "endstream", "endobj",
            "trailer", "<< /Root 1 0 R >>", "%%EOF", ""
        }));
        var ocr = new[] { new PdfStaticFormTextEvidence(1, "Name", 10D, 100D, 70D, 120D, 1D) };

        PdfStaticFormRecognitionReport report = PdfDocument.Load(source).Forms.RecognizeStaticLayout(ocrText: ocr);
        Assert.Empty(report.Proposals);
        Assert.Contains(report.Diagnostics, static diagnostic => diagnostic.Code == "unsupported-effect");
    }

    [Fact]
    public void RepaintedOutlineDoesNotBecomeAFieldCandidate() {
        OfficeShape outline = Box(140D, 20D);
        outline.FillColor = null;
        OfficeShape white = OfficeShape.Rectangle(400D, 300D);
        white.FillColor = OfficeColor.White;
        white.StrokeColor = null;
        byte[] source = PdfDocument.Create(new PdfOptions { PageWidth = 400, PageHeight = 300 })
            .Canvas(canvas => canvas.Shape(outline, 100D, 28D).Shape(white, 0D, 0D)
                .Text("Name:", 20D, 28D, 70D, 20D)).ToBytes();

        PdfStaticFormRecognitionReport report = PdfDocument.Load(source).Forms.RecognizeStaticLayout();

        Assert.Empty(report.Proposals);
        Assert.Contains(report.Diagnostics, static diagnostic => diagnostic.Code == "occluded-outline");
    }

    [Fact]
    public void RepaintedNativeLabelDoesNotSupportAFieldCandidate() {
        OfficeShape white = OfficeShape.Rectangle(100D, 30D);
        white.FillColor = OfficeColor.White;
        white.StrokeColor = null;
        byte[] source = PdfDocument.Create(new PdfOptions { PageWidth = 400, PageHeight = 300 })
            .Canvas(canvas => canvas.Text("Name:", 20D, 28D, 70D, 20D)
                .Shape(white, 10D, 20D).Shape(Box(140D, 20D), 100D, 28D)).ToBytes();

        Assert.Empty(PdfDocument.Load(source).Forms.RecognizeStaticLayout().Proposals);
    }

    [Fact]
    public void RepaintedUnderlineDoesNotBecomeAFieldCandidate() {
        OfficeShape underline = OfficeShape.Line(0D, 0D, 140D, 0D);
        underline.StrokeColor = OfficeColor.Black;
        OfficeShape white = OfficeShape.Rectangle(400D, 300D);
        white.FillColor = OfficeColor.White;
        white.StrokeColor = null;
        byte[] source = PdfDocument.Create(new PdfOptions { PageWidth = 400, PageHeight = 300 })
            .Canvas(canvas => canvas.Shape(underline, 100D, 48D).Shape(white, 0D, 0D)
                .Text("Name:", 20D, 28D, 70D, 20D)).ToBytes();

        PdfStaticFormRecognitionReport report = PdfDocument.Load(source).Forms.RecognizeStaticLayout();

        Assert.Empty(report.Proposals);
        Assert.Contains(report.Diagnostics, static diagnostic => diagnostic.Code == "occluded-outline");
    }

    [Fact]
    public void SuggestedNamesReserveFieldAncestorsAcrossUnselectedPages() {
        byte[] sourceBytes = PdfDocument.Create(new PdfOptions { PageWidth = 400, PageHeight = 300 })
            .Canvas(canvas => canvas.Text("Customer:", 20D, 28D, 85D, 20D).Shape(Box(140D, 20D), 125D, 28D))
            .PageBreak()
            .Paragraph(paragraph => paragraph.Text("Existing field lives here"))
            .ToBytes();
        PdfAcroFormEditResult withExistingField = PdfDocument.Load(sourceBytes).Forms.Edit(edit => edit.Create(
            new PdfFormFieldCreateOptions { Name = "customer.name", PageNumber = 2, X = 20D, Y = 20D, Width = 100D, Height = 20D }));

        PdfStaticFormRecognitionReport report = withExistingField.ToDocument().Forms.RecognizeStaticLayout(
            new PdfStaticFormRecognitionOptions { PageSelection = PdfPageSelection.From(1) });

        PdfStaticFormFieldProposal proposal = Assert.Single(report.Proposals);
        Assert.Equal("customer_2", proposal.SuggestedName);
        Assert.Contains(report.ApplySelected(new[] { proposal.Index }).Fields, field => field.Name == "customer_2");
    }

    [Fact]
    public void GeneratedEncryptedSourceRetainsSnapshotAndHonorsMutationGate() {
        PdfDocument source = PdfDocument.Create(new PdfOptions { PageWidth = 400, PageHeight = 300 }
            .SetEncryption(new PdfStandardEncryptionOptions("open") { OwnerPassword = "owner" }))
            .Canvas(canvas => canvas.Text("Name:", 20D, 28D, 70D, 20D).Shape(Box(140D, 20D), 100D, 28D));

        PdfStaticFormRecognitionReport report = source.Forms.RecognizeStaticLayout();

        Assert.Single(report.Proposals);
        Assert.Throws<PdfMutationBlockedException>(() => report.ApplySelected(new[] { 0 }));
    }

    [Fact]
    public void ProposedWidgetLeavesSourceMarksVisible() {
        PdfStaticFormRecognitionReport report = PdfDocument.Load(CreateStaticForm()).Forms.RecognizeStaticLayout();

        Assert.All(report.Proposals, proposal => {
            PdfFormFieldStyle style = proposal.ToCreateOptions().Style!;
            Assert.Null(style.BackgroundColor);
            Assert.Null(style.BorderColor);
            Assert.Equal(0D, style.BorderWidth);
        });
    }

    [Fact]
    public void LowConfidenceCandidatesRemainVisibleInDiagnostics() {
        PdfStaticFormRecognitionReport report = PdfDocument.Load(CreateStaticForm()).Forms.RecognizeStaticLayout(
            new PdfStaticFormRecognitionOptions { MinimumConfidence = 0.8D });

        Assert.DoesNotContain(report.Proposals, static proposal => proposal.Kind == PdfFormFieldCreationKind.CheckBox);
        Assert.Contains(report.Diagnostics, static diagnostic => diagnostic.Code == "low-confidence");
    }

    [Fact]
    public void ContainedValueTextAndGradientPaintAreNotEmptyFields() {
        OfficeShape gradientBox = Box(100D, 28D);
        gradientBox.FillColor = null;
        gradientBox.FillGradient = OfficeLinearGradient.Horizontal(OfficeColor.Red, OfficeColor.Blue);
        byte[] source = PdfDocument.Create(new PdfOptions { PageWidth = 400D, PageHeight = 300D })
            .Canvas(canvas => canvas
                .Text("Name:", 20D, 28D, 70D, 20D)
                .Shape(Box(100D, 28D), 140D, 20D)
                .Text("Alice", 150D, 25D, 60D, 20D)
                .Text("Status:", 20D, 78D, 70D, 20D)
                .Shape(gradientBox, 140D, 70D))
            .ToBytes();

        PdfStaticFormRecognitionReport report = PdfDocument.Load(source).Forms.RecognizeStaticLayout();

        Assert.Empty(report.Proposals);
    }

    [Fact]
    public void ColoredOpaqueRepaintDoesNotLeaveNativeLabelEvidence() {
        OfficeShape cover = OfficeShape.Rectangle(95D, 42D);
        cover.FillColor = OfficeColor.Red;
        cover.StrokeColor = null;
        byte[] source = PdfDocument.Create(new PdfOptions { PageWidth = 400D, PageHeight = 300D })
            .Canvas(canvas => canvas
                .Text("Name:", 20D, 28D, 70D, 20D)
                .Shape(Box(140D, 20D), 100D, 28D)
                .Shape(cover, 0D, 10D))
            .ToBytes();

        PdfStaticFormRecognitionReport report = PdfDocument.Load(source).Forms.RecognizeStaticLayout();

        Assert.Empty(report.Proposals);
    }

    [Fact]
    public void WhiteOutlineOnDarkBackdropRemainsAFieldCandidate() {
        OfficeShape backdrop = OfficeShape.Rectangle(400D, 300D);
        backdrop.FillColor = OfficeColor.Black;
        backdrop.StrokeColor = null;
        OfficeShape outline = Box(140D, 20D);
        outline.FillColor = null;
        outline.StrokeColor = OfficeColor.White;
        byte[] source = PdfDocument.Create(new PdfOptions { PageWidth = 400D, PageHeight = 300D })
            .Canvas(canvas => canvas.Shape(backdrop, 0D, 0D).Shape(outline, 100D, 28D))
            .ToBytes();
        var label = new[] { new PdfStaticFormTextEvidence(1, "Name", 20D, 28D, 70D, 48D, 1D) };

        PdfStaticFormRecognitionReport report = PdfDocument.Load(source).Forms.RecognizeStaticLayout(ocrText: label);

        Assert.Single(report.Proposals);
    }

    [Fact]
    public void LabelOcclusionComparisonsConsumeCandidateScanBudget() {
        OfficeShape fill = OfficeShape.Rectangle(10D, 10D);
        fill.FillColor = OfficeColor.Red;
        fill.StrokeColor = null;
        byte[] source = PdfDocument.Create(new PdfOptions { PageWidth = 400D, PageHeight = 300D })
            .Canvas(canvas => canvas.Text("Name", 20D, 28D, 70D, 20D)
                .Shape(fill, 300D, 100D).Shape(fill, 320D, 100D)).ToBytes();

        PdfReadLimitException limit = Assert.Throws<PdfReadLimitException>(() =>
            PdfDocument.Load(source).Forms.RecognizeStaticLayout(
                new PdfStaticFormRecognitionOptions { MaxCandidateScanWork = 1 }));

        Assert.Equal(PdfReadLimitKind.UnderstandingArtifacts, limit.Kind);
    }

    [Fact]
    public void CheckedStaticBoxIsNotProposedAsAnEmptyCheckbox() {
        OfficeShape firstMark = OfficeShape.Line(0D, 0D, 9D, 9D);
        firstMark.StrokeColor = OfficeColor.Black;
        OfficeShape secondMark = OfficeShape.Line(0D, 9D, 9D, 0D);
        secondMark.StrokeColor = OfficeColor.Black;
        byte[] source = PdfDocument.Create(new PdfOptions { PageWidth = 400D, PageHeight = 300D })
            .Canvas(canvas => canvas
                .Shape(Box(15D, 15D), 100D, 75D)
                .Shape(firstMark, 103D, 78D)
                .Shape(secondMark, 103D, 78D)
                .Text("I agree", 125D, 72D, 100D, 20D))
            .ToBytes();

        PdfStaticFormRecognitionReport report = PdfDocument.Load(source).Forms.RecognizeStaticLayout();

        Assert.Empty(report.Proposals);
        Assert.Contains(report.Diagnostics, static diagnostic => diagnostic.Code == "occupied-field");
    }

    [Theory]
    [InlineData(9D, 0D)]
    [InlineData(0D, 9D)]
    public void AxisAlignedStrokeInsideStaticBoxCountsAsAnOccupiedField(double deltaX, double deltaY) {
        OfficeShape mark = OfficeShape.Line(0D, 0D, deltaX, deltaY);
        mark.StrokeColor = OfficeColor.Black;
        byte[] source = PdfDocument.Create(new PdfOptions { PageWidth = 400D, PageHeight = 300D })
            .Canvas(canvas => canvas
                .Shape(Box(15D, 15D), 100D, 75D)
                .Shape(mark, 103D, 78D)
                .Text("I agree", 125D, 72D, 100D, 20D))
            .ToBytes();

        PdfStaticFormRecognitionReport report = PdfDocument.Load(source).Forms.RecognizeStaticLayout();

        Assert.Empty(report.Proposals);
        Assert.Contains(report.Diagnostics, static diagnostic => diagnostic.Code == "occupied-field");
    }

    [Fact]
    public void ImageRenderedValueOccupiesOutlinedField() {
        byte[] image = PdfPngTestImages.CreateRgbPng(20, 20);
        byte[] source = PdfDocument.Create(new PdfOptions { PageWidth = 400D, PageHeight = 300D })
            .Canvas(canvas => canvas
                .Text("Name:", 20D, 28D, 70D, 20D)
                .Shape(Box(140D, 20D), 100D, 28D)
                .Image(image, 110D, 30D, 16D, 16D))
            .ToBytes();

        PdfStaticFormRecognitionReport report = PdfDocument.Load(source).Forms.RecognizeStaticLayout();

        Assert.Empty(report.Proposals);
        Assert.Contains(report.Diagnostics, static diagnostic => diagnostic.Code == "occupied-field");
    }

    [Fact]
    public void EdgeNearCheckmarkAndSmallFilledDotBothOccupyAStaticBox() {
        OfficeShape mark = OfficeShape.Line(0D, 0D, 13D, 13D);
        mark.StrokeColor = OfficeColor.Black;
        OfficeShape dot = OfficeShape.Rectangle(3D, 3D);
        dot.FillColor = OfficeColor.Black;
        dot.StrokeColor = null;
        byte[] source = PdfDocument.Create(new PdfOptions { PageWidth = 400D, PageHeight = 300D })
            .Canvas(canvas => canvas
                .Shape(Box(15D, 15D), 100D, 75D)
                .Shape(mark, 101D, 76D)
                .Text("First", 125D, 72D, 100D, 20D)
                .Shape(Box(15D, 15D), 100D, 115D)
                .Shape(dot, 106D, 121D)
                .Text("Second", 125D, 112D, 100D, 20D))
            .ToBytes();

        PdfStaticFormRecognitionReport report = PdfDocument.Load(source).Forms.RecognizeStaticLayout();

        Assert.Empty(report.Proposals);
        Assert.Equal(2, report.Diagnostics.Count(static diagnostic => diagnostic.Code == "occupied-field"));
    }

    [Fact]
    public void StrokeCrossingCheckboxEdgeOccupiesTheField() {
        OfficeShape mark = OfficeShape.Line(0D, 0D, 30D, 8D);
        mark.StrokeColor = OfficeColor.Black;
        byte[] source = PdfDocument.Create(new PdfOptions { PageWidth = 400D, PageHeight = 300D })
            .Canvas(canvas => canvas
                .Shape(Box(15D, 15D), 100D, 75D)
                .Shape(mark, 90D, 78D)
                .Text("I agree", 125D, 72D, 100D, 20D))
            .ToBytes();

        PdfStaticFormRecognitionReport report = PdfDocument.Load(source).Forms.RecognizeStaticLayout();

        Assert.Empty(report.Proposals);
        Assert.Contains(report.Diagnostics, static diagnostic => diagnostic.Code == "occupied-field");
    }

    [Fact]
    public void LargePaintedBackdropDoesNotOccupyATransparentField() {
        OfficeShape backdrop = OfficeShape.Rectangle(400D, 300D);
        backdrop.FillColor = OfficeColor.Blue;
        backdrop.StrokeColor = null;
        OfficeShape field = Box(140D, 20D);
        field.FillColor = null;
        byte[] source = PdfDocument.Create(new PdfOptions { PageWidth = 400D, PageHeight = 300D })
            .Canvas(canvas => canvas
                .Shape(backdrop, 0D, 0D)
                .Text("Name:", 20D, 28D, 70D, 20D)
                .Shape(field, 100D, 28D))
            .ToBytes();

        PdfStaticFormRecognitionReport report = PdfDocument.Load(source).Forms.RecognizeStaticLayout();

        Assert.Single(report.Proposals);
    }

    [Fact]
    public void LargeForegroundPanelCrossingAFieldPreventsAnEmptyProposal() {
        OfficeShape field = Box(140D, 20D);
        field.FillColor = null;
        OfficeShape panel = OfficeShape.Rectangle(200D, 15D);
        panel.FillColor = OfficeColor.Blue;
        panel.StrokeColor = null;
        byte[] source = PdfDocument.Create(new PdfOptions { PageWidth = 400D, PageHeight = 300D })
            .Canvas(canvas => canvas.Text("Name:", 20D, 28D, 70D, 20D)
                .Shape(field, 100D, 28D).Shape(panel, 95D, 35D)).ToBytes();

        Assert.Empty(PdfDocument.Load(source).Forms.RecognizeStaticLayout().Proposals);
    }

    [Fact]
    public void EarlierPanelCrossingOnlyPartOfAFieldPreventsAnEmptyProposal() {
        OfficeShape panel = OfficeShape.Rectangle(200D, 15D);
        panel.FillColor = OfficeColor.Blue;
        panel.StrokeColor = null;
        OfficeShape field = Box(140D, 20D);
        field.FillColor = null;
        byte[] source = PdfDocument.Create(new PdfOptions { PageWidth = 400D, PageHeight = 300D })
            .Canvas(canvas => canvas.Text("Name:", 20D, 28D, 70D, 20D)
                .Shape(panel, 95D, 35D).Shape(field, 100D, 28D)).ToBytes();

        Assert.Empty(PdfDocument.Load(source).Forms.RecognizeStaticLayout().Proposals);
    }

    [Fact]
    public void LaterWhiteFillInsideAnOutlinedFieldKeepsTheEmptyProposal() {
        OfficeShape field = Box(140D, 20D);
        field.FillColor = null;
        OfficeShape white = OfficeShape.Rectangle(120D, 16D);
        white.FillColor = OfficeColor.White;
        white.StrokeColor = null;
        byte[] source = PdfDocument.Create(new PdfOptions { PageWidth = 400D, PageHeight = 300D })
            .Canvas(canvas => canvas.Text("Name:", 20D, 28D, 70D, 20D)
                .Shape(field, 100D, 28D).Shape(white, 110D, 30D)).ToBytes();

        Assert.Single(PdfDocument.Load(source).Forms.RecognizeStaticLayout().Proposals);
    }

    [Fact]
    public void ContentOrderKeysResolveRepaintWhenDoubleOrdersTie() {
        PdfContentOrderKey earlier = PdfContentOrderKey.Root.Append(12).Append(4);
        PdfContentOrderKey later = PdfContentOrderKey.Root.Append(12).Append(5);

        Assert.True(PdfStaticFormRecognizer.IsLater(1D, later, 1D, earlier));
        Assert.False(PdfStaticFormRecognizer.IsLater(1D, earlier, 1D, later));
    }

    [Fact]
    public void WhiteOutlineOnWhitePageIsNotProposed() {
        OfficeShape box = Box(140D, 20D);
        box.StrokeColor = OfficeColor.White;
        byte[] source = PdfDocument.Create(new PdfOptions { PageWidth = 400D, PageHeight = 300D })
            .Canvas(canvas => canvas.Text("Name:", 20D, 28D, 70D, 20D).Shape(box, 100D, 28D))
            .ToBytes();

        PdfStaticFormRecognitionReport report = PdfDocument.Load(source).Forms.RecognizeStaticLayout();

        Assert.Empty(report.Proposals);
        Assert.Contains(report.Diagnostics, static diagnostic => diagnostic.Code == "invisible-outline");
    }

    [Fact]
    public void TransparentNativeLabelDoesNotSupportAProposal() {
        const string content = "1 w 100 205 120 20 re S q /GS1 gs BT /F1 12 Tf 20 208 Td (Name) Tj ET Q";
        byte[] source = System.Text.Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.4", "1 0 obj", "<< /Type /Catalog /Pages 2 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 400 300] /Resources << /Font << /F1 5 0 R >> /ExtGState << /GS1 6 0 R >> >> /Contents 4 0 R >>", "endobj",
            "4 0 obj", "<< /Length " + content.Length + " >>", "stream", content, "endstream", "endobj",
            "5 0 obj", "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>", "endobj",
            "6 0 obj", "<< /Type /ExtGState /ca 0 >>", "endobj",
            "trailer", "<< /Root 1 0 R /Size 7 >>", "%%EOF", ""
        }));

        PdfStaticFormRecognitionReport report = PdfDocument.Load(source).Forms.RecognizeStaticLayout();

        Assert.Empty(report.Proposals);
    }

    [Fact]
    public void LongNativeValueInsideOutlinePreventsAnEmptyFieldProposal() {
        string value = new string('A', 90);
        string content = "1 w 100 205 140 20 re S BT /F1 12 Tf 20 208 Td (Name) Tj ET BT /F1 4 Tf 105 208 Td (" + value + ") Tj ET";
        byte[] source = System.Text.Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.4", "1 0 obj", "<< /Type /Catalog /Pages 2 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 400 300] /Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >>", "endobj",
            "4 0 obj", "<< /Length " + content.Length + " >>", "stream", content, "endstream", "endobj",
            "5 0 obj", "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>", "endobj",
            "trailer", "<< /Root 1 0 R /Size 6 >>", "%%EOF", ""
        }));

        Assert.Empty(PdfDocument.Load(source).Forms.RecognizeStaticLayout().Proposals);
    }

    [Fact]
    public void SoftMaskedNativeLabelDoesNotSupportAVisibleField() {
        const string content = "1 w 100 205 120 20 re S q /GS1 gs BT /F1 12 Tf 20 208 Td (Name) Tj ET Q";
        byte[] source = System.Text.Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.7", "1 0 obj", "<< /Type /Catalog /Pages 2 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 400 300] /Resources << /Font << /F1 5 0 R >> /ExtGState << /GS1 6 0 R >> >> /Contents 4 0 R >>", "endobj",
            "4 0 obj", "<< /Length " + content.Length + " >>", "stream", content, "endstream", "endobj",
            "5 0 obj", "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>", "endobj",
            "6 0 obj", "<< /Type /ExtGState /SMask << /S /Alpha /G 7 0 R >> >>", "endobj",
            "7 0 obj", "<< /Type /XObject /Subtype /Form /BBox [0 0 400 300] /Group << /S /Transparency >> /Length 0 >>", "stream", "", "endstream", "endobj",
            "trailer", "<< /Root 1 0 R /Size 8 >>", "%%EOF", ""
        }));

        Assert.Empty(PdfDocument.Load(source).Forms.RecognizeStaticLayout().Proposals);
    }

    [Fact]
    public void OpaqueFieldFillCanCoverAnEarlierMark() {
        OfficeShape mark = OfficeShape.Line(0D, 0D, 9D, 9D);
        mark.StrokeColor = OfficeColor.Black;
        byte[] source = PdfDocument.Create(new PdfOptions { PageWidth = 400D, PageHeight = 300D })
            .Canvas(canvas => canvas
                .Shape(mark, 103D, 78D)
                .Shape(Box(15D, 15D), 100D, 75D)
                .Text("I agree", 125D, 72D, 100D, 20D))
            .ToBytes();

        PdfStaticFormRecognitionReport report = PdfDocument.Load(source).Forms.RecognizeStaticLayout();

        Assert.Single(report.Proposals);
    }

    [Fact]
    public void NonTruncatingRectangularClipKeepsStaticFieldCandidate() {
        const string content = "q 0 0 400 300 re W n 1 w 100 205 15 15 re S BT /F1 12 Tf 125 208 Td (I agree) Tj ET Q\n";
        byte[] source = System.Text.Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj", "<< /Type /Catalog /Pages 2 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 400 300] /Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >>", "endobj",
            "4 0 obj", "<< /Length " + System.Text.Encoding.ASCII.GetByteCount(content) + " >>", "stream", content.TrimEnd('\n'), "endstream", "endobj",
            "5 0 obj", "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>", "endobj",
            "trailer", "<< /Root 1 0 R /Size 6 >>", "%%EOF", string.Empty
        }));

        PdfStaticFormRecognitionReport report = PdfDocument.Load(source).Forms.RecognizeStaticLayout();

        Assert.Single(report.Proposals);
    }

    [Fact]
    public void NativeLabelOutsideItsClipDoesNotSupportAField() {
        const string content = "1 w 100 205 120 20 re S q 0 0 5 5 re W n BT /F1 12 Tf 20 208 Td (Name) Tj ET Q";
        byte[] source = System.Text.Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.7", "1 0 obj", "<< /Type /Catalog /Pages 2 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 400 300] /Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >>", "endobj",
            "4 0 obj", "<< /Length " + content.Length + " >>", "stream", content, "endstream", "endobj",
            "5 0 obj", "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>", "endobj",
            "trailer", "<< /Root 1 0 R /Size 6 >>", "%%EOF", ""
        }));

        Assert.Empty(PdfDocument.Load(source).Forms.RecognizeStaticLayout().Proposals);
    }

    [Fact]
    public void TextClippedAwayFromAFieldDoesNotOccupyIt() {
        const string content = "1 w 100 205 120 20 re S BT /F1 12 Tf 60 208 Td (Name) Tj ET q 0 0 5 5 re W n BT /F1 12 Tf 110 208 Td (Value) Tj ET Q";
        byte[] source = System.Text.Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.7", "1 0 obj", "<< /Type /Catalog /Pages 2 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 400 300] /Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >>", "endobj",
            "4 0 obj", "<< /Length " + content.Length + " >>", "stream", content, "endstream", "endobj",
            "5 0 obj", "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>", "endobj",
            "trailer", "<< /Root 1 0 R /Size 6 >>", "%%EOF", ""
        }));

        PdfStaticFormRecognitionReport report = PdfDocument.Load(source).Forms.RecognizeStaticLayout();
        Assert.True(report.Proposals.Count == 1,
            "Expected one empty-field proposal; diagnostics: " + string.Join(", ", report.Diagnostics.Select(static d => d.Code)));
    }

    [Theory]
    [InlineData("1 w 103 208 m 112 217 l S")]
    [InlineData("0 0 0 rg 101 206 13 13 re f")]
    public void PaintClippedOutsideCheckboxDoesNotOccupyIt(string clippedPaint) {
        string content = $"q 0 0 10 10 re W n {clippedPaint} Q 1 w 100 205 15 15 re S BT /F1 12 Tf 125 208 Td (I agree) Tj ET\n";
        byte[] source = System.Text.Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj", "<< /Type /Catalog /Pages 2 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 400 300] /Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >>", "endobj",
            "4 0 obj", "<< /Length " + System.Text.Encoding.ASCII.GetByteCount(content) + " >>", "stream", content.TrimEnd('\n'), "endstream", "endobj",
            "5 0 obj", "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>", "endobj",
            "trailer", "<< /Root 1 0 R /Size 6 >>", "%%EOF", string.Empty
        }));

        PdfStaticFormRecognitionReport report = PdfDocument.Load(source).Forms.RecognizeStaticLayout();

        Assert.Single(report.Proposals);
    }

    [Fact]
    public void ExistingLinkAnnotationBlocksOverlappingProposal() {
        byte[] source = PdfDocument.Create(new PdfOptions { PageWidth = 400D, PageHeight = 300D })
            .Canvas(canvas => canvas.Text("Name:", 20D, 28D, 70D, 20D).Shape(Box(140D, 20D), 100D, 28D))
            .ToBytes();
        PdfDocument annotated = PdfDocument.Load(source).Annotations.Add(new PdfAnnotationCreateOptions {
            Subtype = "Link", LinkUri = "https://example.com", Rectangle = new[] { 100D, 252D, 240D, 272D }
        }).ToDocument();

        PdfStaticFormRecognitionReport report = annotated.Forms.RecognizeStaticLayout();

        Assert.Empty(report.Proposals);
        Assert.Contains(report.Diagnostics, static diagnostic => diagnostic.Code == "existing-annotation");
    }

    [Fact]
    public void DiagnosticLimitReportsTruncation() {
        PdfStaticFormRecognitionReport report = PdfDocument.Load(CreateStaticForm()).Forms.RecognizeStaticLayout(
            new PdfStaticFormRecognitionOptions { MinimumConfidence = 1D, MaxDiagnostics = 1 });

        Assert.Empty(report.Proposals);
        Assert.Equal("diagnostics-truncated", Assert.Single(report.Diagnostics).Code);
    }

    [Fact]
    public void CandidateScanWorkLimitAppliesBeforeAProposalIsAccepted() {
        Assert.Throws<PdfReadLimitException>(() => PdfDocument.Load(CreateStaticForm()).Forms.RecognizeStaticLayout(
            new PdfStaticFormRecognitionOptions { MaxCandidateScanWork = 1 }));
    }

    [Fact]
    public void RightToLeftLabelsSuggestRightmostTabFirst() {
        byte[] source = PdfDocument.Create(new PdfOptions { PageWidth = 400, PageHeight = 300 })
            .Canvas(canvas => canvas.Shape(Box(80D, 20D), 70D, 80D).Shape(Box(80D, 20D), 250D, 80D))
            .ToBytes();
        var labels = new[] {
            new PdfStaticFormTextEvidence(1, "\u05E9\u05DD", 70D, 45D, 140D, 65D, 1D),
            new PdfStaticFormTextEvidence(1, "\u05E2\u05D9\u05E8", 250D, 45D, 320D, 65D, 1D)
        };

        PdfStaticFormRecognitionReport report = PdfDocument.Load(source).Forms.RecognizeStaticLayout(ocrText: labels);

        Assert.Equal(2, report.Proposals.Count);
        Assert.True(report.Proposals[0].VisualBounds.Left > report.Proposals[1].VisualBounds.Left);
    }

    [Fact]
    public void RightToLeftTextFieldAcceptsLabelOnItsRight() {
        byte[] source = PdfDocument.Create(new PdfOptions { PageWidth = 400D, PageHeight = 300D })
            .Canvas(canvas => canvas.Shape(Box(140D, 20D), 80D, 80D)).ToBytes();
        var labels = new[] { new PdfStaticFormTextEvidence(1, "\u05E9\u05DD", 230D, 80D, 300D, 100D, 1D) };

        PdfStaticFormRecognitionReport report = PdfDocument.Load(source).Forms.RecognizeStaticLayout(ocrText: labels);

        Assert.Single(report.Proposals);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SoftMaskedOutlineIsNotProposedAsAVisibleField(bool validMask) {
        const string content = "q /GS1 gs 100 205 120 20 re S Q\n";
        byte[] source = System.Text.Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.7", "1 0 obj", "<< /Type /Catalog /Pages 2 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 400 300] /Resources << /ExtGState << /GS1 5 0 R >> >> /Contents 4 0 R >>", "endobj",
            "4 0 obj", "<< /Length " + content.Length + " >>", "stream", content.TrimEnd('\n'), "endstream", "endobj",
            "5 0 obj", "<< /Type /ExtGState /SMask << /S /Alpha /G 6 0 R >> >>", "endobj",
            "6 0 obj", "<< " + (validMask ? "/Type /XObject " : string.Empty) + "/Subtype /Form /BBox [0 0 400 300] /Group << /S /Transparency >> /Length 0 >>", "stream", "", "endstream", "endobj",
            "trailer", "<< /Root 1 0 R /Size 7 >>", "%%EOF", string.Empty
        }));
        var label = new[] { new PdfStaticFormTextEvidence(1, "Name", 10D, 70D, 80D, 90D, 1D) };

        PdfStaticFormRecognitionReport report = PdfDocument.Load(source).Forms.RecognizeStaticLayout(ocrText: label);

        Assert.Empty(report.Proposals);
    }

    private static byte[] CreateStaticForm() {
        OfficeShape textBox = Box(140D, 20D);
        OfficeShape checkBox = Box(15D, 15D);
        return PdfDocument.Create(new PdfOptions { PageWidth = 400, PageHeight = 300 })
            .Canvas(canvas => canvas
                .Text("Name:", 20D, 28D, 70D, 20D)
                .Shape(textBox, 100D, 28D)
                .Shape(checkBox, 100D, 75D)
                .Text("I agree", 125D, 72D, 100D, 20D))
            .ToBytes();
    }

    private static OfficeShape Box(double width, double height) {
        OfficeShape box = OfficeShape.Rectangle(width, height);
        box.FillColor = OfficeColor.White;
        box.StrokeColor = OfficeColor.Black;
        box.StrokeWidth = 1D;
        return box;
    }
}
