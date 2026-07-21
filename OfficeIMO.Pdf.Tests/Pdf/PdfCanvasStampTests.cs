using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfCanvasStampTests {
    [Fact]
    public void ContentStampsGeneralVisualCanvasWithPageContext() {
        byte[] target = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Existing page body"))
            .ToBytes();
        int callbackCount = 0;

        PdfDocument stamped = PdfDocument.Open(target).Stamp.Content((canvas, context) => {
            callbackCount++;
            canvas.Text("Canvas page " + context.PageNumber, 36D, 36D, 220D, 30D, fontSize: 14D)
                .Table(new[] {
                    new[] { PdfTableCell.TextCell("Name"), PdfTableCell.TextCell("Value") },
                    new[] { PdfTableCell.TextCell("Mode"), PdfTableCell.RichTextCell(new[] { TextRun.Bolded("General visual canvas") }) }
                }, 36D, 90D, Math.Min(420D, context.Width - 72D), 100D)
                .Image(PdfPngTestImages.CreateRgbPng(20, 80, 180), 36D, 210D, 24D, 24D, alternativeText: "Blue marker");
        });

        string text = stamped.Read.Text();
        Assert.Equal(1, callbackCount);
        Assert.Contains("Existing page body", text, StringComparison.Ordinal);
        Assert.Contains("Canvas page 1", text, StringComparison.Ordinal);
        Assert.Contains("General visual canvas", text, StringComparison.Ordinal);
        Assert.False(PdfInspector.Probe(stamped.ToBytes()).HasEncryption);
    }

    [Fact]
    public void ContentSupportsAuthenticatedEncryptedTargetsAndReportsIgnoredRestrictions() {
        byte[] encrypted = CreateRestrictedPdf("open", "owner", "Encrypted existing body");
        var options = new PdfReadOptions {
            Password = "open",
            PermissionPolicy = PdfPermissionPolicy.IgnoreRestrictions
        };
        PdfDocument target = PdfDocument.Open(encrypted, options);

        PdfOperationResult<PdfDocument> result = target.Stamp.TryContent(
            (canvas, _) => canvas.Text("Authorized canvas stamp", 30D, 30D, 260D, 30D),
            options: options);

        Assert.True(result.Succeeded, string.Join(Environment.NewLine, result.Diagnostics));
        PdfMutationPlan plan = Assert.IsType<PdfMutationPlan>(result.MutationPlan);
        Assert.True(plan.PermissionRestrictionsIgnored);
        Assert.Contains("Input.PermissionRestrictionsIgnored", plan.Warnings);
        PdfDocument output = result.RequireValue();
        Assert.False(PdfInspector.Probe(output.ToBytes()).HasEncryption);
        Assert.Contains("Encrypted existing body", output.Read.Text(), StringComparison.Ordinal);
        Assert.Contains("Authorized canvas stamp", output.Read.Text(), StringComparison.Ordinal);
    }

    [Fact]
    public void ContentStampsManyPagesWithPageSpecificCanvasContent() {
        byte[] target = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Body one"))
            .PageBreak()
            .Paragraph(paragraph => paragraph.Text("Body two"))
            .PageBreak()
            .Paragraph(paragraph => paragraph.Text("Body three"))
            .ToBytes();
        int callbackCount = 0;

        byte[] stamped = PdfDocument.Open(target).Stamp.Content((canvas, context) => {
            callbackCount++;
            canvas.Text("Stamp page " + context.PageNumber, 30D, 30D, 180D, 24D);
        }).ToBytes();

        PdfReadDocument readback = PdfReadDocument.Open(stamped);
        Assert.Equal(3, callbackCount);
        Assert.Equal(3, readback.Pages.Count);
        for (int pageIndex = 0; pageIndex < readback.Pages.Count; pageIndex++) {
            string pageText = readback.Pages[pageIndex].ExtractText();
            Assert.Contains("Body " + new[] { "one", "two", "three" }[pageIndex], pageText, StringComparison.Ordinal);
            Assert.Contains("Stamp page " + (pageIndex + 1), pageText, StringComparison.Ordinal);
        }
    }

    [Fact]
    public void ContentPreservesRepeatedPageSelectorLayersInOneRewrite() {
        byte[] target = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Existing body"))
            .ToBytes();
        int callbackCount = 0;

        byte[] stamped = PdfDocument.Open(target).Stamp.Content(
            (canvas, _) => {
                callbackCount++;
                canvas.Text("Layer " + callbackCount, 30D, 30D + (callbackCount * 30D), 140D, 24D);
            },
            new PdfCanvasStampOptions {
                TargetPages = PdfPageSelector.Parse("1,1")
            }).ToBytes();

        string text = PdfReadDocument.Open(stamped).Pages[0].ExtractText();
        Assert.Equal(2, callbackCount);
        Assert.Contains("Existing body", text, StringComparison.Ordinal);
        Assert.Contains("Layer 1", text, StringComparison.Ordinal);
        Assert.Contains("Layer 2", text, StringComparison.Ordinal);
    }

    [Fact]
    public void ContentRejectsInteractiveAnnotationsBecauseItIsVisualOnly() {
        PdfDocument target = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Existing page"));

        NotSupportedException exception = Assert.Throws<NotSupportedException>(() =>
            target.Stamp.Content(canvas => canvas.Clip(0D, 0D, 100D, 100D, nested =>
                nested.TextAnnotation("Interactive note", 10D, 10D))));

        Assert.Contains("visual content only", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void ContentRejectsInteractiveMetadataThatWouldNotSurviveVisualImport() {
        PdfDocument target = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Existing page"));

        NotSupportedException linkException = Assert.Throws<NotSupportedException>(() =>
            target.Stamp.Content(canvas => canvas.Text(
                new[] { TextRun.Link("Link", "https://example.com") },
                10D,
                10D,
                120D,
                30D)));
        NotSupportedException formException = Assert.Throws<NotSupportedException>(() =>
            target.Stamp.Content(canvas => canvas.Table(
                new[] { new[] { PdfTableCell.WithFormFields("Field", new[] { PdfTableCellFormField.TextField("Entry") }) } },
                10D,
                50D,
                160D,
                50D)));

        Assert.Contains("visual content only", linkException.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("visual content only", formException.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void ContentRejectsTaggedStructureThatVisualImportCannotPreserve() {
        PdfDocument target = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Existing page"));
        var taggedOptions = new PdfCanvasStampOptions().UseRenderingOptions(
            new PdfOptions().SetTaggedStructureMode(PdfTaggedStructureMode.CatalogMarkers));

        NotSupportedException renderingException = Assert.Throws<NotSupportedException>(() =>
            target.Stamp.Content(canvas => canvas.Text("Tagged text", 10D, 10D, 120D, 30D), taggedOptions));
        NotSupportedException figureException = Assert.Throws<NotSupportedException>(() =>
            target.Stamp.Content(canvas => canvas.Figure("Tagged figure", nested =>
                nested.Text("Figure", 10D, 10D, 120D, 30D))));
        NotSupportedException structureException = Assert.Throws<NotSupportedException>(() =>
            target.Stamp.Content(canvas => canvas.Structure(PdfCanvasStructureRole.Paragraph, nested =>
                nested.Text("Structured paragraph", 10D, 10D, 180D, 30D))));

        Assert.Contains("structure", renderingException.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("structure", figureException.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("structure", structureException.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void ContentIgnoresDocumentChromeFromRenderingOptions() {
        var renderingOptions = new PdfOptions {
            ShowHeader = true,
            HeaderFormat = "Rendering-options header",
            ShowPageNumbers = true,
            FooterFormat = "Rendering-options footer",
            BackgroundColor = PdfColor.FromRgb(240, 240, 240)
        };
        var stampOptions = new PdfCanvasStampOptions().UseRenderingOptions(renderingOptions);
        PdfDocument target = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Existing body"));

        PdfDocument stamped = target.Stamp.Content(
            canvas => canvas.Text("Callback content", 36D, 36D, 220D, 30D),
            stampOptions);

        string text = stamped.Read.Text();
        Assert.Contains("Existing body", text, StringComparison.Ordinal);
        Assert.Contains("Callback content", text, StringComparison.Ordinal);
        Assert.DoesNotContain("Rendering-options header", text, StringComparison.Ordinal);
        Assert.DoesNotContain("Rendering-options footer", text, StringComparison.Ordinal);
    }

    [Fact]
    public void OverlayPageReadsEncryptedSourceWithItsOwnPasswordPolicy() {
        byte[] target = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Target page")).ToBytes();
        byte[] source = CreateRestrictedPdf("source-open", "source-owner", "Encrypted overlay source");
        var overlayOptions = new PdfPageOverlayOptions {
            SourceReadOptions = new PdfReadOptions {
                Password = "source-open",
                PermissionPolicy = PdfPermissionPolicy.IgnoreRestrictions
            }
        };

        PdfDocument stamped = PdfDocument.Open(target).Stamp.OverlayPage(source, overlayOptions);

        Assert.Contains("Target", stamped.Read.Text(), StringComparison.Ordinal);
        Assert.Contains("Encrypted", stamped.Read.Text(), StringComparison.Ordinal);
        Assert.Contains("overlay source", stamped.Read.Text(), StringComparison.Ordinal);
    }

    [Fact]
    public void ContentCarriesRegisteredNamedFontsThroughRenderingOptions() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        const string familyName = "Canvas Stamp Family";
        var renderingOptions = new PdfOptions()
            .RegisterNamedFontFamily(new PdfEmbeddedFontFamily(familyName, File.ReadAllBytes(fontPath)));
        var stampOptions = new PdfCanvasStampOptions()
            .UseRenderingOptions(renderingOptions);
        byte[] target = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Existing body"))
            .ToBytes();

        PdfDocument stamped = PdfDocument.Open(target).Stamp.Content(
            canvas => canvas.Text(
                new[] { TextRun.Normal("Named font stamp", fontFamily: familyName) },
                36D,
                36D,
                220D,
                30D),
            stampOptions);

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(stamped.ToBytes()));
        Assert.Contains(
            pdf.GetPage(1).Letters,
            letter => letter.Value == "N" &&
                      letter.FontName.Contains("CanvasStampFamily", StringComparison.OrdinalIgnoreCase));
    }

    private static byte[] CreateRestrictedPdf(string userPassword, string ownerPassword, string text) {
        var encryption = new PdfStandardEncryptionOptions(userPassword) {
            OwnerPassword = ownerPassword,
            AllowedPermissions = PdfStandardPermissions.None
        };
        return PdfDocument.Create(new PdfOptions().SetEncryption(encryption))
            .Paragraph(paragraph => paragraph.Text(text))
            .ToBytes();
    }
}
