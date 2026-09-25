using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfNUpImpositionTests {
    [Fact]
    public void TwoUpCreatesVectorSheetsWithSourceToCellMapping() {
        PdfDocument source = PdfDocument.Create(new PdfOptions { PageSize = new PageSize(240, 180) })
            .Paragraph(paragraph => paragraph.Text("First page"))
            .PageBreak()
            .Paragraph(paragraph => paragraph.Text("Second page"))
            .PageBreak()
            .Paragraph(paragraph => paragraph.Text("Third page"));

        PdfImpositionResult result = PdfDocument.Load(source.ToBytes()).Pages.ImposeNUp(
            new PdfNUpOptions(new PageSize(600, 400), columns: 2, rows: 1));

        Assert.Equal(2, PdfInspector.Inspect(result.Bytes).PageCount);
        Assert.Equal(new[] { 1, 1, 2 }, result.Placements.Select(static placement => placement.SheetPageNumber));
        Assert.Equal(new[] { 0, 1, 0 }, result.Placements.Select(static placement => placement.Column));
        Assert.Equal(18, result.Placements[0].Cell.Left);
        Assert.Equal(304.5, result.Placements[1].Cell.Left);
        Assert.Equal(18, result.Placements[0].Cell.Bottom);
        Assert.Equal(18, result.Placements[1].Cell.Bottom);
        Assert.Equal(277.5, result.Placements[0].Cell.Width);
        Assert.Equal(364, result.Placements[0].Cell.Height);
        Assert.Contains("First page", result.ToDocument().Reader.Text(PdfPageSelection.From(1)));
        Assert.Contains("Second page", result.ToDocument().Reader.Text(PdfPageSelection.From(1)));
        Assert.Contains("Third page", result.ToDocument().Reader.Text(PdfPageSelection.From(2)));
    }

    [Fact]
    public void InteractiveSourceRequiresExplicitVisualOnlyPolicy() {
        byte[] source = PdfDocument.Create().TextAnnotation("Review note")
            .Paragraph(paragraph => paragraph.Text("Annotated source")).ToBytes();
        PdfDocument document = PdfDocument.Load(source);
        var options = new PdfNUpOptions(new PageSize(600, 400), 2, 1);

        Assert.Throws<NotSupportedException>(() => document.Pages.ImposeNUp(options));
        options.AllowSourceFeatureLoss = true;
        PdfImpositionResult result = document.Pages.ImposeNUp(options);
        Assert.Single(result.ToDocument().Reader.Pages());
        Assert.Empty(PdfInspector.Inspect(result.Bytes).Annotations);
        Assert.True(result.SourceFeatureLoss.HasFlag(PdfImpositionSourceFeatureLoss.Annotations));

        PdfDocument formDocument = PdfDocument.Load(PdfDocument.Create().TextField("ReviewedBy", value: "Ada").ToBytes());
        options.AllowSourceFeatureLoss = false;
        Assert.Throws<NotSupportedException>(() => formDocument.Pages.ImposeNUp(options));
        options.AllowSourceFeatureLoss = true;
        PdfImpositionResult formResult = formDocument.Pages.ImposeNUp(options);
        Assert.False(PdfInspector.Inspect(formResult.Bytes).HasForms);
        Assert.True(formResult.SourceFeatureLoss.HasFlag(PdfImpositionSourceFeatureLoss.Forms));
    }

    [Fact]
    public void BookletPadsToFourAndMapsBothSidesOfDuplexSheets() {
        PdfDocument source = PdfDocument.Create(new PdfOptions { PageSize = new PageSize(240, 180) });
        for (int page = 1; page <= 5; page++) {
            if (page > 1) source.PageBreak();
            int current = page;
            source.Paragraph(paragraph => paragraph.Text("Leaf " + current));
        }

        PdfImpositionResult result = PdfDocument.Load(source.ToBytes()).Pages.ImposeBooklet(
            new PdfBookletOptions(new PageSize(600, 400)));

        Assert.Equal(4, PdfInspector.Inspect(result.Bytes).PageCount);
        Assert.Equal(new[] { (1, 1, 1), (2, 2, 0), (3, 3, 1), (4, 4, 0), (5, 4, 1) },
            result.Placements.Select(static placement => (placement.SourcePageNumber, placement.SheetPageNumber, placement.Column)));
        PdfDocument imposed = result.ToDocument();
        Assert.Contains("Leaf 1", imposed.Reader.Text(PdfPageSelection.From(1)));
        Assert.Contains("Leaf 2", imposed.Reader.Text(PdfPageSelection.From(2)));
        Assert.Contains("Leaf 3", imposed.Reader.Text(PdfPageSelection.From(3)));
        Assert.Contains("Leaf 4", imposed.Reader.Text(PdfPageSelection.From(4)));
        Assert.Contains("Leaf 5", imposed.Reader.Text(PdfPageSelection.From(4)));
    }

    [Fact]
    public void RightToLeftBookletReversesBothCellsOnEachSheetSide() {
        PdfDocument source = PdfDocument.Create(new PdfOptions { PageSize = new PageSize(240, 180) });
        for (int page = 1; page <= 8; page++) {
            if (page > 1) source.PageBreak();
            int current = page;
            source.Paragraph(paragraph => paragraph.Text("Leaf " + current));
        }

        PdfImpositionResult result = PdfDocument.Load(source.ToBytes()).Pages.ImposeBooklet(
            new PdfBookletOptions(new PageSize(600, 400)) { RightToLeft = true });

        Assert.Equal(new[] { 1, 8, 7, 2, 3, 6, 5, 4 }, result.Placements.Select(static placement => placement.SourcePageNumber));
        Assert.Equal(new[] { 1, 1, 2, 2, 3, 3, 4, 4 }, result.Placements.Select(static placement => placement.SheetPageNumber));
    }

    [Fact]
    public void SignedSourceRequiresExplicitUnsignedDerivativeForEitherImposition() {
        byte[] source = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Signed sheet source")).ToBytes();
        PdfExternalSignaturePreparation preparation = PdfIncrementalUpdater.PrepareExternalSignature(
            source, new PdfExternalSignatureOptions { FieldName = "Approval", ReservedSignatureContentsBytes = 512 });
        byte[] signed = PdfIncrementalUpdater.ApplyExternalSignature(preparation, new byte[] { 0x30, 0x01, 0x00 });
        PdfDocument document = PdfDocument.Load(signed);
        var nupOptions = new PdfNUpOptions(new PageSize(600, 400), 2, 1) { AllowSourceFeatureLoss = true };
        var bookletOptions = new PdfBookletOptions(new PageSize(600, 400)) { AllowSourceFeatureLoss = true };

        Assert.Throws<NotSupportedException>(() => document.Pages.ImposeNUp(nupOptions));
        Assert.Throws<NotSupportedException>(() => document.Pages.ImposeBooklet(bookletOptions));

        nupOptions.SignaturePolicy = PdfImpositionSignaturePolicy.CreateUnsignedDerivative;
        bookletOptions.SignaturePolicy = PdfImpositionSignaturePolicy.CreateUnsignedDerivative;
        PdfImpositionResult nup = document.Pages.ImposeNUp(nupOptions);
        PdfImpositionResult booklet = document.Pages.ImposeBooklet(bookletOptions);
        Assert.Equal(1, nup.RemovedSignatureCount);
        Assert.Equal(1, booklet.RemovedSignatureCount);
        Assert.True(nup.SourceFeatureLoss.HasFlag(PdfImpositionSourceFeatureLoss.Signatures));
        Assert.Contains("Signed sheet source", nup.ToDocument().Reader.Text());
        Assert.Contains("Signed sheet source", booklet.ToDocument().Reader.Text());
        Assert.False(PdfInspector.Inspect(nup.Bytes).HasSignatures);
    }

    [Fact]
    public void LayeredSourceIsRejectedBeforeSheetCreation() {
        PdfDocument document = PdfDocument.Load(PdfOptionalContentSupport.BuildOptionalContentMetadataPdf());
        var options = new PdfNUpOptions(new PageSize(600, 400), 2, 1) { AllowSourceFeatureLoss = true };

        Assert.Throws<NotSupportedException>(() => document.Pages.ImposeNUp(options));
        Assert.Throws<NotSupportedException>(() => document.Pages.ImposeBooklet(
            new PdfBookletOptions(new PageSize(600, 400)) { AllowSourceFeatureLoss = true }));
    }

    [Fact]
    public void VisualPageImportStillHonorsUserPasswordExtractionPermissions() {
        var encryption = new PdfStandardEncryptionOptions("open") {
            OwnerPassword = "owner",
            AllowedPermissions = PdfStandardPermissions.Print
        };
        byte[] source = PdfDocument.Create(new PdfOptions().SetEncryption(encryption))
            .Paragraph(paragraph => paragraph.Text("Restricted source")).ToBytes();
        PdfDocument document = PdfDocument.Load(source, new PdfLoadOptions { Password = "open" });

        Assert.Throws<PdfPermissionDeniedException>(() => document.Pages.ImposeNUp(
            new PdfNUpOptions(new PageSize(600, 400), 2, 1)));
    }
}
