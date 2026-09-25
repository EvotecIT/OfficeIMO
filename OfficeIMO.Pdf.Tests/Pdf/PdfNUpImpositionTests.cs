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
    public void SelectedProductionBoxesRequireFeatureLossApproval() {
        byte[] source = System.Text.Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj", "<< /Type /Catalog /Pages 2 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /TrimBox [10 10 190 190] /BleedBox [5 5 195 195] /ArtBox [20 20 180 180] /Contents 4 0 R >>", "endobj",
            "4 0 obj", "<< /Length 0 >>", "stream", string.Empty, "endstream", "endobj",
            "trailer", "<< /Root 1 0 R /Size 5 >>", "%%EOF"
        }));
        PdfDocument document = PdfDocument.Load(source);
        var options = new PdfNUpOptions(new PageSize(600, 400), 2, 1);

        Assert.Throws<NotSupportedException>(() => document.Pages.ImposeNUp(options));
        options.AllowSourceFeatureLoss = true;
        Assert.True(document.Pages.ImposeNUp(options).SourceFeatureLoss.HasFlag(PdfImpositionSourceFeatureLoss.PageFeatures));
    }

    [Theory]
    [InlineData("/TrimBox null", "")]
    [InlineData("/TrimBox 5 0 R", "5 0 obj\nnull\nendobj")]
    [InlineData("/Metadata 5 0 R", "5 0 obj\n6 0 R\nendobj\n6 0 obj\nnull\nendobj")]
    [InlineData("/PieceInfo 5 0 R", "5 0 obj\n6 0 R\nendobj\n6 0 obj\nnull\nendobj")]
    public void NullPageEntriesDoNotRequireFeatureLossApproval(string pageEntry, string extraObject) {
        byte[] source = System.Text.Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.7", "1 0 obj", "<< /Type /Catalog /Pages 2 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] " + pageEntry + " /Contents 4 0 R >>", "endobj",
            "4 0 obj", "<< /Length 0 >>", "stream", string.Empty, "endstream", "endobj",
            extraObject, "trailer", "<< /Root 1 0 R /Size 7 >>", "%%EOF", string.Empty
        }));

        PdfImpositionResult result = PdfDocument.Load(source).Pages.ImposeNUp(
            new PdfNUpOptions(new PageSize(600, 400), 2, 1));

        Assert.Equal(PdfImpositionSourceFeatureLoss.None, result.SourceFeatureLoss);
    }

    [Theory]
    [InlineData("/Producer (Other tool)")]
    [InlineData("/Creator (Authoring tool)")]
    [InlineData("/CustomKey (Review value)")]
    public void RawInfoEntriesRequireMetadataLossApproval(string infoEntry) {
        byte[] source = System.Text.Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.7", "1 0 obj", "<< /Type /Catalog /Pages 2 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>", "endobj",
            "4 0 obj", "<< /Length 0 >>", "stream", string.Empty, "endstream", "endobj",
            "5 0 obj", "<< " + infoEntry + " >>", "endobj",
            "trailer", "<< /Root 1 0 R /Info 5 0 R /Size 6 >>", "%%EOF", string.Empty
        }));
        PdfDocument document = PdfDocument.Load(source);
        var options = new PdfNUpOptions(new PageSize(600, 400), 2, 1);

        Assert.Throws<NotSupportedException>(() => document.Pages.ImposeNUp(options));
        options.AllowSourceFeatureLoss = true;
        Assert.True(document.Pages.ImposeNUp(options).SourceFeatureLoss.HasFlag(PdfImpositionSourceFeatureLoss.DocumentMetadata));
    }

    [Fact]
    public void RepeatedPageContentCannotExceedAggregateImpositionBudget() {
        string content = "q\n%" + new string('x', 8192) + "\nQ\n";
        byte[] source = System.Text.Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj", "<< /Type /Catalog /Pages 2 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>", "endobj",
            "4 0 obj", "<< /Length " + content.Length + " >>", "stream", content, "endstream", "endobj",
            "trailer", "<< /Root 1 0 R /Size 5 >>", "%%EOF"
        }));
        var options = new PdfNUpOptions(new PageSize(600, 400), 2, 1) { MaxOutputBytes = 12000 };
        PdfDocument document = PdfDocument.Load(source);

        Assert.Throws<InvalidDataException>(() => document.Pages.ImposeNUp(options, PdfPageSelection.From(1, 1)));
    }

    [Fact]
    public void UnselectedInteractivePagesDoNotRequireFeatureLossApproval() {
        byte[] source = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Selected page"))
            .PageBreak()
            .TextAnnotation("Review note")
            .TextField("ReviewedBy", value: "Ada")
            .ToBytes();
        PdfDocument document = PdfDocument.Load(source);
        var layout = new PdfNUpOptions(new PageSize(600, 400), 2, 1);

        PdfImpositionResult nup = document.Pages.ImposeNUp(layout, PdfPageSelection.From(1));
        PdfImpositionResult booklet = document.Pages.ImposeBooklet(
            new PdfBookletOptions(new PageSize(600, 400)), PdfPageSelection.From(1));

        Assert.Equal(PdfImpositionSourceFeatureLoss.None, nup.SourceFeatureLoss);
        Assert.Equal(PdfImpositionSourceFeatureLoss.None, booklet.SourceFeatureLoss);
        Assert.Contains("Selected page", nup.ToDocument().Reader.Text());
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void DocumentLevelFormsRequireFeatureLossApproval(bool xfa) {
        byte[] source = System.Text.Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj", "<< /Type /Catalog /Pages 2 0 R /AcroForm 5 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>", "endobj",
            "4 0 obj", "<< /Length 0 >>", "stream", "", "endstream", "endobj",
            "5 0 obj", xfa ? "<< /Fields [] /XFA (packet) >>" : "<< /Fields [6 0 R] >>", "endobj",
            "6 0 obj", "<< /FT /Tx /T (InvisibleValue) /V (Ada) >>", "endobj",
            "trailer", "<< /Root 1 0 R /Size 7 >>", "%%EOF"
        }));
        PdfDocument document = PdfDocument.Load(source);
        var layout = new PdfNUpOptions(new PageSize(600, 400), 2, 1);

        Assert.Throws<NotSupportedException>(() => document.Pages.ImposeNUp(layout));
        layout.AllowSourceFeatureLoss = true;
        PdfImpositionResult result = document.Pages.ImposeNUp(layout);

        Assert.True(result.SourceFeatureLoss.HasFlag(PdfImpositionSourceFeatureLoss.Forms));
        Assert.False(PdfInspector.Inspect(result.Bytes).HasForms);
    }

    [Fact]
    public void UnreadableSelectedAnnotationStillRequiresFeatureLossApproval() {
        byte[] source = System.Text.Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj", "<< /Type /Catalog /Pages 2 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R /Annots [5 0 R] >>", "endobj",
            "4 0 obj", "<< /Length 0 >>", "stream", "", "endstream", "endobj",
            "5 0 obj", "<< /Type /Annot /Subtype /Text /Contents (Unreadable geometry) >>", "endobj",
            "trailer", "<< /Root 1 0 R /Size 6 >>", "%%EOF"
        }));
        PdfDocument document = PdfDocument.Load(source);
        var layout = new PdfNUpOptions(new PageSize(600, 400), 2, 1);

        Assert.Throws<NotSupportedException>(() => document.Pages.ImposeNUp(layout));
        layout.AllowSourceFeatureLoss = true;
        PdfImpositionResult result = document.Pages.ImposeNUp(layout);

        Assert.True(result.SourceFeatureLoss.HasFlag(PdfImpositionSourceFeatureLoss.Annotations));
    }

    [Fact]
    public void EmbeddedFileLossRequiresExplicitApproval() {
        byte[] source = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Sheet")).ToBytes();
        byte[] attached = PdfAttachmentEditor.Add(source,
            new PdfEmbeddedFile("payload.txt", System.Text.Encoding.UTF8.GetBytes("payload"))).ToBytes();
        PdfDocument document = PdfDocument.Load(attached);
        var layout = new PdfNUpOptions(new PageSize(600, 400), 2, 1);

        Assert.Throws<NotSupportedException>(() => document.Pages.ImposeNUp(layout));
        layout.AllowSourceFeatureLoss = true;
        PdfImpositionResult result = document.Pages.ImposeNUp(layout);

        Assert.True(result.SourceFeatureLoss.HasFlag(PdfImpositionSourceFeatureLoss.EmbeddedFiles));
        Assert.False(PdfInspector.Inspect(result.Bytes).HasEmbeddedFiles);
    }

    [Fact]
    public void OutputIntentLossRequiresExplicitApproval() {
        PdfDocument document = PdfDocument.Load(PdfDocument.Create(new PdfOptions().SetSrgbOutputIntent())
            .Paragraph(paragraph => paragraph.Text("Sheet")).ToBytes());
        var layout = new PdfNUpOptions(new PageSize(600, 400), 2, 1);

        Assert.Throws<NotSupportedException>(() => document.Pages.ImposeNUp(layout));
        layout.AllowSourceFeatureLoss = true;
        PdfImpositionResult result = document.Pages.ImposeNUp(layout);

        Assert.True(result.SourceFeatureLoss.HasFlag(PdfImpositionSourceFeatureLoss.OutputIntents));
        Assert.False(PdfInspector.Inspect(result.Bytes).HasOutputIntents);
    }

    [Fact]
    public void OutlineOnlySourceRequiresExplicitLossApproval() {
        PdfDocument document = PdfDocument.Load(PdfDocument.Create(new PdfOptions { CreateOutlineFromHeadings = true })
            .H1("Bookmark")
            .Paragraph(paragraph => paragraph.Text("Sheet"))
            .ToBytes());
        Assert.True(PdfInspector.Inspect(document.ToBytes()).HasOutlines);
        var layout = new PdfNUpOptions(new PageSize(600, 400), 2, 1);

        Assert.Throws<NotSupportedException>(() => document.Pages.ImposeNUp(layout));
        layout.AllowSourceFeatureLoss = true;
        PdfImpositionResult result = document.Pages.ImposeNUp(layout);

        Assert.True(result.SourceFeatureLoss.HasFlag(PdfImpositionSourceFeatureLoss.Outlines));
        Assert.False(PdfInspector.Inspect(result.Bytes).HasOutlines);
    }

    [Fact]
    public void DocumentAndPagePresentationLossIsReportedTogether() {
        byte[] source = System.Text.Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj", "<< /Type /Catalog /Pages 2 0 R /Outlines 5 0 R /PageLabels << /Nums [0 << /S /D /P (A-) >>] >> /Names << /Dests << /Names [(start) [3 0 R /Fit]] >> >> /ViewerPreferences << /HideToolbar true >> /Lang (en-US) >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Dur 2 /Contents 4 0 R >>", "endobj",
            "4 0 obj", "<< /Length 0 >>", "stream", string.Empty, "endstream", "endobj",
            "5 0 obj", "<< /Type /Outlines /First 6 0 R /Last 6 0 R /Count 1 >>", "endobj",
            "6 0 obj", "<< /Title (Start) /Parent 5 0 R /Dest [3 0 R /Fit] >>", "endobj",
            "7 0 obj", "<< /Title (Source title) >>", "endobj",
            "trailer", "<< /Root 1 0 R /Info 7 0 R /Size 8 >>", "%%EOF", string.Empty
        }));
        PdfDocument document = PdfDocument.Load(source);
        PdfDocumentInfo input = PdfInspector.Inspect(source);
        Assert.True(input.HasOutlines);
        Assert.True(input.HasPageLabels);
        Assert.True(input.HasNamedDestinations);
        Assert.True(input.HasViewerPreferences);
        Assert.Equal("Source title", input.Metadata.Title);
        Assert.Equal(2D, input.Pages[0].DurationSeconds);
        var layout = new PdfNUpOptions(new PageSize(600, 400), 2, 1);

        Assert.Throws<NotSupportedException>(() => document.Pages.ImposeNUp(layout));
        layout.AllowSourceFeatureLoss = true;
        PdfImpositionResult result = document.Pages.ImposeNUp(layout);

        Assert.True(result.SourceFeatureLoss.HasFlag(PdfImpositionSourceFeatureLoss.Outlines));
        Assert.True(result.SourceFeatureLoss.HasFlag(PdfImpositionSourceFeatureLoss.PageLabels));
        Assert.True(result.SourceFeatureLoss.HasFlag(PdfImpositionSourceFeatureLoss.NamedDestinations));
        Assert.True(result.SourceFeatureLoss.HasFlag(PdfImpositionSourceFeatureLoss.CatalogFeatures));
        Assert.True(result.SourceFeatureLoss.HasFlag(PdfImpositionSourceFeatureLoss.DocumentMetadata));
        Assert.True(result.SourceFeatureLoss.HasFlag(PdfImpositionSourceFeatureLoss.PageFeatures));
    }

    [Fact]
    public void EncryptedSourceRequiresExplicitApprovalBeforeCreatingUnencryptedSheets() {
        byte[] source = PdfDocument.Create(new PdfOptions().SetEncryption(new PdfStandardEncryptionOptions("reader") {
            OwnerPassword = "owner"
        })).Paragraph(paragraph => paragraph.Text("Confidential sheet")).ToBytes();
        PdfDocument document = PdfDocument.Load(source, new PdfLoadOptions { Password = "owner" });
        var layout = new PdfNUpOptions(new PageSize(600, 400), 2, 1);

        Assert.Throws<NotSupportedException>(() => document.Pages.ImposeNUp(layout));
        layout.AllowSourceFeatureLoss = true;
        PdfImpositionResult result = document.Pages.ImposeNUp(layout);

        Assert.True(result.SourceFeatureLoss.HasFlag(PdfImpositionSourceFeatureLoss.Encryption));
        Assert.False(PdfInspector.Inspect(result.Bytes).Security.HasEncryption);
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
        var nupOptions = new PdfNUpOptions(new PageSize(600, 400), 2, 1);
        var bookletOptions = new PdfBookletOptions(new PageSize(600, 400));

        Assert.Throws<NotSupportedException>(() => document.Pages.ImposeNUp(nupOptions));
        Assert.Throws<NotSupportedException>(() => document.Pages.ImposeBooklet(bookletOptions));

        nupOptions.SignaturePolicy = PdfImpositionSignaturePolicy.CreateUnsignedDerivative;
        bookletOptions.SignaturePolicy = PdfImpositionSignaturePolicy.CreateUnsignedDerivative;
        PdfImpositionResult nup = document.Pages.ImposeNUp(nupOptions);
        PdfImpositionResult booklet = document.Pages.ImposeBooklet(bookletOptions);
        Assert.Equal(1, nup.RemovedSignatureCount);
        Assert.Equal(1, booklet.RemovedSignatureCount);
        Assert.True(nup.SourceFeatureLoss.HasFlag(PdfImpositionSourceFeatureLoss.Signatures));
        Assert.False(nup.SourceFeatureLoss.HasFlag(PdfImpositionSourceFeatureLoss.Annotations));
        Assert.False(nup.SourceFeatureLoss.HasFlag(PdfImpositionSourceFeatureLoss.Forms));
        Assert.Contains("Signed sheet source", nup.ToDocument().Reader.Text());
        Assert.Contains("Signed sheet source", booklet.ToDocument().Reader.Text());
        Assert.False(PdfInspector.Inspect(nup.Bytes).HasSignatures);
    }

    [Fact]
    public void UnsignedDerivativeStillRequiresApprovalForOtherInteractiveFeatures() {
        byte[] source = PdfDocument.Create().TextField("Name").TextAnnotation("Review note").ToBytes();
        PdfExternalSignaturePreparation preparation = PdfIncrementalUpdater.PrepareExternalSignature(
            source, new PdfExternalSignatureOptions { FieldName = "Approval", ReservedSignatureContentsBytes = 512 });
        byte[] signed = PdfIncrementalUpdater.ApplyExternalSignature(preparation, new byte[] { 0x30, 0x01, 0x00 });
        PdfDocument document = PdfDocument.Load(signed);
        var options = new PdfNUpOptions(new PageSize(600, 400), 2, 1) {
            SignaturePolicy = PdfImpositionSignaturePolicy.CreateUnsignedDerivative
        };

        Assert.Throws<NotSupportedException>(() => document.Pages.ImposeNUp(options));
        options.AllowSourceFeatureLoss = true;
        PdfImpositionResult result = document.Pages.ImposeNUp(options);

        Assert.True(result.SourceFeatureLoss.HasFlag(PdfImpositionSourceFeatureLoss.Signatures));
        Assert.True(result.SourceFeatureLoss.HasFlag(PdfImpositionSourceFeatureLoss.Forms));
        Assert.True(result.SourceFeatureLoss.HasFlag(PdfImpositionSourceFeatureLoss.Annotations));
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
