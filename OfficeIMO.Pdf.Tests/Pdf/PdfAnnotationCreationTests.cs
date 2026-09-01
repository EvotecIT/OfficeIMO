using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfAnnotationCreationTests {
    [Fact]
    public void MoveAnnotation_TranslatesRectangleAndLineGeometryTogether() {
        byte[] source = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Move annotation")).ToBytes();
        PdfAnnotationEditResult added = PdfDocument.Load(source).Annotations.Add(new PdfAnnotationCreateOptions {
            Subtype = "Line",
            Rectangle = new[] { 40D, 50D, 180D, 100D },
            Line = new[] { 45D, 55D, 175D, 95D },
            GenerateAppearance = true
        });
        PdfAnnotation annotation = Assert.Single(added.ToDocument().Inspect().GetAnnotationsBySubtype("Line"));

        PdfAnnotationEditResult moved = added.ToDocument().Annotations.Move(annotation.ObjectNumber!.Value, 25D, -10D);
        PdfAnnotation result = Assert.Single(moved.ToDocument().Inspect().GetAnnotationsBySubtype("Line"));

        Assert.Equal(65D, result.X1, 3);
        Assert.Equal(40D, result.Y1, 3);
        Assert.Equal(205D, result.X2, 3);
        Assert.Equal(90D, result.Y2, 3);
        Assert.Equal(new[] { 70D, 45D, 200D, 85D }, result.LineCoordinates);
        Assert.True(result.HasNormalAppearance);
    }

    [Fact]
    public void ResizeAnnotation_ScalesRectangleAndMarkupGeometryTogether() {
        byte[] source = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Resize annotation")).ToBytes();
        PdfAnnotationEditResult added = PdfDocument.Load(source).Annotations.Add(new PdfAnnotationCreateOptions {
            Subtype = "Highlight",
            Rectangle = new[] { 40D, 50D, 140D, 70D },
            QuadPoints = new[] { 40D, 70D, 140D, 70D, 40D, 50D, 140D, 50D },
            GenerateAppearance = true
        });
        PdfAnnotation annotation = Assert.Single(added.ToDocument().Inspect().GetAnnotationsBySubtype("Highlight"));

        PdfAnnotationEditResult resized = added.ToDocument().Annotations.Resize(
            annotation.ObjectNumber!.Value,
            new PdfPageRectangle(20D, 30D, 220D, 70D));
        PdfAnnotation result = Assert.Single(resized.ToDocument().Inspect().GetAnnotationsBySubtype("Highlight"));

        Assert.Equal(200D, result.Width, 3);
        Assert.Equal(40D, result.Height, 3);
        Assert.Equal(new[] { 20D, 70D, 220D, 70D, 20D, 30D, 220D, 30D }, result.QuadPoints);
        Assert.True(result.HasNormalAppearance);
    }

    [Fact]
    public void ResizeAnnotation_ScalesFreeTextRectangleDifferences() {
        byte[] source = BuildFreeTextRectangleDifferencePdf();
        PdfAnnotation annotation = Assert.Single(PdfInspector.Inspect(source).GetAnnotationsBySubtype("FreeText"));

        PdfAnnotationEditResult resized = PdfDocument.Load(source).Annotations.Resize(
            annotation.ObjectNumber!.Value,
            new PdfPageRectangle(10D, 20D, 210D, 100D));
        PdfAnnotation result = Assert.Single(resized.ToDocument().Inspect().GetAnnotationsBySubtype("FreeText"));

        Assert.Equal(200D, result.Width, 3);
        Assert.Equal(80D, result.Height, 3);
        Assert.Equal(new[] { 20D, 8D, 40D, 16D }, result.RectangleDifferences);
        Assert.True(result.HasNormalAppearance);
    }

    [Fact]
    public void AddAnnotation_CreatesUriLinkWithReadback() {
        byte[] source = PdfDocument.Create().Paragraph(p => p.Text("Existing page")).ToBytes();

        PdfAnnotationEditResult result = PdfDocument.Load(source).Annotations.Add(new PdfAnnotationCreateOptions {
            Subtype = "Link",
            Rectangle = new[] { 40D, 50D, 180D, 80D },
            Contents = "OfficeIMO",
            LinkUri = "https://officeimo.com"
        });

        PdfLinkAnnotation link = Assert.Single(PdfInspector.Inspect(result.Bytes).GetLinkAnnotationsByUri("https://officeimo.com"));
        Assert.Equal(1, link.PageNumber);
        Assert.Equal("OfficeIMO", link.Contents);
        Assert.Equal(40D, link.X1);
        Assert.Equal(80D, link.Y2);
    }

    [Theory]
    [InlineData("Title")]
    [InlineData("IconName")]
    [InlineData("InReplyToObjectNumber")]
    [InlineData("ReplyType")]
    [InlineData("ReviewState")]
    [InlineData("Subject")]
    [InlineData("Intent")]
    [InlineData("CreatePopup")]
    [InlineData("PopupRectangle")]
    [InlineData("PopupOpen")]
    public void AddLinkAnnotation_RejectsMarkupOnlyOptions(string optionName) {
        byte[] source = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Link validation")).ToBytes();
        var options = new PdfAnnotationCreateOptions {
            Subtype = "Link",
            Rectangle = new[] { 40D, 50D, 180D, 80D },
            LinkUri = "https://officeimo.com"
        };
        switch (optionName) {
            case "Title": options.Title = "Author"; break;
            case "IconName": options.IconName = "Comment"; break;
            case "InReplyToObjectNumber": options.InReplyToObjectNumber = 1; break;
            case "ReplyType": options.ReplyType = "R"; break;
            case "ReviewState": options.ReviewState = PdfAnnotationReviewState.Accepted; break;
            case "Subject": options.Subject = "Review"; break;
            case "Intent": options.Intent = "Link"; break;
            case "CreatePopup": options.CreatePopup = true; break;
            case "PopupRectangle": options.PopupRectangle = new[] { 10D, 10D, 30D, 30D }; break;
            case "PopupOpen": options.PopupOpen = true; break;
            default: throw new ArgumentOutOfRangeException(nameof(optionName));
        }

        ArgumentException exception = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Load(source).Annotations.Add(options));
        Assert.Contains("markup-only", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void LinkAnnotation_MoveAndResizeUpdateRectangleWithoutAppearanceRegeneration() {
        byte[] source = PdfDocument.Create().Paragraph(p => p.Text("Existing page")).ToBytes();
        PdfAnnotationEditResult added = PdfDocument.Load(source).Annotations.Add(new PdfAnnotationCreateOptions {
            Subtype = "Link",
            Rectangle = new[] { 40D, 50D, 180D, 80D },
            LinkUri = "https://officeimo.com"
        });
        PdfAnnotation annotation = Assert.Single(added.ToDocument().Inspect().GetAnnotationsBySubtype("Link"));

        PdfAnnotationEditResult moved = added.ToDocument().Annotations.Move(annotation.ObjectNumber!.Value, 10D, 15D);
        PdfAnnotation movedAnnotation = Assert.Single(moved.ToDocument().Inspect().GetAnnotationsBySubtype("Link"));
        PdfAnnotationEditResult resized = moved.ToDocument().Annotations.Resize(
            movedAnnotation.ObjectNumber!.Value,
            new PdfPageRectangle(25D, 35D, 225D, 75D));
        PdfLinkAnnotation link = Assert.Single(
            PdfInspector.Inspect(resized.Bytes).GetLinkAnnotationsByUri("https://officeimo.com"));

        Assert.Equal(25D, link.X1, 3);
        Assert.Equal(35D, link.Y1, 3);
        Assert.Equal(225D, link.X2, 3);
        Assert.Equal(75D, link.Y2, 3);
    }

    [Fact]
    public void LinkAnnotation_MoveAndResizePreserveCustomNormalAppearance() {
        byte[] source = BuildLinkAppearanceAnnotationPdf();
        PdfAnnotation annotation = Assert.Single(PdfInspector.Inspect(source).GetAnnotationsBySubtype("Link"));

        PdfAnnotationEditResult moved = PdfDocument.Load(source).Annotations.Move(annotation.ObjectNumber!.Value, 10D, 15D);
        PdfAnnotation movedAnnotation = Assert.Single(moved.ToDocument().Inspect().GetAnnotationsBySubtype("Link"));
        PdfAnnotationEditResult resized = moved.ToDocument().Annotations.Resize(
            movedAnnotation.ObjectNumber!.Value,
            new PdfPageRectangle(25D, 35D, 225D, 75D));
        PdfAnnotation result = Assert.Single(resized.ToDocument().Inspect().GetAnnotationsBySubtype("Link"));

        Assert.True(result.HasNormalAppearance);
        Assert.Contains("1 0 0 RG", Encoding.ASCII.GetString(resized.Bytes), StringComparison.Ordinal);
    }

    [Fact]
    public void VisualAnnotation_MoveAndResizePreserveAuthoredNormalAppearance() {
        byte[] source = BuildSquareAppearanceAnnotationPdf();
        PdfAnnotation annotation = Assert.Single(PdfInspector.Inspect(source).GetAnnotationsBySubtype("Square"));

        PdfAnnotationEditResult moved = PdfDocument.Load(source).Annotations.Move(annotation.ObjectNumber!.Value, 10D, 15D);
        PdfAnnotation movedAnnotation = Assert.Single(moved.ToDocument().Inspect().GetAnnotationsBySubtype("Square"));
        PdfAnnotationEditResult resized = moved.ToDocument().Annotations.Resize(
            movedAnnotation.ObjectNumber!.Value,
            new PdfPageRectangle(25D, 35D, 145D, 155D));

        string raw = Encoding.ASCII.GetString(resized.Bytes);
        PdfAnnotation result = Assert.Single(resized.ToDocument().Inspect().GetAnnotationsBySubtype("Square"));
        Assert.True(result.HasNormalAppearance);
        Assert.Contains("0.125 0.25 0.5 rg", raw, StringComparison.Ordinal);
        var (objects, _) = PdfSyntax.ParseObjects(resized.Bytes);
        PdfStream appearance = Assert.Single(
            objects.Values.Select(static item => item.Value).OfType<PdfStream>(),
            static stream => Encoding.ASCII.GetString(stream.Data).Contains("0.125 0.25 0.5 rg", StringComparison.Ordinal));
        PdfArray boundingBox = Assert.IsType<PdfArray>(appearance.Dictionary.Items["BBox"]);
        Assert.Equal(new[] { 0D, 0D, 60D, 60D }, boundingBox.Items.Cast<PdfNumber>().Select(static number => number.Value));
    }

    [Fact]
    public void VisualAnnotation_MoveRegeneratesAppearanceWhenSelectedStateIsMissing() {
        byte[] source = BuildSquareAppearanceStateMismatchPdf();
        PdfAnnotation annotation = Assert.Single(PdfInspector.Inspect(source).GetAnnotationsBySubtype("Square"));
        Assert.False(annotation.HasNormalAppearance);

        PdfAnnotationEditResult moved = PdfDocument.Load(source).Annotations.Move(
            annotation.ObjectNumber!.Value,
            10D,
            15D);

        PdfAnnotation result = Assert.Single(moved.ToDocument().Inspect().GetAnnotationsBySubtype("Square"));
        Assert.True(result.HasNormalAppearance);
        var (objects, _) = PdfSyntax.ParseObjects(moved.Bytes);
        PdfDictionary annotationDictionary = Assert.IsType<PdfDictionary>(objects[result.ObjectNumber!.Value].Value);
        PdfDictionary appearanceDictionary = Assert.IsType<PdfDictionary>(annotationDictionary.Items["AP"]);
        PdfReference normalAppearance = Assert.IsType<PdfReference>(appearanceDictionary.Items["N"]);
        PdfStream regeneratedAppearance = Assert.IsType<PdfStream>(objects[normalAppearance.ObjectNumber].Value);
        Assert.DoesNotContain(
            "0.125 0.25 0.5 rg",
            Encoding.ASCII.GetString(regeneratedAppearance.Data),
            StringComparison.Ordinal);
    }

    [Fact]
    public void AddAnnotation_CreatesLineGeometryAppearanceAndPopupOnExistingPage() {
        byte[] source = PdfDocument.Create().Paragraph(p => p.Text("Existing page")).ToBytes();

        PdfAnnotationEditResult result = PdfDocument.Load(source).Annotations.Add(new PdfAnnotationCreateOptions {
            Subtype = "Line",
            Rectangle = new[] { 40D, 50D, 180D, 100D },
            Line = new[] { 40D, 50D, 180D, 100D },
            LineStartEnding = "OpenArrow",
            LineEndEnding = "ClosedArrow",
            Contents = "Review line",
            Color = new[] { 0.8D, 0.1D, 0.1D },
            CreatePopup = true,
            PopupOpen = true
        });
        PdfDocumentInfo info = PdfInspector.Inspect(result.Bytes);
        PdfAnnotation line = Assert.Single(info.GetAnnotationsBySubtype("Line"));

        Assert.Equal(new[] { 40D, 50D, 180D, 100D }, line.LineCoordinates);
        Assert.Equal("OpenArrow", line.LineStartEnding);
        Assert.Equal("ClosedArrow", line.LineEndEnding);
        Assert.True(line.HasNormalAppearance);
        Assert.Single(info.GetAnnotationsBySubtype("Popup"));
        Assert.Contains("/Open true", Encoding.ASCII.GetString(result.Bytes), StringComparison.Ordinal);
        Assert.NotNull(result.RewritePreservationReport);
    }

    [Fact]
    public void AddAnnotation_CreatesReplyRelationship() {
        byte[] source = PdfDocument.Create().TextAnnotation("Parent").Paragraph(p => p.Text("Existing page")).ToBytes();
        int parentObject = Assert.Single(PdfInspector.Inspect(source).GetAnnotationsBySubtype("Text")).ObjectNumber!.Value;

        PdfAnnotationEditResult result = PdfAnnotationEditor.AddAnnotation(source, new PdfAnnotationCreateOptions {
            Subtype = "Text",
            Rectangle = new[] { 70D, 70D, 90D, 90D },
            Contents = "Reply",
            InReplyToObjectNumber = parentObject,
            ReplyType = "R",
            IconName = "Comment"
        });

        string raw = Encoding.ASCII.GetString(result.Bytes);
        Assert.Equal(2, PdfInspector.Inspect(result.Bytes).GetAnnotationsBySubtype("Text").Count);
        Assert.Contains("/IRT ", raw, StringComparison.Ordinal);
        Assert.Contains("/RT /R", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void AddAnnotation_PreservesPageAndReplyTargetGenerations() {
        byte[] source = BuildNonZeroGenerationAnnotationPdf();
        PdfExternalSignaturePreparation preparation = PdfIncrementalUpdater.PrepareExternalSignature(source, new PdfExternalSignatureOptions {
            Profile = PdfSignatureProfile.Certification,
            CertificationPermission = PdfCertificationPermissionLevel.FormFillingAnnotationsAndSignatures,
            FieldName = "GenerationCertification",
            ReservedSignatureContentsBytes = 512
        });
        byte[] signed = PdfIncrementalUpdater.ApplyExternalSignature(preparation, new byte[] { 0x30, 0x01, 0x00 });

        PdfAnnotationEditResult result = PdfAnnotationEditor.AddAnnotation(signed, new PdfAnnotationCreateOptions {
            Subtype = "Text",
            Rectangle = new[] { 70D, 70D, 90D, 90D },
            Contents = "Generation-aware reply",
            InReplyToObjectNumber = 6,
            CreatePopup = true
        });

        string raw = Encoding.ASCII.GetString(result.Bytes);
        PdfDocumentInfo info = PdfInspector.Inspect(result.Bytes);
        Assert.Equal(PdfMutationExecutionMode.AppendOnly, result.MutationPlan.ExecutionMode);
        Assert.Equal(2, info.GetAnnotationsBySubtype("Text").Count);
        Assert.Single(info.GetAnnotationsBySubtype("Popup"));
        Assert.Contains("/P 3 1 R", raw, StringComparison.Ordinal);
        Assert.Contains("/IRT 6 2 R", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void AddAnnotation_UsesAppendOnlyRevisionWhenCertificationAllowsAnnotations() {
        byte[] source = PdfDocument.Create().Paragraph(p => p.Text("Certified page")).ToBytes();
        PdfExternalSignaturePreparation preparation = PdfIncrementalUpdater.PrepareExternalSignature(source, new PdfExternalSignatureOptions {
            Profile = PdfSignatureProfile.Certification,
            CertificationPermission = PdfCertificationPermissionLevel.FormFillingAnnotationsAndSignatures,
            FieldName = "Certification",
            ReservedSignatureContentsBytes = 512
        });
        byte[] signed = PdfIncrementalUpdater.ApplyExternalSignature(preparation, new byte[] { 0x30, 0x01, 0x00 });

        PdfAnnotationEditResult result = PdfAnnotationEditor.AddAnnotation(signed, new PdfAnnotationCreateOptions {
            Subtype = "Text",
            Contents = "Append-only review",
            IconName = "Note"
        });

        Assert.Equal(PdfMutationExecutionMode.AppendOnly, result.MutationPlan.ExecutionMode);
        Assert.True(result.SignatureMutationReport!.IsPreservedAppendOnlyMutation);
        Assert.True(result.Bytes.AsSpan(0, signed.Length).SequenceEqual(signed));
        Assert.Equal("Append-only review", Assert.Single(PdfInspector.Inspect(result.Bytes).GetAnnotationsBySubtype("Text")).Contents);
    }

    [Fact]
    public void AppendOnlyAnnotationProofAllowsItsOwnedRevisionBeyondTheSourceBudget() {
        byte[] source = PdfDocument.Create().Paragraph(p => p.Text("Certified page")).ToBytes();
        PdfExternalSignaturePreparation preparation = PdfIncrementalUpdater.PrepareExternalSignature(source, new PdfExternalSignatureOptions {
            Profile = PdfSignatureProfile.Certification,
            CertificationPermission = PdfCertificationPermissionLevel.FormFillingAnnotationsAndSignatures,
            FieldName = "Certification",
            ReservedSignatureContentsBytes = 512
        });
        byte[] signed = PdfIncrementalUpdater.ApplyExternalSignature(preparation, new byte[] { 0x30, 0x01, 0x00 });
        var readOptions = new PdfLoadOptions {
            Limits = new PdfReadLimits { MaxInputBytes = signed.Length }
        };

        PdfAnnotationEditResult result = PdfAnnotationEditor.AddAnnotation(
            signed,
            new PdfAnnotationCreateOptions {
                Subtype = "Text",
                Contents = "Append-only proof"
            },
            readOptions);

        Assert.True(result.Bytes.LongLength > readOptions.Limits.MaxInputBytes);
        Assert.True(result.SignatureMutationReport!.After.ObjectGraphParsed);
        Assert.True(result.SignatureMutationReport.IsPreservedAppendOnlyMutation);
    }

    [Fact]
    public void FluentAnnotationEdits_UseStoredOwnerCredentialsForEncryptedPdf() {
        byte[] source = PdfDocument.Create(new PdfOptions().SetEncryption("open", "owner"))
            .Paragraph(paragraph => paragraph.Text("Encrypted annotations"))
            .ToBytes();
        var readOptions = new PdfLoadOptions { Password = "owner" };

        PdfAnnotationEditResult added = PdfDocument.Load(source, readOptions).Annotations.Add(new PdfAnnotationCreateOptions {
            Subtype = "Text",
            Contents = "Encrypted note"
        });
        PdfAnnotation annotation = Assert.Single(PdfInspector.Inspect(added.Bytes, readOptions).GetAnnotationsBySubtype("Text"));
        PdfAnnotationEditResult updated = added.ToDocument().Annotations.Update(
            annotation.ObjectNumber!.Value,
            new PdfAnnotationUpdateOptions {
                Contents = "Updated encrypted note",
                AllowResidualDataInAppendOnly = true
            });
        PdfAnnotationEditResult removed = updated.ToDocument().Annotations.Remove(
            new PdfAnnotationRemovalOptions {
                ObjectNumber = annotation.ObjectNumber,
                AllowResidualDataInAppendOnly = true
            });

        Assert.Equal(PdfMutationExecutionMode.AppendOnly, added.MutationPlan.ExecutionMode);
        Assert.Equal("Updated encrypted note", Assert.Single(PdfInspector.Inspect(updated.Bytes, readOptions).GetAnnotationsBySubtype("Text")).Contents);
        Assert.Empty(PdfInspector.Inspect(removed.Bytes, readOptions).GetAnnotationsBySubtype("Text"));
        Assert.True(removed.Bytes.AsSpan(0, updated.Bytes.Length).SequenceEqual(updated.Bytes));
    }

    [Fact]
    public void StaticAnnotationResult_RetainsExplicitOwnerCredentialsForToDocument() {
        byte[] source = PdfDocument.Create(new PdfOptions().SetEncryption("open", "owner"))
            .Paragraph(paragraph => paragraph.Text("Static encrypted annotation"))
            .ToBytes();
        var readOptions = new PdfLoadOptions { Password = "owner" };

        PdfAnnotationEditResult result = PdfAnnotationEditor.AddAnnotation(
            source,
            new PdfAnnotationCreateOptions {
                Subtype = "Text",
                Contents = "Static encrypted note"
            },
            readOptions);

        PdfDocument edited = result.ToDocument();
        PdfAnnotation annotation = Assert.Single(edited.Inspect().GetAnnotationsBySubtype("Text"));

        Assert.Equal("Static encrypted note", annotation.Contents);
        Assert.True(edited.Inspect().Security.HasOwnerAuthorization);
    }

    [Fact]
    public void FluentAnnotationFlatten_UsesStoredOwnerCredentialsDuringPreflight() {
        byte[] source = PdfDocument.Create(new PdfOptions().SetEncryption("open", "owner"))
            .FreeTextAnnotation("Encrypted annotation", 120, 30)
            .Paragraph(paragraph => paragraph.Text("Encrypted flatten"))
            .ToBytes();

        PdfMutationBlockedException exception = Assert.Throws<PdfMutationBlockedException>(() =>
            PdfDocument.Load(source, new PdfLoadOptions { Password = "owner" }).Annotations.Flatten());

        Assert.True(exception.Plan.Preflight.CanRead);
        Assert.True(exception.Plan.Preflight.Probe.Security.HasOwnerAuthorization);
        Assert.Contains("FullRewrite.Encryption", exception.Plan.BlockerCodes);
    }

    [Fact]
    public void FlattenAnnotations_FlattensOnlySelectedObjectThroughFluentSurface() {
        byte[] source = PdfDocument.Create()
            .FreeTextAnnotation("Flatten me", 120, 30)
            .HighlightAnnotation("Keep me", 120, 14)
            .Paragraph(p => p.Text("Existing page"))
            .ToBytes();
        int freeTextObject = Assert.Single(PdfInspector.Inspect(source).GetAnnotationsBySubtype("FreeText")).ObjectNumber!.Value;

        PdfAnnotationEditResult result = PdfDocument.Load(source).Annotations.Flatten(new PdfAnnotationFlattenOptions { ObjectNumber = freeTextObject });
        PdfDocumentInfo info = PdfInspector.Inspect(result.Bytes);

        Assert.Equal(1, result.AffectedAnnotationCount);
        Assert.Empty(info.GetAnnotationsBySubtype("FreeText"));
        Assert.Single(info.GetAnnotationsBySubtype("Highlight"));
        Assert.NotNull(result.RewritePreservationReport);
    }

    [Fact]
    public void FlattenAnnotations_ReservesGeneratedContentBudgets() {
        byte[] source = BuildEmptyAppearanceAnnotationPdf();
        var readOptions = new PdfLoadOptions {
            Limits = new PdfReadLimits {
                MaxInputBytes = source.LongLength,
                MaxRawStreamBytes = 1,
                MaxDecodedStreamBytes = 1,
                MaxTotalDecodedStreamBytes = 1,
                MaxPageContentBytes = 1,
                MaxRetainedContentBytes = 1,
                MaxContentOperations = 1,
                MaxContentOperands = 1
            }
        };

        PdfAnnotationEditResult result = PdfAnnotationEditor.FlattenAnnotations(source, options: null, readOptions);

        Assert.Equal(1, result.AffectedAnnotationCount);
        Assert.Empty(result.ToDocument().Reader.Annotations());
    }

    [Fact]
    public void UpdateAnnotation_ReservesTheSerializedPrimaryAnnotation() {
        byte[] source = PdfDocument.Create()
            .TextAnnotation("Short")
            .Paragraph(paragraph => paragraph.Text("Annotation update"))
            .ToBytes();
        int objectNumber = Assert.Single(PdfInspector.Inspect(source).GetAnnotationsBySubtype("Text")).ObjectNumber!.Value;
        var readOptions = new PdfLoadOptions {
            Limits = new PdfReadLimits { MaxObjectCharacters = 512, MaxTokensPerObject = 512 }
        };

        PdfAnnotationEditResult result = PdfAnnotationEditor.UpdateAnnotation(
            source,
            objectNumber,
            new PdfAnnotationUpdateOptions { Contents = new string('x', 4096) },
            readOptions);

        Assert.Equal(4096, Assert.Single(result.ToDocument().Reader.Annotations()).Contents!.Length);
    }

    private static byte[] BuildFreeTextRectangleDifferencePdf() => Encoding.ASCII.GetBytes(
        "%PDF-1.7\n" +
        "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n" +
        "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] >>\nendobj\n" +
        "3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 200] /Contents 4 0 R /Annots [5 0 R] >>\nendobj\n" +
        "4 0 obj\n<< /Length 0 >>\nstream\n\nendstream\nendobj\n" +
        "5 0 obj\n<< /Type /Annot /Subtype /FreeText /Rect [20 40 120 80] /RD [10 4 20 8] /Contents (Resize me) /DA (/Helvetica 10 Tf 0 g) >>\nendobj\n" +
        "trailer\n<< /Root 1 0 R /Size 6 >>\nstartxref\n0\n%%EOF\n");

    private static byte[] BuildLinkAppearanceAnnotationPdf() => Encoding.ASCII.GetBytes(
        "%PDF-1.7\n" +
        "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n" +
        "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] >>\nendobj\n" +
        "3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R /Annots [5 0 R] >>\nendobj\n" +
        "4 0 obj\n<< /Length 0 >>\nstream\n\nendstream\nendobj\n" +
        "5 0 obj\n<< /Type /Annot /Subtype /Link /Rect [20 20 80 40] /A << /S /URI /URI (https://officeimo.com) >> /AP << /N 6 0 R >> >>\nendobj\n" +
        "6 0 obj\n<< /Type /XObject /Subtype /Form /BBox [0 0 60 20] /Length 8 >>\nstream\n1 0 0 RG\nendstream\nendobj\n" +
        "trailer\n<< /Root 1 0 R /Size 7 >>\nstartxref\n0\n%%EOF\n");

    private static byte[] BuildSquareAppearanceAnnotationPdf() => Encoding.ASCII.GetBytes(
        "%PDF-1.7\n" +
        "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n" +
        "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] >>\nendobj\n" +
        "3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R /Annots [5 0 R] >>\nendobj\n" +
        "4 0 obj\n<< /Length 0 >>\nstream\n\nendstream\nendobj\n" +
        "5 0 obj\n<< /Type /Annot /Subtype /Square /Rect [20 20 80 80] /AP << /N 6 0 R >> >>\nendobj\n" +
        "6 0 obj\n<< /Type /XObject /Subtype /Form /BBox [0 0 60 60] /Length 33 >>\nstream\n0.125 0.25 0.5 rg 0 0 60 60 re f\nendstream\nendobj\n" +
        "trailer\n<< /Root 1 0 R /Size 7 >>\nstartxref\n0\n%%EOF\n");

    private static byte[] BuildSquareAppearanceStateMismatchPdf() => Encoding.ASCII.GetBytes(
        "%PDF-1.7\n" +
        "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n" +
        "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] >>\nendobj\n" +
        "3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R /Annots [5 0 R] >>\nendobj\n" +
        "4 0 obj\n<< /Length 0 >>\nstream\n\nendstream\nendobj\n" +
        "5 0 obj\n<< /Type /Annot /Subtype /Square /Rect [20 20 80 80] /AP << /N << /On 6 0 R >> >> /AS /Missing >>\nendobj\n" +
        "6 0 obj\n<< /Type /XObject /Subtype /Form /BBox [0 0 60 60] /Length 33 >>\nstream\n0.125 0.25 0.5 rg 0 0 60 60 re f\nendstream\nendobj\n" +
        "trailer\n<< /Root 1 0 R /Size 7 >>\nstartxref\n0\n%%EOF\n");

    private static byte[] BuildEmptyAppearanceAnnotationPdf() => Encoding.ASCII.GetBytes(
        "%PDF-1.7\n" +
        "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n" +
        "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] >>\nendobj\n" +
        "3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R /Annots [5 0 R] >>\nendobj\n" +
        "4 0 obj\n<< /Length 0 >>\nstream\n\nendstream\nendobj\n" +
        "5 0 obj\n<< /Type /Annot /Subtype /Square /Rect [20 20 80 80] /AP << /N 6 0 R >> >>\nendobj\n" +
        "6 0 obj\n<< /Type /XObject /Subtype /Form /BBox [0 0 60 60] /Length 0 >>\nstream\n\nendstream\nendobj\n" +
        "trailer\n<< /Root 1 0 R /Size 7 >>\nstartxref\n0\n%%EOF\n");

    private static byte[] BuildNonZeroGenerationAnnotationPdf() {
        var objects = new[] {
            (Number: 1, Generation: 0, Body: "<< /Type /Catalog /Pages 2 0 R >>"),
            (Number: 2, Generation: 0, Body: "<< /Type /Pages /Count 1 /Kids [3 1 R] >>"),
            (Number: 3, Generation: 1, Body: "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R /Annots [6 2 R] >>"),
            (Number: 4, Generation: 0, Body: "<< /Length 0 >>\nstream\n\nendstream"),
            (Number: 6, Generation: 2, Body: "<< /Type /Annot /Subtype /Text /Rect [20 20 40 40] /Contents (Parent) >>")
        };
        var builder = new StringBuilder("%PDF-1.7\n");
        var offsets = new Dictionary<int, (int Offset, int Generation)>();
        foreach ((int number, int generation, string body) in objects) {
            offsets[number] = (Encoding.ASCII.GetByteCount(builder.ToString()), generation);
            builder.Append(number).Append(' ').Append(generation).Append(" obj\n")
                .Append(body).Append("\nendobj\n");
        }

        int xrefOffset = Encoding.ASCII.GetByteCount(builder.ToString());
        builder.Append("xref\n0 7\n0000000000 65535 f \n");
        for (int number = 1; number < 7; number++) {
            if (offsets.TryGetValue(number, out (int Offset, int Generation) entry)) {
                builder.Append(entry.Offset.ToString("D10", System.Globalization.CultureInfo.InvariantCulture))
                    .Append(' ')
                    .Append(entry.Generation.ToString("D5", System.Globalization.CultureInfo.InvariantCulture))
                    .Append(" n \n");
            } else {
                builder.Append("0000000000 00000 f \n");
            }
        }

        builder.Append("trailer\n<< /Root 1 0 R /Size 7 >>\nstartxref\n")
            .Append(xrefOffset)
            .Append("\n%%EOF\n");
        return Encoding.ASCII.GetBytes(builder.ToString());
    }
}
