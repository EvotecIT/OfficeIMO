using System.Globalization;
using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfImageEditorTests {
    [Fact]
    public void AddFindAndRemoveAffectOnlyTheSelectedPlacement() {
        byte[] source = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Keep this text"))
            .ToBytes();
        byte[] annotated = PdfDocument.Open(source).Annotations.Add(new PdfAnnotationCreateOptions {
            PageNumber = 1,
            Subtype = "Square",
            Rectangle = new[] { 35D, 75D, 75D, 105D }
        }).Bytes;
        PdfDocument first = PdfDocument.Open(annotated).Images.Add(
            new PdfPageRegion(1, 40D, 80D, 30D, 20D),
            PdfPngTestImages.CreateRgbPng(255, 0, 0)).Document;
        PdfDocument second = first.Images.Add(
            new PdfPageRegion(1, 120D, 80D, 30D, 20D),
            PdfPngTestImages.CreateRgbPng(0, 255, 0)).Document;

        PdfImagePlacement selected = Assert.Single(second.Images.Find(new PdfPageRegion(1, 35D, 75D, 40D, 30D)));
        PdfImageEditResult removed = second.Images.Remove(selected);
        PdfImagePlacement remaining = Assert.Single(removed.Document.Images.Placements());

        Assert.Equal(1, removed.AffectedCount);
        Assert.InRange(remaining.X, 119.99D, 120.01D);
        Assert.Contains("Keep this text", removed.Document.Read.Text(), StringComparison.Ordinal);
        Assert.Single(removed.Document.Read.AnnotationsBySubtype("Square"));
    }

    [Fact]
    public void ReplacePreservesPortableGeometryAndUsesReplacementPixels() {
        byte[] source = PdfStamper.StampImage(
            CreateTextPdf(),
            PdfPngTestImages.CreateRgbPng(255, 0, 0),
            new PdfImageStampOptions {
                PageNumbers = new[] { 1 },
                X = 80D,
                Y = 140D,
                Width = 60D,
                Height = 30D,
                RotationDegrees = 30D
            });
        PdfDocument document = PdfDocument.Open(source);
        PdfImagePlacement original = Assert.Single(document.Images.Placements());
        byte[] originalPayload = Assert.Single(document.Read.Images()).Bytes;

        PdfImageEditResult result = document.Images.Replace(original, PdfPngTestImages.CreateRgbPng(0, 0, 255));
        PdfImagePlacement replacement = Assert.Single(result.Document.Images.Placements());
        byte[] replacementPayload = Assert.Single(result.Document.Read.Images()).Bytes;

        Assert.Equal(original.A, replacement.A, 2);
        Assert.Equal(original.B, replacement.B, 2);
        Assert.Equal(original.C, replacement.C, 2);
        Assert.Equal(original.D, replacement.D, 2);
        Assert.Equal(original.E, replacement.E, 2);
        Assert.Equal(original.F, replacement.F, 2);
        Assert.NotEqual(Convert.ToBase64String(originalPayload), Convert.ToBase64String(replacementPayload));
        Assert.Contains("Image editor proof", result.Document.Read.Text(), StringComparison.Ordinal);
    }

    [Fact]
    public void MovePreservesPortableTransformAndMovesByRequestedOffset() {
        PdfDocument source = PdfDocument.Open(CreateTextPdf()).Images.Add(
            new PdfPageRegion(1, 50D, 90D, 40D, 25D),
            PdfPngTestImages.CreateRgbPng(10, 80, 160)).Document;
        PdfImagePlacement original = Assert.Single(source.Images.Placements());

        PdfImageEditResult result = source.Images.Move(original, 35D, -15D, new PdfImageEditOptions {
            Layer = PdfImageEditLayer.BehindExistingContent
        });
        PdfImagePlacement moved = Assert.Single(result.Document.Images.Placements());

        Assert.Equal(original.E + 35D, moved.E, 2);
        Assert.Equal(original.F - 15D, moved.F, 2);
        Assert.Equal(original.A, moved.A, 2);
        Assert.Equal(original.D, moved.D, 2);
        Assert.Contains("Image editor proof", result.Document.Read.Text(), StringComparison.Ordinal);
    }

    [Fact]
    public void RemoveFailsClosedForAnAmbiguousExactPlacement() {
        byte[] source = BuildRawImagePdf(
            "q 40 0 0 20 20 30 cm /Im0 Do Q\n" +
            "q 40 0 0 20 20 30 cm /Im0 Do Q\n");
        PdfDocument document = PdfDocument.Open(source);
        PdfImagePlacement selected = document.Images.Placements()[0];

        InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() => document.Images.Remove(selected));

        Assert.Contains("ambiguous", exception.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(2, document.Images.Placements().Count);
    }

    [Fact]
    public void ReplaceAndMoveFailClosedWhenPlacementSemanticsCannotBeReproduced() {
        byte[] clipped = BuildRawImagePdf("q 0 0 15 15 re W n 40 0 0 20 20 30 cm /Im0 Do Q\n");
        byte[] skewed = BuildRawImagePdf("q 40 5 0 20 20 30 cm /Im0 Do Q\n");
        PdfDocument clippedDocument = PdfDocument.Open(clipped);
        PdfDocument skewedDocument = PdfDocument.Open(skewed);

        Assert.Throws<NotSupportedException>(() => clippedDocument.Images.Replace(
            Assert.Single(clippedDocument.Images.Placements()),
            PdfPngTestImages.CreateRgbPng(0, 0, 255)));
        Assert.Throws<NotSupportedException>(() => skewedDocument.Images.Move(
            Assert.Single(skewedDocument.Images.Placements()),
            10D,
            10D));
    }

    [Fact]
    public void RemoveSupportsSkewedXObjectsButInlineImagesFailClosed() {
        byte[] skewed = BuildRawImagePdf("q 40 5 0 20 20 30 cm /Im0 Do Q\n");
        byte[] inline = BuildRawInlineImagePdf("q 40 0 0 20 20 30 cm BI /W 1 /H 1 /BPC 8 /CS /RGB ID ", new byte[] { 255, 0, 0 }, " EI Q\n");
        PdfDocument skewedDocument = PdfDocument.Open(skewed);
        PdfDocument inlineDocument = PdfDocument.Open(inline);

        PdfImageEditResult removed = skewedDocument.Images.Remove(Assert.Single(skewedDocument.Images.Placements()));

        Assert.Empty(removed.Document.Images.Placements());
        Assert.Throws<NotSupportedException>(() => inlineDocument.Images.Remove(Assert.Single(inlineDocument.Images.Placements())));
    }

    [Fact]
    public void RemoveClonesARepeatedFormAndKeepsTheOtherImageInvocation() {
        byte[] source = BuildRepeatedFormImagePdf();
        PdfDocument document = PdfDocument.Open(source);
        PdfImagePlacement left = document.Images.Placements().OrderBy(static placement => placement.X).First();

        PdfImageEditResult result = document.Images.Remove(left);
        PdfImagePlacement remaining = Assert.Single(result.Document.Images.Placements());

        Assert.InRange(remaining.X, 119.99D, 120.01D);
    }

    [Fact]
    public void ExactRemovalKeepsSameNamedImageFromAnotherFormResourceContext() {
        PdfDocument document = PdfDocument.Open(BuildCollidingFormImagePdf());
        PdfImagePlacement selected = document.Images.Placements().Single(static placement => placement.ObjectNumber == 8);

        PdfImageEditResult result = document.Images.Remove(selected);
        PdfImagePlacement remaining = Assert.Single(result.Document.Images.Placements());

        Assert.Equal(2, Assert.Single(result.Document.Read.Images()).Width);
    }

    [Fact]
    public void ExactRemovalKeepsSameBoundsPlacementWithDifferentTransform() {
        byte[] source = BuildRawImagePdf(
            "q 40 0 0 40 20 30 cm /Im0 Do Q\n" +
            "q 0 40 -40 0 60 30 cm /Im0 Do Q\n");
        PdfDocument document = PdfDocument.Open(source);
        PdfImagePlacement selected = document.Images.Placements().Single(static placement => Math.Abs(placement.A - 40D) < 0.01D);

        PdfImageEditResult result = document.Images.Remove(selected);
        PdfImagePlacement remaining = Assert.Single(result.Document.Images.Placements());

        Assert.InRange(remaining.B, 39.99D, 40.01D);
    }

    [Fact]
    public void ExactRemovalCarriesGraphicsStateAcrossContentStreamArray() {
        PdfDocument document = PdfDocument.Open(BuildSplitContentImagePdf());

        PdfImageEditResult result = document.Images.Remove(Assert.Single(document.Images.Placements()));

        Assert.Empty(result.Document.Images.Placements());
    }

    [Fact]
    public void ReplaceRejectsNonNormalBlendModeAndMoveRejectsJpegDecodeSemantics() {
        byte[] blended = BuildRawImagePdf(
            "/GS1 gs q 40 0 0 20 20 30 cm /Im0 Do Q\n",
            additionalResources: "/ExtGState << /GS1 << /BM /Multiply >> >>");
        byte[] decodedJpeg = BuildRawImagePdf(
            "q 40 0 0 20 20 30 cm /Im0 Do Q\n",
            imageBytes: new byte[] { 0xFF, 0xD8, 0xFF, 0xD9 },
            imageEntries: "/ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Decode [1 0 1 0 1 0]");
        PdfDocument blendedDocument = PdfDocument.Open(blended);
        PdfDocument jpegDocument = PdfDocument.Open(decodedJpeg);

        Assert.Throws<NotSupportedException>(() => blendedDocument.Images.Replace(
            Assert.Single(blendedDocument.Images.Placements()),
            PdfPngTestImages.CreateRgbPng(0, 0, 255)));
        Assert.Throws<NotSupportedException>(() => jpegDocument.Images.Move(
            Assert.Single(jpegDocument.Images.Placements()),
            10D,
            0D));
    }

    [Fact]
    public void SoftMaskNoneClearsAnActiveGraphicsStateMaskForLaterImageEditing() {
        byte[] source = BuildRawImagePdf(
            "/GSActive gs /GSClear gs q 40 0 0 20 20 30 cm /Im0 Do Q\n",
            additionalResources: "/ExtGState << /GSActive << /SMask << >> >> /GSClear << /SMask /None >> >>");
        PdfDocument document = PdfDocument.Open(source);

        PdfImageEditResult result = document.Images.Replace(
            Assert.Single(document.Images.Placements()),
            PdfPngTestImages.CreateRgbPng(0, 0, 255));

        Assert.Single(result.Document.Images.Placements());
    }

    [Fact]
    public void DegeneratePlacementCannotReportSuccessfulRemoval() {
        PdfDocument document = PdfDocument.Open(BuildRawImagePdf("q 0 0 0 20 20 30 cm /Im0 Do Q\n"));

        NotSupportedException exception = Assert.Throws<NotSupportedException>(() =>
            document.Images.Remove(Assert.Single(document.Images.Placements())));

        Assert.Contains("degenerate", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void PlacementFromAnotherDocumentIsRejected() {
        PdfDocument first = PdfDocument.Open(CreateTextPdf()).Images.Add(
            new PdfPageRegion(1, 20D, 30D, 30D, 20D),
            PdfPngTestImages.CreateRgbPng(255, 0, 0)).Document;
        PdfDocument second = PdfDocument.Open(CreateTextPdf()).Images.Add(
            new PdfPageRegion(1, 120D, 130D, 30D, 20D),
            PdfPngTestImages.CreateRgbPng(255, 0, 0)).Document;

        InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
            second.Images.Remove(Assert.Single(first.Images.Placements())));

        Assert.Contains("does not originate", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void ReplacementRejectsImageMasksAndOptionalContentMembership() {
        PdfDocument imageMask = PdfDocument.Open(BuildRawImagePdf(
            "q 40 0 0 20 20 30 cm /Im0 Do Q\n",
            imageBytes: new byte[] { 0x80 },
            imageEntries: "/ImageMask true /BitsPerComponent 1"));
        PdfDocument optionalContent = PdfDocument.Open(BuildRawImagePdf(
            "q 40 0 0 20 20 30 cm /Im0 Do Q\n",
            imageEntries: "/ColorSpace /DeviceRGB /BitsPerComponent 8 /OC << /Type /OCG /Name (Layer) >>"));

        Assert.Throws<NotSupportedException>(() => imageMask.Images.Replace(
            Assert.Single(imageMask.Images.Placements()),
            PdfPngTestImages.CreateRgbPng(0, 0, 255)));
        Assert.Throws<NotSupportedException>(() => optionalContent.Images.Replace(
            Assert.Single(optionalContent.Images.Placements()),
            PdfPngTestImages.CreateRgbPng(0, 0, 255)));
    }

    [Fact]
    public void DestructiveEditsRejectTaggedAndHiddenOptionalContentContexts() {
        PdfDocument tagged = PdfDocument.Open(BuildRawImagePdf(
            "/Figure << /MCID 0 >> BDC q 40 0 0 20 20 30 cm /Im0 Do Q EMC\n"));
        PdfDocument hiddenSibling = PdfDocument.Open(BuildRawImagePdf(
            "q 40 0 0 20 20 30 cm /Im0 Do Q\n/OC /Hidden BDC q 40 0 0 20 120 30 cm /Im0 Do Q EMC\n"));

        Assert.Throws<NotSupportedException>(() => tagged.Images.Remove(Assert.Single(tagged.Images.Placements())));
        PdfImagePlacement visible = hiddenSibling.Images.Placements().OrderBy(static placement => placement.X).First();
        Assert.Throws<NotSupportedException>(() => hiddenSibling.Images.Remove(visible));
    }

    [Fact]
    public void DestructiveEditsRejectMarkedContentThatInvokesAContainingForm() {
        PdfDocument document = PdfDocument.Open(BuildMarkedFormImagePdf());

        NotSupportedException exception = Assert.Throws<NotSupportedException>(() =>
            document.Images.Remove(Assert.Single(document.Images.Placements())));

        Assert.Contains("marked content", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void XObjectRemovalSkipsUnrelatedInlineImagePayloadOperators() {
        PdfDocument document = PdfDocument.Open(BuildRawImagePdf(
            "BI /W 1 /H 1 /BPC 8 /CS /RGB ID q /Im0 Do Q EI\nq 40 0 0 20 20 30 cm /Im0 Do Q\n"));
        PdfImagePlacement xObject = document.Images.Placements().Single(static placement => placement.ObjectNumber > 0);

        PdfImageEditResult result = document.Images.Remove(xObject);

        PdfImagePlacement remaining = Assert.Single(result.Document.Images.Placements());
        Assert.NotNull(remaining.InlineImageStream);
    }

    [Fact]
    public void SharedFormAcrossContentStreamsFailsClosed() {
        PdfDocument document = PdfDocument.Open(BuildCrossStreamRepeatedFormImagePdf());
        PdfImagePlacement selected = document.Images.Placements().OrderBy(static placement => placement.X).First();

        NotSupportedException exception = Assert.Throws<NotSupportedException>(() => document.Images.Remove(selected));

        Assert.Contains("multiple content streams", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void SharedPageContentStreamThatInvokesAFormFailsClosed() {
        PdfDocument document = PdfDocument.Open(BuildSharedPageContentFormImagePdf());
        PdfImagePlacement selected = document.Images.Placements().Single(static placement => placement.PageNumber == 1);

        NotSupportedException exception = Assert.Throws<NotSupportedException>(() => document.Images.Remove(selected));

        Assert.Contains("shared page content", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void ImageEditorCoordinatesAreRelativeToNonzeroPageBoxOrigin() {
        byte[] source = BuildRawImagePdf(
            string.Empty,
            pageEntries: "/CropBox [100 100 500 700]");

        PdfDocument added = PdfDocument.Open(source).Images.Add(
            new PdfPageRegion(1, 0D, 0D, 30D, 20D),
            PdfPngTestImages.CreateRgbPng(255, 0, 0)).Document;
        PdfImagePlacement placement = Assert.Single(added.Images.Placements());
        PdfImagePlacement found = Assert.Single(added.Images.Find(new PdfPageRegion(1, 0D, 0D, 35D, 25D)));

        Assert.InRange(placement.X, -0.01D, 0.01D);
        Assert.InRange(placement.Y, -0.01D, 0.01D);
        Assert.Equal(placement.X, found.X, 3);
        Assert.Equal(placement.Y, found.Y, 3);
    }

    [Fact]
    public void ImageEditsRecordPageContentMutation() {
        PdfImageEditResult result = PdfDocument.Open(CreateTextPdf()).Images.Add(
            new PdfPageRegion(1, 20D, 30D, 30D, 20D),
            PdfPngTestImages.CreateRgbPng(255, 0, 0));

        PdfPipelineStep mutation = Assert.Single(result.Document.Pipeline.Steps, static step => step.Kind == PdfPipelineStepKind.Mutation);
        Assert.Equal("Image", mutation.Operation);
        Assert.Equal(PdfMutationOperation.ModifyPageContent, mutation.MutationOperation);
    }

    [Fact]
    public void ReplaceCarriesExactInputBudgetAcrossRemovalAndStamping() {
        PdfDocument prepared = PdfDocument.Open(CreateTextPdf()).Images.Add(
            new PdfPageRegion(1, 40D, 80D, 30D, 20D),
            PdfPngTestImages.CreateRgbPng(255, 0, 0)).Document;
        byte[] source = prepared.ToBytes();
        var readOptions = new PdfReadOptions { Limits = new PdfReadLimits { MaxInputBytes = source.Length } };
        PdfDocument bounded = PdfDocument.Open(source, readOptions);

        PdfImageEditResult result = bounded.Images.Replace(
            Assert.Single(bounded.Images.Placements()),
            PdfPngTestImages.CreateRgbPng(32, 32));

        Assert.Single(result.Document.Images.Placements());
        Assert.Single(result.Document.Read.Images());
    }

    [Fact]
    public void PublicOptionsValidateCoordinatesAndPageSelection() {
        PdfDocument document = PdfDocument.Open(CreateTextPdf());

        Assert.Throws<ArgumentOutOfRangeException>(() => document.Images.Find(new PdfPageRegion(2, 0D, 0D, 10D, 10D)));
        Assert.Throws<ArgumentOutOfRangeException>(() => document.Images.Move(
            new PdfImagePlacement(1, "Im0", 1, 0, 1D, 0D, 0D, 1D, 0D, 0D, 0D, 0D, 1D, 1D),
            double.NaN,
            0D));
        Assert.Throws<ArgumentOutOfRangeException>(() => new PdfImageEditOptions { Layer = (PdfImageEditLayer)999 });
    }

    private static byte[] CreateTextPdf() => PdfDocument.Create()
        .Paragraph(paragraph => paragraph.Text("Image editor proof"))
        .ToBytes();

    private static byte[] BuildRawImagePdf(
        string content,
        string additionalResources = "",
        byte[]? imageBytes = null,
        string imageEntries = "/ColorSpace /DeviceRGB /BitsPerComponent 8",
        string pageEntries = "") {
        byte[] contentBytes = Encoding.ASCII.GetBytes(content);
        imageBytes ??= new byte[] { 255, 0, 0 };
        using var output = new MemoryStream();
        WriteAscii(output, "%PDF-1.7\n");
        WriteAscii(output, "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n");
        WriteAscii(output, "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] >>\nendobj\n");
        WriteAscii(output, "3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 120] " + pageEntries + " /Resources << /XObject << /Im0 5 0 R >> " + additionalResources + " >> /Contents 4 0 R >>\nendobj\n");
        WriteAscii(output, "4 0 obj\n<< /Length " + contentBytes.Length.ToString(CultureInfo.InvariantCulture) + " >>\nstream\n");
        output.Write(contentBytes, 0, contentBytes.Length);
        WriteAscii(output, "endstream\nendobj\n");
        WriteAscii(output, "5 0 obj\n<< /Type /XObject /Subtype /Image /Width 1 /Height 1 " + imageEntries + " /Length " + imageBytes.Length.ToString(CultureInfo.InvariantCulture) + " >>\nstream\n");
        output.Write(imageBytes, 0, imageBytes.Length);
        WriteAscii(output, "\nendstream\nendobj\ntrailer\n<< /Root 1 0 R /Size 6 >>\n%%EOF\n");
        return output.ToArray();
    }

    private static byte[] BuildRawInlineImagePdf(string prefix, byte[] imageBytes, string suffix) {
        byte[] prefixBytes = Encoding.ASCII.GetBytes(prefix);
        byte[] suffixBytes = Encoding.ASCII.GetBytes(suffix);
        int contentLength = prefixBytes.Length + imageBytes.Length + suffixBytes.Length;
        using var output = new MemoryStream();
        WriteAscii(output, "%PDF-1.7\n");
        WriteAscii(output, "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n");
        WriteAscii(output, "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] >>\nendobj\n");
        WriteAscii(output, "3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 120] /Contents 4 0 R >>\nendobj\n");
        WriteAscii(output, "4 0 obj\n<< /Length " + contentLength.ToString(CultureInfo.InvariantCulture) + " >>\nstream\n");
        output.Write(prefixBytes, 0, prefixBytes.Length);
        output.Write(imageBytes, 0, imageBytes.Length);
        output.Write(suffixBytes, 0, suffixBytes.Length);
        WriteAscii(output, "endstream\nendobj\ntrailer\n<< /Root 1 0 R /Size 5 >>\n%%EOF\n");
        return output.ToArray();
    }

    private static byte[] BuildRepeatedFormImagePdf() {
        const string pageContent = "q 1 0 0 1 20 30 cm /Fx Do Q\nq 1 0 0 1 120 30 cm /Fx Do Q\n";
        const string formContent = "q 10 0 0 10 0 0 cm /ImShared Do Q\n";
        const string imageBytes = "abc";
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj", "<< /Type /Catalog /Pages 2 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] /Resources << /XObject << /Fx 6 0 R >> >> >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /Contents 5 0 R >>", "endobj",
            "5 0 obj", "<< /Length " + Encoding.ASCII.GetByteCount(pageContent).ToString(CultureInfo.InvariantCulture) + " >>", "stream", pageContent.TrimEnd('\n'), "endstream", "endobj",
            "6 0 obj", "<< /Type /XObject /Subtype /Form /BBox [0 0 20 20] /Resources << /XObject << /ImShared 7 0 R >> >> /Length " + Encoding.ASCII.GetByteCount(formContent).ToString(CultureInfo.InvariantCulture) + " >>", "stream", formContent.TrimEnd('\n'), "endstream", "endobj",
            "7 0 obj", "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Length 3 >>", "stream", imageBytes, "endstream", "endobj",
            "trailer", "<< /Root 1 0 R >>", "%%EOF"
        }) + "\n";
        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildCrossStreamRepeatedFormImagePdf() {
        const string firstContent = "q 1 0 0 1 20 30 cm /Fx Do Q\n";
        const string secondContent = "q 1 0 0 1 120 30 cm /Fx Do Q\n";
        const string formContent = "q 10 0 0 10 0 0 cm /ImShared Do Q\n";
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj", "<< /Type /Catalog /Pages 2 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 120] /Resources << /XObject << /Fx 6 0 R >> >> /Contents [4 0 R 5 0 R] >>", "endobj",
            "4 0 obj", "<< /Length " + Encoding.ASCII.GetByteCount(firstContent).ToString(CultureInfo.InvariantCulture) + " >>", "stream", firstContent.TrimEnd('\n'), "endstream", "endobj",
            "5 0 obj", "<< /Length " + Encoding.ASCII.GetByteCount(secondContent).ToString(CultureInfo.InvariantCulture) + " >>", "stream", secondContent.TrimEnd('\n'), "endstream", "endobj",
            "6 0 obj", "<< /Type /XObject /Subtype /Form /BBox [0 0 20 20] /Resources << /XObject << /ImShared 7 0 R >> >> /Length " + Encoding.ASCII.GetByteCount(formContent).ToString(CultureInfo.InvariantCulture) + " >>", "stream", formContent.TrimEnd('\n'), "endstream", "endobj",
            "7 0 obj", "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Length 3 >>", "stream", "abc", "endstream", "endobj",
            "trailer", "<< /Root 1 0 R /Size 8 >>", "%%EOF"
        }) + "\n";
        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildCollidingFormImagePdf() {
        const string pageContent = "q /F1 Do Q\nq /F2 Do Q\n";
        const string formContent = "q 40 0 0 20 20 30 cm /Im0 Do Q\n";
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj", "<< /Type /Catalog /Pages 2 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 120] /Resources << /XObject << /F1 6 0 R /F2 7 0 R >> >> /Contents 5 0 R >>", "endobj",
            "5 0 obj", "<< /Length " + Encoding.ASCII.GetByteCount(pageContent).ToString(CultureInfo.InvariantCulture) + " >>", "stream", pageContent.TrimEnd('\n'), "endstream", "endobj",
            "6 0 obj", "<< /Type /XObject /Subtype /Form /BBox [0 0 200 120] /Resources << /XObject << /Im0 8 0 R >> >> /Length " + Encoding.ASCII.GetByteCount(formContent).ToString(CultureInfo.InvariantCulture) + " >>", "stream", formContent.TrimEnd('\n'), "endstream", "endobj",
            "7 0 obj", "<< /Type /XObject /Subtype /Form /BBox [0 0 200 120] /Resources << /XObject << /Im0 9 0 R >> >> /Length " + Encoding.ASCII.GetByteCount(formContent).ToString(CultureInfo.InvariantCulture) + " >>", "stream", formContent.TrimEnd('\n'), "endstream", "endobj",
            "8 0 obj", "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Length 3 >>", "stream", "abc", "endstream", "endobj",
            "9 0 obj", "<< /Type /XObject /Subtype /Image /Width 2 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Length 6 >>", "stream", "xyzuvw", "endstream", "endobj",
            "trailer", "<< /Root 1 0 R /Size 10 >>", "%%EOF"
        }) + "\n";
        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildSplitContentImagePdf() {
        const string first = "q 40 0 0 20 20 30 cm\n";
        const string second = "/Im0 Do Q\n";
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj", "<< /Type /Catalog /Pages 2 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 120] /Resources << /XObject << /Im0 6 0 R >> >> /Contents [4 0 R 5 0 R] >>", "endobj",
            "4 0 obj", "<< /Length " + Encoding.ASCII.GetByteCount(first).ToString(CultureInfo.InvariantCulture) + " >>", "stream", first.TrimEnd('\n'), "endstream", "endobj",
            "5 0 obj", "<< /Length " + Encoding.ASCII.GetByteCount(second).ToString(CultureInfo.InvariantCulture) + " >>", "stream", second.TrimEnd('\n'), "endstream", "endobj",
            "6 0 obj", "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Length 3 >>", "stream", "abc", "endstream", "endobj",
            "trailer", "<< /Root 1 0 R /Size 7 >>", "%%EOF"
        }) + "\n";
        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildMarkedFormImagePdf() {
        const string pageContent = "/Figure << /MCID 0 >> BDC /Fx Do EMC\n";
        const string formContent = "q 40 0 0 20 20 30 cm /Im0 Do Q\n";
        string pdf = string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj", "<< /Type /Catalog /Pages 2 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 120] /Resources << /XObject << /Fx 6 0 R >> >> /Contents 4 0 R >>", "endobj",
            "4 0 obj", "<< /Length " + Encoding.ASCII.GetByteCount(pageContent).ToString(CultureInfo.InvariantCulture) + " >>", "stream", pageContent.TrimEnd('\n'), "endstream", "endobj",
            "6 0 obj", "<< /Type /XObject /Subtype /Form /BBox [0 0 200 120] /Resources << /XObject << /Im0 7 0 R >> >> /Length " + Encoding.ASCII.GetByteCount(formContent).ToString(CultureInfo.InvariantCulture) + " >>", "stream", formContent.TrimEnd('\n'), "endstream", "endobj",
            "7 0 obj", "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Length 3 >>", "stream", "abc", "endstream", "endobj",
            "trailer", "<< /Root 1 0 R /Size 8 >>", "%%EOF"
        }) + "\n";
        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildSharedPageContentFormImagePdf() {
        const string pageContent = "/Fx Do\n";
        const string formContent = "q 40 0 0 20 20 30 cm /Im0 Do Q\n";
        string pdf = string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj", "<< /Type /Catalog /Pages 2 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 2 /Kids [3 0 R 4 0 R] /MediaBox [0 0 200 120] /Resources 8 0 R >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /Contents 5 0 R >>", "endobj",
            "4 0 obj", "<< /Type /Page /Parent 2 0 R /Contents 5 0 R >>", "endobj",
            "5 0 obj", "<< /Length " + Encoding.ASCII.GetByteCount(pageContent).ToString(CultureInfo.InvariantCulture) + " >>", "stream", pageContent.TrimEnd('\n'), "endstream", "endobj",
            "6 0 obj", "<< /Type /XObject /Subtype /Form /BBox [0 0 200 120] /Resources << /XObject << /Im0 7 0 R >> >> /Length " + Encoding.ASCII.GetByteCount(formContent).ToString(CultureInfo.InvariantCulture) + " >>", "stream", formContent.TrimEnd('\n'), "endstream", "endobj",
            "7 0 obj", "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Length 3 >>", "stream", "abc", "endstream", "endobj",
            "8 0 obj", "<< /XObject << /Fx 6 0 R >> >>", "endobj",
            "trailer", "<< /Root 1 0 R /Size 9 >>", "%%EOF"
        }) + "\n";
        return Encoding.ASCII.GetBytes(pdf);
    }

    private static void WriteAscii(Stream stream, string value) {
        byte[] bytes = Encoding.ASCII.GetBytes(value);
        stream.Write(bytes, 0, bytes.Length);
    }
}
