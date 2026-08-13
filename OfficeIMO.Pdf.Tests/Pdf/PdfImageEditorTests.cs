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
    public void AddInFrontIsolatesTheExistingPageGraphicsState() {
        byte[] source = BuildRawImagePdf("2 0 0 2 15 10 cm 0 0 5 5 re W n\n");

        byte[] result = PdfDocument.Open(source).Images.Add(
            new PdfPageRegion(1, 40D, 40D, 20D, 20D),
            PdfPngTestImages.CreateRgbPng(0, 0, 255)).Document.ToBytes();
        Dictionary<int, PdfIndirectObject> objects = PdfSyntax.ParseObjects(result).Map;
        PdfDictionary page = Assert.IsType<PdfDictionary>(objects[3].Value);
        PdfArray contents = Assert.IsType<PdfArray>(page.Items["Contents"]);
        string[] streams = contents.Items
            .Select(item => Assert.IsType<PdfStream>(PdfObjectLookup.Resolve(objects, item)))
            .Select(stream => PdfEncoding.Latin1GetString(stream.Data))
            .ToArray();

        Assert.Equal("q\n", streams[0]);
        Assert.Contains("2 0 0 2 15 10 cm", streams[1], StringComparison.Ordinal);
        Assert.Equal("\nQ\n", streams[2]);
        Assert.Contains(" Do", streams[3], StringComparison.Ordinal);
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
    public void DirectPlacementsWithDifferentStreamIdentityAreDistinct() {
        var first = new PdfImagePlacement(1, "Im0", 0, 101, 40, 0, 0, 20, 20, 30, 20, 30, 40, 20);
        var second = new PdfImagePlacement(1, "Im0", 0, 202, 40, 0, 0, 20, 20, 30, 20, 30, 40, 20);

        Assert.False(PdfImageEditor.SamePlacementIdentity(first, second));
    }

    [Fact]
    public void ExactRemovalCarriesGraphicsStateAcrossContentStreamArray() {
        PdfDocument document = PdfDocument.Open(BuildSplitContentImagePdf());

        PdfImageEditResult result = document.Images.Remove(Assert.Single(document.Images.Placements()));

        Assert.Empty(result.Document.Images.Placements());
    }

    [Fact]
    public void DestructiveEditsCarryMarkedContentAcrossContentStreamArray() {
        PdfDocument document = PdfDocument.Open(BuildSplitMarkedContentImagePdf());

        NotSupportedException exception = Assert.Throws<NotSupportedException>(() =>
            document.Images.Remove(Assert.Single(document.Images.Placements())));

        Assert.Contains("marked content", exception.Message, StringComparison.OrdinalIgnoreCase);
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
    public void MoveRejectsSourceInterpolationThatRestampingCannotPreserve() {
        PdfDocument document = PdfDocument.Open(BuildRawImagePdf(
            "q 40 0 0 20 20 30 cm /Im0 Do Q\n",
            imageEntries: "/ColorSpace /DeviceRGB /BitsPerComponent 8 /Interpolate true"));

        NotSupportedException exception = Assert.Throws<NotSupportedException>(() =>
            document.Images.Move(Assert.Single(document.Images.Placements()), 10D, 0D));

        Assert.Contains("interpolation", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void ReplaceRejectsUnsupportedAuthoredBlendMode() {
        byte[] source = BuildRawImagePdf(
            "/GS1 gs q 40 0 0 20 20 30 cm /Im0 Do Q\n",
            additionalResources: "/ExtGState << /GS1 << /BM /ProducerSpecific >> >>");
        PdfDocument document = PdfDocument.Open(source);

        Assert.Throws<NotSupportedException>(() => document.Images.Replace(
            Assert.Single(document.Images.Placements()),
            PdfPngTestImages.CreateRgbPng(0, 0, 255)));
    }

    [Fact]
    public void MoveRejectsUnsupportedImagePaintState() {
        byte[] source = BuildRawImagePdf(
            "/GS1 gs q 40 0 0 20 20 30 cm /Im0 Do Q\n",
            additionalResources: "/ExtGState << /GS1 << /op true /OPM 1 >> >>");
        PdfDocument document = PdfDocument.Open(source);

        NotSupportedException exception = Assert.Throws<NotSupportedException>(() =>
            document.Images.Move(Assert.Single(document.Images.Placements()), 10D, 0D));

        Assert.Contains("paint effect", exception.Message, StringComparison.OrdinalIgnoreCase);
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
    public void MoveRejectsOptionalContentInheritedFromContainingForm() {
        PdfDocument document = PdfDocument.Open(BuildMarkedFormImagePdf("/OC << /Type /OCG /Name (Layer) >>"));

        NotSupportedException exception = Assert.Throws<NotSupportedException>(() =>
            document.Images.Move(Assert.Single(document.Images.Placements()), 10D, 0D));

        Assert.Contains("optional-content", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void PortableImageEditsRejectImagesInsideTransparencyGroups() {
        PdfDocument document = PdfDocument.Open(BuildMarkedFormImagePdf(
            "/Group << /S /Transparency /I true /CS /DeviceRGB >>"));

        NotSupportedException exception = Assert.Throws<NotSupportedException>(() =>
            document.Images.Move(Assert.Single(document.Images.Placements()), 10D, 0D));

        Assert.Contains("transparency group", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void MarkedContentCheckUsesTheSelectedResourceScope() {
        PdfDocument document = PdfDocument.Open(BuildCollidingFormImagePdf(markSecondImage: true));
        PdfImagePlacement selected = document.Images.Placements().Single(static placement => placement.ObjectNumber == 8);

        PdfImageEditResult result = document.Images.Remove(selected);

        Assert.Single(result.Document.Images.Placements());
    }

    [Fact]
    public void ImageValidationHonorsConfiguredDecodedStreamLimit() {
        byte[] source = BuildRawImagePdf(new string(' ', 96) + "q 40 0 0 20 20 30 cm /Im0 Do Q\n");
        PdfDocument document = PdfDocument.Open(source);
        PdfImagePlacement placement = Assert.Single(document.Images.Placements());
        var options = new PdfReadOptions { Limits = new PdfReadLimits { MaxDecodedStreamBytes = 64 } };

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() =>
            document.Images.Move(placement, 10D, 0D, readOptions: options));

        Assert.Equal(PdfReadLimitKind.DecodedStreamBytes, exception.Kind);
    }

    [Fact]
    public void MovableImageExtractionHonorsConfiguredDecodedStreamLimit() {
        byte[] encodedImage = Encoding.ASCII.GetBytes(new string('0', 200) + ">");
        byte[] source = BuildRawImagePdf(
            "q 40 0 0 20 20 30 cm /Im0 Do Q\n",
            imageBytes: encodedImage,
            imageEntries: "/ColorSpace /DeviceGray /BitsPerComponent 8 /Filter /ASCIIHexDecode");
        PdfDocument document = PdfDocument.Open(source);
        PdfImagePlacement placement = Assert.Single(document.Images.Placements());
        var options = new PdfReadOptions {
            Limits = new PdfReadLimits { MaxDecodedStreamBytes = 64 }
        };

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() =>
            document.Images.Move(placement, 10D, 0D, readOptions: options));

        Assert.Equal(PdfReadLimitKind.DecodedStreamBytes, exception.Kind);
    }

    [Fact]
    public void ImageValidationHonorsConfiguredContentNestingLimit() {
        string nestedOperand = new string('[', 129) + "0" + new string(']', 129);
        string unusedFormContent = nestedOperand + " n";
        string unusedForm = "6 0 obj\n<< /Type /XObject /Subtype /Form /BBox [0 0 1 1] /Length " +
            Encoding.ASCII.GetByteCount(unusedFormContent).ToString(CultureInfo.InvariantCulture) +
            " >>\nstream\n" + unusedFormContent + "\nendstream\nendobj\n";
        byte[] source = BuildRawImagePdf(
            "q 40 0 0 20 20 30 cm /Im0 Do Q\n",
            additionalObjects: unusedForm);
        var options = new PdfReadOptions { Limits = new PdfReadLimits { MaxContentNestingDepth = 256 } };
        PdfDocument document = PdfDocument.Open(source, options);

        PdfImageEditResult result = document.Images.Move(
            Assert.Single(document.Images.Placements()),
            10D,
            0D,
            readOptions: options);

        Assert.Single(result.Document.Images.Placements());
    }

    [Fact]
    public void ImageValidationDoesNotDecodeUnreachableStreams() {
        string unusedStream = "6 0 obj\n<< /Length 129 /Filter /ASCIIHexDecode >>\nstream\n" +
            new string('4', 128) + ">\nendstream\nendobj\n";
        byte[] source = BuildRawImagePdf(
            "q 40 0 0 20 20 30 cm /Im0 Do Q\n",
            additionalObjects: unusedStream);
        var options = new PdfReadOptions { Limits = new PdfReadLimits { MaxDecodedStreamBytes = 48 } };
        PdfDocument document = PdfDocument.Open(source, options);

        PdfImageEditResult result = document.Images.Move(
            Assert.Single(document.Images.Placements()),
            10D,
            0D,
            readOptions: options);

        Assert.Single(result.Document.Images.Placements());
    }

    [Fact]
    public void ImageValidationBoundsDocumentWideRetainedContentSeparately() {
        PdfDocument document = PdfDocument.Open(BuildSharedResourceLessFormImagePdf());
        PdfImagePlacement placement = document.Images.Placements().First();
        var options = new PdfReadOptions { Limits = new PdfReadLimits { MaxPageContentBytes = 100, MaxRetainedContentBytes = 40 } };

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() =>
            document.Images.Move(placement, 10D, 0D, readOptions: options));

        Assert.Equal(PdfReadLimitKind.RetainedContentBytes, exception.Kind);
        Assert.Equal(40, exception.Limit);
        Assert.True(exception.Actual > exception.Limit);
    }

    [Fact]
    public void ImageValidationAppliesContentByteBudgetPerPage() {
        PdfDocument document = PdfDocument.Open(BuildSharedResourceLessFormImagePdf());
        PdfImagePlacement placement = document.Images.Placements().First();
        var options = new PdfReadOptions { Limits = new PdfReadLimits { MaxPageContentBytes = 45 } };

        NotSupportedException exception = Assert.Throws<NotSupportedException>(() =>
            document.Images.Move(placement, 10D, 0D, readOptions: options));

        Assert.Contains("multiple content streams", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void PortableImageEditsRejectAuthoredRenderingIntent() {
        byte[] source = BuildRawImagePdf(
            "q 40 0 0 20 20 30 cm /Im0 Do Q\n",
            imageEntries: "/ColorSpace /DeviceRGB /BitsPerComponent 8 /Intent /Perceptual");
        PdfDocument document = PdfDocument.Open(source);
        PdfImagePlacement placement = Assert.Single(document.Images.Placements());

        NotSupportedException moveException = Assert.Throws<NotSupportedException>(() =>
            document.Images.Move(placement, 10D, 0D));
        NotSupportedException replaceException = Assert.Throws<NotSupportedException>(() =>
            document.Images.Replace(placement, PdfPngTestImages.CreateRgbPng(0, 0, 255)));

        Assert.Contains("rendering intent", moveException.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("rendering intent", replaceException.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void PortableImageEditsRejectAlternatePresentations() {
        byte[] source = BuildRawImagePdf(
            "q 40 0 0 20 20 30 cm /Im0 Do Q\n",
            imageEntries: "/ColorSpace /DeviceRGB /BitsPerComponent 8 /Alternates []");
        PdfDocument document = PdfDocument.Open(source);
        PdfImagePlacement placement = Assert.Single(document.Images.Placements());

        NotSupportedException moveException = Assert.Throws<NotSupportedException>(() =>
            document.Images.Move(placement, 10D, 0D));
        NotSupportedException replaceException = Assert.Throws<NotSupportedException>(() =>
            document.Images.Replace(placement, PdfPngTestImages.CreateRgbPng(0, 0, 255)));

        Assert.Contains("alternate", moveException.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("alternate", replaceException.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void RemovingPageImageRetainsSharedImageInvokedByTilingPattern() {
        PdfDocument document = PdfDocument.Open(BuildPatternSharedImagePdf());
        PdfImagePlacement selected = Assert.Single(document.Images.Placements());

        PdfImageEditResult result = document.Images.Remove(selected);
        string output = PdfEncoding.Latin1GetString(result.Document.ToBytes());

        Assert.Contains("/ImPattern", output, StringComparison.Ordinal);
    }

    [Fact]
    public void RemovingPageImageRetainsSharedImageInvokedByType3CharProc() {
        PdfDocument document = PdfDocument.Open(BuildType3SharedImagePdf());
        PdfImagePlacement selected = Assert.Single(document.Images.Placements());

        PdfImageEditResult result = document.Images.Remove(selected);
        string output = PdfEncoding.Latin1GetString(result.Document.ToBytes());

        Assert.Contains("/ImGlyph", output, StringComparison.Ordinal);
    }

    [Fact]
    public void PortableImageEditsRejectContentRenderingIntent() {
        byte[] source = BuildRawImagePdf(
            "q /Perceptual ri 40 0 0 20 20 30 cm /Im0 Do Q\n");
        PdfDocument document = PdfDocument.Open(source);
        PdfImagePlacement placement = Assert.Single(document.Images.Placements());

        NotSupportedException moveException = Assert.Throws<NotSupportedException>(() =>
            document.Images.Move(placement, 10D, 0D));
        NotSupportedException replaceException = Assert.Throws<NotSupportedException>(() =>
            document.Images.Replace(placement, PdfPngTestImages.CreateRgbPng(0, 0, 255)));

        Assert.Contains("rendering intent", moveException.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("rendering intent", replaceException.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void DestructiveEditsRejectStructParentOnSelectedImage() {
        PdfDocument document = PdfDocument.Open(BuildRawImagePdf(
            "q 40 0 0 20 20 30 cm /Im0 Do Q\n",
            imageEntries: "/ColorSpace /DeviceRGB /BitsPerComponent 8 /StructParent 0"));

        NotSupportedException exception = Assert.Throws<NotSupportedException>(() =>
            document.Images.Remove(Assert.Single(document.Images.Placements())));

        Assert.Contains("structure tree", exception.Message, StringComparison.OrdinalIgnoreCase);
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
    public void SharedOuterFormThatInvokesImageFormFailsClosed() {
        PdfDocument document = PdfDocument.Open(BuildSharedOuterFormImagePdf());
        PdfImagePlacement selected = document.Images.Placements().Single(static placement => placement.PageNumber == 1);

        NotSupportedException exception = Assert.Throws<NotSupportedException>(() => document.Images.Remove(selected));

        Assert.Contains("multiple content streams", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void SharedResourceLessFormUsesEachInvokingPageResourceContextAndFailsClosed() {
        PdfDocument document = PdfDocument.Open(BuildSharedResourceLessFormImagePdf());
        PdfImagePlacement selected = document.Images.Placements().Single(static placement => placement.PageNumber == 1);

        NotSupportedException exception = Assert.Throws<NotSupportedException>(() => document.Images.Remove(selected));

        Assert.Contains("multiple content streams", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void RemovePreservesImageResourceUsedByResourceLessDescendantForm() {
        PdfDocument document = PdfDocument.Open(BuildDirectAndInheritedFormImagePdf());
        PdfImagePlacement selected = document.Images.Placements().Single(static placement => placement.X < 50D);

        PdfImageEditResult result = document.Images.Remove(selected);

        PdfImagePlacement remaining = Assert.Single(result.Document.Images.Placements());
        Assert.InRange(remaining.X, 99.99D, 100.01D);
    }

    [Fact]
    public void RemoveFromResourceLessFormPreservesDirectPageImageResource() {
        PdfDocument document = PdfDocument.Open(BuildDirectAndInheritedFormImagePdf());
        PdfImagePlacement selected = document.Images.Placements().Single(static placement => placement.X > 50D);

        PdfImageEditResult result = document.Images.Remove(selected);

        PdfImagePlacement remaining = Assert.Single(result.Document.Images.Placements());
        Assert.InRange(remaining.X, 19.99D, 20.01D);
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
    public void ImageMutationsPreserveAnExistingAcroFormCatalogGraph() {
        byte[] form = PdfDocument.Open(CreateTextPdf()).Forms.Edit(edit => edit.Create(new PdfFormFieldCreateOptions {
            Name = "customer.notes",
            Kind = PdfFormFieldCreationKind.Text,
            X = 72D,
            Y = 600D,
            Width = 180D,
            Height = 24D,
            Value = "kept"
        })).ToBytes();

        PdfDocument added = PdfDocument.Open(form).Images.Add(
            new PdfPageRegion(1, 40D, 80D, 30D, 20D),
            PdfPngTestImages.CreateRgbPng(255, 0, 0)).Document;
        PdfDocument replaced = added.Images.Replace(
            Assert.Single(added.Images.Placements()),
            PdfPngTestImages.CreateRgbPng(0, 0, 255)).Document;
        PdfDocument moved = replaced.Images.Move(Assert.Single(replaced.Images.Placements()), 20D, 0D).Document;
        PdfDocument removed = moved.Images.Remove(Assert.Single(moved.Images.Placements())).Document;

        foreach (PdfDocument document in new[] { added, replaced, moved, removed }) {
            PdfFormField field = Assert.Single(document.Inspect().FormFields);
            Assert.Equal("customer.notes", field.Name);
            Assert.Equal("kept", field.Value);
        }
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
        string pageEntries = "",
        string additionalObjects = "") {
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
        WriteAscii(output, "\nendstream\nendobj\n");
        WriteAscii(output, additionalObjects);
        WriteAscii(output, "trailer\n<< /Root 1 0 R /Size 7 >>\n%%EOF\n");
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

    private static byte[] BuildPatternSharedImagePdf() {
        const string pageContent = "q 40 0 0 20 20 30 cm /ImPage Do Q /Pattern cs /P1 scn 0 0 200 120 re f\n";
        const string patternContent = "q 5 0 0 5 0 0 cm /ImPattern Do Q\n";
        return Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj", "<< /Type /Catalog /Pages 2 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 120] /Resources << /XObject << /ImPage 6 0 R >> /Pattern << /P1 7 0 R >> >> /Contents 4 0 R >>", "endobj",
            "4 0 obj", "<< /Length " + Encoding.ASCII.GetByteCount(pageContent).ToString(CultureInfo.InvariantCulture) + " >>", "stream", pageContent.TrimEnd('\n'), "endstream", "endobj",
            "6 0 obj", "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Length 3 >>", "stream", "abc", "endstream", "endobj",
            "7 0 obj", "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 5 /YStep 5 /Resources << /XObject << /ImPattern 6 0 R >> >> /Length " + Encoding.ASCII.GetByteCount(patternContent).ToString(CultureInfo.InvariantCulture) + " >>", "stream", patternContent.TrimEnd('\n'), "endstream", "endobj",
            "trailer", "<< /Root 1 0 R /Size 8 >>", "%%EOF", string.Empty
        }));
    }

    private static byte[] BuildType3SharedImagePdf() {
        const string pageContent = "q 40 0 0 20 20 30 cm /ImPage Do Q BT /FType3 18 Tf 80 50 Td (A) Tj ET\n";
        const string glyphContent = "500 0 d0 q 500 0 0 700 0 0 cm /ImGlyph Do Q\n";
        return Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj", "<< /Type /Catalog /Pages 2 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 120] /Resources << /XObject << /ImPage 6 0 R >> /Font << /FType3 7 0 R >> >> /Contents 4 0 R >>", "endobj",
            "4 0 obj", "<< /Length " + Encoding.ASCII.GetByteCount(pageContent).ToString(CultureInfo.InvariantCulture) + " >>", "stream", pageContent.TrimEnd('\n'), "endstream", "endobj",
            "6 0 obj", "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Length 3 >>", "stream", "abc", "endstream", "endobj",
            "7 0 obj", "<< /Type /Font /Subtype /Type3 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 8 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /XObject << /ImGlyph 6 0 R >> >> >>", "endobj",
            "8 0 obj", "<< /Length " + Encoding.ASCII.GetByteCount(glyphContent).ToString(CultureInfo.InvariantCulture) + " >>", "stream", glyphContent.TrimEnd('\n'), "endstream", "endobj",
            "trailer", "<< /Root 1 0 R /Size 9 >>", "%%EOF", string.Empty
        }));
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

    private static byte[] BuildCollidingFormImagePdf(bool markSecondImage = false) {
        const string pageContent = "q /F1 Do Q\nq /F2 Do Q\n";
        const string formContent = "q 40 0 0 20 20 30 cm /Im0 Do Q\n";
        string secondFormContent = markSecondImage ? "/Figure BMC " + formContent.TrimEnd('\n') + " EMC\n" : formContent;
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj", "<< /Type /Catalog /Pages 2 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 120] /Resources << /XObject << /F1 6 0 R /F2 7 0 R >> >> /Contents 5 0 R >>", "endobj",
            "5 0 obj", "<< /Length " + Encoding.ASCII.GetByteCount(pageContent).ToString(CultureInfo.InvariantCulture) + " >>", "stream", pageContent.TrimEnd('\n'), "endstream", "endobj",
            "6 0 obj", "<< /Type /XObject /Subtype /Form /BBox [0 0 200 120] /Resources << /XObject << /Im0 8 0 R >> >> /Length " + Encoding.ASCII.GetByteCount(formContent).ToString(CultureInfo.InvariantCulture) + " >>", "stream", formContent.TrimEnd('\n'), "endstream", "endobj",
            "7 0 obj", "<< /Type /XObject /Subtype /Form /BBox [0 0 200 120] /Resources << /XObject << /Im0 9 0 R >> >> /Length " + Encoding.ASCII.GetByteCount(secondFormContent).ToString(CultureInfo.InvariantCulture) + " >>", "stream", secondFormContent.TrimEnd('\n'), "endstream", "endobj",
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

    private static byte[] BuildSplitMarkedContentImagePdf() {
        const string first = "/Figure << /MCID 0 >> BDC\n";
        const string second = "q 40 0 0 20 20 30 cm /Im0 Do Q EMC\n";
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

    private static byte[] BuildMarkedFormImagePdf(string formDictionaryEntries = "") {
        const string pageContent = "/Figure << /MCID 0 >> BDC /Fx Do EMC\n";
        const string formContent = "q 40 0 0 20 20 30 cm /Im0 Do Q\n";
        string pdf = string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj", "<< /Type /Catalog /Pages 2 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 120] /Resources << /XObject << /Fx 6 0 R >> >> /Contents 4 0 R >>", "endobj",
            "4 0 obj", "<< /Length " + Encoding.ASCII.GetByteCount(pageContent).ToString(CultureInfo.InvariantCulture) + " >>", "stream", pageContent.TrimEnd('\n'), "endstream", "endobj",
            "6 0 obj", "<< /Type /XObject /Subtype /Form /BBox [0 0 200 120] " + formDictionaryEntries + " /Resources << /XObject << /Im0 7 0 R >> >> /Length " + Encoding.ASCII.GetByteCount(formContent).ToString(CultureInfo.InvariantCulture) + " >>", "stream", formContent.TrimEnd('\n'), "endstream", "endobj",
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

    private static byte[] BuildSharedOuterFormImagePdf() {
        const string firstPageContent = "/Outer Do\n";
        const string secondPageContent = "/Outer Do\n";
        const string outerContent = "/Inner Do\n";
        const string innerContent = "q 40 0 0 20 20 30 cm /Im0 Do Q\n";
        string pdf = string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj", "<< /Type /Catalog /Pages 2 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 2 /Kids [3 0 R 4 0 R] /MediaBox [0 0 200 120] /Resources << /XObject << /Outer 7 0 R >> >> >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /Contents 5 0 R >>", "endobj",
            "4 0 obj", "<< /Type /Page /Parent 2 0 R /Contents 6 0 R >>", "endobj",
            "5 0 obj", "<< /Length " + Encoding.ASCII.GetByteCount(firstPageContent).ToString(CultureInfo.InvariantCulture) + " >>", "stream", firstPageContent.TrimEnd('\n'), "endstream", "endobj",
            "6 0 obj", "<< /Length " + Encoding.ASCII.GetByteCount(secondPageContent).ToString(CultureInfo.InvariantCulture) + " >>", "stream", secondPageContent.TrimEnd('\n'), "endstream", "endobj",
            "7 0 obj", "<< /Type /XObject /Subtype /Form /BBox [0 0 200 120] /Resources << /XObject << /Inner 8 0 R >> >> /Length " + Encoding.ASCII.GetByteCount(outerContent).ToString(CultureInfo.InvariantCulture) + " >>", "stream", outerContent.TrimEnd('\n'), "endstream", "endobj",
            "8 0 obj", "<< /Type /XObject /Subtype /Form /BBox [0 0 200 120] /Resources << /XObject << /Im0 9 0 R >> >> /Length " + Encoding.ASCII.GetByteCount(innerContent).ToString(CultureInfo.InvariantCulture) + " >>", "stream", innerContent.TrimEnd('\n'), "endstream", "endobj",
            "9 0 obj", "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Length 3 >>", "stream", "abc", "endstream", "endobj",
            "trailer", "<< /Root 1 0 R /Size 10 >>", "%%EOF"
        }) + "\n";
        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildSharedResourceLessFormImagePdf() {
        const string pageContent = "/Fm Do\n";
        const string formContent = "q 40 0 0 20 20 30 cm /Im Do Q\n";
        string pdf = string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj", "<< /Type /Catalog /Pages 2 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 2 /Kids [3 0 R 4 0 R] /MediaBox [0 0 200 120] /Resources << /XObject << /Fm 7 0 R /Im 8 0 R >> >> >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /Contents 5 0 R >>", "endobj",
            "4 0 obj", "<< /Type /Page /Parent 2 0 R /Contents 6 0 R >>", "endobj",
            "5 0 obj", "<< /Length " + Encoding.ASCII.GetByteCount(pageContent).ToString(CultureInfo.InvariantCulture) + " >>", "stream", pageContent.TrimEnd('\n'), "endstream", "endobj",
            "6 0 obj", "<< /Length " + Encoding.ASCII.GetByteCount(pageContent).ToString(CultureInfo.InvariantCulture) + " >>", "stream", pageContent.TrimEnd('\n'), "endstream", "endobj",
            "7 0 obj", "<< /Type /XObject /Subtype /Form /BBox [0 0 200 120] /Length " + Encoding.ASCII.GetByteCount(formContent).ToString(CultureInfo.InvariantCulture) + " >>", "stream", formContent.TrimEnd('\n'), "endstream", "endobj",
            "8 0 obj", "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Length 3 >>", "stream", "abc", "endstream", "endobj",
            "trailer", "<< /Root 1 0 R /Size 9 >>", "%%EOF"
        }) + "\n";
        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildDirectAndInheritedFormImagePdf() {
        const string pageContent = "q 40 0 0 20 20 30 cm /Im Do Q /Fm Do\n";
        const string formContent = "q 40 0 0 20 100 30 cm /Im Do Q\n";
        string pdf = string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj", "<< /Type /Catalog /Pages 2 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 200 120] /Resources << /XObject << /Fm 6 0 R /Im 7 0 R >> >> >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /Contents 4 0 R >>", "endobj",
            "4 0 obj", "<< /Length " + Encoding.ASCII.GetByteCount(pageContent).ToString(CultureInfo.InvariantCulture) + " >>", "stream", pageContent.TrimEnd('\n'), "endstream", "endobj",
            "6 0 obj", "<< /Type /XObject /Subtype /Form /BBox [0 0 200 120] /Length " + Encoding.ASCII.GetByteCount(formContent).ToString(CultureInfo.InvariantCulture) + " >>", "stream", formContent.TrimEnd('\n'), "endstream", "endobj",
            "7 0 obj", "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Length 3 >>", "stream", "abc", "endstream", "endobj",
            "trailer", "<< /Root 1 0 R /Size 8 >>", "%%EOF"
        }) + "\n";
        return Encoding.ASCII.GetBytes(pdf);
    }

    private static void WriteAscii(Stream stream, string value) {
        byte[] bytes = Encoding.ASCII.GetBytes(value);
        stream.Write(bytes, 0, bytes.Length);
    }
}
