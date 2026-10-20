using System.Threading;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfEditCancellationTests {
    [Fact]
    public void StampOptionsCancelPublicAndResultRoutesBeforeReadingAdditionalInputs() {
        byte[] source = CreatePdf();
        PdfDocument target = PdfDocument.Load(source);
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        CancellationToken token = cancellation.Token;
        var text = new PdfTextStampOptions { CancellationToken = token };
        var image = new PdfImageStampOptions { CancellationToken = token };
        var overlay = new PdfPageOverlayOptions { CancellationToken = token };
        var canvas = new PdfCanvasStampOptions { CancellationToken = token };
        using var imageInput = new TrackingInput(new byte[] { 1, 2, 3 });
        using var pdfInput = new TrackingInput(source);
        int callbacks = 0;
        Action[] routes = {
            () => target.Stamp.Text("Draft", text),
            () => target.Stamp.TextResult("Draft", text),
            () => target.Stamp.TextWatermark("Draft", text),
            () => target.Stamp.TextWatermarkResult("Draft", text),
            () => target.Stamp.Image(imageInput, image),
            () => target.Stamp.ImageResult(imageInput, image),
            () => target.Stamp.ImageWatermark(imageInput, image),
            () => target.Stamp.ImageWatermarkResult(imageInput, image),
            () => target.Stamp.OverlayPage(pdfInput, overlay),
            () => target.Stamp.OverlayPageResult(pdfInput, overlay),
            () => target.Stamp.UnderlayPage(pdfInput, overlay),
            () => target.Stamp.UnderlayPageResult(pdfInput, overlay),
            () => target.Stamp.Content(_ => callbacks++, canvas),
            () => target.Stamp.ContentResult((_, _) => callbacks++, canvas),
            () => target.Stamp.Watermark(new PdfWatermarkOptions { CancellationToken = token }),
            () => target.Stamp.ReadWatermarks(null, token)
        };
        foreach (Action route in routes) {
            OperationCanceledException error = Assert.ThrowsAny<OperationCanceledException>(route);
            Assert.Equal(token, error.CancellationToken);
        }
        Assert.Equal(0, imageInput.BytesRead);
        Assert.Equal(0, pdfInput.BytesRead);
        Assert.Equal(0, callbacks);
        Assert.Equal(source, target.ToBytes());
    }

    [Theory]
    [InlineData("text")]
    [InlineData("text-watermark")]
    [InlineData("image")]
    [InlineData("image-watermark")]
    public void StampInputCancellationLeavesDestinationUntouchedAndCallerStreamsOpen(string operation) {
        byte[] pdf = CreatePdf();
        using var cancellation = new CancellationTokenSource();
        using var input = new TrackingInput(pdf, cancellation.Cancel);
        using var output = new MemoryStream();
        output.WriteByte(42);
        var text = new PdfTextStampOptions { CancellationToken = cancellation.Token };
        var image = new PdfImageStampOptions { CancellationToken = cancellation.Token };
        byte[] png = PdfPngTestImages.CreateRgbPng(255, 0, 0);

        Assert.ThrowsAny<OperationCanceledException>(() => {
            if (operation == "text") PdfStamper.StampText(input, output, "Draft", text);
            else if (operation == "text-watermark") PdfStamper.WatermarkText(input, output, "Draft", text);
            else if (operation == "image") PdfStamper.StampImage(input, output, png, image);
            else PdfStamper.WatermarkImage(input, output, png, image);
        });

        Assert.InRange(input.BytesRead, 1, pdf.Length - 1);
        Assert.True(input.CanRead);
        Assert.True(output.CanWrite);
        Assert.Equal(new byte[] { 42 }, output.ToArray());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void OverlaySourcesHonorRemainingInputLimitsAndCancellation(bool underlay) {
        byte[] source = CreatePdf("Imported content");
        byte[] targetBytes = CreatePdf();
        PdfDocument target = PdfDocument.Load(targetBytes);
        var exact = new PdfLoadOptions { Limits = new PdfReadLimits { MaxInputBytes = source.Length } };
        var options = new PdfPageOverlayOptions { SourceReadOptions = exact };
        byte[] prefixed = new byte[source.Length + 3];
        Array.Copy(source, 0, prefixed, 3, source.Length);
        using (var input = new TrackingInput(prefixed, start: 3)) {
            PdfDocument result = underlay ? target.Stamp.UnderlayPage(input, options) : target.Stamp.OverlayPage(input, options);
            Assert.Contains("Imported content", result.Reader.Text(), StringComparison.Ordinal);
            Assert.Contains("Existing content", result.Reader.Text(), StringComparison.Ordinal);
            Assert.Equal(source.Length, input.BytesRead);
            Assert.True(input.CanRead);
        }
        options.SourceReadOptions = new PdfLoadOptions { Limits = new PdfReadLimits { MaxInputBytes = source.Length - 1 } };
        using (var input = new TrackingInput(prefixed, start: 3)) {
            Assert.Throws<PdfReadLimitException>(() => {
                if (underlay) target.Stamp.UnderlayPage(input, options); else target.Stamp.OverlayPage(input, options);
            });
            Assert.True(input.CanRead);
        }
        using var cancellation = new CancellationTokenSource();
        using var cancelledInput = new TrackingInput(source, cancellation.Cancel);
        options.SourceReadOptions = exact;
        options.CancellationToken = cancellation.Token;
        Assert.ThrowsAny<OperationCanceledException>(() => {
            if (underlay) target.Stamp.UnderlayPage(cancelledInput, options); else target.Stamp.OverlayPage(cancelledInput, options);
        });
        Assert.InRange(cancelledInput.BytesRead, 1, source.Length - 1);
        Assert.True(cancelledInput.CanRead);
        Assert.Equal(targetBytes, target.ToBytes());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void CancellationInsideCanvasCallbackStopsBeforeNextPageAndLeavesSourceUnchanged(bool resultRoute) {
        byte[] source = PdfDocument.Create().Compose(compose => {
            compose.Page(page => page.Content(content => content.Paragraph(paragraph => paragraph.Text("First"))));
            compose.Page(page => page.Content(content => content.Paragraph(paragraph => paragraph.Text("Second"))));
        }).ToBytes();
        PdfDocument target = PdfDocument.Load(source);
        using var cancellation = new CancellationTokenSource();
        var options = new PdfCanvasStampOptions { CancellationToken = cancellation.Token };
        int callbacks = 0;
        void Build(PdfPageCanvas canvas, PdfStampPageContext _) {
            callbacks++;
            canvas.Text("Draft", 20D, 20D, 100D, 25D);
            cancellation.Cancel();
        }
        Assert.ThrowsAny<OperationCanceledException>(() => {
            if (resultRoute) target.Stamp.ContentResult(Build, options); else target.Stamp.Content(Build, options);
        });
        Assert.Equal(1, callbacks);
        Assert.Equal(source, target.ToBytes());
    }

    [Fact]
    public void AnnotationOptionAndExplicitTokenRoutesCancelWithoutChangingSource() {
        PdfDocument original = PdfDocument.Load(CreatePdf());
        PdfDocument target = original.Annotations.Add(new PdfAnnotationCreateOptions { Contents = "Review" }).ToDocument();
        byte[] source = target.ToBytes();
        int number = Assert.Single(target.Annotations.GetForEditing()).ObjectNumber!.Value;
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        CancellationToken token = cancellation.Token;
        Action[] routes = {
            () => target.Annotations.Add(new PdfAnnotationCreateOptions { CancellationToken = token }),
            () => target.Annotations.AddStamp(new PdfStampAnnotationOptions { CancellationToken = token }),
            () => target.Annotations.Update(number, new PdfAnnotationUpdateOptions { Contents = "Changed", CancellationToken = token }),
            () => target.Annotations.Remove(new PdfAnnotationRemovalOptions { CancellationToken = token }),
            () => target.Annotations.Flatten(new PdfAnnotationFlattenOptions { CancellationToken = token }),
            () => target.Annotations.AddReply(number, "Reply", new PdfAnnotationReplyOptions { CancellationToken = token }),
            () => target.Annotations.Move(number, 5D, 5D, token),
            () => target.Annotations.Resize(number, new PdfPageRectangle(10D, 10D, 60D, 60D), token),
            () => target.Annotations.MoveMany(new[] { number }, 5D, 5D, token),
            () => target.Annotations.MoveManyVisual(new[] { number }, 5D, 5D, token),
            () => target.Annotations.CopyMany(new[] { number }, 5D, 5D, token),
            () => target.Annotations.CopyManyVisual(new[] { number }, 5D, 5D, token),
            () => target.Annotations.Group(new[] { number }, token),
            () => target.Annotations.Ungroup(new[] { number }, token),
            () => target.Annotations.Arrange(new[] { number }, PdfAnnotationOrderChange.Raise, token),
            () => target.Annotations.RemoveMany(new[] { number }, false, token),
            () => target.Annotations.SetReviewState(number, PdfAnnotationReviewState.Accepted, PdfMutationExecutionPreference.Automatic, false, token),
            () => target.Annotations.GetForEditing(token),
            () => target.Annotations.GetEditingInteractions(1, token),
            () => target.Annotations.GetReviewCatalog(token),
            () => target.FlattenVisualAnnotations(token),
            () => target.FlattenVisualAnnotationsResult(null, token)
        };
        foreach (Action route in routes) {
            OperationCanceledException error = Assert.ThrowsAny<OperationCanceledException>(route);
            Assert.Equal(token, error.CancellationToken);
        }
        Assert.Equal(source, target.ToBytes());
    }

    [Fact]
    public void CancelledFileEditsPreserveExistingDestination() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-edit-cancel-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try {
            string input = Path.Combine(directory, "input.pdf"), output = Path.Combine(directory, "output.pdf");
            File.WriteAllBytes(input, CreatePdf());
            byte[] sentinel = new byte[] { 42, 43, 44 };
            File.WriteAllBytes(output, sentinel);
            using var cancellation = new CancellationTokenSource();
            cancellation.Cancel();
            Assert.ThrowsAny<OperationCanceledException>(() => PdfStamper.StampText(input, output, "Draft", new PdfTextStampOptions { CancellationToken = cancellation.Token }));
            Assert.Equal(sentinel, File.ReadAllBytes(output));
            Assert.ThrowsAny<OperationCanceledException>(() => PdfAnnotationEditor.RemoveAnnotations(input, output, new PdfAnnotationRemovalOptions { CancellationToken = cancellation.Token }));
            Assert.Equal(sentinel, File.ReadAllBytes(output));
            Assert.Equal(2, Directory.GetFiles(directory).Length);
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void InvalidOutputDoesNotConsumeStampInput() {
        using var input = new TrackingInput(CreatePdf());
        using var output = new MemoryStream(new byte[] { 42 }, writable: false);
        Assert.Throws<ArgumentException>(() => PdfStamper.StampText(input, output, "Draft"));
        Assert.Equal(0, input.BytesRead);
        Assert.True(input.CanRead);
    }

    [Fact]
    public void CancellationDuringOutputWriteIsObservedWithoutClosingCallerStream() {
        using var cancellation = new CancellationTokenSource();
        using var output = new CancellingOutput(cancellation);
        Assert.ThrowsAny<OperationCanceledException>(() => PdfStamper.StampText(CreatePdf(), output, "Draft",
            new PdfTextStampOptions { CancellationToken = cancellation.Token }));
        Assert.True(output.Length > 0);
        Assert.True(output.CanWrite);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void BatesCancellationStopsBeforeRenderingAndLeavesInputsUnchanged(bool cancelDuringEnumeration) {
        byte[] source = CreatePdf();
        var input = new PdfBatesDocument(source);
        using var cancellation = new CancellationTokenSource();
        int enumerated = 0;
        IEnumerable<PdfBatesDocument> Inputs() {
            enumerated++;
            yield return input;
            cancellation.Cancel();
            enumerated++;
            yield return input;
        }
        if (!cancelDuringEnumeration) cancellation.Cancel();
        Assert.ThrowsAny<OperationCanceledException>(() => PdfBatesNumberer.Apply(Inputs(),
            new PdfBatesNumberingOptions { CancellationToken = cancellation.Token }));
        Assert.Equal(cancelDuringEnumeration ? 2 : 0, enumerated);
        Assert.Equal(source, input.GetBytes());
    }

    private static byte[] CreatePdf(string text = "Existing content") => PdfDocument.Create().Paragraph(paragraph => paragraph.Text(text)).ToBytes();

    private sealed class TrackingInput : MemoryStream {
        private readonly Action? _afterRead;
        internal TrackingInput(byte[] bytes, Action? afterRead = null, int start = 0) : base(bytes) { _afterRead = afterRead; Position = start; }
        internal int BytesRead { get; private set; }
        public override bool CanSeek => false;
        public override int Read(byte[] buffer, int offset, int count) {
            int read = base.Read(buffer, offset, Math.Min(count, 7));
            BytesRead += read;
            if (read > 0) _afterRead?.Invoke();
            return read;
        }
    }

    private sealed class CancellingOutput : MemoryStream {
        private readonly CancellationTokenSource _cancellation;
        internal CancellingOutput(CancellationTokenSource cancellation) { _cancellation = cancellation; }
        public override void Write(byte[] buffer, int offset, int count) { base.Write(buffer, offset, count); _cancellation.Cancel(); }
    }
}
