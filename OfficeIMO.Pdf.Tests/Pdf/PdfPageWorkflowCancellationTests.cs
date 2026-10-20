using OfficeIMO.Pdf;
using System.Threading;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfPageWorkflowCancellationTests {
    [Fact]
    public void CanceledPageEditsDoNotChangeTheSource() {
        byte[] source = CreateSource();
        PdfDocument document = PdfDocument.Load(source);
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        CancellationToken token = cancellation.Token;
        PdfPageSelection selection = PdfPageSelection.From(1);
        PdfPageSelector selector = PdfPageSelector.Parse("1");
        Action[] edits = {
            () => document.Pages.Delete(token, 1),
            () => document.Pages.Delete(selection, token),
            () => document.Pages.Delete(selector, token),
            () => document.Pages.Delete("1", token),
            () => document.Pages.Reorder(token, 2, 1),
            () => document.Pages.Duplicate(selection, token),
            () => document.Pages.Move(3, selector, token),
            () => document.Pages.Rotate(90, "1", token),
            () => document.Pages.RotateBy(90, selection, token),
            () => document.Pages.SetCropBox(20, 20, 400, 700, token, 1),
            () => document.Pages.Resize(new PdfPageResizeOptions(PageSizes.Letter) { CancellationToken = token }, 1),
            () => document.Pages.CropAndTranslate(20, 20, 400, 700, token, 1),
            () => document.Pages.DestructiveCrop(20, 20, 400, 700, new PdfDestructiveCropOptions { CancellationToken = token }, 1),
            () => document.Pages.ApplyAttributes(token, PdfPageAttributeEdit.RotateBy(90, selection)),
            () => document.Pages.ReorderWithMapping(token, 2, 1),
            () => document.Pages.BookmarkPageRanges(token),
            () => document.Pages.SplitByBookmarks(token),
            () => document.Bookmarks.Validate(token),
            () => document.Bookmarks.Edit(_ => throw new InvalidOperationException("Canceled edit must not run."), token),
            () => document.Pages.SplitForProductionResult(new PdfProductionSplitOptions { MaximumPagesPerPart = 1, CancellationToken = token })
        };

        foreach (Action edit in edits) {
            Assert.Throws<OperationCanceledException>(edit);
            Assert.Equal(source, document.ToBytes());
        }
    }

    [Fact]
    public void CanceledPageResultRoutesThrowInsteadOfReturningFailureReports() {
        byte[] source = CreateSource();
        PdfDocument document = PdfDocument.Load(source);
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        CancellationToken token = cancellation.Token;
        PdfPageSelection selection = PdfPageSelection.From(1);
        PdfPageSelector selector = PdfPageSelector.Parse("last");
        Action[] edits = {
            () => document.Pages.DeleteResult(selection, token),
            () => document.Pages.DeleteResult(selector, token),
            () => document.Pages.DeleteResult("1", token),
            () => document.Pages.ReorderResult(selection, token),
            () => document.Pages.DuplicateResult(selector, token),
            () => document.Pages.MoveResult(3, "1", token),
            () => document.Pages.RotateResult(90, selection, token),
            () => document.Pages.RotateByResult(90, selector, token),
            () => document.Pages.SetPageBoxResult(PdfPageBoundaryBox.CropBox, 20, 20, 400, 700, selection, token),
            () => document.Pages.ApplyAttributesResult(new[] { PdfPageAttributeEdit.RotateBy(90, selection) }, cancellationToken: token),
            () => document.Pages.ExtractResult(selection, token),
            () => document.Pages.ExtractResult(selector, token),
            () => document.Pages.ExtractResult("1", token),
            () => document.Pages.SplitResult(token),
            () => document.Pages.SplitResult(1, token),
            () => document.Pages.SplitResult(new[] { selection }, token),
            () => document.Pages.SplitResult(new[] { PdfPageRange.From(1, 1) }, token),
            () => document.Pages.SplitResult("1", token),
            () => document.Pages.SplitByBookmarksResult(token)
        };

        foreach (Action edit in edits) {
            Assert.Throws<OperationCanceledException>(edit);
            Assert.Equal(source, document.ToBytes());
        }
    }

    [Fact]
    public void SplitResultPropagatesCancellationRaisedDuringRangeEnumeration() {
        byte[] source = CreateSource();
        PdfDocument document = PdfDocument.Load(source);
        using var cancellation = new CancellationTokenSource();
        IEnumerable<PdfPageRange> Ranges() {
            yield return PdfPageRange.From(1, 1);
            cancellation.Cancel();
            yield return PdfPageRange.From(2, 2);
        }

        Assert.Throws<OperationCanceledException>(() => document.Pages.SplitResult(Ranges(), cancellation.Token));
        Assert.Equal(source, document.ToBytes());
    }

    [Fact]
    public void BookmarkEditCancellationAfterTheCallbackDoesNotReturnOrAdoptPartialEdits() {
        byte[] source = CreateSource();
        PdfDocument document = PdfDocument.Load(source);
        using var cancellation = new CancellationTokenSource();
        Assert.Throws<OperationCanceledException>(() => document.Bookmarks.Edit(session => {
            session.Rename(session.Roots[0].Id, "Canceled title");
            cancellation.Cancel();
        }, cancellation.Token));
        Assert.Equal(source, document.ToBytes());
        Assert.Equal("First section", document.Inspect().Outlines[0].Title);

        PdfBookmarkEditResult edited = document.Bookmarks.Edit(session => session.Rename(session.Roots[0].Id, "Saved title"), CancellationToken.None);
        Assert.Equal("Saved title", edited.Outlines[0].Title);
        Assert.Empty(edited.ToDocument().Bookmarks.Validate(CancellationToken.None));
    }

    [Fact]
    public void ResizeHonorsBothItsOptionTokenAndTheExplicitOwnerToken() {
        byte[] source = CreateSource();
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        var options = new PdfPageResizeOptions(PageSizes.Letter) { CancellationToken = cancellation.Token };
        Assert.Throws<OperationCanceledException>(() => PdfPageEditor.ResizePages(source, options, null, CancellationToken.None));
        options.CancellationToken = CancellationToken.None;
        Assert.Throws<OperationCanceledException>(() => PdfPageEditor.ResizePages(source, options, null, cancellation.Token));
    }

    [Fact]
    public void ProductionSplittingKeepsPageCoverageWithCancellationEnabled() {
        byte[] source = CreateSource();
        using var cancellation = new CancellationTokenSource();
        PdfDocument document = PdfDocument.Load(source);
        PdfProductionSplitResult result = document.Pages.SplitForProductionResult(new PdfProductionSplitOptions {
            TargetPartSizeBytes = 1,
            BoundaryText = "section",
            CancellationToken = cancellation.Token
        });
        Assert.Equal(new[] { 1, 2 }, result.Parts.SelectMany(part => part.SourcePages));
        Assert.All(result.Parts, part => Assert.Equal(1, part.ToDocument().Inspect().PageCount));
        Assert.Equal(source, document.ToBytes());
    }

    [Fact]
    public void PageUtilitiesUseTheRemainingStreamAndKeepCallerStreamsOpen() {
        byte[] source = CreateSource();
        byte[] prefixed = new byte[source.Length + 7];
        Buffer.BlockCopy(source, 0, prefixed, 7, source.Length);
        var options = new PdfLoadOptions { Limits = new PdfReadLimits { MaxInputBytes = source.Length } };
        foreach (var route in StreamRoutes(options)) {
            foreach (bool seekable in new[] { true, false }) {
                using var input = new ControlledInput(prefixed, 7, seekable);
                using var output = new MemoryStream();
                route(input, output, CancellationToken.None);
                Assert.True(input.CanRead);
                Assert.True(output.CanWrite);
                Assert.Equal(prefixed.Length, input.BytesConsumed);
                Assert.InRange(PdfDocument.Load(output.ToArray()).Inspect().PageCount, 1, 3);
            }
        }
    }

    [Fact]
    public void PageUtilitiesRejectOversizedInputBeforeChangingAnyDestination() {
        byte[] source = CreateSource();
        var options = new PdfLoadOptions { Limits = new PdfReadLimits { MaxInputBytes = source.Length - 1 } };
        foreach (var route in StreamRoutes(options)) {
            foreach (bool seekable in new[] { true, false }) {
                using var input = new ControlledInput(source, 0, seekable);
                using var output = new MemoryStream();
                output.WriteByte(42);
                Assert.Throws<PdfReadLimitException>(() => route(input, output, CancellationToken.None));
                Assert.Equal(new byte[] { 42 }, output.ToArray());
                Assert.True(input.CanRead);
                if (seekable) Assert.Equal(0, input.BytesConsumed);
            }
        }
    }

    [Fact]
    public void PageUtilitiesCancelDuringInputAndAfterTheLastOutputWrite() {
        byte[] source = CreateSource();
        foreach (var route in StreamRoutes(null)) {
            using var readingCancellation = new CancellationTokenSource();
            using var input = new ControlledInput(source, 0, false, readingCancellation);
            using var output = new MemoryStream();
            output.WriteByte(42);
            Assert.Throws<OperationCanceledException>(() => route(input, output, readingCancellation.Token));
            Assert.Equal(new byte[] { 42 }, output.ToArray());
            Assert.True(input.CanRead);

            using var writingCancellation = new CancellationTokenSource();
            using var completeInput = new MemoryStream(source);
            using var cancelingOutput = new CancelingOutput(writingCancellation);
            Assert.Throws<OperationCanceledException>(() => route(completeInput, cancelingOutput, writingCancellation.Token));
            Assert.True(cancelingOutput.CanWrite);
        }
    }

    [Fact]
    public void CanceledPageFileOperationsDoNotOpenMissingInputOrReplaceExistingOutput() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-page-cancel-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        string input = Path.Combine(directory, "missing.pdf");
        string output = Path.Combine(directory, "existing.pdf");
        try {
            File.WriteAllBytes(output, new byte[] { 42 });
            using var cancellation = new CancellationTokenSource();
            cancellation.Cancel();
            Action[] routes = {
                () => PdfPageEditor.DeletePages(input, output, null, cancellation.Token, 1),
                () => PdfPageEditor.ReorderPages(input, output, null, cancellation.Token, 2, 1),
                () => PdfPageEditor.DuplicatePages(input, output, null, cancellation.Token, 1),
                () => PdfPageEditor.MovePages(input, output, 3, null, cancellation.Token, 1),
                () => PdfPageEditor.RotatePages(input, output, 90, null, cancellation.Token, 1)
            };
            foreach (Action route in routes) {
                Assert.Throws<OperationCanceledException>(route);
                Assert.Equal(new byte[] { 42 }, File.ReadAllBytes(output));
            }
        } finally {
            Directory.Delete(directory, true);
        }
    }

    private static Action<Stream, Stream, CancellationToken>[] StreamRoutes(PdfLoadOptions? options) => new Action<Stream, Stream, CancellationToken>[] {
        (input, output, token) => PdfPageEditor.DeletePages(input, output, options, token, 1),
        (input, output, token) => PdfPageEditor.ReorderPages(input, output, options, token, 2, 1),
        (input, output, token) => PdfPageEditor.DuplicatePages(input, output, options, token, 1),
        (input, output, token) => PdfPageEditor.MovePages(input, output, 3, options, token, 1),
        (input, output, token) => PdfPageEditor.RotatePages(input, output, 90, options, token, 1)
    };

    private static byte[] CreateSource() => PdfDocument.Create(new PdfOptions { CreateOutlineFromHeadings = true })
        .H1("First section").Paragraph(p => p.Text("First content"))
        .PageBreak().H1("Second section").Paragraph(p => p.Text("Second content")).ToBytes();

    private sealed class ControlledInput : Stream {
        private readonly byte[] _bytes;
        private readonly bool _seekable;
        private readonly CancellationTokenSource? _cancellation;
        private int _position;
        public ControlledInput(byte[] bytes, int position, bool seekable, CancellationTokenSource? cancellation = null) {
            _bytes = bytes; _position = position; _seekable = seekable; _cancellation = cancellation;
        }
        public int BytesConsumed => _position;
        public override bool CanRead => true;
        public override bool CanSeek => _seekable;
        public override bool CanWrite => false;
        public override long Length => _seekable ? _bytes.Length : throw new NotSupportedException();
        public override long Position { get => _seekable ? _position : throw new NotSupportedException(); set => throw new NotSupportedException(); }
        public override int Read(byte[] buffer, int offset, int count) {
            int length = Math.Min(Math.Min(count, 113), _bytes.Length - _position);
            Buffer.BlockCopy(_bytes, _position, buffer, offset, length);
            _position += length;
            if (length > 0) _cancellation?.Cancel();
            return length;
        }
        public override void Flush() { }
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    }

    private sealed class CancelingOutput : MemoryStream {
        private readonly CancellationTokenSource _cancellation;
        public CancelingOutput(CancellationTokenSource cancellation) => _cancellation = cancellation;
        public override void Write(byte[] buffer, int offset, int count) {
            base.Write(buffer, offset, count);
            _cancellation.Cancel();
        }
    }
}
