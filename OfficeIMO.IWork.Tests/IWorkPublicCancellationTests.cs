using System.Threading;
using OfficeIMO.IWork;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Theory]
    [InlineData(IWorkDocumentKind.Pages, "nim-iwork/simple.pages")]
    [InlineData(IWorkDocumentKind.Numbers, "nim-iwork/simple.numbers")]
    [InlineData(IWorkDocumentKind.Keynote, "nim-iwork/simple.key")]
    public void Public_conversion_cancels_during_package_read_and_keeps_caller_stream_open(
        IWorkDocumentKind kind, string fixture) {
        using var cancellation = new CancellationTokenSource();
        using var input = new CancellingReadStream(File.ReadAllBytes(Fixture(fixture)), cancellation);

        Assert.Throws<OperationCanceledException>(() => ConvertForCancellation(input, kind, cancellation.Token));
        Assert.True(input.CanRead);
    }

    [Theory]
    [InlineData(IWorkDocumentKind.Pages, "nim-iwork/simple.pages")]
    [InlineData(IWorkDocumentKind.Numbers, "nim-iwork/simple.numbers")]
    [InlineData(IWorkDocumentKind.Keynote, "nim-iwork/simple.key")]
    public void Opened_source_cancellation_governs_later_projection_and_conversion(
        IWorkDocumentKind kind, string fixture) {
        using var cancellation = new CancellationTokenSource();
        IWorkSourceDocument source = IWorkSourceDocument.Open(Fixture(fixture), kind, null, cancellation.Token);
        cancellation.Cancel();

        Assert.Throws<OperationCanceledException>(() => Project(source));
        Assert.Throws<OperationCanceledException>(() => ConvertForCancellation(source));
    }

    [Fact]
    public void Cancelled_byte_input_is_rejected_before_projection() {
        byte[] bytes = File.ReadAllBytes(Fixture("nim-iwork/simple.pages"));
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => IWorkSourceDocument.Open(bytes, null, cancellation.Token));
        Assert.Throws<OperationCanceledException>(() => IWorkSourceDocument.Open(bytes,
            IWorkDocumentKind.Pages, null, cancellation.Token));
    }

    [Theory]
    [InlineData(IWorkDocumentKind.Pages, "nim-iwork/simple.pages")]
    [InlineData(IWorkDocumentKind.Numbers, "nim-iwork/simple.numbers")]
    [InlineData(IWorkDocumentKind.Keynote, "nim-iwork/simple.key")]
    public void Reusing_loaded_source_with_an_independent_token_preserves_other_operations(
        IWorkDocumentKind kind, string fixture) {
        using var opening = new CancellationTokenSource();
        IWorkSourceDocument source = IWorkSourceDocument.Open(Fixture(fixture), kind, null, opening.Token);
        Project(source); // Exercise reuse after the shared parsed-message cache has been populated.
        using var operation = new CancellationTokenSource();
        IWorkSourceDocument view = source.WithCancellation(operation.Token);
        operation.Cancel();
        Assert.Throws<OperationCanceledException>(() => Project(view));
        Assert.Throws<OperationCanceledException>(() => ConvertForCancellation(view));
        Project(source);
        opening.Cancel();
        Assert.Throws<OperationCanceledException>(() => Project(source));
        Project(source.WithCancellation(CancellationToken.None));
        Assert.Throws<OperationCanceledException>(() => source.WithCancellation(operation.Token));
    }

    private static void Project(IWorkSourceDocument source) {
        switch (source.Kind) {
            case IWorkDocumentKind.Pages: source.ReadPages(); break;
            case IWorkDocumentKind.Numbers: source.ReadNumbers(); break;
            case IWorkDocumentKind.Keynote: source.ReadKeynote(); break;
        }
    }

    private static void ConvertForCancellation(IWorkSourceDocument source) {
        using IDisposable result = source.Kind switch {
            IWorkDocumentKind.Pages => WordIWorkConverter.ToWordDocumentResult(source),
            IWorkDocumentKind.Numbers => ExcelIWorkConverter.ToExcelDocumentResult(source),
            IWorkDocumentKind.Keynote => PowerPointIWorkConverter.ToPowerPointPresentationResult(source),
            _ => throw new ArgumentOutOfRangeException()
        };
    }

    private static void ConvertForCancellation(Stream input, IWorkDocumentKind kind, CancellationToken token) {
        using IDisposable result = kind switch {
            IWorkDocumentKind.Pages => WordIWorkConverter.ConvertPagesToWordResult(input, null, null, token),
            IWorkDocumentKind.Numbers => ExcelIWorkConverter.ConvertNumbersToExcelResult(input, null, null, token),
            IWorkDocumentKind.Keynote => PowerPointIWorkConverter.ConvertKeynoteToPowerPointResult(input, null, null, token),
            _ => throw new ArgumentOutOfRangeException()
        };
    }
}
