using System.Threading;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfFormUtilityCancellationTests {
    private static readonly IReadOnlyDictionary<string, string> Values = new Dictionary<string, string> { ["Value"] = "QgjH" };

    [Theory]
    [InlineData("fill")]
    [InlineData("flatten")]
    [InlineData("fill-flatten")]
    public void DocumentFormOperationsCancelWithoutChangingSource(string operation) {
        byte[] original = PdfFormPlacementFixture.CreateAuthored("text");
        PdfDocument document = PdfDocument.Load(original);
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        var options = new PdfFormFillerOptions { CancellationToken = cancellation.Token };

        Assert.ThrowsAny<OperationCanceledException>(() => operation switch {
            "fill" => document.Forms.Fill(Values, options),
            "flatten" => document.Forms.Flatten(options),
            _ => document.Forms.FillAndFlatten(Values, options)
        });
        Assert.ThrowsAny<OperationCanceledException>(() => operation switch {
            "fill" => document.Forms.FillResult(Values, options, null),
            "flatten" => document.Forms.FlattenResult(options, null),
            _ => document.Forms.FillAndFlattenResult(Values, options, null)
        });
        Assert.Equal(original, document.ToBytes());
    }

    [Theory]
    [InlineData("fill", false)]
    [InlineData("flatten", false)]
    [InlineData("fill-flatten", false)]
    [InlineData("fill", true)]
    [InlineData("flatten", true)]
    [InlineData("fill-flatten", true)]
    public void StreamUtilitiesReadFromCurrentPositionAndKeepCallerOwnership(string operation, bool nonseekable) {
        byte[] source = PdfFormPlacementFixture.CreateAuthored("text");
        byte[] prefixed = new byte[source.Length + 7];
        Array.Copy(source, 0, prefixed, 7, source.Length);
        using var input = new TrackingInput(prefixed, nonseekable, start: 7);
        var readOptions = new PdfLoadOptions { Limits = new PdfReadLimits { MaxInputBytes = source.Length } };

        byte[] output = Run(operation, input, null, readOptions);

        Assert.True(input.CanRead);
        Assert.Equal(source.Length, input.BytesRead);
        if (operation == "fill") Assert.Equal("QgjH", Assert.Single(PdfDocument.Load(output).Forms.GetForInteraction().Fields).Value);
        else Assert.Empty(PdfDocument.Load(output).Forms.GetForInteraction().Fields);
    }

    [Theory]
    [InlineData("fill", false)]
    [InlineData("flatten", false)]
    [InlineData("fill-flatten", false)]
    [InlineData("fill", true)]
    [InlineData("flatten", true)]
    [InlineData("fill-flatten", true)]
    public void StreamUtilitiesEnforceCallerInputBudget(string operation, bool nonseekable) {
        byte[] source = PdfFormPlacementFixture.CreateAuthored("text");
        using var input = new TrackingInput(source, nonseekable);
        var options = new PdfLoadOptions { Limits = new PdfReadLimits { MaxInputBytes = source.Length - 1 } };

        Assert.Throws<PdfReadLimitException>(() => Run(operation, input, null, options));

        Assert.True(input.CanRead);
        Assert.InRange(input.BytesRead, 0, source.Length);
        if (!nonseekable) Assert.Equal(0, input.BytesRead);
    }

    [Theory]
    [InlineData("fill")]
    [InlineData("flatten")]
    [InlineData("fill-flatten")]
    public void StreamUtilityCancellationDuringAdmissionLeavesDestinationUntouched(string operation) {
        using var cancellation = new CancellationTokenSource();
        using var input = new TrackingInput(PdfFormPlacementFixture.CreateAuthored("text"), nonseekable: true,
            afterRead: cancellation.Cancel);
        using var output = new MemoryStream(new byte[] { 1, 2, 3 }, writable: true);
        var options = new PdfFormFillerOptions { CancellationToken = cancellation.Token };

        Assert.ThrowsAny<OperationCanceledException>(() => {
            if (operation == "fill") PdfFormFiller.FillFields(input, output, Values, options);
            else if (operation == "flatten") PdfFormFiller.FlattenFields(input, output, options);
            else PdfFormFiller.FillAndFlattenFields(input, output, Values, options);
        });

        Assert.True(input.BytesRead > 0);
        Assert.True(input.CanRead);
        Assert.True(output.CanWrite);
        Assert.Equal(new byte[] { 1, 2, 3 }, output.ToArray());
    }

    [Fact]
    public void StreamOutputObservesCancellationTriggeredByItsFinalWrite() {
        using var cancellation = new CancellationTokenSource();
        using var output = new CancellingOutput(cancellation);
        var options = new PdfFormFillerOptions { CancellationToken = cancellation.Token };
        byte[] source = PdfFormPlacementFixture.CreateAuthored("text");

        Assert.ThrowsAny<OperationCanceledException>(() => PdfFormFiller.FillFields(source, output, Values, options));

        Assert.True(output.Length > 0);
        Assert.True(output.CanWrite);
        Assert.Equal("HgjQ", Assert.Single(PdfDocument.Load(source).Forms.GetForInteraction().Fields).Value);
    }

    [Theory]
    [InlineData("fill")]
    [InlineData("flatten")]
    [InlineData("fill-flatten")]
    public void PathUtilitiesRejectBeforeOpeningCancelledInputOrReplacingDestination(string operation) {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-form-cancel-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try {
            string input = Path.Combine(directory, "missing.pdf");
            string output = Path.Combine(directory, "existing.pdf");
            File.WriteAllBytes(output, new byte[] { 1, 2, 3 });
            using var cancellation = new CancellationTokenSource();
            cancellation.Cancel();
            var options = new PdfFormFillerOptions { CancellationToken = cancellation.Token };

            Assert.ThrowsAny<OperationCanceledException>(() => {
                if (operation == "fill") PdfFormFiller.FillFields(input, output, Values, options);
                else if (operation == "flatten") PdfFormFiller.FlattenFields(input, output, options);
                else PdfFormFiller.FillAndFlattenFields(input, output, Values, options);
            });

            Assert.Equal(new byte[] { 1, 2, 3 }, File.ReadAllBytes(output));
            Assert.Single(Directory.GetFiles(directory));
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void ExplicitAndOptionsTokensBothCancelTypedFill(bool cancelOptionsToken) {
        using var explicitCancellation = new CancellationTokenSource();
        using var optionsCancellation = new CancellationTokenSource();
        (cancelOptionsToken ? optionsCancellation : explicitCancellation).Cancel();
        byte[] source = PdfFormPlacementFixture.CreateAuthored("text");
        var values = new Dictionary<string, PdfFormFieldValue> { ["Value"] = PdfFormFieldValue.From("QgjH") };
        var options = new PdfFormFillerOptions { CancellationToken = optionsCancellation.Token };

        Assert.ThrowsAny<OperationCanceledException>(() =>
            PdfFormFiller.FillFields(source, values, options, null, explicitCancellation.Token));
    }

    private static byte[] Run(string operation, Stream stream, PdfFormFillerOptions? formOptions, PdfLoadOptions readOptions) => operation switch {
        "fill" => PdfFormFiller.FillFields(stream, Values, formOptions, readOptions),
        "flatten" => PdfFormFiller.FlattenFields(stream, formOptions, readOptions),
        _ => PdfFormFiller.FillAndFlattenFields(stream, Values, formOptions, readOptions)
    };

    private sealed class TrackingInput : MemoryStream {
        private readonly bool _nonseekable;
        private readonly Action? _afterRead;
        internal int BytesRead { get; private set; }
        internal TrackingInput(byte[] bytes, bool nonseekable, int start = 0, Action? afterRead = null) : base(bytes, writable: false) {
            _nonseekable = nonseekable;
            _afterRead = afterRead;
            base.Position = start;
        }
        public override bool CanSeek => !_nonseekable;
        public override long Length => _nonseekable ? throw new NotSupportedException() : base.Length;
        public override long Position { get => _nonseekable ? throw new NotSupportedException() : base.Position; set { if (_nonseekable) throw new NotSupportedException(); base.Position = value; } }
        public override int Read(byte[] buffer, int offset, int count) {
            int read = base.Read(buffer, offset, count);
            BytesRead += read;
            if (read > 0) _afterRead?.Invoke();
            return read;
        }
    }

    private sealed class CancellingOutput : MemoryStream {
        private readonly CancellationTokenSource _cancellation;
        internal CancellingOutput(CancellationTokenSource cancellation) => _cancellation = cancellation;
        public override void Write(byte[] buffer, int offset, int count) {
            base.Write(buffer, offset, count);
            _cancellation.Cancel();
        }
    }
}
