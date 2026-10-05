using System.Diagnostics;
using OfficeIMO.Excel.IWork;
using OfficeIMO.PowerPoint.IWork;
using OfficeIMO.Word.IWork;

namespace OfficeIMO.IWork.Benchmarks;

public sealed partial class IWorkRuntimeWorkload {
    private IWorkKeynoteDocument? _nativeCopyDocument;
    /// <summary>Gets the deterministic encoded package hash for the owned one-slide cancellation model.</summary>
    public string? NativeCopyInputSha256 { get; private set; }
    /// <summary>Gets the encoded size of the owned one-slide cancellation model.</summary>
    public long NativeCopyInputBytes { get; private set; }
    /// <summary>Prepares operation-specific input and releases previous results outside measurement.</summary>
    public void Prepare(string operation) {
        ReleaseResults();
        if (operation == "CancelDuringNativeCopy" && _kind == IWorkDocumentKind.Keynote && _nativeCopyDocument is null) {
            _nativeCopyDocument = IWorkKeynoteDocument.Create();
            _nativeCopyDocument.AddSlide().AddText(new string('a', 150_000), 0, 0, 960, 540);
            byte[] bytes = _nativeCopyDocument.SaveBytes();
            NativeCopyInputBytes = bytes.LongLength;
            NativeCopyInputSha256 = Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(bytes));
        }
    }
    /// <summary>Gets whether cancellation was observed after the operation performed I/O.</summary>
    public bool CancellationObserved { get; private set; }
    /// <summary>Gets bytes read or written before requesting cancellation.</summary>
    public long CancellationProcessedBytes { get; private set; }
    /// <summary>Gets all operation I/O observed before cancellation was reported, including any subsequent I/O.</summary>
    public long CancellationTotalIoBytes { get; private set; }
    /// <summary>Gets time from the synchronous I/O cancellation request to its observed exception.</summary>
    public double CancellationLatencyMs { get; private set; }

    private void ExecuteCancellation(string operation) {
        using var cancellation = new CancellationTokenSource();
        using var stream = operation == "CancelDuringNativeCopy"
            ? new CancellationIoStream(cancellation)
            : new CancellationIoStream(_input, cancellation);
        try {
            switch (operation) {
                case "CancelDuringLoad":
                    IWorkSourceDocument.Open(stream, _kind, null, cancellation.Token);
                    break;
                case "CancelDuringConvert":
                    using (IDisposable result = _kind switch {
                        IWorkDocumentKind.Pages => WordIWorkConverter.ConvertPagesToWordResult(stream, null, null, cancellation.Token),
                        IWorkDocumentKind.Numbers => ExcelIWorkConverter.ConvertNumbersToExcelResult(stream, null, null, cancellation.Token),
                        IWorkDocumentKind.Keynote => PowerPointIWorkConverter.ConvertKeynoteToPowerPointResult(stream, null, null, cancellation.Token),
                        _ => throw new InvalidOperationException()
                    }) { }
                    break;
                case "CancelDuringNativeCopy" when _kind == IWorkDocumentKind.Keynote:
                    (_nativeCopyDocument ?? throw new InvalidOperationException("Prepare the native-copy operation first."))
                        .Save(stream, cancellationToken: cancellation.Token);
                    break;
                default: throw new ArgumentException("Unknown or incompatible cancellation operation.", nameof(operation));
            }
            throw new InvalidDataException("The operation completed without observing cancellation.");
        } catch (OperationCanceledException) {
            long observed = Stopwatch.GetTimestamp();
            if (!cancellation.IsCancellationRequested || stream.RequestedAt == 0 || stream.ProcessedBytes == 0
                || !stream.CanRead || !stream.CanWrite && operation == "CancelDuringNativeCopy")
                throw new InvalidDataException("The operation did not preserve the caller stream after active cancellation.");
            CancellationObserved = true;
            CancellationProcessedBytes = stream.RequestedProcessedBytes;
            CancellationTotalIoBytes = stream.ProcessedBytes;
            CancellationLatencyMs = Stopwatch.GetElapsedTime(stream.RequestedAt, observed).TotalMilliseconds;
            if (operation == "CancelDuringNativeCopy" && stream.Length != 65_536)
                throw new InvalidDataException("Native copying continued beyond its first cancelled chunk.");
        }
    }

    private sealed class CancellationIoStream : MemoryStream {
        private readonly CancellationTokenSource _cancellation;
        internal long RequestedAt { get; private set; }
        internal long RequestedProcessedBytes { get; private set; }
        internal long ProcessedBytes { get; private set; }
        internal CancellationIoStream(CancellationTokenSource cancellation) => _cancellation = cancellation;
        internal CancellationIoStream(byte[] bytes, CancellationTokenSource cancellation) : base(bytes, writable: false) => _cancellation = cancellation;
        public override int Read(byte[] buffer, int offset, int count) {
            int read = base.Read(buffer, offset, count); RequestAfterIo(read); return read;
        }
        public override int Read(Span<byte> buffer) {
            int read = base.Read(buffer); RequestAfterIo(read); return read;
        }
        public override void Write(byte[] buffer, int offset, int count) {
            base.Write(buffer, offset, count); RequestAfterIo(count);
        }
        private void RequestAfterIo(int count) {
            ProcessedBytes += count;
            if (count <= 0 || RequestedAt != 0) return;
            RequestedProcessedBytes = ProcessedBytes;
            RequestedAt = Stopwatch.GetTimestamp();
            _cancellation.Cancel();
        }
    }
}
