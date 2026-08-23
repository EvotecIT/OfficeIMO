using System.Diagnostics;
using System.Reflection;

namespace OfficeIMO.Drawing.Benchmarks;

/// <summary>Collects isolated process-memory evidence for materialized and streamed encoding.</summary>
internal static class ImagePeakMemoryEvidence {
    private sealed record EvidenceCase(string ScenarioId, OfficeImageExportFormat Format);

    private sealed record EvidenceResult(
        string ScenarioId,
        OfficeImageExportFormat Format,
        string Mode,
        long EncodedBytes,
        long ManagedAllocatedBytes,
        long PeakWorkingSetDelta,
        long PeakPrivateBytesDelta);

    private static readonly EvidenceCase[] Cases = {
        new(ImageBenchmarkScenarios.Tiny.Id, OfficeImageExportFormat.Png),
        new(ImageBenchmarkScenarios.Screenshot.Id, OfficeImageExportFormat.Png),
        new(ImageBenchmarkScenarios.Scan.Id, OfficeImageExportFormat.Tiff),
        new(ImageBenchmarkScenarios.AlphaGraphic.Id, OfficeImageExportFormat.Webp),
        new(ImageBenchmarkScenarios.HighEntropy.Id, OfficeImageExportFormat.Png),
        new(ImageBenchmarkScenarios.Photo.Id, OfficeImageExportFormat.Jpeg),
        new(ImageBenchmarkScenarios.VeryLarge.Id, OfficeImageExportFormat.Tiff)
    };

    internal static void Validate(TextWriter writer, string[] scenarioFilters) {
        HashSet<string>? filters = scenarioFilters.Length == 0
            ? null
            : new HashSet<string>(scenarioFilters, StringComparer.OrdinalIgnoreCase);
        writer.WriteLine("Isolated encode memory evidence (deltas after the source image is resident):");
        writer.WriteLine("Scenario       Format API             Bytes      Managed    Peak working   Peak private");
        foreach (EvidenceCase evidenceCase in Cases) {
            if (filters != null && !filters.Contains(evidenceCase.ScenarioId)) continue;
            (long materializedBytes, long streamedBytes) = ValidateEvidenceCase(evidenceCase);
            EvidenceResult materialized = RunIsolated(evidenceCase, "materialized");
            EvidenceResult streamed = RunIsolated(evidenceCase, "stream");
            WriteRow(writer, materialized);
            WriteRow(writer, streamed);
            if (materialized.EncodedBytes != materializedBytes || streamed.EncodedBytes != streamedBytes) {
                throw new InvalidOperationException(
                    $"{evidenceCase.ScenarioId} {evidenceCase.Format} memory evidence did not match the validated output lengths.");
            }
        }
    }

    private static (long MaterializedBytes, long StreamedBytes) ValidateEvidenceCase(
        EvidenceCase evidenceCase) {
        ImageBenchmarkScenario scenario = ImageBenchmarkScenarios.Get(evidenceCase.ScenarioId);
        OfficeRasterImage image = scenario.CreateImage();
        OfficeRasterEncodingOptions options = CreateOptions();
        byte[] materialized = OfficeRasterImageEncoder.Encode(image, evidenceCase.Format, options);
        using var destination = new MemoryStream();
        OfficeRasterImageEncoder.EncodeTo(image, evidenceCase.Format, destination, options);
        byte[] streamed = destination.ToArray();

        if (!OfficeRasterImageDecoder.TryDecode(materialized, out OfficeRasterImage? expected) || expected == null ||
            !OfficeRasterImageDecoder.TryDecode(streamed, out OfficeRasterImage? actual) || actual == null) {
            throw new InvalidOperationException(
                $"{evidenceCase.ScenarioId} {evidenceCase.Format} could not be decoded before memory measurement.");
        }
        if (expected.Width != scenario.Width || expected.Height != scenario.Height ||
            actual.Width != scenario.Width || actual.Height != scenario.Height ||
            !expected.GetPixels().AsSpan().SequenceEqual(actual.GetPixels())) {
            throw new InvalidOperationException(
                $"{evidenceCase.ScenarioId} {evidenceCase.Format} streamed output failed pre-measurement validation.");
        }
        return (materialized.LongLength, streamed.LongLength);
    }

    internal static void RunWorker(
        string scenarioId,
        string formatText,
        string mode,
        TextWriter writer) {
        if (!Enum.TryParse(formatText, ignoreCase: true, out OfficeImageExportFormat format) ||
            format == OfficeImageExportFormat.Svg) {
            throw new ArgumentException("The memory worker format is invalid.", nameof(formatText));
        }
        bool materialize = mode.Equals("materialized", StringComparison.OrdinalIgnoreCase);
        if (!materialize && !mode.Equals("stream", StringComparison.OrdinalIgnoreCase)) {
            throw new ArgumentException("The memory worker mode is invalid.", nameof(mode));
        }

        ImageBenchmarkScenario scenario = ImageBenchmarkScenarios.Get(scenarioId);
        OfficeRasterImage image = scenario.CreateImage();
        OfficeRasterEncodingOptions options = CreateOptions();
        GC.Collect(GC.MaxGeneration, GCCollectionMode.Forced, blocking: true, compacting: true);
        GC.WaitForPendingFinalizers();
        GC.Collect(GC.MaxGeneration, GCCollectionMode.Forced, blocking: true, compacting: true);

        using var sampler = new ProcessMemorySampler();
        sampler.Start();
        long allocatedBefore = GC.GetAllocatedBytesForCurrentThread();
        byte[]? output = null;
        long encodedBytes;
        if (materialize) {
            output = OfficeRasterImageEncoder.Encode(image, format, options);
            encodedBytes = output.LongLength;
        } else {
            var destination = new CountingWriteStream();
            OfficeRasterImageEncoder.EncodeTo(image, format, destination, options);
            encodedBytes = destination.BytesWritten;
        }
        long allocated = GC.GetAllocatedBytesForCurrentThread() - allocatedBefore;
        sampler.Stop();
        GC.KeepAlive(output);
        GC.KeepAlive(image);

        writer.WriteLine(string.Join(
            "\t",
            scenarioId,
            format,
            materialize ? "materialized" : "stream",
            encodedBytes,
            allocated,
            sampler.PeakWorkingSetDelta,
            sampler.PeakPrivateBytesDelta));
    }

    private static EvidenceResult RunIsolated(EvidenceCase evidenceCase, string mode) {
        ProcessStartInfo startInfo = CreateSelfStartInfo();
        startInfo.ArgumentList.Add("--memory-worker");
        startInfo.ArgumentList.Add(evidenceCase.ScenarioId);
        startInfo.ArgumentList.Add(evidenceCase.Format.ToString());
        startInfo.ArgumentList.Add(mode);
        startInfo.UseShellExecute = false;
        startInfo.RedirectStandardOutput = true;
        startInfo.RedirectStandardError = true;
        startInfo.CreateNoWindow = true;

        using Process process = Process.Start(startInfo) ??
            throw new InvalidOperationException("The image memory worker could not be started.");
        string output = process.StandardOutput.ReadToEnd();
        string error = process.StandardError.ReadToEnd();
        process.WaitForExit();
        if (process.ExitCode != 0) {
            throw new InvalidOperationException(
                $"The image memory worker failed for {evidenceCase.ScenarioId} {evidenceCase.Format} {mode}: {error.Trim()}");
        }

        string line = output.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries).LastOrDefault() ?? string.Empty;
        string[] fields = line.Split('\t');
        if (fields.Length != 7 ||
            !Enum.TryParse(fields[1], out OfficeImageExportFormat format) ||
            !long.TryParse(fields[3], out long encodedBytes) ||
            !long.TryParse(fields[4], out long managedAllocated) ||
            !long.TryParse(fields[5], out long peakWorkingSet) ||
            !long.TryParse(fields[6], out long peakPrivateBytes)) {
            throw new InvalidOperationException("The image memory worker returned an invalid result: " + line);
        }
        return new EvidenceResult(fields[0], format, fields[2], encodedBytes, managedAllocated, peakWorkingSet, peakPrivateBytes);
    }

    private static ProcessStartInfo CreateSelfStartInfo() {
        string executable = Environment.ProcessPath ??
            throw new InvalidOperationException("The current benchmark executable path is unavailable.");
        string assemblyPath = Assembly.GetExecutingAssembly().Location;
        var startInfo = new ProcessStartInfo(executable);
        if (string.Equals(
                Path.GetFileNameWithoutExtension(executable),
                "dotnet",
                StringComparison.OrdinalIgnoreCase)) {
            startInfo.ArgumentList.Add(assemblyPath);
        }
        return startInfo;
    }

    private static OfficeRasterEncodingOptions CreateOptions() => new() {
        DpiX = 144D,
        DpiY = 120D,
        Png = new OfficePngEncodeOptions {
            Compression = OfficePngCompression.Optimal
        },
        Jpeg = new OfficeJpegEncodeOptions {
            Quality = 85,
            Subsampling = OfficeJpegSubsampling.Y420,
            Background = OfficeColor.White
        },
        Tiff = new OfficeTiffEncodeOptions {
            Compression = OfficeTiffCompression.PackBits
        }
    };

    private static void WriteRow(TextWriter writer, EvidenceResult result) {
        writer.WriteLine(
            $"{result.ScenarioId,-14} {result.Format,-6} {result.Mode,-12} " +
            $"{result.EncodedBytes,12:N0} {result.ManagedAllocatedBytes,12:N0} " +
            $"{result.PeakWorkingSetDelta,14:N0} {result.PeakPrivateBytesDelta,14:N0}");
    }

    private sealed class CountingWriteStream : Stream {
        internal long BytesWritten { get; private set; }
        public override bool CanRead => false;
        public override bool CanSeek => false;
        public override bool CanWrite => true;
        public override long Length => BytesWritten;

        public override long Position {
            get => BytesWritten;
            set => throw new NotSupportedException();
        }

        public override void Flush() { }
        public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => BytesWritten = checked(BytesWritten + count);
        public override void WriteByte(byte value) => BytesWritten = checked(BytesWritten + 1);
    }

    private sealed class ProcessMemorySampler : IDisposable {
        private readonly ManualResetEventSlim _started = new(false);
        private readonly Thread _thread;
        private volatile bool _stop;
        private long _baselineWorkingSet;
        private long _baselinePrivateBytes;
        private long _peakWorkingSet;
        private long _peakPrivateBytes;

        internal ProcessMemorySampler() {
            _thread = new Thread(Sample) {
                IsBackground = true,
                Name = "OfficeIMO image memory sampler"
            };
        }

        internal long PeakWorkingSetDelta => Math.Max(0L, _peakWorkingSet - _baselineWorkingSet);
        internal long PeakPrivateBytesDelta => Math.Max(0L, _peakPrivateBytes - _baselinePrivateBytes);

        internal void Start() {
            _thread.Start();
            _started.Wait();
        }

        internal void Stop() {
            _stop = true;
            _thread.Join();
        }

        public void Dispose() {
            if (_thread.IsAlive) Stop();
            _started.Dispose();
        }

        private void Sample() {
            using Process process = Process.GetCurrentProcess();
            process.Refresh();
            _baselineWorkingSet = process.WorkingSet64;
            _baselinePrivateBytes = process.PrivateMemorySize64;
            _peakWorkingSet = _baselineWorkingSet;
            _peakPrivateBytes = _baselinePrivateBytes;
            _started.Set();
            while (!_stop) {
                Record(process);
                Thread.Sleep(1);
            }
            Record(process);
        }

        private void Record(Process process) {
            process.Refresh();
            _peakWorkingSet = Math.Max(_peakWorkingSet, process.WorkingSet64);
            _peakPrivateBytes = Math.Max(_peakPrivateBytes, process.PrivateMemorySize64);
        }
    }
}
