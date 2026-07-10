namespace OfficeIMO.Reader.Ocr.Process;

/// <summary>Owns a started OCR process and its redirected readers.</summary>
internal sealed class OfficeOcrStartedProcess : IDisposable {
    private bool _disposed;

    internal OfficeOcrStartedProcess(System.Diagnostics.Process process, TextReader standardOutput, TextReader standardError) {
        Process = process;
        StandardOutput = standardOutput;
        StandardError = standardError;
    }

    internal System.Diagnostics.Process Process { get; }

    internal TextReader StandardOutput { get; }

    internal TextReader StandardError { get; }

    internal void CloseRedirectedStreams() {
        try { StandardOutput.Dispose(); } catch (ObjectDisposedException) { }
        try { StandardError.Dispose(); } catch (ObjectDisposedException) { }
    }

    public void Dispose() {
        if (_disposed) return;
        _disposed = true;
        CloseRedirectedStreams();
        Process.Dispose();
    }
}
