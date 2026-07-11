namespace OfficeIMO.Reader.Ocr.Tesseract;

/// <summary>Configures the optional Tesseract command-line OCR engine.</summary>
public sealed class TesseractOcrEngineOptions {
    /// <summary>Executable path or name. Defaults to <c>tesseract</c>.</summary>
    public string ExecutablePath { get; set; } = "tesseract";

    /// <summary>Default Tesseract language expression, such as <c>eng</c> or <c>eng+pol</c>.</summary>
    public string? Language { get; set; } = "eng";

    /// <summary>Optional tessdata directory passed through <c>--tessdata-dir</c>.</summary>
    public string? TessdataDirectory { get; set; }

    /// <summary>Optional OCR engine mode from 0 through 3.</summary>
    public int? EngineMode { get; set; }

    /// <summary>Optional page segmentation mode from 0 through 13.</summary>
    public int? PageSegmentationMode { get; set; }

    /// <summary>Optional input DPI passed through <c>--dpi</c>.</summary>
    public int? Dpi { get; set; }

    /// <summary>Additional direct Tesseract arguments inserted before the <c>tsv</c> output config.</summary>
    public IReadOnlyList<string> AdditionalArguments { get; set; } = Array.Empty<string>();

    /// <summary>Optional parent directory for isolated per-request temporary folders.</summary>
    public string? TemporaryDirectory { get; set; }

    /// <summary>Maximum recognition process duration. Defaults to two minutes.</summary>
    public TimeSpan Timeout { get; set; } = TimeSpan.FromMinutes(2);

    /// <summary>Maximum TSV result size. Defaults to 32 MiB.</summary>
    public long MaxOutputBytes { get; set; } = 32L * 1024L * 1024L;

    /// <summary>Maximum input payload size accepted by direct engine calls. Defaults to 25 MiB.</summary>
    public long MaxInputBytes { get; set; } = 25L * 1024L * 1024L;

    /// <summary>Maximum retained process output characters.</summary>
    public int MaxProcessOutputCharacters { get; set; } = 64 * 1024;

    /// <summary>Whether isolated temporary files are retained after recognition.</summary>
    public bool KeepTemporaryFiles { get; set; }
}
