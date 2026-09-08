namespace OfficeIMO.Ocr;

/// <summary>Operation requested through the shared provider execution and timeout boundary.</summary>
public enum OcrOperation {
    /// <summary>Recognizes text and positioned spans.</summary>
    RecognizeText = 0,
    /// <summary>Detects a corrective quarter-turn without recognizing the full document.</summary>
    DetectOrientation = 1
}

/// <summary>Provider evidence for a corrective image rotation; consumers decide whether to apply it.</summary>
public sealed class OcrOrientationResult {
    /// <summary>Clockwise correction to make text upright. Must be zero, 90, 180, or 270 degrees.</summary>
    public int ClockwiseRotationDegrees { get; set; }
    /// <summary>Provider-normalized evidence from zero through one. Calibration remains provider-specific.</summary>
    public double Confidence { get; set; }
    /// <summary>Optional script name reported by the provider; it does not establish a document language.</summary>
    public string? Script { get; set; }
}
