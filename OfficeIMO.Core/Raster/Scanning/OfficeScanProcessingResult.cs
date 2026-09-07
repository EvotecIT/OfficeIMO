using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

/// <summary>One applied or deliberately skipped scan-processing decision.</summary>
public sealed class OfficeScanProcessingStep {
    internal OfficeScanProcessingStep(string operation, bool applied, string message) {
        Operation = operation; Applied = applied; Message = message;
    }
    /// <summary>Stable operation identifier such as deskew, background, color, or downsample.</summary>
    public string Operation { get; }
    /// <summary>Whether this decision changed image samples or geometry.</summary>
    public bool Applied { get; }
    /// <summary>Explanation of the decision, including low-confidence skips.</summary>
    public string Message { get; }
}

/// <summary>Immutable geometry and quality evidence for a processed scan.</summary>
public sealed class OfficeScanProcessingReport {
    internal OfficeScanProcessingReport(int sourceWidth, int sourceHeight, int width, int height,
        OfficeTransform transform, double detectedSkew, double appliedSkew, double confidence,
        double foregroundFraction, bool probablyBlank, long workingBytes, List<OfficeScanProcessingStep> steps) {
        SourceWidth = sourceWidth; SourceHeight = sourceHeight; Width = width; Height = height;
        SourceToProcessed = transform; ProcessedToSource = transform.Invert();
        DetectedSkewDegrees = detectedSkew; AppliedDeskewDegrees = appliedSkew; DeskewConfidence = confidence;
        ForegroundFraction = foregroundFraction; IsProbablyBlank = probablyBlank; EstimatedPeakWorkingBytes = workingBytes;
        Steps = Array.AsReadOnly(steps.ToArray());
    }
    /// <summary>Source width in pixels.</summary>
    public int SourceWidth { get; }
    /// <summary>Source height in pixels.</summary>
    public int SourceHeight { get; }
    /// <summary>Processed width in pixels.</summary>
    public int Width { get; }
    /// <summary>Processed height in pixels.</summary>
    public int Height { get; }
    /// <summary>Maps top-left source pixel-edge coordinates into the processed image.</summary>
    public OfficeTransform SourceToProcessed { get; }
    /// <summary>Maps processed pixel-edge coordinates, including OCR geometry, back to the source.</summary>
    public OfficeTransform ProcessedToSource { get; }
    /// <summary>Detected clockwise line skew after explicit quarter-turns, before confidence filtering.</summary>
    public double DetectedSkewDegrees { get; }
    /// <summary>Clockwise correction actually applied; zero when deskew was skipped.</summary>
    public double AppliedDeskewDegrees { get; }
    /// <summary>Relative separation of the selected projection score from competing angles.</summary>
    public double DeskewConfidence { get; }
    /// <summary>Fraction of processed samples below the measured foreground threshold.</summary>
    public double ForegroundFraction { get; }
    /// <summary>Conservative blank-page suggestion. The processor never removes a page.</summary>
    public bool IsProbablyBlank { get; }
    /// <summary>Conservative managed-buffer accounting used by the operation's resource gate.</summary>
    public long EstimatedPeakWorkingBytes { get; }
    /// <summary>Applied and skipped transformations in execution order.</summary>
    public IReadOnlyList<OfficeScanProcessingStep> Steps { get; }
}

/// <summary>A separately owned processed image and the evidence needed to review it.</summary>
public sealed class OfficeScanProcessingResult {
    internal OfficeScanProcessingResult(OfficeRasterImage image, OfficeScanProcessingReport report) { Image = image; Report = report; }
    /// <summary>Processed image. Its pixel buffer is independent of the source image.</summary>
    public OfficeRasterImage Image { get; }
    /// <summary>Geometry, quality measurements, and transformation decisions.</summary>
    public OfficeScanProcessingReport Report { get; }
}

/// <summary>Scan processing could not satisfy a configured resource limit; the source remains unchanged.</summary>
public sealed class OfficeScanProcessingLimitException : InvalidOperationException {
    internal OfficeScanProcessingLimitException(string message) : base(message) { }
}
