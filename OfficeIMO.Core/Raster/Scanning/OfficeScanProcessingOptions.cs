using System;

namespace OfficeIMO.Drawing;

/// <summary>Output color treatment for an explicitly requested scan-cleanup operation.</summary>
public enum OfficeScanColorMode {
    /// <summary>Retains color while applying the selected geometric and background corrections.</summary>
    PreserveColor,
    /// <summary>Produces opaque grayscale samples.</summary>
    Grayscale,
    /// <summary>Produces opaque black or white samples using a measured or explicit threshold.</summary>
    Bilevel
}

/// <summary>Bounded, opt-in scan cleanup. The supplied source image is never modified.</summary>
public sealed class OfficeScanProcessingOptions {
    /// <summary>Explicit clockwise quarter-turns, applied before deskew. Must be zero through three.</summary>
    public int ClockwiseQuarterTurns { get; set; }
    /// <summary>Additional clockwise rotation, from minus fifteen through fifteen degrees, applied with any detected deskew correction.</summary>
    public double StraightenDegrees { get; set; }
    /// <summary>Input sample mapped to black, from zero through 254. Must be less than WhitePoint.</summary>
    public int BlackPoint { get; set; }
    /// <summary>Input sample mapped to white, from one through 255. Must exceed BlackPoint.</summary>
    public int WhitePoint { get; set; } = 255;
    /// <summary>Midtone gamma, from 0.1 through 10. Values above one lighten midtones.</summary>
    public double Gamma { get; set; } = 1D;
    /// <summary>Detects a small text-line skew and corrects it only when the confidence threshold is met.</summary>
    public bool Deskew { get; set; } = true;
    /// <summary>Largest absolute skew considered, from one through fifteen degrees.</summary>
    public double MaximumDeskewAngleDegrees { get; set; } = 7D;
    /// <summary>Minimum relative projection-score separation required to apply detected skew.</summary>
    public double MinimumDeskewConfidence { get; set; } = 0.02D;
    /// <summary>Normalizes locally estimated paper brightness while retaining darker foreground samples.</summary>
    public bool NormalizeBackground { get; set; } = true;
    /// <summary>Radius, in processed pixels, of the local paper-brightness window. Must be one through 128.</summary>
    public int BackgroundRadius { get; set; } = 15;
    /// <summary>Color conversion applied after geometry and background normalization.</summary>
    public OfficeScanColorMode ColorMode { get; set; } = OfficeScanColorMode.Grayscale;
    /// <summary>Optional bilevel threshold. Null selects a global histogram threshold; values are one through 254.</summary>
    public int? BilevelThreshold { get; set; }
    /// <summary>Optional longest-side limit. Images are downsampled proportionally and are never enlarged.</summary>
    public int? MaximumDimension { get; set; }
    /// <summary>Maximum pixels in the source or any output canvas.</summary>
    public long MaximumPixels { get; set; } = 20_000_000L;
    /// <summary>Maximum accounted managed bytes, including source, output, analysis, and resampling buffers.</summary>
    public long MaximumWorkingBytes { get; set; } = 256L * 1024L * 1024L;
    /// <summary>Maximum sampled foreground points used by deskew analysis.</summary>
    public int MaximumAnalysisSamples { get; set; } = 32_000;
    /// <summary>Maximum point projections evaluated during deskew analysis.</summary>
    public long MaximumAnalysisOperations { get; set; } = 10_000_000L;

    /// <summary>Creates an independent snapshot.</summary>
    public OfficeScanProcessingOptions Clone() => (OfficeScanProcessingOptions)MemberwiseClone();

    /// <summary>Validates operation settings before allocating buffers or requesting provider work.</summary>
    public void Validate() {
        if (ClockwiseQuarterTurns < 0 || ClockwiseQuarterTurns > 3) throw new ArgumentOutOfRangeException(nameof(ClockwiseQuarterTurns));
        if (!Finite(StraightenDegrees) || Math.Abs(StraightenDegrees) > 15D) throw new ArgumentOutOfRangeException(nameof(StraightenDegrees));
        if (BlackPoint < 0 || BlackPoint > 254) throw new ArgumentOutOfRangeException(nameof(BlackPoint));
        if (WhitePoint <= BlackPoint || WhitePoint > 255) throw new ArgumentOutOfRangeException(nameof(WhitePoint));
        if (!Finite(Gamma) || Gamma < 0.1D || Gamma > 10D) throw new ArgumentOutOfRangeException(nameof(Gamma));
        if (!Finite(MaximumDeskewAngleDegrees) || MaximumDeskewAngleDegrees < 1D || MaximumDeskewAngleDegrees > 15D)
            throw new ArgumentOutOfRangeException(nameof(MaximumDeskewAngleDegrees));
        if (!Finite(MinimumDeskewConfidence) || MinimumDeskewConfidence < 0D || MinimumDeskewConfidence > 1D)
            throw new ArgumentOutOfRangeException(nameof(MinimumDeskewConfidence));
        if (BackgroundRadius < 1 || BackgroundRadius > 128) throw new ArgumentOutOfRangeException(nameof(BackgroundRadius));
        if (ColorMode < OfficeScanColorMode.PreserveColor || ColorMode > OfficeScanColorMode.Bilevel) throw new ArgumentOutOfRangeException(nameof(ColorMode));
        if (BilevelThreshold.HasValue && (BilevelThreshold < 1 || BilevelThreshold > 254)) throw new ArgumentOutOfRangeException(nameof(BilevelThreshold));
        if (MaximumDimension.HasValue && MaximumDimension <= 0) throw new ArgumentOutOfRangeException(nameof(MaximumDimension));
        if (MaximumPixels <= 0 || MaximumPixels > OfficeRasterGuards.MaximumPixels) throw new ArgumentOutOfRangeException(nameof(MaximumPixels));
        if (MaximumWorkingBytes <= 0 || MaximumWorkingBytes > OfficeRasterGuards.MaximumDecodedBytes) throw new ArgumentOutOfRangeException(nameof(MaximumWorkingBytes));
        if (MaximumAnalysisSamples < 1 || MaximumAnalysisSamples > 1_000_000) throw new ArgumentOutOfRangeException(nameof(MaximumAnalysisSamples));
        if (MaximumAnalysisOperations <= 0) throw new ArgumentOutOfRangeException(nameof(MaximumAnalysisOperations));
    }

    private static bool Finite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);
}