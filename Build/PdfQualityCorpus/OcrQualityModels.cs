using OfficeIMO.Ocr;

namespace OfficeIMO.PdfQualityCorpus;

internal sealed class OcrQualityManifest {
    public int Version { get; set; }
    public string Labels { get; set; } = "";
    public string LabelsSha256 { get; set; } = "";
    public List<OcrQualitySource> Sources { get; set; } = new();
}

internal sealed class OcrQualitySource {
    public string File { get; set; } = "";
    public string Sha256 { get; set; } = "";
    public string Authority { get; set; } = "";
    public string License { get; set; } = "";
}

internal sealed class OcrQualityLabel {
    public string Id { get; set; } = "";
    public string File { get; set; } = "";
    public string DocumentClass { get; set; } = "";
    public string Language { get; set; } = "eng";
    public int Page { get; set; } = 1;
    public double[]? Region { get; set; }
    public string Expected { get; set; } = "";
    public double MaximumCer { get; set; } = .02;
    public double MaximumWer { get; set; } = .05;
    public bool Cleanup { get; set; }
}

internal sealed class OcrQualityMeasurement {
    public string Id { get; init; } = "";
    public string DocumentClass { get; init; } = "";
    public string SourceSha256 { get; init; } = "";
    public string RasterSha256 { get; init; } = "";
    public string BaselineTextSha256 { get; init; } = "";
    public string SelectedTextSha256 { get; init; } = "";
    public int Repetition { get; init; }
    public string Outcome { get; init; } = "completed";
    public string? FailureType { get; init; }
    public ScanTextAccuracy? Baseline { get; init; }
    public ScanTextAccuracy? Selected { get; init; }
    public bool GoldLimitsMet { get; init; }
    public bool ReviewRecommended { get; init; } = true;
    public bool HasDisagreement { get; init; }
    public bool RetryIncomplete { get; init; }
    public bool GeometryWithinRaster { get; init; }
    public bool SourceUnchanged { get; init; }
    public double ElapsedMilliseconds { get; init; }
    public long HostPeakWorkingSetBytes { get; init; }
    public int SelectedAttempt { get; init; }
    public IReadOnlyList<OcrAttemptAssessment> Attempts { get; init; } = Array.Empty<OcrAttemptAssessment>();
    public List<OcrThresholdObservation> Thresholds { get; init; } = new();
}

internal sealed class OcrThresholdObservation {
    public double MinimumWordConfidence { get; init; }
    public double MaximumUncertainWordFraction { get; init; }
    public bool MeetsThresholds { get; init; }
    public bool GoldLimitsMet { get; init; }
    public int LowConfidenceWords { get; init; }
    public int UnknownConfidenceWords { get; init; }
}
