namespace OfficeIMO.Word.Rtf;

/// <summary>Controls RTF-to-Word conversion.</summary>
public sealed class RtfWordExportOptions {
    /// <summary>Source path recorded in read diagnostics when converting a parsed read result.</summary>
    public string? SourcePath { get; set; }

    /// <summary>Throws <see cref="OfficeConversionException"/> when the conversion reports loss.</summary>
    public bool FailOnLoss { get; set; }

    /// <summary>Creates an independent copy of these settings.</summary>
    public RtfWordExportOptions Clone() => new RtfWordExportOptions { SourcePath = SourcePath, FailOnLoss = FailOnLoss };
}
