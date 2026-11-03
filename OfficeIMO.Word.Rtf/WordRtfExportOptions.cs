using System.Text;
using OfficeIMO.Rtf;

namespace OfficeIMO.Word.Rtf;

/// <summary>Controls Word-to-RTF conversion and the RTF output written by <c>SaveAsRtf</c>.</summary>
public sealed class WordRtfExportOptions {
    /// <summary>RTF writer settings used when the converted document is saved.</summary>
    public RtfSaveOptions? SaveOptions { get; set; }

    /// <summary>Encoding used when the RTF is saved. Null uses the RTF writer default.</summary>
    public Encoding? Encoding { get; set; }

    /// <summary>Throws <see cref="OfficeConversionException"/> when the conversion reports loss.</summary>
    public bool FailOnLoss { get; set; }

    /// <summary>Creates an independent copy of these settings.</summary>
    public WordRtfExportOptions Clone() => new WordRtfExportOptions { SaveOptions = SaveOptions, Encoding = Encoding, FailOnLoss = FailOnLoss };
}
