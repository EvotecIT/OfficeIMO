using OfficeIMO.IWork;

namespace OfficeIMO.Excel.IWork;

/// <summary>Controls Numbers-to-Excel conversion.</summary>
public sealed class NumbersExcelExportOptions {
    /// <summary>Shared iWork projection settings. Null uses the defaults.</summary>
    public IWorkConversionOptions? ConversionOptions { get; set; }

    /// <summary>Throws <see cref="OfficeConversionException"/> when the conversion reports loss.</summary>
    public bool FailOnLoss { get; set; }

    /// <summary>Creates a copy of these settings; the projection settings are shared.</summary>
    public NumbersExcelExportOptions Clone() => new NumbersExcelExportOptions { ConversionOptions = ConversionOptions, FailOnLoss = FailOnLoss };
}
