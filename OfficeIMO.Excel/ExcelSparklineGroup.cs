using DocumentFormat.OpenXml.Office2010.Excel;

namespace OfficeIMO.Excel;

/// <summary>A group of worksheet sparklines sharing display settings.</summary>
public sealed class ExcelSparklineGroup {
    private readonly SparklineGroup _group;
    internal ExcelSparklineGroup(ExcelSheet sheet, SparklineGroup group) {
        Sheet = sheet;
        _group = group;
    }
    /// <summary>Gets the worksheet containing the sparklines.</summary>
    public ExcelSheet Sheet { get; }
    /// <summary>Gets or sets the rendering type for every sparkline in the group.</summary>
    public ExcelSparklineType Type {
        get => _group.Type?.Value.ToOfficeEnum() ?? ExcelSparklineType.Line;
        set {
            if (value != ExcelSparklineType.Line && value != ExcelSparklineType.Column && value != ExcelSparklineType.Stacked)
                throw new ArgumentOutOfRangeException(nameof(value));
            Sheet.RangeEdit(() => _group.Type = value.ToOpenXml());
        }
    }
}
