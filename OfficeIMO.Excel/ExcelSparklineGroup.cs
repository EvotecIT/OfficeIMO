using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.Excel;

/// <summary>A group of worksheet sparklines sharing display settings.</summary>
public sealed class ExcelSparklineGroup {
    private readonly OfficeXmlNode _liveGroup;
    private readonly string _formula;
    private readonly string _location;
    private ExcelSparklineType _type;

    internal ExcelSparklineGroup(ExcelSheet sheet, OfficeXmlNode group, string formula, string location, ExcelSparklineType type) {
        Sheet = sheet;
        _liveGroup = group.IsSdk ? group : default;
        _formula = formula;
        _location = location;
        _type = type;
    }

    /// <summary>Gets the worksheet containing the sparklines.</summary>
    public ExcelSheet Sheet { get; }

    /// <summary>Gets or sets the rendering type for every sparkline in the group.</summary>
    public ExcelSparklineType Type {
        get => _liveGroup.IsNone ? Sheet.ReadSparklineGroupType(_formula, _location, _type) : ExcelSheet.ReadSparklineType(_liveGroup);
        set {
            string typeName = ExcelSheet.SparklineTypeName(value);
            if (_liveGroup.IsNone) Sheet.EditSparklineGroupType(_formula, _location, typeName);
            else Sheet.RangeEdit(() => _liveGroup.SetAttribute("type", typeName));
            _type = value;
        }
    }
}
