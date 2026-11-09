using OfficeIMO;
using System.Management.Automation;
using OfficeIMO.Word;

namespace PSWriteOffice.Services.Word;

internal sealed class WordTableConditionModel
{
    public WordTableConditionModel(ScriptBlock filterScript, WordTableStyle? tableStyle, OfficeColor? backgroundColor, OfficeColor? color, bool? bold, int? columnIndex)
    {
        FilterScript = filterScript ?? throw new ArgumentNullException(nameof(filterScript));
        TableStyle = tableStyle;
        BackgroundColor = backgroundColor;
        Color = color; Bold = bold; ColumnIndex = columnIndex;
    }

    public ScriptBlock FilterScript { get; }
    public WordTableStyle? TableStyle { get; }
    public OfficeColor? BackgroundColor { get; }
    public OfficeColor? Color { get; }
    public bool? Bold { get; }
    public int? ColumnIndex { get; }
}
