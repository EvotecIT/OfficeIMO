using System;
using System.Management.Automation;
using OfficeIMO.Word;

namespace PSWriteOffice.Services.Word;

internal sealed class WordTableConditionModel
{
    public WordTableConditionModel(ScriptBlock filterScript, WordTableStyle? tableStyle, string? backgroundColor)
    {
        FilterScript = filterScript ?? throw new ArgumentNullException(nameof(filterScript));
        TableStyle = tableStyle;
        BackgroundColor = backgroundColor;
    }

    public ScriptBlock FilterScript { get; }
    public WordTableStyle? TableStyle { get; }
    public string? BackgroundColor { get; }
}
