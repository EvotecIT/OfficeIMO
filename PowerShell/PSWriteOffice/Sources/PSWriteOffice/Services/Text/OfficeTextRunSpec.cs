namespace PSWriteOffice.Services.Text;

/// <summary>PowerShell-friendly rich text run specification used by document adapters.</summary>
public sealed class OfficeTextRunSpec
{
    /// <summary>Text to append for this run.</summary>
    public string Text { get; set; } = string.Empty;

    /// <summary>Run kind such as text, line break, or tab.</summary>
    public string? Kind { get; set; }

    /// <summary>Render the run in bold.</summary>
    public bool Bold { get; set; }

    /// <summary>Render the run in italics.</summary>
    public bool Italic { get; set; }

    /// <summary>Render the run with underline.</summary>
    public bool Underline { get; set; }

    /// <summary>Optional underline style name when the target format supports it.</summary>
    public string? UnderlineStyle { get; set; }

    /// <summary>Render the run with strikethrough.</summary>
    public bool Strike { get; set; }

    /// <summary>Text color. Named colors and hexadecimal colors are accepted.</summary>
    public string? Color { get; set; }

    /// <summary>Run background or highlight color. Named colors and hexadecimal colors are accepted; None clears the background or highlight where supported.</summary>
    public string? BackgroundColor { get; set; }

    /// <summary>Font size in points.</summary>
    public double? FontSize { get; set; }

    /// <summary>Font name, family, or target-specific font identifier.</summary>
    public string? FontName { get; set; }

    /// <summary>Target-specific baseline name, such as superscript or subscript.</summary>
    public string? Baseline { get; set; }

    /// <summary>URI link target when supported by the target format.</summary>
    public string? LinkUri { get; set; }

    /// <summary>Named destination or bookmark target when supported by the target format.</summary>
    public string? LinkDestinationName { get; set; }

    /// <summary>Optional link tooltip or annotation contents.</summary>
    public string? LinkContents { get; set; }

    /// <summary>PDF tab leader style name.</summary>
    public string? TabLeader { get; set; }

    /// <summary>Tab alignment name.</summary>
    public string? TabAlignment { get; set; }

    internal bool IsLineBreak => OfficeTextRunParser.NormalizeKind(Kind) is "linebreak" or "break" or "br";

    internal bool IsTab => OfficeTextRunParser.NormalizeKind(Kind) == "tab";
}
