namespace PSWriteOffice.Services.Word;

/// <summary>Defines how HTML fragments are inserted.</summary>
public enum HtmlImportMode
{
    /// <summary>Parse HTML into OpenXML elements.</summary>
    Parse,
    /// <summary>Embed the HTML payload as-is.</summary>
    AsIs
}
