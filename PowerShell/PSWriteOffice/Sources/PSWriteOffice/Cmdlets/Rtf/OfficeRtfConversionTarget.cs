namespace PSWriteOffice.Cmdlets.Rtf;

/// <summary>Supported document targets for RTF conversion.</summary>
public enum OfficeRtfConversionTarget
{
    /// <summary>Convert RTF content to a Word document.</summary>
    Word,

    /// <summary>Convert RTF content to HTML through OfficeIMO.Word.Html.</summary>
    Html,

    /// <summary>Convert RTF content to PDF through OfficeIMO.Rtf.Pdf.</summary>
    Pdf,

    /// <summary>Convert RTF content to Markdown through OfficeIMO.Rtf.Markdown.</summary>
    Markdown
}
