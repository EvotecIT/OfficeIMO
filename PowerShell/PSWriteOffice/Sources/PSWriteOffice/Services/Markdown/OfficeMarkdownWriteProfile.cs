namespace PSWriteOffice.Services.Markdown;

/// <summary>Friendly Markdown writer profiles exposed by PSWriteOffice.</summary>
public enum OfficeMarkdownWriteProfile
{
    /// <summary>OfficeIMO-flavored Markdown with rich round-trip hints.</summary>
    OfficeIMO,

    /// <summary>Portable Markdown for stricter hosts and broad compatibility.</summary>
    Portable,

    /// <summary>Markdown that emits raw HTML for images to preserve image metadata.</summary>
    HtmlImage
}
