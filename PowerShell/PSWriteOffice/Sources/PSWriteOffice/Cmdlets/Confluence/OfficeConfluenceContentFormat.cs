namespace PSWriteOffice.Cmdlets.Confluence;

/// <summary>Input representation accepted by Confluence page publishing cmdlets.</summary>
public enum OfficeConfluenceContentFormat
{
    /// <summary>OfficeIMO Markdown.</summary>
    Markdown,

    /// <summary>HTML or Confluence storage-format XHTML.</summary>
    Html,

    /// <summary>Atlas Document Format JSON.</summary>
    AtlasDocFormat,

    /// <summary>Confluence storage-format XHTML.</summary>
    Storage
}
