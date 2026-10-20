namespace OfficeIMO.Html;

/// <summary>Identifies the fragmentation contexts to which an authored or structural keep applies.</summary>
[Flags]
internal enum HtmlFragmentationScope {
    None = 0,
    Page = 1,
    Column = 2,
    All = Page | Column
}
