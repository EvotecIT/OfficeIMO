namespace OfficeIMO.Word;

public partial class WordDocument {
    private WordStyleCollection? _styles;
    private WordTheme? _theme;

    /// <summary>Gets the styles this document defines, to list, inspect, create, change, remove and import them.</summary>
    public WordStyleCollection Styles => _styles ??= new WordStyleCollection(this);

    /// <summary>Gets the document theme: its colour scheme and its heading and body fonts.</summary>
    public WordTheme Theme => _theme ??= new WordTheme(this);
}
