namespace OfficeIMO.Visio;

public partial class VisioDocument {
    /// <summary>Adds a default-sized page and configures the created native page.</summary>
    public VisioPage AddPage(string name, Action<VisioPage> configure) => AddPage(name).Configure(configure);
    /// <summary>Adds a background page and configures the created native page.</summary>
    public VisioPage AddBackgroundPage(string name, Action<VisioPage> configure) => AddBackgroundPage(name).Configure(configure);
    /// <summary>Configures an existing page selected by index.</summary>
    public VisioPage GetPage(int index, Action<VisioPage> configure) => GetPage(index).Configure(configure);
    /// <summary>Configures an existing page selected by name.</summary>
    public VisioPage GetPage(string name, Action<VisioPage> configure, StringComparison comparison = StringComparison.OrdinalIgnoreCase) => GetPage(name, comparison).Configure(configure);
    /// <summary>Configures a page belonging to this document.</summary>
    public VisioPage GetPage(VisioPage page, Action<VisioPage> configure) {
        ArgumentNullException.ThrowIfNull(page);
        if (!Pages.Contains(page)) throw new InvalidOperationException("The page does not belong to this document.");
        return page.Configure(configure);
    }
    /// <summary>Gets or adds a named page and configures its native model.</summary>
    public VisioPage GetOrAddPage(string name, Action<VisioPage> configure, StringComparison comparison = StringComparison.OrdinalIgnoreCase) => GetOrAddPage(name, comparison: comparison).Configure(configure);
    /// <summary>Configures the document's native metadata.</summary>
    public VisioDocument ConfigureProperties(Action<VisioDocumentProperties> configure) { ArgumentNullException.ThrowIfNull(configure); configure(Properties); return this; }
}
