namespace OfficeIMO.Visio;

public partial class VisioDocument {
    /// <summary>Gets an existing page by zero-based index.</summary>
    public VisioPage GetPage(int pageIndex) {
        if (pageIndex < 0 || pageIndex >= Pages.Count) throw new ArgumentOutOfRangeException(nameof(pageIndex));
        return Pages[pageIndex];
    }

    /// <summary>Gets an existing page by name without adding a page.</summary>
    public VisioPage GetPage(string name, StringComparison comparison = StringComparison.OrdinalIgnoreCase) {
        if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("Page name cannot be null or whitespace.", nameof(name));
        return Pages.FirstOrDefault(page => string.Equals(page.Name, name, comparison))
            ?? throw new InvalidOperationException($"Page '{name}' was not found.");
    }

    /// <summary>Gets an existing named page or creates one with the supplied physical dimensions.</summary>
    public VisioPage GetOrAddPage(string name, OfficeLength? width = null, OfficeLength? height = null,
        VisioMeasurementUnit scaleUnit = VisioMeasurementUnit.Inches, int? id = null,
        StringComparison comparison = StringComparison.OrdinalIgnoreCase) {
        if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("Page name cannot be null or whitespace.", nameof(name));
        return Pages.FirstOrDefault(page => string.Equals(page.Name, name, comparison))
            ?? AddPage(name, width, height, scaleUnit, id);
    }
}
