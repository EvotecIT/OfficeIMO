namespace OfficeIMO.Visio.Fluent;

public partial class VisioFluentDocument {
    /// <summary>Configures an existing page by zero-based index.</summary>
    public VisioFluentDocument GetPage(int pageIndex, Action<VisioFluentPage> configure) {
        ConfigurePage(_document.GetPage(pageIndex), configure);
        return this;
    }

    /// <summary>Configures an existing page by name.</summary>
    public VisioFluentDocument GetPage(string name, Action<VisioFluentPage> configure,
        StringComparison comparison = StringComparison.OrdinalIgnoreCase) {
        ConfigurePage(_document.GetPage(name, comparison), configure);
        return this;
    }

    /// <summary>Configures a page belonging to this document.</summary>
    public VisioFluentDocument GetPage(VisioPage page, Action<VisioFluentPage> configure) {
        if (page == null) throw new ArgumentNullException(nameof(page));
        if (!_document.Pages.Contains(page)) throw new InvalidOperationException("The page does not belong to this document.");
        ConfigurePage(page, configure);
        return this;
    }

    /// <summary>Gets or adds a named page and optionally configures it.</summary>
    public VisioFluentDocument GetOrAddPage(string name, OfficeLength? width = null, OfficeLength? height = null,
        VisioMeasurementUnit scaleUnit = VisioMeasurementUnit.Inches, int? id = null,
        StringComparison comparison = StringComparison.OrdinalIgnoreCase, Action<VisioFluentPage>? configure = null) {
        ConfigurePage(_document.GetOrAddPage(name, width, height, scaleUnit, id, comparison), configure);
        return this;
    }

    /// <summary>Gets or adds a default-sized named page and configures it.</summary>
    public VisioFluentDocument GetOrAddPage(string name, Action<VisioFluentPage> configure,
        StringComparison comparison = StringComparison.OrdinalIgnoreCase) =>
        GetOrAddPage(name, width: null, comparison: comparison, configure: configure);

    private void ConfigurePage(VisioPage page, Action<VisioFluentPage>? configure) {
        var builder = new VisioFluentPage(this, page);
        configure?.Invoke(builder);
        builder.RebuildShapeIndex();
    }
}
