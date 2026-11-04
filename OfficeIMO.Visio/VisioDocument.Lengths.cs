namespace OfficeIMO.Visio;

public partial class VisioDocument {
    /// <summary>Adds a page with physical dimensions and explicit native drawing-scale units.</summary>
    public VisioPage AddPage(string name, OfficeLength? width = null, OfficeLength? height = null, VisioMeasurementUnit scaleUnit = VisioMeasurementUnit.Inches, int? id = null) {
        VisioPage page = AddPageInches(name, width?.ToInches() ?? 8.26771653543307, height?.ToInches() ?? 11.69291338582677, VisioMeasurementUnit.Inches, id);
        page.DefaultUnit = scaleUnit;
        page.ScaleMeasurementUnit = scaleUnit;
        return page;
    }

    /// <summary>Adds a page with physical dimensions and explicit native drawing-scale units.</summary>
    public VisioPage AddBackgroundPage(string name, OfficeLength? width = null, OfficeLength? height = null, VisioMeasurementUnit scaleUnit = VisioMeasurementUnit.Inches, int? id = null) {
        VisioPage page = AddBackgroundPageInches(name, width?.ToInches() ?? 8.26771653543307, height?.ToInches() ?? 11.69291338582677, VisioMeasurementUnit.Inches, id);
        page.DefaultUnit = scaleUnit;
        page.ScaleMeasurementUnit = scaleUnit;
        return page;
    }
}
