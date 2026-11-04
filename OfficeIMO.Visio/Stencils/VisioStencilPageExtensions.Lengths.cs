namespace OfficeIMO.Visio.Stencils;

public static partial class VisioStencilPageExtensions {
    /// <summary>Creates the shape using explicit physical coordinates and dimensions.</summary>
    public static VisioShape AddStencilShape(this VisioPage page, VisioStencilShape stencil, string id, OfficeLength x, OfficeLength y, string? text = null) {
        ArgumentNullException.ThrowIfNull(page);
        return AddStencilShapeInches(page, stencil, id, VisioMeasurementUnitExtensions.FromInches(x.ToInches(), page.DefaultUnit), VisioMeasurementUnitExtensions.FromInches(y.ToInches(), page.DefaultUnit), text);
    }

    /// <summary>Creates the shape using explicit physical coordinates and dimensions.</summary>
    public static VisioShape AddStencilShape(this VisioPage page, VisioStencilShape stencil, string id, OfficeLength x, OfficeLength y, OfficeLength width, OfficeLength height, string? text = null) {
        ArgumentNullException.ThrowIfNull(page);
        return AddStencilShapeInches(page, stencil, id, VisioMeasurementUnitExtensions.FromInches(x.ToInches(), page.DefaultUnit), VisioMeasurementUnitExtensions.FromInches(y.ToInches(), page.DefaultUnit), VisioMeasurementUnitExtensions.FromInches(width.ToInches(), page.DefaultUnit), VisioMeasurementUnitExtensions.FromInches(height.ToInches(), page.DefaultUnit), text);
    }

    /// <summary>Creates the shape using explicit physical coordinates and dimensions.</summary>
    public static VisioShape AddStencilShape(this VisioPage page, VisioStencilCatalog catalog, string stencilIdOrName, string id, OfficeLength x, OfficeLength y, string? text = null) {
        ArgumentNullException.ThrowIfNull(page);
        return AddStencilShapeInches(page, catalog, stencilIdOrName, id, VisioMeasurementUnitExtensions.FromInches(x.ToInches(), page.DefaultUnit), VisioMeasurementUnitExtensions.FromInches(y.ToInches(), page.DefaultUnit), text);
    }

    /// <summary>Creates the shape using explicit physical coordinates and dimensions.</summary>
    public static VisioShape AddStencilShape(this VisioPage page, VisioStencilCatalog catalog, string stencilIdOrName, string id, OfficeLength x, OfficeLength y, OfficeLength width, OfficeLength height, string? text = null) {
        ArgumentNullException.ThrowIfNull(page);
        return AddStencilShapeInches(page, catalog, stencilIdOrName, id, VisioMeasurementUnitExtensions.FromInches(x.ToInches(), page.DefaultUnit), VisioMeasurementUnitExtensions.FromInches(y.ToInches(), page.DefaultUnit), VisioMeasurementUnitExtensions.FromInches(width.ToInches(), page.DefaultUnit), VisioMeasurementUnitExtensions.FromInches(height.ToInches(), page.DefaultUnit), text);
    }

    /// <summary>Creates the shape using explicit physical coordinates and dimensions.</summary>
    public static VisioShape AddStencilShape(this VisioPage page, string stencilIdOrName, string id, OfficeLength x, OfficeLength y, string? text = null) {
        ArgumentNullException.ThrowIfNull(page);
        return AddStencilShapeInches(page, stencilIdOrName, id, VisioMeasurementUnitExtensions.FromInches(x.ToInches(), page.DefaultUnit), VisioMeasurementUnitExtensions.FromInches(y.ToInches(), page.DefaultUnit), text);
    }

    /// <summary>Creates the shape using explicit physical coordinates and dimensions.</summary>
    public static VisioShape AddStencilShape(this VisioPage page, string stencilIdOrName, string id, OfficeLength x, OfficeLength y, OfficeLength width, OfficeLength height, string? text = null) {
        ArgumentNullException.ThrowIfNull(page);
        return AddStencilShapeInches(page, stencilIdOrName, id, VisioMeasurementUnitExtensions.FromInches(x.ToInches(), page.DefaultUnit), VisioMeasurementUnitExtensions.FromInches(y.ToInches(), page.DefaultUnit), VisioMeasurementUnitExtensions.FromInches(width.ToInches(), page.DefaultUnit), VisioMeasurementUnitExtensions.FromInches(height.ToInches(), page.DefaultUnit), text);
    }
}
