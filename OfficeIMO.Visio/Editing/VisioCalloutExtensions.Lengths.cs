namespace OfficeIMO.Visio;

public static partial class VisioCalloutExtensions {
    /// <summary>Creates the shape using explicit physical coordinates and dimensions.</summary>
    public static VisioShape AddCallout(this VisioPage page, VisioShape target, string text, OfficeLength pinX, OfficeLength pinY, VisioCalloutOptions? options = null) {
        ArgumentNullException.ThrowIfNull(page);
        return AddCalloutInches(page, target, text, VisioMeasurementUnitExtensions.FromInches(pinX.ToInches(), page.DefaultUnit), VisioMeasurementUnitExtensions.FromInches(pinY.ToInches(), page.DefaultUnit), options);
    }

    /// <summary>Creates the shape using explicit physical coordinates and dimensions.</summary>
    public static VisioShape AddCallout(this VisioPage page, VisioShape target, string text, VisioSide placement, OfficeLength? gap = null, VisioCalloutOptions? options = null) {
        ArgumentNullException.ThrowIfNull(page);
        return AddCalloutInches(page, target, text, placement, VisioMeasurementUnitExtensions.FromInches((gap ?? OfficeLength.Inches(0.35D)).ToInches(), page.DefaultUnit), options);
    }

    /// <summary>Creates the shape using explicit physical coordinates and dimensions.</summary>
    public static VisioShape AddCallout(this VisioPage page, VisioShape target, string id, string text, OfficeLength pinX, OfficeLength pinY, VisioCalloutOptions? options = null) {
        ArgumentNullException.ThrowIfNull(page);
        return AddCalloutInches(page, target, id, text, VisioMeasurementUnitExtensions.FromInches(pinX.ToInches(), page.DefaultUnit), VisioMeasurementUnitExtensions.FromInches(pinY.ToInches(), page.DefaultUnit), options);
    }

    /// <summary>Creates the shape using explicit physical coordinates and dimensions.</summary>
    public static VisioShape AddCallout(this VisioPage page, VisioShape target, string id, string text, VisioSide placement, OfficeLength? gap = null, VisioCalloutOptions? options = null) {
        ArgumentNullException.ThrowIfNull(page);
        return AddCalloutInches(page, target, id, text, placement, VisioMeasurementUnitExtensions.FromInches((gap ?? OfficeLength.Inches(0.35D)).ToInches(), page.DefaultUnit), options);
    }
}
