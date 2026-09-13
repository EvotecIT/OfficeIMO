using OfficeIMO.Drawing;

namespace OfficeIMO.Project;

public sealed partial class ProjectView {
    private readonly OfficeRenderingProfile? _typography;
    private OfficeRasterCanvas? _measurement;

    private ProjectView(ProjectView source, OfficeRenderingProfile typography) {
        Title = source.Title; ModelRevision = source.ModelRevision; Layout = source.Layout; Kind = source.Kind;
        Columns = source.Columns; Rows = source.Rows; Buckets = source.Buckets; Links = source.Links; Report = source.Report;
        _typography = typography;
    }

    /// <summary>Measures and renders with the same supplied font faces and shaping profile. The original view remains immutable and can be rendered concurrently.</summary>
    public IReadOnlyList<ProjectViewPage> Render(OfficeRenderingProfile typography, CancellationToken cancellationToken = default) {
        if (typography == null) throw new ArgumentNullException(nameof(typography));
        return new ProjectView(this, typography).Render(cancellationToken);
    }

    private double MeasureWidth(string? value, double size, bool bold = false) {
        // The canvas owns the same font resolution and shaping path used by PNG rendering.
        _measurement ??= new OfficeRasterCanvas(new OfficeRasterImage(1, 1), font: null, fonts: _typography?.Fonts,
            textShapingProvider: _typography?.TextShapingProvider ?? OfficeManagedTextShapingProvider.Instance,
            textShapingLanguage: _typography?.TextShapingLanguage);
        return _measurement.MeasureText(value, size, "Arial", bold ? OfficeFontStyle.Bold : OfficeFontStyle.Regular);
    }
}
