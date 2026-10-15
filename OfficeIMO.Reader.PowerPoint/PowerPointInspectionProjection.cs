using OfficeIMO.PowerPoint;
using OfficeIMO.Reader.FormatInternals;
using System.Globalization;

namespace OfficeIMO.Reader.PowerPoint;

internal static class PowerPointInspectionProjection {
    internal static OfficeDocumentReadResult Apply(PowerPointPresentation presentation, ReaderOptions options,
        OfficeDocumentReadResult result, CancellationToken cancellationToken) {
        if (options.Inspection == null) return result;
        var builder = new ReaderInspectionMetadataBuilder(options, "powerpoint", cancellationToken);
        string? path = result.Source.Path;
        if (builder.Options.IncludeComments) {
            cancellationToken.ThrowIfCancellationRequested();
            foreach (PowerPointReviewComment comment in presentation.InspectReviewComments(cancellationToken).Comments) {
                builder.Add(OfficeDocumentInspectionCategories.Comment, comment.Kind.ToString(), comment.Text, "text",
                    new ReaderLocation { Path = path, Slide = comment.SlideNumber, SourceBlockKind = "comment" }, comment.Id,
                    ReaderInspectionMetadataBuilder.Attributes(
                        ("author", comment.AuthorName), ("authorId", comment.AuthorId), ("parentId", comment.ParentId),
                        ("date", comment.Created?.ToString("O", CultureInfo.InvariantCulture)), ("status", comment.Status),
                        ("shapeId", comment.ShapeId?.ToString(CultureInfo.InvariantCulture)),
                        ("x", comment.X?.ToString(CultureInfo.InvariantCulture)), ("y", comment.Y?.ToString(CultureInfo.InvariantCulture))));
            }
            builder.Coverage("comments", "complete", "Classic and modern comments in the loaded presentation model.", new ReaderLocation { Path = path });
        }
        if (builder.Options.IncludeVisibility) for (int index = 0; index < presentation.Slides.Count; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            PowerPointSlide slide = presentation.Slides[index];
            var location = new ReaderLocation { Path = path, Slide = index + 1, SourceBlockKind = "slide" };
            if (slide.Hidden) builder.Add(OfficeDocumentInspectionCategories.Visibility, "Slide", "hidden", "visibility-state", location,
                attributes: ReaderInspectionMetadataBuilder.Attributes(("name", slide.Name)));
            foreach (PowerPointShape shape in slide.EnumerateShapesDeep(slide.Shapes, includeHidden: true)) {
                if (!shape.Hidden) continue;
                builder.Add(OfficeDocumentInspectionCategories.Visibility, "Shape", "hidden", "visibility-state",
                    new ReaderLocation { Path = path, Slide = index + 1, SourceBlockKind = "shape" }, shape.Id?.ToString(CultureInfo.InvariantCulture),
                    ReaderInspectionMetadataBuilder.Attributes(("name", shape.Name)));
            }
        }
        if (builder.Options.IncludePackageContent) {
            NativePackageInspectionProjection.Add(builder, presentation.SourcePackageSecurityReport, presentation.SourcePackageSecurityScope, path);
            if (options.Projection == ReaderProjection.TextAndMetadata) {
                NativeInspectionProjection.AddBoundedPayloads(builder, presentation.InspectEmbeddedPayloads, path, cancellationToken);
                NativeInspectionProjection.AddVba(builder, presentation.InspectVbaProjectMetadata(), path);
            } else {
                NativeInspectionProjection.AddPayloads(builder, presentation.GetEmbeddedPayloads(), path);
                NativeInspectionProjection.AddVba(builder, presentation.InspectVbaProject(), path);
            }
        }
        if (builder.Options.IncludeContentSafety)
            NativeInspectionProjection.InspectContentSafety(builder, safety =>
                presentation.InspectLoadedContentSafety(safety, cancellationToken), path);
        return builder.Complete(result);
    }
}
