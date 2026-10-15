using OfficeIMO.Reader.FormatInternals;
using OfficeIMO.Word;
using System.Globalization;

namespace OfficeIMO.Reader.Word;

internal static class WordInspectionProjection {
    internal static OfficeDocumentReadResult Apply(WordDocument document, ReaderOptions options,
        OfficeDocumentReadResult result, CancellationToken cancellationToken) {
        if (options.Inspection == null) return result;
        var builder = new ReaderInspectionMetadataBuilder(options, "word", cancellationToken);
        string? path = result.Source.Path;
        if (builder.Options.IncludeComments || builder.Options.IncludeRevisions) {
            cancellationToken.ThrowIfCancellationRequested();
            WordReviewInfo review = document.InspectReview(cancellationToken);
            if (builder.Options.IncludeComments) foreach (WordCommentInfo comment in review.Comments) {
                builder.Add(OfficeDocumentInspectionCategories.Comment, "Comment", comment.Text, "text",
                    new ReaderLocation { Path = path, SourceBlockKind = "comment", BlockAnchor = comment.Id }, comment.Id,
                    ReaderInspectionMetadataBuilder.Attributes(
                        ("author", comment.Author), ("initials", comment.Initials),
                        ("date", comment.DateTime?.ToString("O", CultureInfo.InvariantCulture)),
                        ("paragraphId", comment.ParaId), ("parentParagraphId", comment.ParentParaId),
                        ("isResolved", comment.IsResolved?.ToString().ToLowerInvariant()),
                        ("isReply", comment.IsReply ? "true" : "false"),
                        ("targetPartUri", comment.TargetPartUri), ("story", comment.TargetLocationKind?.ToString()),
                        ("targetText", comment.TargetText)));
            }
            if (builder.Options.IncludeRevisions) foreach (WordRevisionInfo revision in review.Revisions) {
                builder.Add(OfficeDocumentInspectionCategories.Revision, revision.RevisionType.ToString(), revision.AffectedText, "text",
                    new ReaderLocation { Path = path, SourceBlockKind = "revision", BlockAnchor = revision.Id }, revision.Id,
                    ReaderInspectionMetadataBuilder.Attributes(
                        ("author", revision.Author), ("date", revision.DateTime?.ToString("O", CultureInfo.InvariantCulture)),
                        ("partUri", revision.PartUri), ("story", revision.LocationKind.ToString()),
                        ("elementName", revision.ElementName), ("locationText", revision.LocationText),
                        ("acceptedOrRejected", "false")));
            }
            foreach (string limitation in review.UnsupportedMetadata)
                builder.Diagnostic("INSPECTION_REVIEW_UNSUPPORTED", limitation, new ReaderLocation { Path = path });
            builder.Coverage("review", review.UnsupportedMetadata.Count == 0 ? "complete" : "partial",
                "Comments and revisions in supported native stories; selection options apply.", new ReaderLocation { Path = path });
        }
        if (builder.Options.IncludePackageContent) {
            NativePackageInspectionProjection.Add(builder, document.SourcePackageSecurityReport, document.SourcePackageSecurityScope, path);
            if (options.Projection == ReaderProjection.TextAndMetadata) {
                NativeInspectionProjection.AddBoundedPayloads(builder, document.InspectEmbeddedPayloads, path, cancellationToken);
                NativeInspectionProjection.AddVba(builder, document.InspectVbaProjectMetadata(), path);
            } else {
                NativeInspectionProjection.AddPayloads(builder, document.GetEmbeddedPayloads(), path);
                NativeInspectionProjection.AddVba(builder, document.InspectVbaProject(), path);
            }
        }
        if (builder.Options.IncludeContentSafety)
            NativeInspectionProjection.InspectContentSafety(builder, safety =>
                document.InspectLoadedContentSafety(safety, cancellationToken), path);
        else if (builder.Options.IncludeVisibility)
            builder.Coverage("visibility", "not-inspected", "Word run concealment requires IncludeContentSafety; no hidden state is inferred from extracted text.", new ReaderLocation { Path = path });
        return builder.Complete(result);
    }
}
