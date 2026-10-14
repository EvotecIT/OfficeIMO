using OfficeIMO.Pdf;

namespace OfficeIMO.Studio.Features.Workspace;

internal sealed partial class PdfWorkspace {
    private byte[]? _annotationMetadataSource;
    private IReadOnlyList<PdfAnnotation>? _annotationMetadata;

    internal IReadOnlyList<PdfAnnotation> AnnotationMetadata {
        get {
            if (DocumentInfo is { } info) return info.Annotations;
            lock (_capabilityGate) {
                byte[] source = _bytes;
                if (!ReferenceEquals(_annotationMetadataSource, source)) {
                    _annotationMetadata = null;
                    _annotationMetadataSource = source;
                }
                return _annotationMetadata ??= CanEditAnnotations ? LoadDocument(source).Annotations.GetForEditing() : [];
            }
        }
    }
}
