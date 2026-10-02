using System.Threading;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private static PdfFileVersion RequireStructureAssociatedFileVersion(LayoutResult layout, PdfOptions options) {
        if (!layout.Pages.Any(page => page.StructElements.Any(element => element.AssociatedFiles.Count > 0)))
            return options.FileVersion;
        // PDF/A-3 introduced AF with PDF 1.7. Plain PDFs use the PDF 2.0 mechanism;
        // an identification dictionary alone is never treated as conformance proof.
        PdfFileVersion minimum = options.PdfAIdentificationSnapshot?.Part == 3
            ? PdfFileVersion.Pdf17 : PdfFileVersion.Pdf20;
        return PdfFileAssembler.RequireAtLeast(options.FileVersion, minimum);
    }

    private static void ValidateAssociatedFileMetadata(PdfOptions options, PdfFileVersion version) {
        if (version < PdfFileVersion.Pdf20) return;
        foreach (PdfEmbeddedFile file in options.EmbeddedFileSnapshots) {
            if (string.IsNullOrWhiteSpace(file.MimeType))
                throw new InvalidOperationException("A PDF 2.0 associated catalog file requires a MIME type: " + file.FileName);
        }
    }

    private static List<(string FileName, int FileSpecId)> BuildStructureAssociatedFiles(
        IList<byte[]> objects,
        IReadOnlyList<LayoutResult.Page> pages,
        IReadOnlyList<PdfEmbeddedFile> catalogFiles,
        CancellationToken cancellationToken) {
        var entries = new List<(string FileName, int FileSpecId)>();
        var files = new Dictionary<string, (PdfEmbeddedFile File, int Id)>(StringComparer.Ordinal);
        var catalogNames = new HashSet<string>(catalogFiles.Select(file => file.FileName), StringComparer.Ordinal);
        foreach (LayoutResult.Page page in pages) {
            foreach (PageStructElement element in page.StructElements) {
                cancellationToken.ThrowIfCancellationRequested();
                if (element.AssociatedFiles.Count == 0) continue;
                var ids = new List<int>(element.AssociatedFiles.Count);
                foreach (PdfEmbeddedFile file in element.AssociatedFiles) {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (catalogNames.Contains(file.FileName))
                        throw new InvalidOperationException("A structure-associated file conflicts with a catalog attachment name: " + file.FileName);
                    if (files.TryGetValue(file.FileName, out var existing)) {
                        if (!file.HasSameDescriptionAndData(existing.File))
                            throw new InvalidOperationException("Different structure-associated files have the same name: " + file.FileName);
                        ids.Add(existing.Id);
                        continue;
                    }
                    byte[] data = file.DataSnapshot;
                    int streamId = AddStreamObject(objects,
                        PdfEmbeddedFileDictionaryBuilder.BuildEmbeddedFileStreamDictionary(file, data, omitUndatedParameters: true), data);
                    int specId = AddObject(objects, PdfEmbeddedFileDictionaryBuilder.BuildFileSpecificationObject(file, streamId));
                    files.Add(file.FileName, (file, specId));
                    entries.Add((file.FileName, specId));
                    ids.Add(specId);
                }
                element.AssociatedFileIds = ids;
            }
        }
        return entries;
    }
}
