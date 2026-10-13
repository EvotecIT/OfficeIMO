using OfficeIMO.OpenDocument;

namespace OfficeIMO.Reader.OpenDocument;

internal static partial class OpenDocumentReaderAdapter {
    private static OfficeDocumentReadResult CompleteTextAndMetadata(OdfDocument document, string path, ReaderOptions options,
        ReaderOpenDocumentOptions formatOptions, OfficeDocumentReadResult result, CancellationToken token) {
        var metadata = result.Metadata.ToList();
        foreach (string entry in document.BinaryPayloadEntries) {
            token.ThrowIfCancellationRequested();
            metadata.Add(new OfficeDocumentMetadataEntry {
                Id = "odf-binary:" + entry, Category = "odf.payload", Name = "BinaryPayload", Value = "not-requested",
                ValueType = "coverage", Location = new ReaderLocation { Path = path, SourceBlockKind = "package-entry", BlockAnchor = entry },
                Attributes = ReaderInspectionMetadataBuilder.Attributes(("partPath", entry),
                    ("declaredLength", document.GetPackageEntryByteLength(entry).ToString(CultureInfo.InvariantCulture)),
                    ("nativeState", document.GetPackageEntryPayloadState(entry).ToString()), ("contentType", document.GetPackageEntryMediaType(entry)))
            });
        }
        metadata.Add(new OfficeDocumentMetadataEntry { Id = "odf-rich-projection", Category = "odf.projection",
            Name = "RichResources", Value = "not-requested", ValueType = "coverage" });
        result.Metadata = metadata.AsReadOnly();
        result.CapabilitiesUsed = result.CapabilitiesUsed.Where(value => value != "officeimo.reader.opendocument.native-rich-v1")
            .Concat(new[] { "officeimo.reader.opendocument.text-and-metadata" }).ToArray();
        return options.Inspection == null ? result : OpenDocumentInspectionProjection.Apply(document, path, options, formatOptions, result, token);
    }
}
