using OfficeIMO.Core.Internal;
using OfficeIMO.Word.LegacyDoc.Model;

namespace OfficeIMO.Word.LegacyDoc.Write {
    internal static partial class LegacyDocWriter {
        internal static byte[] WriteDocument(WordDocument document, WordSaveOptions? options = null,
            bool isTemplate = false) => WriteDocument(document, out _, options, isTemplate);

        internal static byte[] WriteDocument(WordDocument document, out OfficeConversionReport report,
            WordSaveOptions? options = null, bool isTemplate = false) {
            ArgumentNullException.ThrowIfNull(document);
            var timer = System.Diagnostics.Stopwatch.StartNew();
            using var palette = new LegacyDocPaletteProjection();
            byte[] bytes = WriteDocumentCore(document, options, isTemplate, out int sectionCount);
            report = palette.CreateReport(document, sectionCount, timer.Elapsed);
            if (options?.AllowLoss != true) report.ThrowIfLoss();
            return bytes;
        }

        private static byte[] WriteDocumentCore(
            WordDocument document,
            WordSaveOptions? options,
            bool isTemplate,
            out int sectionCount) {
            ArgumentNullException.ThrowIfNull(document);

            if (document.IsWord69Source) {
                throw new NotSupportedException("Word 6/95 sources support migration to DOCX. Saving them as binary DOC is unavailable.");
            }

            ThrowIfUnsupportedLegacyDocImportState(document, options);

            using LegacyDocSourcePackage source = document.OpenLegacyDocSource();
            LegacyDocSourcePackage? previousSource = t_writingSource;
            t_writingSource = source;
            try {
                sectionCount = LegacyDocWordHelpers.GetDocumentSectionProperties(source).Count;
                return WriteDocumentBody(document, source, options, isTemplate);
            } finally {
                t_writingSource = previousSource;
            }
        }

        private static byte[] WriteDocumentBody(WordDocument document, LegacyDocSourcePackage source, WordSaveOptions? options, bool isTemplate) {
            LegacyDocWritableBody body = BuildBody(document, source);
            byte[] wordDocumentStream = PadToRegularOleStream(CreateWordDocumentStream(body, isTemplate));
            byte[] tableStream = PadToRegularOleStream(CreateTableStream(body));
            IReadOnlyList<OfficeCompoundStream> propertyStreams = LegacyDocPropertySetWriter.CreateDocumentPropertyStreams(document, source);
            var streams = new List<OfficeCompoundStream>(propertyStreams.Count + 2) {
                new OfficeCompoundStream("WordDocument", wordDocumentStream),
                new OfficeCompoundStream("1Table", tableStream)
            };
            if (body.HasPictures) {
                streams.Add(new OfficeCompoundStream("Data", body.PictureData));
            }
            foreach (OfficeCompoundStream propertyStream in propertyStreams) {
                streams.Add(new OfficeCompoundStream(propertyStream.Name, PadToRegularOleStream(propertyStream.Bytes)));
            }

            OfficeCompoundFile? sourceCompoundFile = document.LegacyDocSourceCompoundFile;
            if (sourceCompoundFile == null) {
                return OfficeCompoundFileWriter.Write(streams);
            }

            return OfficeCompoundFileWriter.Rewrite(
                sourceCompoundFile,
                streams.ToDictionary(stream => stream.Name, stream => stream.Bytes, StringComparer.OrdinalIgnoreCase),
                options?.SignatureMutationPolicy == OfficeSignatureMutationPolicy.RemoveInvalidatedSignatures
                    ? GetSignatureEntryPaths(sourceCompoundFile)
                    : null);
        }

        private static string[] GetSignatureEntryPaths(OfficeCompoundFile sourceCompoundFile) =>
            sourceCompoundFile.Entries
                .Where(entry => entry.Name.Equals("_signatures", StringComparison.OrdinalIgnoreCase)
                    || entry.Name.Equals("_xmlsignatures", StringComparison.OrdinalIgnoreCase))
                .Select(entry => entry.Path)
                .ToArray();

    }
}
