using OfficeIMO.Core.Internal;
using OfficeIMO.Word.LegacyDoc.Model;

namespace OfficeIMO.Word.LegacyDoc.Write {
    internal static partial class LegacyDocWriter {
        internal static byte[] WriteDocument(WordDocument document, WordSaveOptions? options = null,
            bool isTemplate = false) => WriteDocument(document, out _, options, isTemplate);

        internal static byte[] WriteDocument(WordDocument document, out OfficeConversionReport report,
            WordSaveOptions? options = null, bool isTemplate = false) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            var timer = System.Diagnostics.Stopwatch.StartNew();
            using var palette = new LegacyDocPaletteProjection();
            byte[] bytes = WriteDocumentCore(document, options, isTemplate);
            report = palette.CreateReport(document, timer.Elapsed);
            if (options?.AllowLoss != true) report.ThrowIfLoss();
            return bytes;
        }

        private static byte[] WriteDocumentCore(
            WordDocument document,
            WordSaveOptions? options = null,
            bool isTemplate = false) {
            if (document == null) throw new ArgumentNullException(nameof(document));

            ThrowIfUnsupportedLegacyDocImportState(document, options);

            LegacyDocWritableBody body = BuildBody(document);
            byte[] wordDocumentStream = PadToRegularOleStream(CreateWordDocumentStream(body, isTemplate));
            byte[] tableStream = PadToRegularOleStream(CreateTableStream(body));
            IReadOnlyList<OfficeCompoundStream> propertyStreams = LegacyDocPropertySetWriter.CreateDocumentPropertyStreams(document);
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
                streams.ToDictionary(stream => stream.Name, stream => stream.Bytes, StringComparer.OrdinalIgnoreCase));
        }

    }
}
