using OfficeIMO.Core.Internal;

namespace OfficeIMO.Excel.LegacyXls.Write {
    internal static partial class LegacyXlsWriter {
        /// <summary>Prepares the compound layout before publishing the validated, owned BIFF buffer.</summary>
        private static void WriteDirectCompoundFile(
            ExcelDocument document,
            DirectTabularWorkbookStream workbook,
            Stream destination) {
            IReadOnlyList<OfficeCompoundStream> properties = LegacyOlePropertySetWriter.CreateDocumentPropertyStreams(document);
            var streams = new List<OfficeCompoundStream>(properties.Count + 1) {
                new OfficeCompoundStream(GetWorkbookStreamName(null), workbook.Buffer, checked((int)workbook.Length))
            };
            streams.AddRange(properties);
            OfficeCompoundWriterLayout layout = OfficeCompoundWriterLayout.Create(streams);
            long length = OfficeCompoundFileWriter.GetSerializedLength(layout);
            if (length > int.MaxValue) {
                throw new NotSupportedException("A compound file returned as a byte array cannot exceed 2 GiB.");
            }
            ExcelDocument.PrepareDestinationStreamForWrite(destination);
            // A single staged write previously reserved the full size. Keep that exact
            // reservation for ordinary memory destinations instead of growing per sector.
            if (destination.GetType() == typeof(MemoryStream)
                && ((MemoryStream)destination).Capacity < length) {
                ((MemoryStream)destination).Capacity = checked((int)length);
            }
            OfficeCompoundFileWriter.Write(destination, layout, rootClassId: null);
        }
    }
}
