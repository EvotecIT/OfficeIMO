using DocumentFormat.OpenXml.Spreadsheet;
using System.Xml;

namespace OfficeIMO.Excel {
    internal sealed partial class ExcelSheetReader {
        /// <summary>Reads the first sheet view without materializing worksheet cells.</summary>
        internal bool ReadRightToLeft(CancellationToken cancellationToken) {
            cancellationToken.ThrowIfCancellationRequested();
            if (_hasSdkWorksheetPart && _wsPart.IsRootElementLoaded) {
                return _wsPart.Worksheet?.GetFirstChild<SheetViews>()?.GetFirstChild<SheetView>()
                    ?.RightToLeft?.Value == true;
            }

            using Stream stream = OpenReaderWorksheetStream(cancellationToken);
            var settings = CreateWorksheetXmlReaderSettings();
            settings.MaxCharactersInDocument = _opt.MaxMetadataPartBytes;
            using XmlReader reader = XmlReader.Create(stream, OpenXmlReadNameTable.WithSchemaNames(settings));
            var worksheet = new WorksheetXmlRowSelector();
            int viewsDepth = -1;
            while (reader.Read()) {
                cancellationToken.ThrowIfCancellationRequested();
                if (worksheet.IsWorksheetChildElement(reader, "sheetData")) return false;
                if (worksheet.IsWorksheetChildElement(reader, "sheetViews")) viewsDepth = reader.Depth;
                if (viewsDepth >= 0 && SpreadsheetXmlContent.IsDirectChildElement(reader, viewsDepth, "sheetView")) {
                    string? value = reader.GetAttribute("rightToLeft");
                    return value != null && XmlConvert.ToBoolean(value);
                }
                if (reader.NodeType == XmlNodeType.EndElement && reader.Depth == viewsDepth) viewsDepth = -1;
            }
            return false;
        }
    }
}
