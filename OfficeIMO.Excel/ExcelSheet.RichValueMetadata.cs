using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private const string RichValueMetadataName = "XLRICHVALUE";
        private const string RichValueMetadataExtensionUri = "{3e2802c4-a4d2-4d8b-9148-e3be6c30e623}";

        private static uint EnsureRichValueMetadataType(Metadata metadata) {
            MetadataTypes types = metadata.MetadataTypes ??= new MetadataTypes();
            List<MetadataType> existing = types.Elements<MetadataType>().ToList();
            int index = existing.FindIndex(type => string.Equals(type.Name?.Value, RichValueMetadataName, StringComparison.OrdinalIgnoreCase));
            if (index < 0) {
                types.Append(new MetadataType {
                    Name = RichValueMetadataName,
                    MinSupportedVersion = 120000U,
                    Copy = true,
                    PasteAll = true,
                    PasteValues = true,
                    Merge = true,
                    SplitFirst = true,
                    RowColumnShift = true,
                    ClearFormats = true,
                    ClearComments = true,
                    Assign = true,
                    Coerce = true
                });
                index = existing.Count;
            }
            types.Count = (uint)types.Elements<MetadataType>().Count();
            return (uint)index + 1U;
        }

        private static FutureMetadata EnsureRichValueFutureMetadata(Metadata metadata) {
            FutureMetadata? future = metadata.Elements<FutureMetadata>()
                .FirstOrDefault(item => string.Equals(item.Name?.Value, RichValueMetadataName, StringComparison.OrdinalIgnoreCase));
            if (future != null) return future;
            future = new FutureMetadata { Name = RichValueMetadataName, Count = 0U };
            OpenXmlElement? metadataBlocks = metadata.ChildElements
                .FirstOrDefault(element => element is CellMetadata || element is ValueMetadata);
            if (metadataBlocks == null) metadata.Append(future); else metadata.InsertBefore(future, metadataBlocks);
            return future;
        }

    }
}
