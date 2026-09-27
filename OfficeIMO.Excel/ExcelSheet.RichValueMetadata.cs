using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using DynamicArray = DocumentFormat.OpenXml.Office2019.Excel.DynamicArray;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private const string RichValueMetadataName = "XLRICHVALUE";
        private const string RichValueMetadataExtensionUri = "{3e2802c4-a4d2-4d8b-9148-e3be6c30e623}";
        private const string DynamicArrayMetadataName = "XLDAPR";
        private const string DynamicArrayMetadataExtensionUri = "{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}";

        private uint EnsureDynamicArrayMetadata() {
            var workbook = _excelDocument.WorkbookPartRoot;
            var part = workbook.CellMetadataPart ?? workbook.AddNewPart<DocumentFormat.OpenXml.Packaging.CellMetadataPart>();
            Metadata metadata = part.Metadata ??= new Metadata();
            MetadataTypes types = metadata.MetadataTypes ??= new MetadataTypes();
            MetadataType[] existingTypes = types.Elements<MetadataType>().ToArray();
            int typeIndex = Array.FindIndex(existingTypes, type => type.Name?.Value == DynamicArrayMetadataName);
            if (typeIndex < 0) {
                typeIndex = existingTypes.Length;
                types.Append(new MetadataType {
                    Name = DynamicArrayMetadataName,
                    MinSupportedVersion = 120000U,
                    Copy = true, PasteAll = true, PasteValues = true, Merge = true,
                    SplitFirst = true, RowColumnShift = true, ClearFormats = true,
                    ClearComments = true, Assign = true, Coerce = true, CellMeta = true
                });
                types.Count = (uint)typeIndex + 1;
            }
            FutureMetadata future = EnsureFutureMetadata(metadata, DynamicArrayMetadataName);
            FutureMetadataBlock[] futureBlocks = future.Elements<FutureMetadataBlock>().ToArray();
            int futureIndex = Array.FindIndex(futureBlocks, block => block.Descendants<DynamicArray.DynamicArrayProperties>()
                .Any(properties => properties.FDynamic?.Value == true && properties.FCollapsed?.Value != true));
            if (futureIndex < 0) {
                futureIndex = futureBlocks.Length;
                var extension = new Extension { Uri = DynamicArrayMetadataExtensionUri };
                extension.Append(new DynamicArray.DynamicArrayProperties { FDynamic = true, FCollapsed = false });
                future.Append(new FutureMetadataBlock(new ExtensionList(extension)));
                future.Count = (uint)futureIndex + 1;
            }
            CellMetadata cells = metadata.GetFirstChild<CellMetadata>() ?? metadata.AppendChild(new CellMetadata());
            MetadataBlock[] blocks = cells.Elements<MetadataBlock>().ToArray();
            int blockIndex = Array.FindIndex(blocks, block => block.Elements<MetadataRecord>().Count() == 1
                && block.Elements<MetadataRecord>().Any(record => record.TypeIndex?.Value == (uint)typeIndex + 1
                    && record.Val?.Value == (uint)futureIndex));
            if (blockIndex < 0) {
                blockIndex = blocks.Length;
                cells.Append(new MetadataBlock(new MetadataRecord {
                    TypeIndex = (uint)typeIndex + 1, Val = (uint)futureIndex
                }));
                cells.Count = (uint)blockIndex + 1;
            }
            part.Metadata.Save();
            return (uint)blockIndex + 1;
        }

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
            return EnsureFutureMetadata(metadata, RichValueMetadataName);
        }

        private static FutureMetadata EnsureFutureMetadata(Metadata metadata, string name) {
            FutureMetadata? future = metadata.Elements<FutureMetadata>()
                .FirstOrDefault(item => string.Equals(item.Name?.Value, name, StringComparison.OrdinalIgnoreCase));
            if (future != null) return future;
            future = new FutureMetadata { Name = name, Count = 0U };
            OpenXmlElement? metadataBlocks = metadata.ChildElements
                .FirstOrDefault(element => element is CellMetadata || element is ValueMetadata);
            if (metadataBlocks == null) metadata.Append(future); else metadata.InsertBefore(future, metadataBlocks);
            return future;
        }

    }
}
