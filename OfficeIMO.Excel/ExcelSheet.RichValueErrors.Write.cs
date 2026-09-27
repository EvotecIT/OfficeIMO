using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Rich = DocumentFormat.OpenXml.Office2019.Excel.RichData;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private bool TryWriteRichFormulaError(Cell cell, string? error) {
            if (error != "#CALC!" && error != "#SPILL!") return false;
            WorkbookPart workbook = _excelDocument.WorkbookPartRoot;
            CellMetadataPart metadataPart = workbook.CellMetadataPart ?? workbook.AddNewPart<CellMetadataPart>();
            Metadata metadata = metadataPart.Metadata ??= new Metadata();
            uint typeIndex = EnsureRichValueMetadataType(metadata);
            FutureMetadata future = EnsureRichValueFutureMetadata(metadata);
            ValueMetadata blocks = metadata.GetFirstChild<ValueMetadata>() ?? metadata.AppendChild(new ValueMetadata());
            RdRichValuePart valuePart = workbook.RdRichValueParts.FirstOrDefault() ?? workbook.AddNewPart<RdRichValuePart>();
            Rich.RichValueData values = valuePart.RichValueData ??= new Rich.RichValueData();
            RdRichValueStructurePart structurePart = workbook.GetPartsOfType<RdRichValueStructurePart>().FirstOrDefault() ?? workbook.AddNewPart<RdRichValueStructurePart>();
            Rich.RichValueStructures structures = structurePart.RichValueStructures ??= new Rich.RichValueStructures();
            bool emptyArray = error == "#CALC!" && cell.CellFormula?.FormulaType?.Value == CellFormulaValues.Array;
            bool changed = false;
            string detailName = emptyArray ? "subType" : "propagated";
            string detailType = emptyArray ? "i" : "b";
            string detailValue = emptyArray ? "3" : "1";
            string errorType = error == "#CALC!" ? "13" : "8";
            var existingStructures = structures.Elements<Rich.RichValueStructure>().ToList();
            int structureIndex = existingStructures.FindIndex(item => item.T?.Value == "_error"
                && item.Elements<Rich.Key>().Select(key => key.N?.Value + ":" + key.GetAttributes()
                    .FirstOrDefault(attribute => attribute.LocalName == "t" && string.IsNullOrEmpty(attribute.NamespaceUri)).Value)
                    .SequenceEqual(new[] { "errorType:i", detailName + ":" + detailType }));
            if (structureIndex < 0) {
                changed = true;
                structureIndex = existingStructures.Count;
                structures.Append(new Rich.RichValueStructure(
                    new Rich.Key { N = "errorType", T = Rich.RichValueValueType.I },
                    new Rich.Key { N = detailName, T = emptyArray ? Rich.RichValueValueType.I : Rich.RichValueValueType.B }) { T = "_error" });
            }
            var existingValues = values.Elements<Rich.RichValue>().ToList();
            int valueIndex = existingValues.FindIndex(item => item.S?.Value == (uint)structureIndex
                && item.Elements<Rich.Value>().Select(value => value.Text).SequenceEqual(new[] { errorType, detailValue }));
            if (valueIndex < 0) {
                changed = true;
                valueIndex = existingValues.Count;
                values.Append(new Rich.RichValue(new Rich.Value(errorType), new Rich.Value(detailValue)) { S = (uint)structureIndex });
            }
            var existingFuture = future.Elements<FutureMetadataBlock>().ToList();
            int futureIndex = existingFuture.FindIndex(item => item.Descendants<Rich.RichValueBlock>().Any(value => value.I?.Value == (uint)valueIndex));
            if (futureIndex < 0) {
                changed = true;
                futureIndex = existingFuture.Count;
                var extension = new Extension { Uri = RichValueMetadataExtensionUri };
                extension.Append(new Rich.RichValueBlock { I = (uint)valueIndex });
                future.Append(new FutureMetadataBlock(new ExtensionList(extension)));
            }
            var existingBlocks = blocks.Elements<MetadataBlock>().ToList();
            int blockIndex = existingBlocks.FindIndex(item => item.Elements<MetadataRecord>().Count() == 1
                && item.Elements<MetadataRecord>().Any(record => record.TypeIndex?.Value == typeIndex && record.Val?.Value == (uint)futureIndex));
            if (blockIndex < 0) {
                changed = true;
                blockIndex = existingBlocks.Count;
                blocks.Append(new MetadataBlock(new MetadataRecord { TypeIndex = typeIndex, Val = (uint)futureIndex }));
            }
            values.Count = (uint)values.Elements<Rich.RichValue>().Count();
            structures.Count = (uint)structures.Elements<Rich.RichValueStructure>().Count();
            future.Count = (uint)future.Elements<FutureMetadataBlock>().Count();
            blocks.Count = (uint)blocks.Elements<MetadataBlock>().Count();
            if (changed) { values.Save(); structures.Save(); metadata.Save(); }
            cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.Error;
            cell.CellValue = new CellValue("#VALUE!");
            cell.ValueMetaIndex = (uint)blockIndex + 1;
            cell.InlineString = null;
            return true;
        }
    }
}
