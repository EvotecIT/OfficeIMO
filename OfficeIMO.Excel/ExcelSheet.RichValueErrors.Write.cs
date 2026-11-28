using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.OpenXml.Internal;
using System.Globalization;
using Rich = DocumentFormat.OpenXml.Office2019.Excel.RichData;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private bool TryWriteRichFormulaError(Cell cell, string? error, DynamicArrayPlan? spillPlan = null) {
            if (error is not ("#CALC!" or "#SPILL!")) return false;
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
            uint metadataIndex = AddRichFormulaError(OfficeXmlNode.FromSdk(future), OfficeXmlNode.FromSdk(blocks),
                OfficeXmlNode.FromSdk(values), OfficeXmlNode.FromSdk(structures), typeIndex, error, emptyArray, spillPlan, out bool changed);
            if (changed) {
                values.Save();
                structures.Save();
                metadata.Save();
            }

            cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.Error;
            cell.CellValue = new CellValue("#VALUE!");
            cell.ValueMetaIndex = metadataIndex;
            cell.InlineString = null;
            return true;
        }

        private static uint AddRichFormulaError(OfficeXmlNode future, OfficeXmlNode blocks, OfficeXmlNode values, OfficeXmlNode structures,
            uint typeIndex, string error, bool emptyArray, DynamicArrayPlan? spillPlan, out bool changed) {
            const string rich = RichValueErrorLookup.RichData;
            string errorType = error == "#CALC!" ? "13" : "8";
            (string Name, string Type, string Value)[] fields = spillPlan != null && error == "#SPILL!"
                ? new[] { ("colOffset", "i", spillPlan.ErrorColumns.ToString(CultureInfo.InvariantCulture)),
                    ("errorType", "i", errorType), ("rwOffset", "i", spillPlan.ErrorRows.ToString(CultureInfo.InvariantCulture)),
                    ("subType", "i", spillPlan.ErrorSubtype.ToString(CultureInfo.InvariantCulture)) }
                : new[] { ("errorType", "i", errorType), emptyArray ? ("subType", "i", "3") : ("propagated", "b", "1") };

            var existingStructures = ReadRichErrorElements(structures, "s");
            int structureIndex = existingStructures.FindIndex(item => item.Attribute("t") == "_error"
                && ReadRichErrorElements(item, "k").Select(key => (key.Attribute("n"), key.Attribute("t")))
                    .SequenceEqual(fields.Select(field => ((string?)field.Name, (string?)field.Type))));
            changed = structureIndex < 0;
            if (structureIndex < 0) {
                structureIndex = existingStructures.Count;
                OfficeXmlNode structure = structures.Append(structures.CreateElement(rich, "s", "xlrd"));
                structure.SetAttribute("t", "_error");
                foreach (var field in fields) {
                    OfficeXmlNode key = structure.Append(structure.CreateElement(rich, "k", "xlrd"));
                    key.SetAttribute("n", field.Name);
                    key.SetAttribute("t", field.Type);
                }
            }

            var existingValues = ReadRichErrorElements(values, "rv");
            int valueIndex = existingValues.FindIndex(item => item.UInt32Attribute("s") == (uint)structureIndex
                && ReadRichErrorElements(item, "v").Select(value => value.InnerText).SequenceEqual(fields.Select(field => field.Value)));
            if (valueIndex < 0) {
                changed = true;
                valueIndex = existingValues.Count;
                OfficeXmlNode value = values.Append(values.CreateElement(rich, "rv", "xlrd"));
                value.SetUInt32Attribute("s", (uint)structureIndex);
                foreach (var field in fields) {
                    value.Append(value.CreateElement(rich, "v", "xlrd")).SetInnerText(field.Value);
                }
            }

            var existingFuture = future.Children(WorksheetMarkup, "bk").ToList();
            int futureIndex = existingFuture.FindIndex(item => RichValueErrorLookup.FindValueBlock(item).UInt32Attribute("i") == (uint)valueIndex);
            if (futureIndex < 0) {
                changed = true;
                futureIndex = existingFuture.Count;
                OfficeXmlNode block = future.Append(future.CreateElement(WorksheetMarkup, "bk"));
                OfficeXmlNode extensions = block.Append(block.CreateElement(WorksheetMarkup, "extLst"));
                OfficeXmlNode extension = extensions.Append(extensions.CreateElement(WorksheetMarkup, "ext"));
                extension.SetAttribute("uri", RichValueMetadataExtensionUri);
                extension.Append(extension.CreateElement(rich, "rvb", "xlrd")).SetUInt32Attribute("i", (uint)valueIndex);
            }

            var existingBlocks = blocks.Children(WorksheetMarkup, "bk").ToList();
            int blockIndex = existingBlocks.FindIndex(item => {
                OfficeXmlNode record = item.Child(WorksheetMarkup, "rc");
                return item.Children(WorksheetMarkup, "rc").Count == 1 && record.UInt32Attribute("t") == typeIndex
                    && record.UInt32Attribute("v") == (uint)futureIndex;
            });
            if (blockIndex < 0) {
                changed = true;
                blockIndex = existingBlocks.Count;
                OfficeXmlNode block = blocks.Append(blocks.CreateElement(WorksheetMarkup, "bk"));
                OfficeXmlNode record = block.Append(block.CreateElement(WorksheetMarkup, "rc"));
                record.SetUInt32Attribute("t", typeIndex);
                record.SetUInt32Attribute("v", (uint)futureIndex);
            }
            values.SetUInt32Attribute("count", (uint)ReadRichErrorElements(values, "rv").Count);
            structures.SetUInt32Attribute("count", (uint)ReadRichErrorElements(structures, "s").Count);
            future.SetUInt32Attribute("count", (uint)future.Children(WorksheetMarkup, "bk").Count);
            blocks.SetUInt32Attribute("count", (uint)blocks.Children(WorksheetMarkup, "bk").Count);
            return (uint)blockIndex + 1U;
        }

        private static List<OfficeXmlNode> ReadRichErrorElements(OfficeXmlNode root, string name) {
            var elements = new List<OfficeXmlNode>();
            foreach (OfficeXmlNode element in root.Children()) {
                if (element.Is(RichValueErrorLookup.RichData, name)) elements.Add(element);
            }
            return elements;
        }
    }
}
