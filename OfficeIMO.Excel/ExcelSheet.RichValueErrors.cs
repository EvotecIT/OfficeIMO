using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Excel.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private RichValueErrorLookup? _richValueErrorLookup;
        private OfficeXmlNode _richValueErrorMetadata;
        private OfficeXmlNode _richValueErrorValues;
        private OfficeXmlNode _richValueErrorStructures;
        private int _richValueErrorMetadataCount, _richValueErrorValueCount, _richValueErrorStructureCount;

        private string? ResolveRichValueError(bool isError, uint? valueMetadataIndex, string? fallback) {
            if (fallback == null || !isError || !valueMetadataIndex.HasValue) return fallback;
            ExcelPartRef workbook = ExcelPartRef.Workbook(_excelDocument);
            // An editable root can change a value without changing any collection count.
            if (_excelDocument.FileOpenAccess != FileAccess.Read && _formulaEvaluationCache == null) {
                return RichValueErrorLookup.FromWorkbook(workbook, RichValueErrorLookup.MaximumPartBytes)
                    .Resolve(valueMetadataIndex.Value, fallback);
            }

            bool initialized = _richValueErrorLookup == null;
            if (initialized) _richValueErrorLookup = RichValueErrorLookup.FromWorkbook(workbook, RichValueErrorLookup.MaximumPartBytes);
            OfficeXmlNode metadata = workbook.FindWorkbookPart(ExcelReadWorkbookParts.Metadata).Root;
            OfficeXmlNode values = workbook.FindWorkbookPart(ExcelReadWorkbookParts.RichValues).Root;
            OfficeXmlNode structures = workbook.FindWorkbookPart(ExcelReadWorkbookParts.RichStructures).Root;
            int metadataCount = CountRichValueChildren(metadata.Child(OfficeXmlNamespaces.SpreadsheetML, "valueMetadata"));
            int valueCount = CountRichValueChildren(values);
            int structureCount = CountRichValueChildren(structures);
            if (initialized || _richValueErrorMetadata != metadata || _richValueErrorValues != values
                || _richValueErrorStructures != structures || _richValueErrorMetadataCount != metadataCount
                || _richValueErrorValueCount != valueCount || _richValueErrorStructureCount != structureCount) {
                if (!initialized) _richValueErrorLookup = RichValueErrorLookup.FromWorkbook(workbook, RichValueErrorLookup.MaximumPartBytes);
                _richValueErrorMetadata = metadata;
                _richValueErrorValues = values;
                _richValueErrorStructures = structures;
                _richValueErrorMetadataCount = metadataCount;
                _richValueErrorValueCount = valueCount;
                _richValueErrorStructureCount = structureCount;
            }
            return _richValueErrorLookup!.Resolve(valueMetadataIndex.Value, fallback);
        }

        private static int CountRichValueChildren(OfficeXmlNode parent) {
            int count = 0;
            foreach (OfficeXmlNode child in parent.Children()) count++;
            return count;
        }
    }
}
