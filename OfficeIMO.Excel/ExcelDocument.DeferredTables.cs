using OfficeIMO.Core.Internal.OpenXml;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Table reads of a workbook that has not opened its SDK package. Recorded cell values never change a table (a
    /// value in a header row opens the SDK), so the tables are read from the package the workbook holds, once per
    /// package.
    /// </summary>
    public partial class ExcelDocument {
        private const string WorksheetPartRelationshipType = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "worksheet";
        // Created on first use: a field initializer would compile into the document's constructors.
        private object? _deferredTablesGate;
        private DeferredTables? _deferredTables;

        private sealed class DeferredTables {
            internal DeferredTables(byte[] bytes, ExcelDeferredWorkbook workbook, int sheetCount) {
                Bytes = bytes;
                Workbook = workbook;
                SheetCount = sheetCount;
            }

            internal byte[] Bytes { get; }
            internal ExcelDeferredWorkbook Workbook { get; }
            internal int SheetCount { get; }
            internal List<string> WorksheetOrder { get; } = new();
            internal Dictionary<string, List<ExcelTablePart>> ByPart { get; } = new(StringComparer.Ordinal);
        }

        /// <summary>The workbook's tables read from its package, or null when the workbook reads them through the model.</summary>
        private DeferredTables? ReadDeferredTables() {
            ExcelDeferredWorkbook? workbook = _deferredWorkbook;
            byte[]? bytes = _deferredPackageBytes;
            if (workbook == null || bytes == null) return null;
            object gate = Volatile.Read(ref _deferredTablesGate) ?? Interlocked.CompareExchange(ref _deferredTablesGate, new object(), null) ?? _deferredTablesGate!;
            lock (gate) {
                DeferredTables? tables = _deferredTables;
                if (tables != null && ReferenceEquals(tables.Bytes, bytes) && ReferenceEquals(tables.Workbook, workbook)
                    && tables.SheetCount == workbook.Sheets.Count) {
                    return tables;
                }
                tables = new DeferredTables(bytes, workbook, workbook.Sheets.Count);
                try {
                    using var source = new MemoryStream(bytes, writable: false);
                    using OfficeOpcPackage package = OfficeOpcPackage.Open(source, leaveOpen: true);
                    foreach (OfficeOpcRelationship relationship in package.GetRelationships(workbook.WorkbookPartName)) {
                        if (string.Equals(relationship.Type, WorksheetPartRelationshipType, StringComparison.Ordinal) && !relationship.IsExternal) {
                            tables.WorksheetOrder.Add(relationship.Id);
                        }
                    }
                    foreach (ExcelDeferredSheet sheet in workbook.Sheets) {
                        if (sheet.IsWorksheet && !tables.ByPart.ContainsKey(sheet.PartName)) {
                            tables.ByPart.Add(sheet.PartName, ExcelTablePart.ReadSheetTables(package, sheet.PartName));
                        }
                    }
                } catch (Exception exception) when (exception is InvalidDataException or System.Xml.XmlException or IOException) {
                    return null;
                }
                return _deferredTables = tables;
            }
        }

        /// <summary>
        /// Lists the tables of every worksheet, or of <paramref name="selected"/> only, as the model lists them: worksheets
        /// in the order the workbook part relates them, tables in the order each worksheet relates them. False when the
        /// workbook reads its tables through the SDK model.
        /// </summary>
        internal bool TryGetDeferredTables(ExcelDeferredSheet? selected, out List<ExcelTableInfo> tables) {
            tables = null!;
            if (ReadDeferredTables() is not { } read) return false;
            List<ExcelDeferredSheet> sheets = read.Workbook.Sheets;
            if (selected != null && !sheets.Contains(selected)) {
                throw new InvalidOperationException(
                    "The worksheet is no longer part of the current workbook package. Obtain a current worksheet from ExcelDocument.Sheets before inspecting or exporting it.");
            }
            var result = new List<ExcelTableInfo>();
            foreach (string relationshipId in read.WorksheetOrder) {
                int index = sheets.FindIndex(sheet => string.Equals(sheet.RelationshipId, relationshipId, StringComparison.Ordinal));
                if (index < 0 || !sheets[index].IsWorksheet) continue;
                ExcelDeferredSheet sheet = sheets[index];
                if (selected != null && !ReferenceEquals(sheet, selected)) continue;
                foreach (ExcelTablePart table in read.ByPart[sheet.PartName]) result.Add(table.ToInfo(sheet.Name, index));
            }
            tables = result;
            return true;
        }

        /// <summary>The table parts of a worksheet that has not opened its SDK package, or null when it reads them through the model.</summary>
        internal List<ExcelTablePart>? TryGetDeferredTableParts(ExcelDeferredSheet sheet) =>
            ReadDeferredTables() is { } read && read.Workbook.Sheets.Contains(sheet) && read.ByPart.TryGetValue(sheet.PartName, out List<ExcelTablePart>? tables)
                ? tables
                : null;
    }
}
