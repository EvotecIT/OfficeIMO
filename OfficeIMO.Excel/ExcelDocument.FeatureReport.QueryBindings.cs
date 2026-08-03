using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        private HashSet<Uri> GetManagedQueryPartUris(
            IReadOnlyList<ExcelQueryBackedTableInfo> bindings) {
            var result = new HashSet<Uri>();
            foreach (ExcelQueryBackedTableInfo binding in bindings) {
                ExcelSheet? bindingSheet = Sheets.FirstOrDefault(sheet => string.Equals(
                    sheet.Name,
                    binding.WorksheetName,
                    StringComparison.OrdinalIgnoreCase));
                TableDefinitionPart? bindingTable = bindingSheet?.WorksheetPart.TableDefinitionParts.FirstOrDefault(part =>
                    string.Equals(
                        part.Table?.Name?.Value ?? part.Table?.DisplayName?.Value,
                        binding.TableName,
                        StringComparison.OrdinalIgnoreCase));
                QueryTablePart? bindingQuery = bindingTable?.QueryTableParts.FirstOrDefault(part =>
                    part.QueryTable?.ConnectionId?.Value == binding.ConnectionId);
                if (bindingQuery != null) result.Add(bindingQuery.Uri);
            }
            return result;
        }
    }
}
