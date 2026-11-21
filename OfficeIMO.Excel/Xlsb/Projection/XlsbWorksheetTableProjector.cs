using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel.Xlsb.Model;

namespace OfficeIMO.Excel.Xlsb.Projection {
    /// <summary>Projects decoded Table parts as table definition parts that keep their native relationship ids.</summary>
    internal static class XlsbWorksheetTableProjector {
        /// <summary>Adds one table definition part per table that the worksheet's tableParts element lists.</summary>
        /// <remarks>The tableParts element itself is a decoded feature element; call this after it is appended.</remarks>
        internal static void Apply(ExcelSheet sheet, XlsbSheet source) {
            ArgumentNullException.ThrowIfNull(sheet);
            ArgumentNullException.ThrowIfNull(source);
            TableParts? listed = source.FeatureElements.OfType<TableParts>().SingleOrDefault();
            if (listed == null) return;
            foreach (TablePart reference in listed.Elements<TablePart>()) {
                XlsbTable table = source.Tables[reference.Id!.Value!];
                TableDefinitionPart part = sheet.WorksheetPart.AddNewPart<TableDefinitionPart>(table.RelationshipId);
                part.Table = (Table)table.Definition.CloneNode(true);
            }
        }

        /// <summary>
        /// Pairs each table part of a loaded worksheet with the source table of the same relationship id. Returns false when
        /// the worksheet lists, adds or removes tables differently from its source, which a native rewrite cannot encode.
        /// </summary>
        internal static bool TryPair(ExcelSheet sheet, XlsbSheet source, out IReadOnlyList<(TableDefinitionPart Part, XlsbTable Source)> pairs) {
            pairs = Array.Empty<(TableDefinitionPart, XlsbTable)>();
            if (!XlsbWorksheetFeatureProjector.Matches<TableParts>(sheet, source)) return false;
            TableParts? listed = source.FeatureElements.OfType<TableParts>().SingleOrDefault();
            var result = new List<(TableDefinitionPart, XlsbTable)>();
            var seen = new HashSet<TableDefinitionPart>();
            foreach (TablePart reference in listed?.Elements<TablePart>() ?? Enumerable.Empty<TablePart>()) {
                string id = reference.Id!.Value!;
                if (!sheet.WorksheetPart.TryGetPartById(id, out OpenXmlPart? part) || part is not TableDefinitionPart definition) return false;
                seen.Add(definition);
                result.Add((definition, source.Tables[id]));
            }
            if (sheet.WorksheetPart.TableDefinitionParts.Any(part => !seen.Contains(part))) return false;
            pairs = result;
            return true;
        }
    }
}
