using OfficeIMO.Excel.Xlsb.Biff12;
using OfficeIMO.Excel.Xlsb.Model;
using OfficeIMO.Excel.Xlsb.Package;

namespace OfficeIMO.Excel.Xlsb;

internal static partial class XlsbWorkbookReader {
    private static void ReadChartSheet(XlsbPackagePartReader parts, string partName, XlsbSheet sheet,
        XlsbImportOptions options, XlsbWorkbook workbook, XlsbRecordReadBudget recordBudget) {
        var records = ReadRecords(parts.ReadPart(partName, options.CancellationToken), options, recordBudget);
        if (records.Count < 2 || records[0].Type != BrtBeginSheet || records[records.Count - 1].Type != BrtEndSheet)
            throw new InvalidDataException($"Chart sheet '{sheet.Name}' lacks its BIFF12 sheet boundaries.");
        foreach (var record in records) {
            options.CancellationToken.ThrowIfCancellationRequested();
            if (record.Type == 651) { // BrtCsProp: flags, tab color, CodeName
                if (sheet.ChartCodeName != null) throw new InvalidDataException("Duplicate chart-sheet properties record.");
                var cursor = new XlsbBinaryCursor(record.Data);
                for (int i = 0; i < 10; i++) cursor.ReadByte();
                sheet.ChartCodeName = cursor.ReadWideString(options.MaxStringCharacters);
                if (cursor.Remaining != 0) throw new InvalidDataException("Unexpected chart properties payload.");
                // The host identity is projected; remaining properties stay in the native source.
                PreserveRecord(options, workbook, partName, record);
            } else if (record.Type == 550) { // BrtDrawing
                if (sheet.DrawingRelationshipId != null) throw new InvalidDataException("Duplicate chart-sheet drawing record.");
                var cursor = new XlsbBinaryCursor(record.Data);
                sheet.DrawingRelationshipId = cursor.ReadWideString(options.MaxStringCharacters);
                if (cursor.Remaining != 0) throw new InvalidDataException("Unexpected chart drawing payload.");
            } else if (record.Type is not BrtBeginSheet and not BrtEndSheet) {
                PreserveRecord(options, workbook, partName, record);
            }
        }
        if (sheet.DrawingRelationshipId == null || !sheet.Relationships.TryGetValue(sheet.DrawingRelationshipId, out var drawing)
            || drawing.IsExternal || drawing.Type != "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing")
            throw new InvalidDataException($"Chart sheet '{sheet.Name}' lacks its internal drawing relationship.");
        ReadRelatedGraph(parts, XlsbPackagePartReader.ResolveTarget(partName, drawing.Target), options, workbook, 0);
    }

    private static void ReadVbaProjectGraph(XlsbPackagePartReader parts, string workbookPartName,
        IReadOnlyDictionary<string, XlsbPackageRelationship> relationships, XlsbImportOptions options, XlsbWorkbook workbook) {
        var matches = relationships.Values.Where(item => item.Type == "http://schemas.microsoft.com/office/2006/relationships/vbaProject").ToArray();
        if (matches.Length == 0) return;
        if (matches.Length != 1 || matches[0].IsExternal) throw new InvalidDataException("The XLSB VBA project must have one internal relationship.");
        string partName = XlsbPackagePartReader.ResolveTarget(workbookPartName, matches[0].Target);
        if (parts.GetContentType(partName) != "application/vnd.ms-office.vbaProject") throw new InvalidDataException("Invalid VBA project content type.");
        workbook.VbaProjectPartName = partName;
        ReadRelatedGraph(parts, partName, options, workbook, 0);
    }

    private static void ReadThemeGraph(XlsbPackagePartReader parts, string workbookPartName,
        IReadOnlyDictionary<string, XlsbPackageRelationship> relationships, XlsbImportOptions options, XlsbWorkbook workbook) {
        var matches = relationships.Values.Where(item => item.Type == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme").ToArray();
        if (matches.Length == 0) return;
        if (matches.Length != 1 || matches[0].IsExternal) throw new InvalidDataException("The XLSB theme must have one internal relationship.");
        string partName = XlsbPackagePartReader.ResolveTarget(workbookPartName, matches[0].Target);
        if (parts.GetContentType(partName) != "application/vnd.openxmlformats-officedocument.theme+xml")
            throw new InvalidDataException("Invalid XLSB theme content type.");
        workbook.ThemePartName = partName;
        ReadRelatedGraph(parts, partName, options, workbook, 0);
    }

    /// <summary>
    /// Decodes the worksheet's Comments part and loads the VML drawing its BrtLegacyDrawing record names, with that
    /// drawing's own related parts, for projection.
    /// </summary>
    private static void ReadComments(XlsbPackagePartReader parts, string sheetPartName, XlsbSheet sheet,
        XlsbImportOptions options, XlsbWorkbook workbook, XlsbRecordReadBudget recordBudget) {
        var matches = sheet.Relationships.Values.Where(item => item.Type == Read.XlsbCommentRelationships.Comments).ToArray();
        if (matches.Length == 0) return;
        if (matches.Length != 1 || matches[0].IsExternal) throw new InvalidDataException($"The XLSB worksheet '{sheet.Name}' must have at most one internal comments relationship.");
        string partName = XlsbPackagePartReader.ResolveTarget(sheetPartName, matches[0].Target);
        if (parts.GetContentType(partName) != Read.XlsbCommentRelationships.CommentsContentType) throw new InvalidDataException($"The XLSB comments part '{partName}' has an invalid content type.");
        if (parts.ReadRelationships(partName, options.CancellationToken).Count != 0) throw new InvalidDataException($"The XLSB comments part '{partName}' must not have relationships.");
        var records = ReadRecords(parts.ReadPart(partName, options.CancellationToken), options, recordBudget);
        XlsbComments comments = Read.XlsbWorksheetCommentReader.Read(records, partName, matches[0].Id);
        if (sheet.LegacyDrawingRelationshipId != null) {
            string vmlPartName = XlsbPackagePartReader.ResolveTarget(sheetPartName, sheet.Relationships[sheet.LegacyDrawingRelationshipId].Target);
            if (parts.GetContentType(vmlPartName) != Read.XlsbCommentRelationships.VmlDrawingContentType) throw new InvalidDataException($"The XLSB VML drawing part '{vmlPartName}' has an invalid content type.");
            ReadRelatedGraph(parts, vmlPartName, options, workbook, 0);
            comments.VmlPartName = vmlPartName;
            comments.VmlRelationshipId = sheet.LegacyDrawingRelationshipId;
        }
        sheet.Comments = comments;
    }

    private static void ReadRelatedGraph(XlsbPackagePartReader parts, string partName,
        XlsbImportOptions options, XlsbWorkbook workbook, int depth) {
        if (workbook.RelatedParts.ContainsKey(partName)) return;
        if (depth > options.MaxRelationshipDepth) throw new InvalidDataException("The XLSB related-part graph exceeds MaxRelationshipDepth.");
        if (workbook.RelatedParts.Count >= options.MaxRelatedParts) throw new InvalidDataException("The XLSB related-part graph exceeds MaxRelatedParts.");
        options.CancellationToken.ThrowIfCancellationRequested();
        var relationships = parts.ReadRelationships(partName, options.CancellationToken);
        workbook.RelatedParts.Add(partName, new(parts.GetContentType(partName), parts.ReadPart(partName, options.CancellationToken), relationships));
        foreach (var relationship in relationships.Values) {
            if (!relationship.IsExternal) ReadRelatedGraph(parts, XlsbPackagePartReader.ResolveTarget(partName, relationship.Target), options, workbook, depth + 1);
        }
    }
}
