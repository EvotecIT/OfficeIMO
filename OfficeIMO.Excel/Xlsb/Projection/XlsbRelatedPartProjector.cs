using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel.Xlsb.Model;
using OfficeIMO.Excel.Xlsb.Package;

namespace OfficeIMO.Excel.Xlsb.Projection;

/// <summary>Projects chart/VBA carriers using their original bounded part graph, without executing linked content.</summary>
internal sealed class XlsbRelatedPartProjector {
    private const string OfficeRelationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
    private readonly ExcelDocument _document;
    private readonly XlsbWorkbook _workbook;
    private readonly System.Threading.CancellationToken _cancellationToken;
    private readonly Dictionary<string, OpenXmlPart> _copied = new(StringComparer.OrdinalIgnoreCase);

    internal XlsbRelatedPartProjector(ExcelDocument document, XlsbWorkbook workbook, System.Threading.CancellationToken cancellationToken) {
        _document = document; _workbook = workbook; _cancellationToken = cancellationToken;
    }

    internal void ProjectChartSheet(XlsbSheet source) {
        WorkbookPart workbook = _document.WorkbookPartRoot;
        var part = workbook.AddNewPart<ChartsheetPart>();
        part.Chartsheet = new Chartsheet();
        if (!string.IsNullOrEmpty(source.ChartCodeName)) part.Chartsheet.Append(new ChartSheetProperties { CodeName = source.ChartCodeName });
        string id = source.DrawingRelationshipId!;
        var drawing = source.Relationships[id];
        Copy(part, XlsbPackagePartReader.ResolveTarget(source.PartName!, drawing.Target), drawing.Type, id);
        part.Chartsheet.Append(new DocumentFormat.OpenXml.Spreadsheet.Drawing { Id = id });
        part.Chartsheet.Save();
        var sheets = _document.WorkbookRoot.Sheets ?? _document.WorkbookRoot.AppendChild(new Sheets());
        uint nextId = sheets.Elements<Sheet>().Select(item => item.SheetId?.Value ?? 0).DefaultIfEmpty().Max() + 1;
        sheets.Append(new Sheet {
            Id = workbook.GetIdOfPart(part), SheetId = nextId, Name = source.Name,
            State = source.State == 1 ? SheetStateValues.Hidden : source.State == 2 ? SheetStateValues.VeryHidden : (SheetStateValues?)null
        });
    }

    internal void ProjectVba() {
        if (!string.IsNullOrEmpty(_workbook.CodeName)) {
            var properties = _document.WorkbookRoot.GetFirstChild<WorkbookProperties>();
            if (properties == null) { properties = new WorkbookProperties(); _document.WorkbookRoot.AddChild(properties, true); }
            properties.CodeName = _workbook.CodeName;
        }
        if (_workbook.VbaProjectPartName == null) return;
        string sourceName = _workbook.VbaProjectPartName;
        _document.AddMacro(_workbook.RelatedParts[sourceName].Bytes);
        var part = _document.WorkbookPartRoot.VbaProjectPart!;
        _copied.Add(sourceName, part);
        _workbook.ProjectionPartUris[sourceName] = part.Uri.OriginalString;
        CopyRelationships(part, sourceName);
    }

    private OpenXmlPart Copy(OpenXmlPartContainer owner, string sourceName, string relationshipType, string id) {
        _cancellationToken.ThrowIfCancellationRequested();
        if (_copied.TryGetValue(sourceName, out var existing)) {
            if (existing.RelationshipType != relationshipType) throw new NotSupportedException("One related part cannot be projected through conflicting relationship types.");
            return owner.AddPart(existing, id);
        }
        XlsbProjectedPart source = _workbook.RelatedParts[sourceName];
        OpenXmlPart target = CreatePart(owner, relationshipType, source.ContentType, id, Path.GetExtension(sourceName));
        _copied.Add(sourceName, target);
        _workbook.ProjectionPartUris[sourceName] = target.Uri.OriginalString;
        using (var input = new MemoryStream(source.Bytes, writable: false)) target.FeedData(input);
        CopyRelationships(target, sourceName);
        return target;
    }

    private void CopyRelationships(OpenXmlPart owner, string sourceName) {
        foreach (var relationship in _workbook.RelatedParts[sourceName].Relationships.Values) {
            _cancellationToken.ThrowIfCancellationRequested();
            if (relationship.IsExternal) {
                var target = new Uri(relationship.Target, UriKind.RelativeOrAbsolute);
                if (relationship.Type == OfficeRelationships + "hyperlink") owner.AddHyperlinkRelationship(target, true, relationship.Id);
                else owner.AddExternalRelationship(relationship.Type, target, relationship.Id);
            } else Copy(owner, XlsbPackagePartReader.ResolveTarget(sourceName, relationship.Target), relationship.Type, relationship.Id);
        }
    }

    private static OpenXmlPart CreatePart(OpenXmlPartContainer owner, string type, string contentType, string id, string extension) {
        if (type == OfficeRelationships + "drawing" && contentType == "application/vnd.openxmlformats-officedocument.drawing+xml") return owner.AddNewPart<DrawingsPart>(id);
        if (type == OfficeRelationships + "chart" && contentType == "application/vnd.openxmlformats-officedocument.drawingml.chart+xml") return owner.AddNewPart<ChartPart>(id);
        if (type == OfficeRelationships + "chartUserShapes" && contentType == "application/vnd.openxmlformats-officedocument.drawingml.chartshapes+xml") return owner.AddNewPart<ChartDrawingPart>(id);
        if (type == OfficeRelationships + "image") return owner.AddNewPart<ImagePart>(contentType, id);
        if (type == OfficeRelationships + "package") return owner.AddNewPart<EmbeddedPackagePart>(contentType, id);
        return owner.AddExtendedPart(type, contentType, string.IsNullOrEmpty(extension) ? ".bin" : extension, id);
    }
}
