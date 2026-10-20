using DocumentFormat.OpenXml.Packaging;

namespace OfficeIMO.Excel.Utilities;

/// <summary>Tracks worksheet-owned package descendants so deleting one sheet cannot remove a shared graph.</summary>
internal static class ExcelWorksheetPartOwnership {
    internal static Dictionary<string, HashSet<string>> Capture(ExcelDocument document) {
        var result = new Dictionary<string, HashSet<string>>(StringComparer.Ordinal);
        foreach (var sheet in document.Sheets) {
            var visited = new HashSet<OpenXmlPart>(); var queue = new Queue<OpenXmlPart>(); queue.Enqueue(sheet.WorksheetPart);
            while (queue.Count > 0) {
                var part = queue.Dequeue(); if (!visited.Add(part)) continue;
                if (!result.TryGetValue(part.Uri.OriginalString, out var owners)) result.Add(part.Uri.OriginalString, owners = new(StringComparer.Ordinal));
                owners.Add(sheet.WorksheetPart.Uri.OriginalString);
                foreach (var child in part.Parts) queue.Enqueue(child.OpenXmlPart);
            }
        }
        return result;
    }
}
