using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Removes empty containers and orphaned references on this worksheet to prevent Excel repairs.
        /// </summary>
        internal void Preflight() {
            WriteLock(() => {
                var ws = _worksheetPart.Worksheet;

                // Remove empty Hyperlinks
                var links = ws.Elements<Hyperlinks>().FirstOrDefault();
                if (links != null && !links.Elements<Hyperlink>().Any()) {
                    ws.RemoveChild(links);
                }

                // Remove empty MergeCells
                var merges = ws.Elements<MergeCells>().FirstOrDefault();
                if (merges != null && !merges.Elements<MergeCell>().Any()) {
                    ws.RemoveChild(merges);
                }

                // Drop orphaned Drawing reference
                var drawing = ws.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Drawing>();
                if (drawing?.Id?.Value is string dId) {
                    try { var _ = _worksheetPart.GetPartById(dId); } catch { ws.RemoveChild(drawing); }
                }

                // Drop orphaned LegacyDrawingHeaderFooter reference
                var legacy = ws.GetFirstChild<LegacyDrawingHeaderFooter>();
                if (legacy?.Id?.Value is string lId) {
                    try { var _ = _worksheetPart.GetPartById(lId); } catch { ws.RemoveChild(legacy); }
                }

                // Clean invalid TableParts and de-duplicate
                var parts = ws.Elements<TableParts>().FirstOrDefault();
                if (parts != null) {
                    var seen = new System.Collections.Generic.HashSet<string>(System.StringComparer.Ordinal);
                    foreach (var tp in parts.Elements<TablePart>().ToList()) {
                        var id = tp.Id?.Value ?? string.Empty;
                        if (string.IsNullOrEmpty(id) || !seen.Add(id)) {
                            tp.Remove();
                            continue;
                        }
                        try { var _ = _worksheetPart.GetPartById(id); } catch { tp.Remove(); }
                    }
                    if (!parts.Elements<TablePart>().Any()) {
                        ws.RemoveChild(parts);
                    } else {
                        parts.Count = (uint)parts.Elements<TablePart>().Count();
                    }
                }

                ws.Save();
            });
        }
    }
}

