using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private void ValidateColumnCommentVmlAnchorCapacity(int firstColumn, int count) {
            XNamespace x = "urn:schemas-microsoft-com:office:excel";
            foreach (VmlDrawingPart vmlPart in _worksheetPart.VmlDrawingParts) {
                XDocument document = LoadOrCreateVmlDocument(vmlPart);
                foreach (XElement clientData in document.Descendants(x + "ClientData")
                    .Where(element => string.Equals(
                        element.Attribute("ObjectType")?.Value,
                        "Note",
                        StringComparison.OrdinalIgnoreCase))) {
                    VmlAnchorPlacement placement = GetVmlAnchorPlacement(clientData, x);
                    if (placement == VmlAnchorPlacement.Absolute
                        || !TryParseVmlAnchor(clientData.Element(x + "Anchor"), out int[] values)) continue;

                    if (placement == VmlAnchorPlacement.OneCell) {
                        int oneBasedFromColumn = values[0] + 1;
                        if (oneBasedFromColumn >= firstColumn
                            && (long)values[4] + count > A1.MaxColumns) {
                            throw new InvalidOperationException(
                                "Column insertion would move a comment note anchor beyond Excel's column limit.");
                        }
                        continue;
                    }

                    int firstSpannedColumn = values[0] + 1;
                    int lastSpannedColumn = values[4];
                    if (lastSpannedColumn < firstSpannedColumn) continue;
                    try {
                        _ = ExcelDocument.TransformColumnReference(
                            ExcelReference.Parse(
                                A1.CellReference(1, firstSpannedColumn) + ":" +
                                A1.CellReference(1, lastSpannedColumn)),
                            firstColumn,
                            lastDeletedColumn: 0,
                            count,
                            deleting: false);
                    } catch (ArgumentOutOfRangeException) {
                        throw new InvalidOperationException(
                            "Column insertion would move a comment note anchor beyond Excel's column limit.");
                    }
                }
            }
        }

        private void ValidateColumnConnectionParameters(int firstColumn, int count, bool deleting) {
            Connections? connections = WorkbookPartRoot.ConnectionsPart?.Connections;
            if (connections == null) return;

            HashSet<uint> connectionIds = GetWorksheetQueryConnectionIds(_worksheetPart);
            foreach (Connection connection in connections.Elements<Connection>()
                .Where(connection => connection.Id?.Value is uint id && connectionIds.Contains(id))) {
                foreach (Parameter parameter in connection.Descendants<Parameter>()) {
                    if (parameter.Cell?.Value is not string reference
                        || !ExcelReference.TryParse(reference, out ExcelReference? parsed)) continue;
                    ExcelReference? mapped;
                    try {
                        mapped = ExcelDocument.TransformColumnReference(
                            parsed!,
                            firstColumn,
                            firstColumn + count - 1,
                            count,
                            deleting);
                    } catch (ArgumentOutOfRangeException) {
                        throw new InvalidOperationException(
                            $"Column insertion would move cell-backed connection parameter '{reference}' beyond the Excel column limit.");
                    }
                    if (deleting && mapped == null) {
                        throw new InvalidOperationException(
                            $"Cannot delete cell-backed connection parameter reference '{reference}'. Update or remove the parameter first.");
                    }
                }
            }
        }

        private void RemapColumnConnectionParameters(
            int firstColumn,
            int count,
            bool deleting,
            CancellationToken cancellationToken) {
            Connections? connections = WorkbookPartRoot.ConnectionsPart?.Connections;
            if (connections == null) return;

            HashSet<uint> connectionIds = GetWorksheetQueryConnectionIds(_worksheetPart);
            bool changed = false;
            foreach (Connection connection in connections.Elements<Connection>()
                .Where(connection => connection.Id?.Value is uint id && connectionIds.Contains(id))) {
                foreach (Parameter parameter in connection.Descendants<Parameter>()) {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (parameter.Cell?.Value is not string reference
                        || !ExcelReference.TryParse(reference, out ExcelReference? parsed)) continue;
                    ExcelReference? mapped = ExcelDocument.TransformColumnReference(
                        parsed!,
                        firstColumn,
                        firstColumn + count - 1,
                        count,
                        deleting);
                    if (mapped == null) parameter.Remove();
                    else parameter.Cell = mapped.ToString();
                    changed = true;
                }
                foreach (Parameters parameters in connection.Elements<Parameters>().ToList()) {
                    uint parameterCount = (uint)parameters.Elements<Parameter>().Count();
                    if (parameterCount == 0U) parameters.Remove();
                    else parameters.Count = parameterCount;
                }
            }
            if (changed) connections.Save();
        }

        private void RemapColumnCommentVml(
            int firstColumn,
            int count,
            bool deleting) {
            VmlDrawingPart? vmlPart = TryGetCommentVmlPart();
            if (vmlPart == null) {
                CleanupCommentArtifacts();
                return;
            }

            XDocument document = LoadOrCreateVmlDocument(vmlPart);
            XElement? root = document.Root;
            if (root == null) return;
            XNamespace v = "urn:schemas-microsoft-com:vml";
            XNamespace x = "urn:schemas-microsoft-com:office:excel";
            bool changed = false;
            foreach (XElement shape in root.Elements(v + "shape").ToList()) {
                XElement? clientData = shape.Element(x + "ClientData");
                if (clientData == null
                    || !string.Equals(clientData.Attribute("ObjectType")?.Value, "Note", StringComparison.OrdinalIgnoreCase)
                    || !TryParseVmlCoordinate(clientData.Element(x + "Row")?.Value, out int zeroBasedRow)
                    || !TryParseVmlCoordinate(clientData.Element(x + "Column")?.Value, out int zeroBasedColumn)) continue;

                ExcelReference point = ExcelReference.Parse(A1.CellReference(zeroBasedRow + 1, zeroBasedColumn + 1));
                ExcelReference? mapped = ExcelDocument.TransformColumnReference(
                    point,
                    firstColumn,
                    firstColumn + count - 1,
                    count,
                    deleting);
                if (mapped == null) {
                    shape.Remove();
                    changed = true;
                    continue;
                }

                int mappedColumn = mapped.Start.Column;
                if (mappedColumn != zeroBasedColumn + 1) {
                    clientData.SetElementValue(
                        x + "Column",
                        (mappedColumn - 1).ToString(CultureInfo.InvariantCulture));
                    changed = true;
                }
                changed |= RemapVmlAnchorColumns(
                    clientData.Element(x + "Anchor"),
                    firstColumn,
                    count,
                    deleting,
                    GetVmlAnchorPlacement(clientData, x));
            }
            if (changed) SaveVmlDocument(vmlPart, document);
            CleanupCommentArtifacts();
        }
    }
}
