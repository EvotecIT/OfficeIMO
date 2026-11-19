using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel.Xlsb.Package;

namespace OfficeIMO.Excel.Xlsb.Write {
    /// <summary>
    /// Links a generated worksheet to its DrawingML drawing with BrtDrawing (MS-XLSB 2.4.354). The drawing part and the
    /// parts it uses (images, charts, chart styles and colors, chart drawings, embedded chart data) are the same XML and
    /// binary parts an XLSX package holds, so they are copied with their content types and relationship ids.
    /// </summary>
    internal static class XlsbWorksheetDrawingWriter {
        internal const int BrtDrawing = 550;
        private const int MaximumRelationshipIdLength = 255;
        private const string OfficeRelationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
        internal const string DrawingRelationshipType = OfficeRelationships + "drawing";

        // Relationship types a worksheet drawing graph may use. Anything else (ChartEx, SmartArt, OLE objects, controls,
        // custom XML) is refused until its XLSB output is proven in Excel.
        private static readonly HashSet<string> InternalRelationshipTypes = new HashSet<string>(StringComparer.Ordinal) {
            OfficeRelationships + "image",
            OfficeRelationships + "chart",
            OfficeRelationships + "chartUserShapes",
            OfficeRelationships + "package",
            OfficeRelationships + "themeOverride",
            "http://schemas.microsoft.com/office/2007/relationships/hdphoto",
            "http://schemas.microsoft.com/office/2011/relationships/chartStyle",
            "http://schemas.microsoft.com/office/2011/relationships/chartColorStyle",
        };

        private static readonly HashSet<string> ExternalRelationshipTypes = new HashSet<string>(StringComparer.Ordinal) {
            OfficeRelationships + "hyperlink",
            OfficeRelationships + "image",
        };

        internal static void Append(XlsbWorksheetFeatureRecords records, XlsbWorksheetFeatureContext context) {
            DrawingsPart? part = Resolve(context.Sheet);
            if (part == null) return;
            if (context.Parts == null) {
                throw new NotSupportedException($"Native XLSB rewriting cannot add a drawing to worksheet '{context.Sheet.Name}'; save as .xlsx to keep its images and charts.");
            }
            string drawingPartName = CopyGraph(context, part, new Dictionary<OpenXmlPart, string>());
            string relationshipId = context.Parts.AddRelationship(DrawingRelationshipType, drawingPartName);
            records.Drawings.Add(Create(relationshipId));
        }

        /// <summary>The worksheet's drawing part, or null when it has no drawing; refuses a drawing reference it cannot resolve.</summary>
        internal static DrawingsPart? Resolve(ExcelSheet sheet) {
            var drawings = sheet.WorksheetPart.Worksheet?.Elements<DocumentFormat.OpenXml.Spreadsheet.Drawing>().ToArray() ?? Array.Empty<DocumentFormat.OpenXml.Spreadsheet.Drawing>();
            DrawingsPart? owned = sheet.WorksheetPart.DrawingsPart;
            if (drawings.Length == 0) {
                if (owned != null) {
                    throw new NotSupportedException($"Native XLSB generation found a drawing part on worksheet '{sheet.Name}' that the worksheet does not reference.");
                }
                return null;
            }
            if (drawings.Length > 1 || drawings[0].HasChildren || drawings[0].GetAttributes().Count != 1
                || drawings[0].Id?.Value is not string id
                || !sheet.WorksheetPart.TryGetPartById(id, out OpenXmlPart? target) || target is not DrawingsPart drawingsPart
                || !ReferenceEquals(drawingsPart, owned)) {
                throw new NotSupportedException($"Native XLSB generation found an unsupported drawing reference on worksheet '{sheet.Name}'.");
            }
            return drawingsPart;
        }

        /// <summary>
        /// On the loaded-XLSB rewrite the native BrtDrawing and drawing parts are copied unchanged, so the worksheet must
        /// still link the drawing projected from them. Content edits inside the graph are refused by the preservation guard.
        /// </summary>
        internal static void ValidateUnchanged(ExcelSheet sheet, OfficeIMO.Excel.Xlsb.Model.XlsbSheet source, OfficeIMO.Excel.Xlsb.Model.XlsbWorkbook workbook) {
            DrawingsPart? part = Resolve(sheet);
            if (source.DrawingPartName == null) {
                if (part != null) {
                    throw new NotSupportedException($"Native XLSB rewriting cannot add a drawing to worksheet '{sheet.Name}'; save as .xlsx to keep its images and charts.");
                }
                return;
            }
            if (part == null || !workbook.ProjectionPartUris.TryGetValue(source.DrawingPartName, out string? projected)
                || !string.Equals(projected, part.Uri.OriginalString, StringComparison.Ordinal)) {
                throw new NotSupportedException($"Native XLSB rewriting cannot remove or replace the drawing on worksheet '{sheet.Name}'; save as .xlsx to keep that change.");
            }
        }

        internal static XlsbGeneratedRecord Create(string relationshipId) {
            if (relationshipId.Length == 0 || relationshipId.Length > MaximumRelationshipIdLength || relationshipId.IndexOf('\0') >= 0) {
                throw new NotSupportedException($"Native XLSB generation cannot encode drawing relationship id '{relationshipId}'.");
            }
            using var payload = new MemoryStream(4 + relationshipId.Length * 2);
            XlsbPayloadWriter.WriteNullableWideString(payload, relationshipId);
            return new XlsbGeneratedRecord(BrtDrawing, payload.ToArray());
        }

        private static string CopyGraph(XlsbWorksheetFeatureContext context, OpenXmlPart part, Dictionary<OpenXmlPart, string> copied) {
            if (copied.TryGetValue(part, out string? existing)) return existing;
            string sheetName = context.Sheet.Name;
            string partName = context.Parts!.Names.Next(NamePrefix(part.Uri.OriginalString, out string extension), extension);
            copied.Add(part, partName);
            var generated = new XlsbGeneratedPart(partName, part.ContentType, ReadContent(part));
            foreach (IdPartPair child in part.Parts) {
                string childType = child.OpenXmlPart.RelationshipType;
                if (!InternalRelationshipTypes.Contains(childType)) {
                    throw new NotSupportedException($"Native XLSB generation does not support drawing content related as '{childType}' on worksheet '{sheetName}'.");
                }
                string childName = CopyGraph(context, child.OpenXmlPart, copied);
                generated.Relationships.Add(new XlsbPackageRelationship(child.RelationshipId, childType,
                    XlsbWorksheetPackageParts.RelativeTarget(partName, childName), isExternal: false));
            }
            foreach (ExternalRelationship external in part.ExternalRelationships) {
                if (!ExternalRelationshipTypes.Contains(external.RelationshipType)) {
                    throw new NotSupportedException($"Native XLSB generation does not support external drawing content related as '{external.RelationshipType}' on worksheet '{sheetName}'.");
                }
                generated.Relationships.Add(new XlsbPackageRelationship(external.Id, external.RelationshipType, external.Uri.OriginalString, isExternal: true));
            }
            foreach (HyperlinkRelationship hyperlink in part.HyperlinkRelationships) {
                if (!hyperlink.IsExternal) {
                    throw new NotSupportedException($"Native XLSB generation does not support an internal drawing hyperlink relationship on worksheet '{sheetName}'.");
                }
                generated.Relationships.Add(new XlsbPackageRelationship(hyperlink.Id, hyperlink.RelationshipType, hyperlink.Uri.OriginalString, isExternal: true));
            }
            if (part.DataPartReferenceRelationships.Any()) {
                throw new NotSupportedException($"Native XLSB generation does not support media data references in the drawing on worksheet '{sheetName}'.");
            }
            context.Parts.AddPart(generated);
            return partName;
        }

        /// <summary>The source part name without its trailing number, so generated names follow the usual xl/charts/chartN.xml layout.</summary>
        private static string NamePrefix(string uri, out string extension) {
            string name = uri.TrimStart('/');
            int slash = name.LastIndexOf('/');
            string folder = slash < 0 ? string.Empty : name.Substring(0, slash + 1);
            string file = name.Substring(slash + 1);
            int dot = file.LastIndexOf('.');
            extension = dot < 0 ? string.Empty : file.Substring(dot);
            string stem = dot < 0 ? file : file.Substring(0, dot);
            int end = stem.Length;
            while (end > 0 && char.IsDigit(stem[end - 1])) end--;
            return folder + stem.Substring(0, end);
        }

        // XML parts are serialized from their loaded DOM so unsaved edits are written; other parts are copied as stored.
        private static byte[] ReadContent(OpenXmlPart part) {
            if (part is not ImagePart and not EmbeddedPackagePart && part.RootElement is DocumentFormat.OpenXml.OpenXmlPartRootElement root) {
                using var xml = new MemoryStream();
                using (var writer = System.Xml.XmlWriter.Create(xml, new System.Xml.XmlWriterSettings { Encoding = new UTF8Encoding(false), CloseOutput = false })) {
                    root.WriteTo(writer);
                }
                return xml.ToArray();
            }
            using Stream input = part.GetStream(FileMode.Open, FileAccess.Read);
            using var output = new MemoryStream();
            input.CopyTo(output);
            return output.ToArray();
        }
    }
}
