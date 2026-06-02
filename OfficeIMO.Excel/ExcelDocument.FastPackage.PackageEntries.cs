using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Diagnostics;
using System.Globalization;
using System.IO.Compression;
using System.Threading;
using System.Xml;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {

        private static void WriteContentTypesEntry(ZipArchive archive, bool hasStyles, bool hasSharedStrings, int worksheetCount, int tableCount) {
            var builder = new System.Text.StringBuilder(512 + worksheetCount * 160 + tableCount * 160);
            builder.Append("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
            builder.Append("<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">");
            builder.Append("<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>");
            builder.Append("<Default Extension=\"xml\" ContentType=\"application/xml\"/>");
            builder.Append("<Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/>");
            builder.Append("<Override PartName=\"/docProps/app.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/>");
            builder.Append("<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>");
            for (int i = 1; i <= worksheetCount; i++) {
                builder.Append("<Override PartName=\"/xl/worksheets/sheet");
                AppendInvariant(builder, i);
                builder.Append(".xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>");
            }

            if (hasStyles) {
                builder.Append("<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>");
            }

            if (hasSharedStrings) {
                builder.Append("<Override PartName=\"/xl/sharedStrings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/>");
            }

            for (int i = 1; i <= tableCount; i++) {
                builder.Append("<Override PartName=\"/xl/tables/table");
                AppendInvariant(builder, i);
                builder.Append(".xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml\"/>");
            }

            builder.Append("</Types>");
            WriteTextEntry(archive, "[Content_Types].xml", builder.ToString());
        }

        private static void WriteWorkbookEntry(ZipArchive archive, FastWorkbookPackageModel model) {
            var entry = archive.CreateEntry("xl/workbook.xml", CompressionLevel.Fastest);
            var worksheets = model.Worksheets;
            using var stream = entry.Open();
            using var writer = CreateFastXmlWriter(stream);
            writer.WriteStartDocument();
            writer.WriteStartElement("workbook", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
            writer.WriteAttributeString("xmlns", "r", null, "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            if (model.FileVersion != null) {
                model.FileVersion.WriteTo(writer);
            }

            if (model.FileSharing != null) {
                model.FileSharing.WriteTo(writer);
            }

            if (model.WorkbookProperties != null) {
                model.WorkbookProperties.WriteTo(writer);
            }

            if (model.WorkbookProtection != null) {
                model.WorkbookProtection.WriteTo(writer);
            }

            if (model.BookViews != null) {
                model.BookViews.WriteTo(writer);
            }

            writer.WriteStartElement("sheets");
            foreach (var worksheet in worksheets) {
                writer.WriteStartElement("sheet");
                writer.WriteAttributeString("name", worksheet.SheetName);
                writer.WriteAttributeString("sheetId", InvariantNumberText.Get(worksheet.SheetId));
                writer.WriteAttributeString("r", "id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships", worksheet.WorkbookRelationshipId);
                if (!string.IsNullOrEmpty(worksheet.SheetState)) {
                    writer.WriteAttributeString("state", worksheet.SheetState);
                }

                writer.WriteEndElement();
            }

            writer.WriteEndElement();
            if (model.DefinedNames != null) {
                model.DefinedNames.WriteTo(writer);
            }

            if (model.CalculationProperties != null) {
                model.CalculationProperties.WriteTo(writer);
            }

            writer.WriteEndElement();
        }

        private static void WriteWorkbookRelationshipsEntry(ZipArchive archive, IReadOnlyList<FastWorksheetPackageModel> worksheets, bool hasStyles, bool hasSharedStrings) {
            var builder = new System.Text.StringBuilder(384 + worksheets.Count * 180);
            builder.Append("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
            builder.Append("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
            foreach (var worksheet in worksheets) {
                builder.Append("<Relationship Id=\"");
                AppendXmlEscaped(builder, worksheet.WorkbookRelationshipId);
                builder.Append("\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"");
                AppendXmlEscaped(builder, worksheet.WorksheetPath.Substring("xl/".Length));
                builder.Append("\"/>");
            }

            if (hasStyles) {
                builder.Append("<Relationship Id=\"rId");
                AppendInvariant(builder, worksheets.Count + 1);
                builder.Append("\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>");
            }

            if (hasSharedStrings) {
                builder.Append("<Relationship Id=\"rId");
                AppendInvariant(builder, worksheets.Count + (hasStyles ? 2 : 1));
                builder.Append("\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\" Target=\"sharedStrings.xml\"/>");
            }

            builder.Append("</Relationships>");
            WriteTextEntry(archive, "xl/_rels/workbook.xml.rels", builder.ToString());
        }

        private static void WriteCorePropertiesEntry(ZipArchive archive) {
            WriteTextEntry(archive, "docProps/core.xml", CreateCorePropertiesXml());
        }

        private static string CreateCorePropertiesXml() {
            return "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
                "<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" " +
                "xmlns:dc=\"http://purl.org/dc/elements/1.1/\" " +
                "xmlns:dcterms=\"http://purl.org/dc/terms/\" " +
                "xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\" " +
                "xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\"/>";
        }

        private static string CreateCorePropertiesXml(SpreadsheetDocument document) {
            var properties = document.PackageProperties;
            var builder = new System.Text.StringBuilder(512);
            builder.Append("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
            builder.Append("<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" ");
            builder.Append("xmlns:dc=\"http://purl.org/dc/elements/1.1/\" ");
            builder.Append("xmlns:dcterms=\"http://purl.org/dc/terms/\" ");
            builder.Append("xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\" ");
            builder.Append("xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">");

            AppendCoreProperty(builder, "dc:title", properties.Title);
            AppendCoreProperty(builder, "dc:subject", properties.Subject);
            AppendCoreProperty(builder, "dc:creator", properties.Creator);
            AppendCoreProperty(builder, "cp:keywords", properties.Keywords);
            AppendCoreProperty(builder, "dc:description", properties.Description);
            AppendCoreProperty(builder, "cp:lastModifiedBy", properties.LastModifiedBy);
            AppendCoreProperty(builder, "cp:revision", properties.Revision);
            AppendCoreProperty(builder, "cp:category", properties.Category);
            AppendCoreProperty(builder, "cp:version", properties.Version);
            AppendCoreProperty(builder, "cp:contentStatus", properties.ContentStatus);
            AppendCoreProperty(builder, "dc:identifier", properties.Identifier);
            AppendCoreProperty(builder, "dc:language", properties.Language);
            AppendCoreDateProperty(builder, "cp:lastPrinted", properties.LastPrinted, includeW3CType: false);
            AppendCoreDateProperty(builder, "dcterms:created", properties.Created, includeW3CType: true);
            AppendCoreDateProperty(builder, "dcterms:modified", properties.Modified, includeW3CType: true);

            builder.Append("</cp:coreProperties>");
            return builder.ToString();
        }

        private static void WriteAppPropertiesEntry(ZipArchive archive) {
            WriteTextEntry(archive, "docProps/app.xml", CreateAppPropertiesXml());
        }

        private static string CreateAppPropertiesXml() {
            return "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
                "<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\" " +
                "xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\">" +
                "<Application>OfficeIMO.Excel</Application>" +
                "</Properties>";
        }

        private static void AppendCoreProperty(System.Text.StringBuilder builder, string elementName, string? value) {
            if (string.IsNullOrEmpty(value)) {
                return;
            }

            builder.Append('<');
            builder.Append(elementName);
            builder.Append('>');
            AppendXmlEscaped(builder, value!);
            builder.Append("</");
            builder.Append(elementName);
            builder.Append('>');
        }

        private static void AppendCoreDateProperty(System.Text.StringBuilder builder, string elementName, DateTime? value, bool includeW3CType) {
            if (!value.HasValue) {
                return;
            }

            builder.Append('<');
            builder.Append(elementName);
            if (includeW3CType) {
                builder.Append(" xsi:type=\"dcterms:W3CDTF\"");
            }

            builder.Append('>');
            AppendXmlEscaped(builder, value.Value.ToUniversalTime().ToString("yyyy-MM-dd'T'HH:mm:ss'Z'", CultureInfo.InvariantCulture));
            builder.Append("</");
            builder.Append(elementName);
            builder.Append('>');
        }

        private static void WriteSharedStringsEntry(ZipArchive archive, SharedStringTable sharedStrings) {
            WriteOpenXmlElementEntry(archive, "xl/sharedStrings.xml", sharedStrings);
        }

        private static void WriteOpenXmlElementEntry(ZipArchive archive, string path, OpenXmlElement element) {
            var entry = archive.CreateEntry(path, CompressionLevel.Fastest);
            using var stream = entry.Open();
            using var writer = CreateFastXmlWriter(stream);
            writer.WriteStartDocument();
            element.WriteTo(writer);
        }

        private static void WriteBinaryEntry(ZipArchive archive, string path, byte[] bytes) {
            var entry = archive.CreateEntry(path, CompressionLevel.Fastest);
            using var stream = entry.Open();
            stream.Write(bytes, 0, bytes.Length);
        }

        private static void WriteRawPartEntry(ZipArchive archive, string path, byte[] bytes) {
            var entry = archive.CreateEntry(path, CompressionLevel.Fastest);
            using var stream = entry.Open();
            stream.Write(bytes, 0, bytes.Length);
        }

        private static void WriteWorksheetRelationshipsEntry(ZipArchive archive, FastWorksheetPackageModel worksheet) {
            var builder = new System.Text.StringBuilder(160 + worksheet.TablePartPaths.Count * 180 + worksheet.HyperlinkRelationships.Count * 220);
            builder.Append("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
            builder.Append("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
            foreach (var item in worksheet.TablePartPaths.OrderBy(static item => item.Key, StringComparer.Ordinal)) {
                builder.Append("<Relationship Id=\"");
                AppendXmlEscaped(builder, item.Key);
                builder.Append("\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/table\" Target=\"");
                AppendXmlEscaped(builder, item.Value);
                builder.Append("\"/>");
            }

            foreach (var relationship in worksheet.HyperlinkRelationships.OrderBy(static item => item.Id, StringComparer.Ordinal)) {
                builder.Append("<Relationship Id=\"");
                AppendXmlEscaped(builder, relationship.Id);
                builder.Append("\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\"");
                AppendXmlEscaped(builder, relationship.Target);
                builder.Append('"');
                if (relationship.IsExternal) {
                    builder.Append(" TargetMode=\"External\"");
                }

                builder.Append("/>");
            }

            builder.Append("</Relationships>");
            WriteTextEntry(archive, worksheet.RelationshipsPath, builder.ToString());
        }

    }
}
