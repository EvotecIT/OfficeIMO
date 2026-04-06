using System.IO.Compression;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using OfficeIMO.Excel;

namespace OfficeIMO.Tests;

internal static class ExcelCompatibilityCorpusBuilder {
    internal static string CreateWorkbook(string filePath, Action<ExcelDocument> configureWorkbook, bool rewriteAppContentTypeToXml = false) {
        if (configureWorkbook == null) throw new ArgumentNullException(nameof(configureWorkbook));

        using (var document = ExcelDocument.Create(filePath)) {
            configureWorkbook(document);
            document.Save();
        }

        if (rewriteAppContentTypeToXml) {
            RewriteContentTypes(filePath, root => {
                XNamespace ns = root.Name.Namespace;
                var appOverride = root.Elements(ns + "Override")
                    .FirstOrDefault(e => string.Equals((string?)e.Attribute("PartName"), "/docProps/app.xml", StringComparison.OrdinalIgnoreCase))
                    ?? throw new InvalidOperationException("Missing /docProps/app.xml override.");
                appOverride.SetAttributeValue("ContentType", "application/xml");
            });
        }

        return filePath;
    }

    private static void RewriteContentTypes(string filePath, Action<XElement> mutateRoot) {
        using var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
        using var archive = new ZipArchive(fileStream, ZipArchiveMode.Update, leaveOpen: false);

        var contentTypes = archive.GetEntry("[Content_Types].xml") ?? throw new InvalidOperationException("Missing content types.");

        string xml;
        using (var entryStream = contentTypes.Open())
        using (var reader = new StreamReader(entryStream, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false))) {
            xml = reader.ReadToEnd();
        }

        var document = XDocument.Parse(xml, LoadOptions.PreserveWhitespace);
        var root = document.Root ?? throw new InvalidOperationException("Missing content types root.");
        mutateRoot(root);

        contentTypes.Delete();
        var replacement = archive.CreateEntry("[Content_Types].xml", CompressionLevel.NoCompression);
        using var replacementStream = replacement.Open();
        var settings = new XmlWriterSettings {
            Encoding = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false),
            Indent = false,
            OmitXmlDeclaration = false,
            NewLineHandling = NewLineHandling.None
        };
        using var writer = XmlWriter.Create(replacementStream, settings);
        document.Save(writer);
        writer.Flush();
    }
}
