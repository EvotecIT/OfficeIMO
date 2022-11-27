using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Xml;
using DocumentFormat.OpenXml.CustomProperties;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using VerifyTests;
using VerifyXunit;
using Hyperlink = DocumentFormat.OpenXml.Wordprocessing.Hyperlink;

namespace OfficeIMO.VerifyTests;

[UsesVerify]
public abstract class VerifyTestBase {
    private const string RowDelimiter = "<!--------------------------------------------------------------------------------------------------------------------->";

    private static readonly XmlWriterSettings XmlWriterSettings = new() {
        Indent = true,
        NewLineOnAttributes = false,
        IndentChars = "  ",
        ConformanceLevel = ConformanceLevel.Document
    };

    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };

    static VerifyTestBase() {
        // To disable Visual Studio popping up on every test execution.
        Environment.SetEnvironmentVariable("DiffEngine_Disabled", "true");
        Environment.SetEnvironmentVariable("Verify_DisableClipboard", "true");
    }

    protected static VerifySettings GetSettings() {
        var settings = new VerifySettings();
        settings.UseDirectory("verified");
        return settings;
    }

    protected static string ToVerifyResult(WordprocessingDocument document) {
        NormalizeWord(document);

        var result = new StringBuilder();

        result.AppendLine(nameof(document.PackageProperties));
        result.AppendLine(RowDelimiter);
        var packageProperties = JsonSerializer.Serialize(document.PackageProperties, JsonOptions);
        result.AppendLine(packageProperties);
        result.AppendLine(RowDelimiter);

        foreach (var id in document.Parts) {
            if (id.OpenXmlPart.RootElement is null)
                continue;
            var xml = FormatXml(id.OpenXmlPart.RootElement.OuterXml);
            result.AppendLine(id.OpenXmlPart.Uri.ToString());
            result.AppendLine(RowDelimiter);
            result.AppendLine(xml);
            result.AppendLine(RowDelimiter);
        }

        return result.ToString();
    }

    private static void NormalizeWord(WordprocessingDocument document) {
        NormalizeDocument(document.MainDocumentPart?.Document);
        NormalizeCustomFilePropertiesPart(document.CustomFilePropertiesPart);
    }

    private static string FormatXml(string value) {
        using var textReader = new StringReader(value);
        using var xmlReader = XmlReader.Create(
            textReader, new XmlReaderSettings { ConformanceLevel = XmlWriterSettings.ConformanceLevel } );
        using var textWriter = new StringWriter();
        using (var xmlWriter = XmlWriter.Create(textWriter, XmlWriterSettings))
            xmlWriter.WriteNode(xmlReader, true);
        return textWriter.ToString();
    }

    private static void NormalizeDocument(Document? document) {
        if (document is null)
            return;

        var i = 1;
        foreach (var hyperlink in document.Descendants<Hyperlink>()) {
            hyperlink.Id = "R" + i.ToString("X8");
            i++;
        }

        i = 1;
        foreach (var headerReference in document.Descendants<HeaderReference>()) {
            headerReference.Id = "R" + i.ToString("X8");
            i++;
        }

        i = 1;
        foreach (var footerReference in document.Descendants<FooterReference>()) {
            footerReference.Id = "R" + i.ToString("X8");
            i++;
        }

        i = 1;
        foreach (var chartReference in document.Descendants<ChartReference>()) {
            chartReference.Id = "R" + i.ToString("X8");
            i++;
        }
    }

    private static void NormalizeCustomFilePropertiesPart(CustomFilePropertiesPart? part) {
        var fileTime = part?.Properties
            .FirstOrDefault(x => ((CustomDocumentProperty?)x)?.VTFileTime != null);
        if (fileTime != null) {
            ((CustomDocumentProperty?) fileTime)!.VTFileTime!.Text =
                DateTimeOffset.MaxValue.ToString("yyyy-MM-ddTHH:mm:ssZ", CultureInfo.InvariantCulture);
        }
    }
}
