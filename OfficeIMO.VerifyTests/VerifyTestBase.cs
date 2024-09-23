using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using System.Xml;
using DocumentFormat.OpenXml.CustomProperties;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using VerifyTests;
using VerifyXunit;
using Formatting = System.Xml.Formatting;
using Hyperlink = DocumentFormat.OpenXml.Wordprocessing.Hyperlink;

namespace OfficeIMO.VerifyTests;

public abstract class VerifyTestBase {
    private const string RowDelimiter = "<!--------------------------------------------------------------------------------------------------------------------->";

    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };
    private static readonly string LastTime =
        DateTimeOffset.MaxValue.ToString("yyyy-MM-ddTHH:mm:ssZ", CultureInfo.InvariantCulture);

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

    protected static async Task<string> ToVerifyResult(WordprocessingDocument document) {
        NormalizeWord(document);

        var result = new StringBuilder();

        result.AppendLine(nameof(document.PackageProperties));
        result.AppendLine(RowDelimiter);
        var packageProperties = JsonSerializer.Serialize(document.PackageProperties, JsonOptions);
        result.AppendLine(packageProperties);
        result.AppendLine(RowDelimiter);

        foreach (var id in document.Parts) {
            var r = await GetVerifyResult(id);
            if (string.IsNullOrEmpty(r)) continue;
            result.Append(r);
        }

        return result.ToString();
    }

    private static async Task<string> GetVerifyResult(IdPartPair id) {
        if (id.OpenXmlPart.RootElement is null)
            return "";

        var result = new StringBuilder();
        var xml = FormatXml(id.OpenXmlPart.RootElement.OuterXml);
        result.AppendLine(id.OpenXmlPart.Uri.ToString());
        result.AppendLine(RowDelimiter);
        result.AppendLine(xml);
        result.AppendLine(RowDelimiter);

        foreach (var part in id.OpenXmlPart.Parts) {
            var r = await GetVerifyResult(part);
            if (string.IsNullOrEmpty(r)) continue;
            result.Append(r);
        }

        return result.ToString();
    }

    private static void NormalizeWord(WordprocessingDocument document) {
        NormalizeDocument(document.MainDocumentPart?.Document);
        NormalizeCustomFilePropertiesPart(document.CustomFilePropertiesPart);
    }

    private static string FormatXml(string value) {
        var xDoc = new XmlDocument();
        xDoc.LoadXml(value);
        xDoc.Normalize();
        xDoc.PreserveWhitespace = true;

        var sb = new StringBuilder();
        using var writer = new StringWriter(sb, CultureInfo.InvariantCulture);
        using var xTarget = new XmlTextWriter(writer);
        xTarget.Formatting = Formatting.Indented;
        xTarget.Indentation = 2;
        xDoc.WriteContentTo(xTarget);

        return sb.ToString();
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

        if (document.MainDocumentPart!.GetPartsOfType<WordprocessingCommentsPart>().Any()) {
            foreach (var comment in document.MainDocumentPart.WordprocessingCommentsPart!.RootElement!.Descendants<Comment>()) {
                comment.Date = DateTime.MaxValue;
            }
        }

        if (document.MainDocumentPart!.GetPartsOfType<NumberingDefinitionsPart>().Any()) {
            i = 1;
            foreach (var nsid in document.MainDocumentPart.NumberingDefinitionsPart!.RootElement!.Descendants<Nsid>()) {
                nsid.Val = i.ToString("X8");
                i++;
            }
        }

        // Normalize RSID in SectionProperties
        i = 1;
        foreach (var sectionProperties in document.Descendants<SectionProperties>()) {
            if (sectionProperties.RsidRPr != null) {
                sectionProperties.RsidRPr = "R" + i.ToString("X8");
                i++;
            }
            if (sectionProperties.RsidR != null) {
                sectionProperties.RsidR = "R" + i.ToString("X8");
                i++;
            }
            if (sectionProperties.RsidDel != null) {
                sectionProperties.RsidDel = "R" + i.ToString("X8");
                i++;
            }
        }
    }

    private static void NormalizeCustomFilePropertiesPart(CustomFilePropertiesPart? part) {
        var fileTime = part?.Properties
            .FirstOrDefault(x => ((CustomDocumentProperty?)x)?.VTFileTime != null);
        if (fileTime is CustomDocumentProperty property) {
            property.VTFileTime!.Text = LastTime;
        }
    }
}
