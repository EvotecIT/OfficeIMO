using DocumentFormat.OpenXml.CustomProperties;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using System.Xml;
using VerifyTests;
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
        var diffDisabled = Environment.GetEnvironmentVariable("DiffEngine_Disabled");
        if (string.IsNullOrEmpty(diffDisabled))
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
        NormalizePart(document.MainDocumentPart);
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

    private static void NormalizePart(OpenXmlPart? part) {
        if (part is null) {
            return;
        }

        NormalizeRootElement(part.RootElement);

        if (part is WordprocessingCommentsPart commentsPart && commentsPart.RootElement != null) {
            foreach (var comment in commentsPart.RootElement.Descendants<Comment>()) {
                comment.Date = DateTime.MaxValue;
            }
        }

        if (part is NumberingDefinitionsPart numberingPart && numberingPart.RootElement != null) {
            var i = 1;
            foreach (var nsid in numberingPart.RootElement.Descendants<Nsid>()) {
                nsid.Val = i.ToString("X8");
                i++;
            }
        }

        foreach (var childPart in part.Parts) {
            NormalizePart(childPart.OpenXmlPart);
        }
    }

    private static void NormalizeRootElement(DocumentFormat.OpenXml.OpenXmlElement? rootElement) {
        if (rootElement is null) {
            return;
        }

        if (rootElement is Document document) {
            NormalizeDocumentReferences(document);
        } else {
            NormalizeRelationshipReferences(rootElement);
        }

        NormalizeDrawingReferences(rootElement);
        NormalizeSectionProperties(rootElement);
    }

    private static void NormalizeDocumentReferences(Document document) {
        NormalizeRelationshipReferences(document);

        var i = 1;
        foreach (var headerReference in document.Descendants<HeaderReference>()) {
            headerReference.Id = "R" + i.ToString("X8");
            i++;
        }

        i = 1;
        foreach (var footerReference in document.Descendants<FooterReference>()) {
            footerReference.Id = "R" + i.ToString("X8");
            i++;
        }
    }

    private static void NormalizeRelationshipReferences(DocumentFormat.OpenXml.OpenXmlElement rootElement) {
        var i = 1;
        foreach (var hyperlink in rootElement.Descendants<Hyperlink>()) {
            hyperlink.Id = "R" + i.ToString("X8");
            i++;
        }

        i = 1;
        foreach (var chartReference in rootElement.Descendants<ChartReference>()) {
            chartReference.Id = "R" + i.ToString("X8");
            i++;
        }

        i = 1;
        foreach (var blip in rootElement.Descendants<Blip>()) {
            if (blip.Embed != null) {
                blip.Embed = "R" + i.ToString("X8");
                i++;
            }

            if (blip.Link != null) {
                blip.Link = "R" + i.ToString("X8");
                i++;
            }
        }
    }

    private static void NormalizeDrawingReferences(DocumentFormat.OpenXml.OpenXmlElement rootElement) {
        var i = 1;
        foreach (var docProperties in rootElement.Descendants<DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties>()) {
            docProperties.Id = (UInt32Value)(uint)i;
            i++;
        }

        i = 1;
        foreach (var anchor in rootElement.Descendants<DocumentFormat.OpenXml.Drawing.Wordprocessing.Anchor>()) {
            anchor.AnchorId = i.ToString("X8");
            anchor.EditId = "E" + i.ToString("X8");
            i++;
        }

        i = 1;
        foreach (var inline in rootElement.Descendants<Inline>()) {
            inline.AnchorId = i.ToString("X8");
            inline.EditId = "E" + i.ToString("X8");
            i++;
        }
    }

    private static void NormalizeSectionProperties(DocumentFormat.OpenXml.OpenXmlElement rootElement) {
        var i = 1;
        foreach (var sectionProperties in rootElement.Descendants<SectionProperties>()) {
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
