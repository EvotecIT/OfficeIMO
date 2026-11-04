using System.IO.Compression;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Xml.Linq;
using OfficeIMO.Markup;
using PSWriteOffice.Cmdlets.Conversion;

namespace PSWriteOffice.Tests;

[Collection(PowerShellRunspaceCollection.Name)]
public sealed class MarkupEmitterCmdletTests {
    [Theory]
    [InlineData(OfficeMarkupProfile.Document, ".docx", "word/document.xml")]
    [InlineData(OfficeMarkupProfile.Presentation, ".pptx", "ppt/slides/slide1.xml")]
    [InlineData(OfficeMarkupProfile.Workbook, ".xlsx", "xl/sharedStrings.xml")]
    public void EmittedScriptsWriteNativeDocuments(OfficeMarkupProfile profile, string extension, string contentEntry) {
        string root = Path.Combine(AppContext.BaseDirectory, "officeimo-markup-script-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            string path = Path.Combine(root, "generated" + extension);
            var document = new OfficeMarkupDocument(profile);
            if (profile == OfficeMarkupProfile.Presentation) {
                document.Blocks.Add(new OfficeMarkupSlideBlock("Generated content"));
            } else if (profile == OfficeMarkupProfile.Workbook) {
                var parsed = OfficeMarkupParser.Parse("""
                    ::range address=Data!A1
                    Generated content,Value
                    Ready,1

                    ::format target=Data!A1:B1 fill=#D9EAD3 color=#112233 bold=true italic=true underline=true border=thin border-color=#445566
                    """, new OfficeMarkupParserOptions { Profile = profile });
                Assert.False(parsed.HasErrors);
                document = parsed.Document;
            } else {
                document.Blocks.Add(new OfficeMarkupParagraphBlock("Generated content"));
            }
            string script = new OfficeMarkupPowerShellEmitter().Emit(document);
            var state = InitialSessionState.CreateDefault();
            state.ImportPSModule(new[] { typeof(ConvertOfficeDocumentCommand).Assembly.Location });
            using var runspace = RunspaceFactory.CreateRunspace(state);
            runspace.Open();
            runspace.SessionStateProxy.SetVariable("FilePath", path);
            using var shell = PowerShell.Create();
            shell.Runspace = runspace;
            shell.AddScript("$ErrorActionPreference = 'Stop'\ntry {\n" + script + "\n} finally {\n" +
                "foreach ($owned in @($document, $workbook, $presentation)) { if ($null -ne $owned) { $owned.Dispose() } }\n}");
            shell.Invoke();
            Assert.Empty(shell.Streams.Error);
            using var package = ZipFile.OpenRead(path);
            using (var source = package.GetEntry(contentEntry)!.Open()) {
                var content = XDocument.Load(source);
                Assert.Contains("Generated content", content.Root!.Value, StringComparison.Ordinal);
            }
            if (profile == OfficeMarkupProfile.Workbook) {
                using var source = package.GetEntry("xl/styles.xml")!.Open();
                var styles = XDocument.Load(source);
                XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
                Assert.Contains(styles.Descendants(ns + "fgColor"), color => (string?)color.Attribute("rgb") == "FFD9EAD3");
                Assert.Contains(styles.Descendants(ns + "color"), color => (string?)color.Attribute("rgb") == "FF112233");
                Assert.Contains(styles.Descendants(ns + "color"), color => (string?)color.Attribute("rgb") == "FF445566");
                Assert.Contains(styles.Descendants(ns + "font"), font =>
                    font.Element(ns + "b") != null && font.Element(ns + "i") != null && font.Element(ns + "u") != null);
            }
        } finally {
            Directory.Delete(root, true);
        }
    }
}
