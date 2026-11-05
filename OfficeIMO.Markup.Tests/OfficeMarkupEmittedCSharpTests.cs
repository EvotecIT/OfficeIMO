#if NET8_0_OR_GREATER
using System.Diagnostics;
using System.IO.Compression;
using System.Xml.Linq;
using Xunit;
using Xunit.Abstractions;

namespace OfficeIMO.Markup.Tests;

public sealed class OfficeMarkupEmittedCSharpTests(ITestOutputHelper output) {
    [Theory]
    [InlineData(OfficeMarkupProfile.Workbook, ".xlsx", "xl/sharedStrings.xml")]
    [InlineData(OfficeMarkupProfile.Presentation, ".pptx", "ppt/slides/slide1.xml")]
    [InlineData(OfficeMarkupProfile.Document, ".docx", "word/document.xml")]
    public void EmittedProgramsCompileAndWriteNativeDocuments(OfficeMarkupProfile profile, string extension, string contentEntry) {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-emitted-csharp-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            string path = Path.Combine(root, "generated" + extension);
            var parsed = OfficeMarkupParser.Parse(Source(profile), new OfficeMarkupParserOptions { Profile = profile });
            Assert.False(parsed.HasErrors);
            string code = new OfficeMarkupCSharpEmitter().Emit(parsed.Document,
                new OfficeMarkupEmitterOptions { FilePathVariable = "@\"" + path.Replace("\"", "\"\"") + "\"" });
            File.WriteAllText(Path.Combine(root, "Program.cs"), code);
            string framework = "net" + Environment.Version.Major + ".0";
            var references = Directory.GetFiles(AppContext.BaseDirectory, "*.dll")
                .Where(file => Path.GetFileName(file).StartsWith("OfficeIMO", StringComparison.Ordinal)
                    || Path.GetFileName(file).StartsWith("DocumentFormat.OpenXml", StringComparison.Ordinal)
                    || Path.GetFileName(file) == "System.IO.Packaging.dll")
                .Select(file => new XElement("Reference", new XAttribute("Include", Path.GetFileNameWithoutExtension(file)),
                    new XElement("HintPath", file)));
            var project = new XElement("Project", new XAttribute("Sdk", "Microsoft.NET.Sdk"),
                new XElement("PropertyGroup", new XElement("TargetFramework", framework), new XElement("OutputType", "Exe"),
                    new XElement("ImplicitUsings", "enable"), new XElement("Nullable", "enable")),
                new XElement("ItemGroup", references));
            File.WriteAllText(Path.Combine(root, "Generated.csproj"), project.ToString());
            File.WriteAllText(Path.Combine(root, "NuGet.Config"), "<configuration><packageSources><clear /></packageSources></configuration>");
            Run(root, "build", "Generated.csproj", "-c", "Release", "-m:2", "-nodeReuse:false", "-p:UseSharedCompilation=false", "-v:minimal");
            Run(root, Path.Combine(root, "bin", "Release", framework, "Generated.dll"));
            using var package = ZipFile.OpenRead(path);
            using (var content = package.GetEntry(contentEntry)!.Open())
                Assert.Contains("Generated content", XDocument.Load(content).Root!.Value, StringComparison.Ordinal);
            if (profile == OfficeMarkupProfile.Workbook) AssertWorkbook(package);
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    private void Run(string root, params string[] arguments) {
        var start = new ProcessStartInfo("dotnet") { WorkingDirectory = root, UseShellExecute = false,
            RedirectStandardOutput = true, RedirectStandardError = true };
        foreach (string argument in arguments) start.ArgumentList.Add(argument);
        start.Environment["MSBUILDDISABLENODEREUSE"] = "1";
        start.Environment["DOTNET_CLI_USE_MSBUILD_SERVER"] = "0";
        using Process process = Process.Start(start)!;
        output.WriteLine("Owned SDK consumer PID: " + process.Id);
        var stdout = process.StandardOutput.ReadToEndAsync();
        var stderr = process.StandardError.ReadToEndAsync();
        if (!process.WaitForExit(120000)) {
            process.Kill();
            process.WaitForExit();
            throw new TimeoutException("The owned generated-code consumer did not finish.");
        }
        Assert.True(process.ExitCode == 0, stdout.GetAwaiter().GetResult() + stderr.GetAwaiter().GetResult());
    }

    private static void AssertWorkbook(ZipArchive package) {
        XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        using var source = package.GetEntry("xl/styles.xml")!.Open();
        var styles = XDocument.Load(source);
        Assert.Contains(styles.Descendants(ns + "font"), font =>
            font.Element(ns + "b") != null && font.Element(ns + "i") != null && font.Element(ns + "u") != null);
        Assert.Contains(styles.Descendants(ns + "fgColor"), color => (string?)color.Attribute("rgb") == "FFD9EAD3");
        Assert.Contains(styles.Descendants(ns + "color"), color => (string?)color.Attribute("rgb") == "FF112233");
        Assert.Contains(styles.Descendants(ns + "color"), color => (string?)color.Attribute("rgb") == "FF445566");
        Assert.Contains(styles.Descendants(ns + "alignment"), alignment =>
            (string?)alignment.Attribute("horizontal") == "center" && (string?)alignment.Attribute("vertical") == "center"
            && (string?)alignment.Attribute("wrapText") == "1");
        Assert.Contains(styles.Descendants(ns + "numFmt"), format => (string?)format.Attribute("formatCode") == "#,##0");
        using var contentTypes = package.GetEntry("[Content_Types].xml")!.Open();
        Assert.Equal(2, XDocument.Load(contentTypes).Descendants().Count(part =>
            (string?)part.Attribute("ContentType") == "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"));
    }

    private static string Source(OfficeMarkupProfile profile) => profile switch {
        OfficeMarkupProfile.Workbook => """
            ::range address=Data!A1
            Generated content,Value,Double
            Revenue,120,240
            Costs,50,100

            ::formula cell=Data!C2
            =B2*2

            ::format target=Data!B2:C3 numberFormat="#,##0" fill=#D9EAD3 color=#112233 bold=true italic=true underline=true align=center valign=middle wrap=true border=thin border-color=#445566

            ::table name="DataTable" range=Data!A1:C3 header=true

            ::chart type=column title="Generated chart" source=Data!DataTable cell=Dashboard!E2 width=480 height=320

            ::chart type=line title="Local chart" source=A1:C3 cell=Data!F20 width=480 height=320
            """,
        OfficeMarkupProfile.Presentation => "# Generated content\n\n" + """
            @slide {
              transition: fade duration=0.6 speed=fast advance-on-click=false advance-after=5
            }

            ::chart type=column title="Generated chart" legend=right labels=true
            Metric,Amount
            Revenue,120
            Costs,50

            ::notes
            Speaker marker
            """,
        _ => "# Generated content\n\nBody marker"
    };
}
#endif
