using System.IO;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Security.Tests;

public sealed class OfficeVbaFailureBoundaryTests {
    [Theory]
    [InlineData("project")]
    [InlineData("source")]
    [InlineData("module")]
    [InlineData("projectName")]
    [InlineData("referenceName")]
    [InlineData("referenceId")]
    [InlineData("unsupported")]
    public void LoadReportsMalformedOrUnavailableTextAsInvalidData(string field) {
        var exception = Assert.Throws<InvalidDataException>(() => OfficeVbaProject.Load(OfficeVbaMalformedTextFixtures.Create(field)));
        Assert.NotNull(exception.InnerException);
        if (field == "unsupported") Assert.IsType<NotSupportedException>(exception.InnerException);
        else Assert.IsType<DecoderFallbackException>(exception.InnerException);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void NormalizationEmitsOneNameAndOneValuePerClassAttribute(bool document) {
        var project = OfficeVbaProject.Create();
        const string supplied = "\r\n  Attribute\tVB_Name = \"InputName\"\r\n"
            + "Attribute\tVB_PredeclaredId = False\r\n  Attribute VB_Exposed = True\r\n"
            + "Attribute\tVB_TemplateDerived = False\r\n";
        OfficeVbaModule module;
        if (document) {
            module = project.AddDocumentModule("ThisWorkbook", "", new Guid("00020819-0000-0000-C000-000000000046"));
            project.SetModuleSource(module.Name, supplied);
        } else {
            module = project.AddModule("Counter", supplied + "  Attribute\tVB_Base = \"0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}\"\r\n", OfficeVbaModuleKind.Class);
        }
        Assert.StartsWith("Attribute VB_Name = \"" + module.Name + "\"", module.Source);
        var attributes = Regex.Matches(module.Source, @"^\s*Attribute\s+(VB_\w+)\s*=", RegexOptions.Multiline | RegexOptions.IgnoreCase)
            .Select(match => match.Groups[1].Value).ToArray();
        Assert.Equal(attributes.Length, attributes.Distinct(StringComparer.OrdinalIgnoreCase).Count());
        Assert.Equal(module.Source, OfficeVbaProject.Load(project.Write().GetBytes()).GetModule(module.Name).Source);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void SourceReplacementRejectsMalformedBaseBeforeChangingDocumentOrDesigner(bool designer, bool import) {
        var project = designer
            ? OfficeVbaProject.Load(File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "TestData", "Vba", "Excel-nested-form.bin")))
            : OfficeVbaProject.Create();
        string name = designer ? "FrmNested" : "ThisWorkbook";
        if (!designer) project.AddDocumentModule(name, "'original\r\n", new Guid("00020819-0000-0000-C000-000000000046"));
        project.AddModule("PendingEdit", "'before\r\n");
        byte[] original = project.Write().GetBytes();
        project = OfficeVbaProject.Load(original);
        string source = project.GetModule(name).Source;
        Assert.NotNull(OfficeVbaText.GetBaseIdentity(source));
        string folder = Path.Combine(Path.GetTempPath(), "OfficeIMO-vba-base-" + Guid.NewGuid().ToString("N"));
        try {
            string? targetFile = null;
            if (import) {
                project.ExportSources(folder);
                var manifest = XDocument.Load(Path.Combine(folder, "vba-project.xml"));
                XElement pending = manifest.Root!.Elements("module").Single(entry => (string?)entry.Attribute("name") == "PendingEdit");
                File.WriteAllText(Path.Combine(folder, (string)pending.Attribute("file")!), "'valid pending edit\r\n");
                pending.Remove();
                manifest.Root.AddFirst(pending);
                manifest.Save(Path.Combine(folder, "vba-project.xml"));
                targetFile = (string)manifest.Root.Elements("module").Single(entry => (string?)entry.Attribute("name") == name).Attribute("file")!;
            }
            foreach (string value in new[] { "False", "\"\"", "\"unterminated", "\"identity\" trailing" }) {
                string replacement = "Attribute VB_Base = " + value + "\r\nOption Explicit\r\n";
                if (import) {
                    File.WriteAllText(Path.Combine(folder, targetFile!), replacement);
                    Assert.Throws<InvalidDataException>(() => project.ImportSources(folder));
                } else Assert.Throws<InvalidDataException>(() => project.SetModuleSource(name, replacement));
                Assert.Equal(source, project.GetModule(name).Source);
                Assert.Contains("'before", project.GetModule("PendingEdit").Source);
                Assert.False(project.HasChanges);
                Assert.Equal(original, project.Write().GetBytes());
            }
            string identity = OfficeVbaText.GetBaseIdentity(source)!;
            string valid = "  Attribute\tVB_Base = \"" + identity + "\"\r\n'valid replacement\r\n";
            if (import) {
                File.WriteAllText(Path.Combine(folder, targetFile!), valid);
                project.ImportSources(folder);
            } else project.SetModuleSource(name, valid);
            var written = OfficeVbaProject.Load(project.Write().GetBytes());
            Assert.Equal(identity, OfficeVbaText.GetBaseIdentity(written.GetModule(name).Source));
            Assert.Contains("'valid replacement", written.GetModule(name).Source);
        } finally { if (Directory.Exists(folder)) Directory.Delete(folder, true); }
    }

}
