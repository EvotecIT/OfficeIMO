using System.IO;
using System.Text.RegularExpressions;

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

}
