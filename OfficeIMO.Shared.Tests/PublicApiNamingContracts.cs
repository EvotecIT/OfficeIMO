using System.Reflection;
using OfficeIMO.Excel;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed class PublicApiNamingContracts {
    [Fact]
    public void ExcelWorksheetApisUseOneCanonicalCasing() {
        MethodInfo[] methods = typeof(ExcelDocument).GetMethods(
            BindingFlags.Public | BindingFlags.Instance | BindingFlags.Static | BindingFlags.DeclaredOnly);

        Assert.DoesNotContain(methods, method => method.Name.Contains("WorkSheet", StringComparison.Ordinal));
        Assert.DoesNotContain(methods.SelectMany(method => method.GetParameters()), parameter =>
            parameter.Name?.Contains("workSheet", StringComparison.Ordinal) == true);
        Assert.Contains(methods, method => method.Name == "AddWorksheet");
        Assert.Contains(methods, method => method.Name == "RemoveWorksheet");
        Assert.Contains(methods, method => method.Name == "CopyWorksheet");
        Assert.Contains(methods, method => method.Name == "CopyWorksheetFrom");
        Assert.Contains(methods, method => method.Name == "ReorderWorksheet");
    }

    [Fact]
    public void RtfHtmlMemoryOutputUsesStreamVocabulary() {
        string[] methodNames = typeof(HtmlRtfConverterExtensions)
            .GetMethods(BindingFlags.Public | BindingFlags.Static)
            .Select(method => method.Name)
            .ToArray();

        Assert.Contains("ToHtmlStream", methodNames);
        Assert.DoesNotContain("ToHtmlMemoryStream", methodNames);
    }
}
