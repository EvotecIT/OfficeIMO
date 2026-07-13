using System.Reflection;
using OfficeIMO.Excel.GoogleSheets;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Excel {
    [Fact]
    public void GoogleSheetsPlanningApisUseBuildVocabulary() {
        string[] names = typeof(ExcelGoogleSheetsExtensions)
            .GetMethods(BindingFlags.Public | BindingFlags.Static | BindingFlags.DeclaredOnly)
            .Select(static method => method.Name)
            .ToArray();

        Assert.Contains("BuildGoogleSheetsPlan", names);
        Assert.Contains("BuildGoogleSheetsBatch", names);
        Assert.DoesNotContain("CreateGoogleSheetsTranslationPlan", names);
        Assert.DoesNotContain("CreateGoogleSheetsBatch", names);
    }
}
