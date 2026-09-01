using System.Reflection;
using OfficeIMO.Excel.Legacy;
using OfficeIMO.Reader.Excel;
using OfficeIMO.Word.Legacy;

namespace OfficeIMO.LegacyImport.Tests;

public sealed class LegacyPublicApiContracts {
    [Fact]
    public void LegacyReportsAndResultsUseTheSharedLossContract() {
        Assert.Contains(typeof(IOfficeConversionReport), typeof(OfficeLegacyImportReport).GetInterfaces());
        Assert.NotNull(typeof(OfficeLegacyImportReport).GetMethod("RequireNoLoss", Type.EmptyTypes));
        Assert.Null(typeof(OfficeLegacyImportReport).GetMethod("RequireStructuredNoLoss", Type.EmptyTypes));

        AssertImportResult(typeof(LegacyWordImportResult));
        AssertImportResult(typeof(LegacySpreadsheetImportResult));
    }

    [Fact]
    public void LegacyDetectionAndImportSupportTheSameInputKinds() {
        AssertImporterInputMatrix(typeof(LegacyWordImporter));
        AssertImporterInputMatrix(typeof(LegacySpreadsheetImporter));
        Assert.NotNull(typeof(OfficeDocumentReaderBuilderExcelExtensions)
            .GetMethod("AddExcelAndLegacyHandlers", BindingFlags.Public | BindingFlags.Static));
    }

    private static void AssertImportResult(Type resultType) {
        Assert.NotNull(resultType.GetProperty("Value"));
        Assert.NotNull(resultType.GetProperty("Report"));
        Assert.NotNull(resultType.GetProperty("HasLoss"));
        Assert.NotNull(resultType.GetMethod("RequireValue", Type.EmptyTypes));
        Assert.NotNull(resultType.GetMethod("RequireNoLoss", Type.EmptyTypes));
        Assert.Null(resultType.GetProperty("Document"));
    }

    private static void AssertImporterInputMatrix(Type importerType) {
        MethodInfo[] methods = importerType.GetMethods(BindingFlags.Public | BindingFlags.Static);
        Type[] inputTypes = { typeof(string), typeof(Stream), typeof(byte[]) };
        foreach (Type inputType in inputTypes) {
            Assert.Contains(methods, method => HasFirstParameter(method, "Detect", inputType));
            Assert.Contains(methods, method => HasFirstParameter(method, "Import", inputType));
        }
    }

    private static bool HasFirstParameter(MethodInfo method, string name, Type parameterType) =>
        method.Name == name
        && method.GetParameters().FirstOrDefault()?.ParameterType == parameterType;
}
