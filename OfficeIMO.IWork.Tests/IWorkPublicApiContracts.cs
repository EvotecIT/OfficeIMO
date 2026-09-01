using System.Reflection;
using OfficeIMO.Excel.IWork;
using OfficeIMO.PowerPoint.IWork;
using OfficeIMO.Word.IWork;

namespace OfficeIMO.IWork.Tests;

public sealed class IWorkPublicApiContracts {
    [Fact]
    public void ConversionReportsAndResultsUseTheSharedLossContract() {
        Assert.Contains(typeof(IOfficeConversionReport), typeof(IWorkConversionReport).GetInterfaces());

        AssertConversionResult(typeof(PagesToWordResult));
        AssertConversionResult(typeof(NumbersToExcelResult));
        AssertConversionResult(typeof(KeynoteToPowerPointResult));
    }

    [Fact]
    public void SourceReadingAndDestinationConversionRemainSeparate() {
        Assert.Null(typeof(IWorkReadOptions).GetProperty("ImportMode"));
        Assert.NotNull(typeof(IWorkConversionOptions).GetProperty("Mode"));

        AssertSourceExtension(typeof(WordIWorkConverter), "ToWordDocumentResult");
        AssertSourceExtension(typeof(ExcelIWorkConverter), "ToExcelDocumentResult");
        AssertSourceExtension(typeof(PowerPointIWorkConverter), "ToPowerPointPresentationResult");

        MethodInfo[] converterMethods = new[] {
                typeof(WordIWorkConverter),
                typeof(ExcelIWorkConverter),
                typeof(PowerPointIWorkConverter)
            }
            .SelectMany(static type => type.GetMethods(BindingFlags.Public | BindingFlags.Static))
            .ToArray();
        Assert.DoesNotContain(converterMethods, static method =>
            method.Name.StartsWith("Load", StringComparison.Ordinal));
    }

    [Fact]
    public void OpenSupportsDetectedStreamsAndByteArrays() {
        MethodInfo[] methods = typeof(IWorkSourceDocument).GetMethods(
            BindingFlags.Public | BindingFlags.Static);

        Assert.Contains(methods, static method => HasLeadingParameters(method, typeof(Stream), typeof(IWorkReadOptions)));
        Assert.Contains(methods, static method => HasLeadingParameters(method, typeof(byte[]), typeof(IWorkReadOptions)));
        Assert.Contains(methods, static method => HasLeadingParameters(method, typeof(Stream), typeof(IWorkDocumentKind)));
        Assert.Contains(methods, static method => HasLeadingParameters(method, typeof(byte[]), typeof(IWorkDocumentKind)));
    }

    [Fact]
    public void DiscardedPreReleaseNamesAreNotExported() {
        AssertTypesAreNotExported(typeof(IWorkSourceDocument).Assembly,
            "IWorkImportMode", "IWorkImportReport");
        AssertTypesAreNotExported(typeof(PagesToWordResult).Assembly,
            "IWorkPagesLoadResult");
        AssertTypesAreNotExported(typeof(NumbersToExcelResult).Assembly,
            "IWorkNumbersLoadResult");
        AssertTypesAreNotExported(typeof(KeynoteToPowerPointResult).Assembly,
            "IWorkKeynoteLoadResult");
    }

    private static void AssertConversionResult(Type resultType) {
        Assert.NotNull(resultType.GetProperty("Value"));
        Assert.NotNull(resultType.GetProperty("Report"));
        Assert.NotNull(resultType.GetProperty("HasLoss"));
        Assert.NotNull(resultType.GetMethod("RequireValue", Type.EmptyTypes));
        Assert.NotNull(resultType.GetMethod("RequireNoLoss", Type.EmptyTypes));
        Assert.Null(resultType.GetProperty("Document"));
        Assert.Null(resultType.GetProperty("ImportReport"));
        Assert.Null(resultType.GetProperty("HasConversionLoss"));
    }

    private static void AssertSourceExtension(Type converterType, string methodName) {
        MethodInfo method = Assert.Single(
            converterType.GetMethods(BindingFlags.Public | BindingFlags.Static),
            candidate => candidate.Name == methodName);
        Assert.Equal(typeof(IWorkSourceDocument), method.GetParameters()[0].ParameterType);
        Assert.NotNull(method.GetCustomAttribute<System.Runtime.CompilerServices.ExtensionAttribute>());
    }

    private static void AssertTypesAreNotExported(Assembly assembly, params string[] names) {
        Type[] exportedTypes = assembly.GetExportedTypes();
        foreach (string name in names) {
            Assert.DoesNotContain(exportedTypes, type => type.Name == name);
        }
    }

    private static bool HasLeadingParameters(MethodInfo method, params Type[] expected) {
        if (method.Name != "Open") return false;
        ParameterInfo[] parameters = method.GetParameters();
        return parameters.Length >= expected.Length
            && expected.Select((type, index) => parameters[index].ParameterType == type).All(static match => match);
    }
}
