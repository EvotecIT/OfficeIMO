using System.Reflection;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfPublicApiContractTests {
    private static readonly string[] InternalEngineTypeNames = {
        "PdfAcroFormEditor",
        "PdfAnnotationEditor",
        "PdfAnnotationFlattener",
        "PdfAttachmentEditor",
        "PdfAttachmentExtractor",
        "PdfBookmarkEditor",
        "PdfComplianceAnalyzer",
        "PdfDebugger",
        "PdfDiagnostics",
        "PdfFormData",
        "PdfFormFiller",
        "PdfImageExtractor",
        "PdfIncrementalUpdater",
        "PdfInspector",
        "PdfLayoutDebugOverlay",
        "PdfMerger",
        "PdfMetadataEditor",
        "PdfMutationPlanner",
        "PdfOcr",
        "PdfOptimizer",
        "PdfPageEditor",
        "PdfPageExtractor",
        "PdfPageImageRenderer",
        "PdfPageImporter",
        "PdfRedactionApplier",
        "PdfRedactionPlanner",
        "PdfRedactionVerification",
        "PdfSanitizer",
        "PdfSecurityEditor",
        "PdfSignatureMutationAnalyzer",
        "PdfSignatureValidator",
        "PdfStamper",
        "PdfTextDiagnostics",
        "PdfTextExtractor",
        "PdfValidator"
    };

    [Fact]
    public void FacadeExposesOneCreateOpenAnalyzeWorkflowWithoutLegacyLoad() {
        MethodInfo[] methods = typeof(PdfDocument).GetMethods(BindingFlags.Public | BindingFlags.Static | BindingFlags.Instance);

        Assert.Contains(methods, method =>
            method.Name == nameof(PdfDocument.Create) &&
            method.IsStatic);
        Assert.Contains(methods, method =>
            method.Name == nameof(PdfDocument.Open) &&
            method.IsStatic &&
            method.GetParameters().FirstOrDefault()?.ParameterType == typeof(byte[]));
        Assert.Contains(methods, method =>
            method.Name == nameof(PdfDocument.Open) &&
            method.IsStatic &&
            method.GetParameters().FirstOrDefault()?.ParameterType == typeof(string));
        Assert.Contains(methods, method =>
            method.Name == nameof(PdfDocument.Open) &&
            method.IsStatic &&
            method.GetParameters().FirstOrDefault()?.ParameterType == typeof(Stream));
        Assert.Contains(methods, method =>
            method.Name == nameof(PdfDocument.OpenAsync) &&
            method.IsStatic);
        Assert.Contains(methods, method =>
            method.Name == nameof(PdfDocument.Analyze) &&
            !method.IsStatic &&
            method.ReturnType == typeof(PdfAnalysisReport));
        Assert.DoesNotContain(methods, method => method.Name == "Load");

        Assert.Equal(typeof(PdfDocumentReader), typeof(PdfDocument).GetProperty(nameof(PdfDocument.Read))?.PropertyType);
        Assert.Equal(typeof(PdfDocumentPages), typeof(PdfDocument).GetProperty(nameof(PdfDocument.Pages))?.PropertyType);
        Assert.Equal(typeof(PdfDocumentForms), typeof(PdfDocument).GetProperty(nameof(PdfDocument.Forms))?.PropertyType);
    }

    [Fact]
    public void FacadeOwnedEnginesAreNotExportedAsDuplicateStaticBrains() {
        Assembly assembly = typeof(PdfDocument).Assembly;
        var exportedNames = assembly.GetExportedTypes()
            .Select(type => type.Name)
            .ToHashSet(StringComparer.Ordinal);

        Assert.All(InternalEngineTypeNames, name => Assert.DoesNotContain(name, exportedNames));
    }

    [Fact]
    public void PublicSurfaceAndRuntimeDependenciesStayBounded() {
        Assembly assembly = typeof(PdfDocument).Assembly;
        Type[] exportedTypes = assembly.GetExportedTypes();
        int publicMemberCount = exportedTypes.Sum(type =>
            type.GetMembers(
                BindingFlags.Public |
                BindingFlags.Instance |
                BindingFlags.Static |
                BindingFlags.DeclaredOnly).Length);

        Assert.InRange(exportedTypes.Length, 1, 480);
        Assert.InRange(publicMemberCount, 1, 9400);

        string[] officeReferences = assembly.GetReferencedAssemblies()
            .Select(reference => reference.Name)
            .Where(name => name != null && name.StartsWith("OfficeIMO.", StringComparison.Ordinal))
            .Cast<string>()
            .OrderBy(name => name, StringComparer.Ordinal)
            .ToArray();
        Assert.Equal(new[] { "OfficeIMO.Drawing" }, officeReferences);
    }
}
