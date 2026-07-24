using System.Reflection;
using System.Threading.Tasks;
using OfficeIMO.Drawing;
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
        Assert.Equal(typeof(PdfPipelineReport), typeof(PdfDocument).GetProperty(nameof(PdfDocument.Pipeline))?.PropertyType);
        Assert.Equal(typeof(IOfficeTextShapingProvider), typeof(PdfOptions).GetProperty(nameof(PdfOptions.TextShapingProvider))?.PropertyType);
        Assert.Equal(typeof(Func<string, IReadOnlyList<int>>), typeof(PdfOptions).GetProperty(nameof(PdfOptions.TextLineBreakCallback))?.PropertyType);

        Assert.Equal(
            2,
            methods.Count(method =>
                method.Name == nameof(PdfDocument.Save) &&
                method.ReturnType == typeof(PdfSaveResult)));
        Assert.Equal(
            2,
            methods.Count(method =>
                method.Name == nameof(PdfDocument.TrySave) &&
                method.ReturnType == typeof(PdfSaveResult)));
        Assert.Equal(
            2,
            methods.Count(method =>
                method.Name == nameof(PdfDocument.SaveAsync) &&
                method.ReturnType == typeof(Task<PdfSaveResult>)));
        Assert.Equal(
            2,
            methods.Count(method =>
                method.Name == nameof(PdfDocument.TrySaveAsync) &&
                method.ReturnType == typeof(Task<PdfSaveResult>)));
    }

    [Fact]
    public void FacadeOwnedEnginesAreNotExportedAsDuplicateStaticBrains() {
        Assembly assembly = typeof(PdfDocument).Assembly;
        var exportedNames = assembly.GetExportedTypes()
            .Select(type => type.Name)
            .ToHashSet(StringComparer.Ordinal);

        Assert.All(InternalEngineTypeNames, name => Assert.DoesNotContain(name, exportedNames));
        Assert.Null(assembly.GetType("OfficeIMO.Pdf.IPdfTextShapingProvider"));
        Assert.Null(assembly.GetType("OfficeIMO.Pdf.PdfTextShapingRequest"));
        Assert.Null(assembly.GetType("OfficeIMO.Pdf.PdfTextShapingResult"));
        Assert.Null(assembly.GetType("OfficeIMO.Pdf.PdfShapedGlyph"));
        Assert.Null(assembly.GetType("OfficeIMO.Pdf.PdfTextDirection"));
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

        Assert.InRange(exportedTypes.Length, 1, 500);
        Assert.InRange(publicMemberCount, 1, 9900);

        string[] officeReferences = assembly.GetReferencedAssemblies()
            .Select(reference => reference.Name)
            .Where(name => name != null && name.StartsWith("OfficeIMO.", StringComparison.Ordinal))
            .Cast<string>()
            .OrderBy(name => name, StringComparer.Ordinal)
            .ToArray();
        Assert.Equal(new[] { "OfficeIMO.Drawing", "OfficeIMO.Security" }, officeReferences);
    }
}
