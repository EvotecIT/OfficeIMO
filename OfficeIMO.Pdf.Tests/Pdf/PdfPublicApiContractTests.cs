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
            method.Name == nameof(PdfDocument.Create) &&
            method.IsStatic &&
            method.GetParameters().FirstOrDefault()?.ParameterType == typeof(Action<PdfCompose>));
        Assert.All(
            methods.Where(method => method.Name == nameof(PdfDocument.Create) && method.IsStatic),
            method => Assert.Equal(typeof(Action<PdfCompose>), method.GetParameters()[0].ParameterType));
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
        Assert.Equal(typeof(PdfDocumentSecurity), typeof(PdfDocument).GetProperty(nameof(PdfDocument.Security))?.PropertyType);
        Assert.Equal(typeof(PdfDocumentRedactions), typeof(PdfDocument).GetProperty(nameof(PdfDocument.Redactions))?.PropertyType);
        Assert.Equal(typeof(PdfDocumentOptimization), typeof(PdfDocument).GetProperty(nameof(PdfDocument.Optimization))?.PropertyType);
        Assert.Equal(typeof(PdfDocumentProof), typeof(PdfDocument).GetProperty(nameof(PdfDocument.Proof))?.PropertyType);
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
    public void CanonicalCreateDslBuildsReadableContentWithoutDirectDocumentAuthoring() {
        PdfDocument document = PdfDocument.Create(pdf => pdf
            .Content(content => content
                .H1("Canonical authoring")
                .Paragraph(paragraph => paragraph.Text("One public composition model."))));

        string text = document.Read.Text();

        Assert.Contains("Canonical authoring", text, StringComparison.Ordinal);
        Assert.Contains("One public composition model.", text, StringComparison.Ordinal);
    }

    [Fact]
    public void CanonicalPageDslBuildsHeadersFootersAndFlowContent() {
        PdfDocument document = PdfDocument.Create(pdf => pdf.Page(page => page
            .Header(header => header.Text("Quarterly review").AlignLeft())
            .Footer(footer => footer.Text("Confidential").AlignCenter())
            .Content(layout => layout.Item(content => content
                .H1("Service report")
                .Paragraph(paragraph => paragraph.Text("Validated operational data."))))));

        string text = document.Read.Text();

        Assert.Contains("Quarterly review", text, StringComparison.Ordinal);
        Assert.Contains("Service report", text, StringComparison.Ordinal);
        Assert.Contains("Validated operational data.", text, StringComparison.Ordinal);
        Assert.Contains("Confidential", text, StringComparison.Ordinal);
    }

    [Fact]
    public void ComposeBuildersAreClosedAndDoNotExposeInertPaddingApi() {
        Type[] builderTypes = {
            typeof(PdfCompose),
            typeof(PdfPageCompose),
            typeof(PdfContentCompose),
            typeof(PdfItemCompose),
            typeof(PdfElementCompose),
            typeof(PdfColumnCompose),
            typeof(PdfRowCompose),
            typeof(PdfRowColumnCompose),
            typeof(PdfTextStyleCompose),
            typeof(PdfHeaderCompose),
            typeof(PdfFooterCompose),
            typeof(HeaderTextBuilder),
            typeof(FooterTextBuilder)
        };

        Assert.All(builderTypes, type => Assert.True(type.IsSealed, type.FullName));
        Assert.Null(typeof(PdfContentCompose).GetMethod("PaddingBottom", BindingFlags.Public | BindingFlags.Instance));
    }

    [Fact]
    public void FacadeDoesNotDuplicateBuilderAuthoringMethods() {
        Type[] authoringBuilders = {
            typeof(PdfCompose),
            typeof(PdfPageCompose),
            typeof(PdfContentCompose),
            typeof(PdfItemCompose),
            typeof(PdfElementCompose)
        };
        HashSet<string> authoringMethodNames = authoringBuilders
            .SelectMany(type => type.GetMethods(BindingFlags.Public | BindingFlags.Instance | BindingFlags.DeclaredOnly))
            .Where(method => !method.IsSpecialName)
            .Select(method => method.Name)
            .ToHashSet(StringComparer.Ordinal);
        MethodInfo[] facadeMethods = typeof(PdfDocument)
            .GetMethods(BindingFlags.Public | BindingFlags.Instance | BindingFlags.DeclaredOnly)
            .Where(method => !method.IsSpecialName)
            .ToArray();

        Assert.DoesNotContain(facadeMethods, method => authoringMethodNames.Contains(method.Name));
    }

    [Fact]
    public void SpecializedOperationsAreGroupedAndAbsentFromTheRootFacade() {
        Type[] capabilityTypes = {
            typeof(PdfDocumentSecurity),
            typeof(PdfDocumentRedactions),
            typeof(PdfDocumentOptimization),
            typeof(PdfDocumentProof)
        };
        string[] groupedRootMethods = {
            "Encrypt",
            "Decrypt",
            "ValidateSignatures",
            "PlanRedactions",
            "ApplyRedactions",
            "AnalyzeOptimization",
            "Optimize",
            "CompareVisual",
            "AssessRewritePreservation"
        };

        Assert.All(capabilityTypes, type => Assert.True(type.IsSealed, type.FullName));
        Assert.All(groupedRootMethods, methodName => Assert.Null(
            typeof(PdfDocument).GetMethod(methodName, BindingFlags.Public | BindingFlags.Instance)));

        PdfDocument document = PdfDocument.Create(pdf => pdf.Content(content => content.Paragraph(paragraph => paragraph.Text("capabilities"))));
        Assert.NotNull(document.Security);
        Assert.NotNull(document.Redactions);
        Assert.NotNull(document.Optimization);
        Assert.NotNull(document.Proof);
    }

    [Fact]
    public void SpecializedCapabilityObjectsExecuteThroughTheirPublicPaths() {
        byte[] bytes = PdfDocument.Create(pdf => pdf.Content(content => content
                .H1("Capability proof")
                .Paragraph(paragraph => paragraph.Text("Sensitive marker"))))
            .ToBytes();
        PdfDocument document = PdfDocument.Open(bytes);

        PdfSignatureValidationReport signatures = document.Security.ValidateSignatures();
        PdfRedactionPlan redactions = document.Redactions.Search(
            new PdfRedactionSearchOptions().AddLiteral("Sensitive marker"));
        PdfOptimizationReport optimization = document.Optimization.Analyze();
        PdfVisualComparisonReport visual = document.Proof.CompareVisual(bytes);

        Assert.True(signatures.ObjectGraphParsed);
        Assert.NotEmpty(redactions.Areas);
        Assert.True(optimization.StreamCount > 0);
        Assert.True(visual.IsMatch);
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
    public void RuntimeDependenciesStayBounded() {
        Assembly assembly = typeof(PdfDocument).Assembly;
        string[] officeReferences = assembly.GetReferencedAssemblies()
            .Select(reference => reference.Name)
            .Where(name => name != null && name.StartsWith("OfficeIMO.", StringComparison.Ordinal))
            .Cast<string>()
            .OrderBy(name => name, StringComparer.Ordinal)
            .ToArray();
        Assert.Equal(new[] { "OfficeIMO.Core" }, officeReferences);
    }

}
