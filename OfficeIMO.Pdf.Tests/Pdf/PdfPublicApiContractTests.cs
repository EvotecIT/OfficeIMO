using System.Reflection;
using System.Threading.Tasks;
using OfficeIMO.Drawing;
using OfficeIMO.Ocr;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Ocr;
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
        "PdfDocumentReader",
        "PdfFormData",
        "PdfFormFiller",
        "PdfImageExtractor",
        "PdfIncrementalUpdater",
        "PdfInspector",
        "PdfJavaScriptEditor",
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
    public void FacadeExposesOneCreateLoadReadAnalyzeWorkflowWithoutLegacyOpenOrResultLoaders() {
        MethodInfo[] methods = typeof(PdfDocument).GetMethods(BindingFlags.Public | BindingFlags.Static | BindingFlags.Instance);

        Assert.Contains(methods, method =>
            method.Name == nameof(PdfDocument.Create) &&
            method.IsStatic);
        Assert.Contains(methods, method =>
            method.Name == nameof(PdfDocument.Create) &&
            method.IsStatic &&
            method.GetParameters().FirstOrDefault()?.ParameterType == typeof(Action<PdfDocumentBuilder>));
        Assert.All(
            methods.Where(method => method.Name == nameof(PdfDocument.Create) && method.IsStatic),
            method => Assert.Equal(typeof(Action<PdfDocumentBuilder>), method.GetParameters()[0].ParameterType));
        Assert.Contains(methods, method =>
            method.Name == nameof(PdfDocument.Load) &&
            method.IsStatic &&
            method.GetParameters().FirstOrDefault()?.ParameterType == typeof(byte[]));
        Assert.Contains(methods, method =>
            method.Name == nameof(PdfDocument.Load) &&
            method.IsStatic &&
            method.GetParameters().FirstOrDefault()?.ParameterType == typeof(string));
        Assert.Contains(methods, method =>
            method.Name == nameof(PdfDocument.Load) &&
            method.IsStatic &&
            method.GetParameters().FirstOrDefault()?.ParameterType == typeof(Stream));
        Assert.Contains(methods, method =>
            method.Name == nameof(PdfDocument.LoadAsync) &&
            method.IsStatic);
        Assert.Contains(methods, method =>
            method.Name == nameof(PdfDocument.Analyze) &&
            !method.IsStatic &&
            method.ReturnType == typeof(PdfAnalysisReport));
        Assert.DoesNotContain(methods, method => method.Name == "Open");
        Assert.Contains(methods, method =>
            method.Name == nameof(PdfDocument.Read) &&
            !method.IsStatic &&
            method.ReturnType == typeof(PdfDocumentReadResult));

        Assert.Null(typeof(PdfDocument).GetProperty("Read", BindingFlags.Public | BindingFlags.Instance));
        Assert.Null(typeof(PdfDocument).GetProperty("Reader", BindingFlags.Public | BindingFlags.Instance));
        Assert.Null(typeof(PdfDocument).GetProperty("Ocr", BindingFlags.Public | BindingFlags.Instance));
        Assert.Equal(typeof(PdfDocumentRenderer), typeof(PdfDocument).GetProperty(nameof(PdfDocument.Render))?.PropertyType);
        Assert.Equal(typeof(PdfDocumentResources), typeof(PdfDocument).GetProperty(nameof(PdfDocument.Resources))?.PropertyType);
        Assert.Equal(typeof(PdfDocumentImageEditor), typeof(PdfDocument).GetProperty(nameof(PdfDocument.Images))?.PropertyType);
        Assert.Equal(typeof(PdfDocumentAttachments), typeof(PdfDocument).GetProperty(nameof(PdfDocument.Attachments))?.PropertyType);
        Assert.Contains(
            typeof(PdfOcrExtensions).GetMethods(BindingFlags.Public | BindingFlags.Static),
            method => method.Name == nameof(PdfOcrExtensions.ReadWithOcrAsync) &&
                method.GetParameters().Length >= 2 &&
                method.GetParameters()[0].ParameterType == typeof(PdfDocument) &&
                method.GetParameters()[1].ParameterType == typeof(IOcrEngine));
        Assert.DoesNotContain(
            typeof(PdfDocumentReadResult).GetMethods(BindingFlags.Public | BindingFlags.Static),
            method => method.Name is "Load" or "LoadPageRanges" or "From" or "FromPageRanges");
        Assert.Null(typeof(PdfDocument).Assembly.GetType("OfficeIMO.Pdf.IPdfOcrProvider"));
        Assert.Null(typeof(PdfDocument).Assembly.GetType("OfficeIMO.Pdf.PdfOcrRequest"));
        Assert.Null(typeof(PdfDocument).Assembly.GetType("OfficeIMO.Pdf.PdfOcrResponse"));
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
            3,
            methods.Count(method =>
                method.Name == nameof(PdfDocument.SaveAsync) &&
                method.ReturnType == typeof(Task<PdfSaveResult>)));
        Assert.Contains(methods, method =>
            method.Name == nameof(PdfDocument.SaveAsync) &&
            method.GetParameters().Select(static parameter => parameter.ParameterType).SequenceEqual(new[] {
                typeof(string),
                typeof(OfficeConversionFileConflictPolicy),
                typeof(System.Threading.CancellationToken)
            }));
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

        string text = document.Read().Text;

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

        string text = string.Join(" ", document.Read().TextBlocks.Select(static block => block.Text));

        Assert.Contains("Quarterly review", text, StringComparison.Ordinal);
        Assert.Contains("Service report", text, StringComparison.Ordinal);
        Assert.Contains("Validated operational data.", text, StringComparison.Ordinal);
        Assert.Contains("Confidential", text, StringComparison.Ordinal);
    }

    [Fact]
    public void ComposeBuildersAreClosedAndDoNotExposeInertPaddingApi() {
        Type[] builderTypes = {
            typeof(PdfDocumentBuilder),
            typeof(PdfPageBuilder),
            typeof(PdfContentBuilder),
            typeof(PdfElementBuilder),
            typeof(PdfRowBuilder),
            typeof(PdfTextStyleBuilder),
            typeof(PdfHeaderBuilder),
            typeof(PdfFooterBuilder),
            typeof(HeaderTextBuilder),
            typeof(FooterTextBuilder)
        };

        Assert.All(builderTypes, type => Assert.True(type.IsSealed, type.FullName));
        Assert.Null(typeof(PdfContentBuilder).GetMethod("PaddingBottom", BindingFlags.Public | BindingFlags.Instance));
        Assert.Null(typeof(PdfContentBuilder).GetMethod("Container", BindingFlags.Public | BindingFlags.Instance));
        Assert.Null(typeof(PdfDocumentBuilder).GetMethod("Defaults", BindingFlags.Public | BindingFlags.Instance));
    }

    [Fact]
    public void AuthoringUsesOneContentReceiverAtEveryNestingBoundary() {
        Type contentCallback = typeof(Action<PdfContentBuilder>);

        Assert.Equal(contentCallback, typeof(PdfDocumentBuilder).GetMethod(nameof(PdfDocumentBuilder.Content))!.GetParameters()[0].ParameterType);
        Assert.Equal(contentCallback, typeof(PdfPageBuilder).GetMethod(nameof(PdfPageBuilder.Content))!.GetParameters()[0].ParameterType);
        Assert.Equal(contentCallback, typeof(PdfContentBuilder).GetMethod(nameof(PdfContentBuilder.Column))!.GetParameters()[0].ParameterType);
        Assert.Equal(typeof(Action<PdfElementBuilder>), typeof(PdfContentBuilder).GetMethod(nameof(PdfContentBuilder.Element))!.GetParameters()[0].ParameterType);
        Assert.Equal(contentCallback, typeof(PdfElementBuilder).GetMethod(nameof(PdfElementBuilder.Content))!.GetParameters()[0].ParameterType);
        Assert.Equal(contentCallback, typeof(PdfRowBuilder).GetMethod(nameof(PdfRowBuilder.Column))!.GetParameters()[1].ParameterType);
    }

    [Fact]
    public void AuthoringExposesSharedTypographyProfilesAtDocumentAndPageScope() {
        MethodInfo documentTypography = typeof(PdfDocumentBuilder).GetMethod(nameof(PdfDocumentBuilder.Typography))!;
        MethodInfo pageTypography = typeof(PdfPageBuilder).GetMethod(nameof(PdfPageBuilder.Typography))!;

        Assert.Equal(typeof(OfficeRenderingProfile), documentTypography.GetParameters()[0].ParameterType);
        Assert.Equal(typeof(OfficeRenderingProfileApplyMode), documentTypography.GetParameters()[1].ParameterType);
        Assert.Equal(typeof(OfficeRenderingProfile), pageTypography.GetParameters()[0].ParameterType);
        Assert.Equal(typeof(OfficeRenderingProfileApplyMode), pageTypography.GetParameters()[1].ParameterType);

        PdfDocument document = PdfDocument.Create(pdf => pdf
            .Typography(OfficeRenderingProfile.Managed)
            .Content(content => content.Text("Profiled text")));

        Assert.Same(OfficeRenderingProfile.Managed.TextShapingProvider, document.Options.TextShapingProvider);
    }

    [Fact]
    public void DocumentTypographyAppliesToPagesCreatedEarlierInTheSameComposition() {
        var profile = new OfficeRenderingProfile(
            "late-document-typography",
            textShapingProvider: OfficeManagedTextShapingProvider.Instance,
            textShapingLanguage: "pl-PL");

        PdfDocument document = PdfDocument.Create(pdf => pdf
            .Page(page => page.Content(content => content.Text("Earlier page")))
            .Typography(profile));
        PageBlock page = Assert.IsType<PageBlock>(Assert.Single(document.Blocks));

        Assert.Same(profile.TextShapingProvider, page.Options.TextShapingProvider);
        Assert.Equal("pl-PL", page.Options.Language);
    }

    [Fact]
    public void IncrementalDocumentTypographyUpdatesExistingPageSnapshots() {
        var profile = new OfficeRenderingProfile(
            "incremental-document-typography",
            textShapingProvider: OfficeManagedTextShapingProvider.Instance,
            textShapingLanguage: "de-DE");
        PdfDocument document = PdfDocument.Create(pdf => pdf
            .Page(page => page.Content(content => content.Text("Existing page"))));

        document.Compose(pdf => pdf.Typography(profile));
        PageBlock page = Assert.IsType<PageBlock>(Assert.Single(document.Blocks));

        Assert.Same(profile.TextShapingProvider, page.Options.TextShapingProvider);
        Assert.Equal("de-DE", page.Options.Language);
    }

    [Fact]
    public void SupersededComposeReceiverTypesAreNotPublic() {
        Assembly assembly = typeof(PdfDocument).Assembly;

        Assert.Null(assembly.GetType("OfficeIMO.Pdf.PdfCompose"));
        Assert.Null(assembly.GetType("OfficeIMO.Pdf.PdfPageCompose"));
        Assert.Null(assembly.GetType("OfficeIMO.Pdf.PdfItemCompose"));
        Assert.Null(assembly.GetType("OfficeIMO.Pdf.PdfContentCompose"));
        Assert.Null(assembly.GetType("OfficeIMO.Pdf.PdfColumnCompose"));
        Assert.Null(assembly.GetType("OfficeIMO.Pdf.PdfElementCompose"));
        Assert.Null(assembly.GetType("OfficeIMO.Pdf.PdfRowColumnCompose"));
        Assert.Null(assembly.GetType("OfficeIMO.Pdf.PdfRowCompose"));
        Assert.Null(assembly.GetType("OfficeIMO.Pdf.PdfHeaderCompose"));
        Assert.Null(assembly.GetType("OfficeIMO.Pdf.PdfFooterCompose"));
        Assert.Null(assembly.GetType("OfficeIMO.Pdf.PdfTextStyleCompose"));
    }

    [Fact]
    public void FacadeDoesNotDuplicateBuilderAuthoringMethods() {
        Type[] authoringBuilders = {
            typeof(PdfDocumentBuilder),
            typeof(PdfPageBuilder),
            typeof(PdfContentBuilder)
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
        PdfDocument document = PdfDocument.Load(bytes);

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
    public void RuntimeDependencyOwnershipStaysBounded() {
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
