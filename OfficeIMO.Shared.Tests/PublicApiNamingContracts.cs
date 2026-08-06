using System.Reflection;
using System.Threading;
using OfficeIMO.CSV;
using OfficeIMO.Data;
using OfficeIMO.Drawing;
using OfficeIMO.Email;
using OfficeIMO.Excel;
using OfficeIMO.Html;
using OfficeIMO.Pdf;
using OfficeIMO.Reader;
using OfficeIMO.Visio;
using OfficeIMO.Word;
using OfficeIMO.Word.Fluent;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed class PublicApiNamingContracts {
    [Fact]
    public void CoreNeutralContractsUsePurposeBasedNamespaces() {
        Type[] rootContracts = {
            typeof(IOfficeConversionReport),
            typeof(OfficeFormatDescriptor),
            typeof(OfficeCompatibilityReport),
            typeof(OfficeCompatibilityFinding),
            typeof(OfficeConversionCapability),
            typeof(OfficeConversionCapabilityCatalog),
            typeof(OfficeCapability),
            typeof(OfficeCapabilityCatalog)
        };
        Type[] dataContracts = {
            typeof(ObjectFlattener),
            typeof(ObjectFlattenerOptions),
            typeof(CollectionColumnMapping),
            typeof(HeaderCase),
            typeof(NullPolicy),
            typeof(CollectionMode)
        };

        Assert.All(rootContracts, static type => Assert.Equal("OfficeIMO", type.Namespace));
        Assert.All(dataContracts, static type => Assert.Equal("OfficeIMO.Data", type.Namespace));
    }

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

    [Fact]
    public void WordTemplateConversionUsesCanonicalExtensionCasing() {
        string[] methodNames = typeof(WordHelpers)
            .GetMethods(BindingFlags.Public | BindingFlags.Static | BindingFlags.DeclaredOnly)
            .Select(method => method.Name)
            .ToArray();

        Assert.Contains("ConvertDotxToDocx", methodNames);
        Assert.DoesNotContain("ConvertDotXtoDocX", methodNames);
    }

    [Theory]
    [InlineData(typeof(EmailDocumentWriter))]
    [InlineData(typeof(EmailMailboxWriter))]
    public void EmailWriterMemoryOutputUsesToBytesVocabulary(Type writerType) {
        string[] methodNames = writerType
            .GetMethods(BindingFlags.Public | BindingFlags.Instance | BindingFlags.DeclaredOnly)
            .Select(static method => method.Name)
            .ToArray();

        Assert.Contains("ToBytes", methodNames);
        Assert.DoesNotContain("WriteToBytes", methodNames);
    }

    [Fact]
    public void ExcelDocumentRemoteLoadsAreAsyncOnly() {
        MethodInfo[] documentMethods = typeof(ExcelDocument).GetMethods(
            BindingFlags.Public | BindingFlags.Static | BindingFlags.DeclaredOnly);

        Assert.Contains(documentMethods, static method =>
            method.Name == "LoadAsync" && method.GetParameters().FirstOrDefault()?.ParameterType == typeof(Uri));
        Assert.DoesNotContain(documentMethods, static method =>
            method.Name == "Load" && method.GetParameters().FirstOrDefault()?.ParameterType == typeof(Uri));
    }

    [Fact]
    public void ExcelPublicApiExposesOnlyCanonicalReadSurfaces() {
        string[] exportedTypeNames = typeof(ExcelDocument).Assembly
            .GetExportedTypes()
            .Select(static type => type.Name)
            .ToArray();
        MethodInfo[] documentMethods = typeof(ExcelDocument).GetMethods(
            BindingFlags.Public | BindingFlags.Instance | BindingFlags.Static | BindingFlags.DeclaredOnly);
        MethodInfo[] sheetMethods = typeof(ExcelSheet).GetMethods(
            BindingFlags.Public | BindingFlags.Instance | BindingFlags.DeclaredOnly);
        PropertyInfo[] sheetProperties = typeof(ExcelSheet).GetProperties(
            BindingFlags.Public | BindingFlags.Instance | BindingFlags.DeclaredOnly);

        Assert.DoesNotContain("ExcelRead", exportedTypeNames);
        Assert.DoesNotContain("ExcelDocumentReader", exportedTypeNames);
        Assert.DoesNotContain("ExcelSheetReader", exportedTypeNames);
        Assert.DoesNotContain("RowEdit", exportedTypeNames);
        Assert.DoesNotContain("CellEdit", exportedTypeNames);
        Assert.DoesNotContain(documentMethods, static method => method.Name == "Read");
        Assert.DoesNotContain(sheetMethods, static method => method.Name == "Rows");
        Assert.Contains(documentMethods, static method => method.Name == "OpenDataReader");
        Assert.Contains(sheetMethods, static method => method.Name == "CreateDataReader");
        Assert.Contains(sheetMethods, static method => method.Name == "RowsAs" && method.IsGenericMethodDefinition);
        Assert.Contains(sheetMethods, static method =>
            method.Name == "RowsAs" && method.GetParameters().Any(parameter =>
                parameter.ParameterType.IsGenericType &&
                parameter.ParameterType.GetGenericTypeDefinition() == typeof(Action<>)));
        Assert.DoesNotContain(sheetMethods, static method => method.Name == "RowsAsStream");
        Assert.DoesNotContain(sheetMethods, static method => method.Name == "GetUsedRangeA1");
        Assert.Contains(sheetProperties, static property => property.Name == "UsedRangeA1");
        Assert.All(
            sheetMethods.Where(static method =>
                method.Name is "RowsAs" or "EnumerateCells" or "EnumerateRange" &&
                method.GetParameters().Any(static parameter => parameter.ParameterType == typeof(CancellationToken))),
            static method => Assert.Contains(
                method.GetParameters(),
                static parameter => parameter.ParameterType == typeof(CancellationToken) &&
                                    parameter.Name == "cancellationToken"));
    }

    [Fact]
    public void CsvPublicApiExposesOnlyCanonicalMappingAndWriterSurfaces() {
        Type[] exportedTypes = typeof(CsvDocument).Assembly.GetExportedTypes();
        string[] exportedTypeNames = exportedTypes.Select(static type => type.Name).ToArray();
        MethodInfo[] documentMethods = typeof(CsvDocument).GetMethods(
            BindingFlags.Public | BindingFlags.Instance | BindingFlags.Static | BindingFlags.DeclaredOnly);
        string[] writerMethodNames = typeof(CsvRowWriter)
            .GetMethods(BindingFlags.Public | BindingFlags.Instance | BindingFlags.DeclaredOnly)
            .Select(static method => method.Name)
            .ToArray();

        Assert.Contains("CsvRowWriter", exportedTypeNames);
        Assert.Contains(typeof(RowMapper<>), typeof(RowMapper<>).Assembly.GetExportedTypes());
        Assert.DoesNotContain("CsvObjectWriter", exportedTypeNames);
        Assert.DoesNotContain("CsvMapper`1", exportedTypeNames);
        Assert.DoesNotContain("CsvFile", exportedTypeNames);
        Assert.DoesNotContain(documentMethods, static method => method.Name == "Materialize");
        Assert.Contains("WriteRow", writerMethodNames);
        Assert.Contains("WriteTextRow", writerMethodNames);
        Assert.DoesNotContain(writerMethodNames, static name => name.Contains("Trusted", StringComparison.Ordinal));
    }

    [Fact]
    public void VisioAsyncLoadsUseOptionsThenCancellationToken() {
        MethodInfo[] loadMethods = typeof(VisioDocument).GetMethods(
            BindingFlags.Public | BindingFlags.Static | BindingFlags.DeclaredOnly)
            .Where(static method => method.Name == "LoadAsync")
            .ToArray();

        Assert.Contains(loadMethods, static method =>
            method.GetParameters().Select(static parameter => parameter.ParameterType).SequenceEqual([
                typeof(string),
                typeof(VisioLoadOptions),
                typeof(CancellationToken)
            ]));
        Assert.Contains(loadMethods, static method =>
            method.GetParameters().Select(static parameter => parameter.ParameterType).SequenceEqual([
                typeof(Stream),
                typeof(VisioLoadOptions),
                typeof(CancellationToken)
            ]));
        Assert.DoesNotContain(loadMethods, static method => {
            Type[] parameterTypes = method.GetParameters()
                .Select(static parameter => parameter.ParameterType)
                .ToArray();
            return parameterTypes.Length >= 2
                   && parameterTypes[1] == typeof(CancellationToken);
        });
    }

    [Fact]
    public void WordRemoteImageApisAreAsyncOnly() {
        string[] documentMethodNames = typeof(WordDocument)
            .GetMethods(BindingFlags.Public | BindingFlags.Instance | BindingFlags.DeclaredOnly)
            .Select(static method => method.Name)
            .ToArray();
        string[] builderMethodNames = typeof(ImageBuilder)
            .GetMethods(BindingFlags.Public | BindingFlags.Instance | BindingFlags.DeclaredOnly)
            .Select(static method => method.Name)
            .ToArray();

        Assert.Contains("AddImageFromUrlAsync", documentMethodNames);
        Assert.DoesNotContain("AddImageFromUrl", documentMethodNames);
        Assert.Contains("AddFromUrlAsync", builderMethodNames);
        Assert.DoesNotContain("AddFromUrl", builderMethodNames);
    }

    [Fact]
    public void ExcelRemoteImageApisAreAsyncOnly() {
        MethodInfo[] sheetMethods = typeof(ExcelSheet).GetMethods(
            BindingFlags.Public | BindingFlags.Instance | BindingFlags.DeclaredOnly);
        MethodInfo[] templateMethods = typeof(ExcelTemplateImage).GetMethods(
            BindingFlags.Public | BindingFlags.Static | BindingFlags.DeclaredOnly);
        MethodInfo[] composerMethods = typeof(OfficeIMO.Excel.Fluent.SheetComposer).GetMethods(
            BindingFlags.Public | BindingFlags.Instance | BindingFlags.DeclaredOnly);

        Assert.Contains(sheetMethods, static method => method.Name == "AddImageFromUrlAsync");
        Assert.Contains(sheetMethods, static method => method.Name == "SetHeaderImageFromUrlAsync");
        Assert.Contains(sheetMethods, static method => method.Name == "SetFooterImageFromUrlAsync");
        Assert.DoesNotContain(sheetMethods, static method =>
            (method.Name.Contains("Image", StringComparison.Ordinal) || method.Name.Contains("Logo", StringComparison.Ordinal))
            && method.Name.Contains("Url", StringComparison.Ordinal)
            && !method.Name.EndsWith("Async", StringComparison.Ordinal));

        Assert.Contains(templateMethods, static method => method.Name == "FromUrlAsync");
        Assert.DoesNotContain(templateMethods, static method => method.Name == "FromUrl");

        Assert.Contains(composerMethods, static method => method.Name == "ImageFromUrlAtAsync");
        Assert.Contains(composerMethods, static method => method.Name == "HeaderLogoFromUrlAsync");
        Assert.DoesNotContain(composerMethods, static method =>
            (method.Name.Contains("Image", StringComparison.Ordinal) || method.Name.Contains("Logo", StringComparison.Ordinal))
            && method.Name.Contains("Url", StringComparison.Ordinal)
            && !method.Name.EndsWith("Async", StringComparison.Ordinal));
    }

    [Fact]
    public void CoreOwnsTheSharedRemoteImageLoader() {
        MethodInfo[] methods = typeof(OfficeRemoteImageLoader).GetMethods(
            BindingFlags.Public | BindingFlags.Static | BindingFlags.DeclaredOnly);

        Assert.NotEmpty(methods);
        Assert.All(methods, static method => Assert.Equal("LoadAsync", method.Name));
    }

    [Fact]
    public void PdfConversionResultUsesValueReportAndLossContract() {
        Type resultType = typeof(PdfDocumentConversionResult);

        Assert.NotNull(resultType.GetProperty("Value"));
        Assert.NotNull(resultType.GetProperty("Report"));
        Assert.NotNull(resultType.GetProperty("HasLoss"));
        Assert.NotNull(resultType.GetMethod("RequireValue", Type.EmptyTypes));
        Assert.NotNull(resultType.GetMethod("RequireNoLoss", Type.EmptyTypes));
    }

    [Fact]
    public void PersistenceOptionsDoNotMixSavingWithApplicationLaunching() {
        Assembly[] assemblies = {
            typeof(WordDocument).Assembly,
            typeof(ExcelDocument).Assembly,
            typeof(OfficeIMO.PowerPoint.PowerPointPresentation).Assembly
        };

        PropertyInfo[] launchProperties = assemblies
            .SelectMany(static assembly => assembly.GetExportedTypes())
            .SelectMany(static type => type.GetProperties(BindingFlags.Public | BindingFlags.Instance | BindingFlags.Static))
            .Where(static property => property.Name == "OpenAfterSave")
            .ToArray();

        Assert.Empty(launchProperties);
    }

    [Fact]
    public void SharedOfficeContractsAreOwnedByCore() {
        Assembly coreAssembly = typeof(OfficeConversionLossPolicy).Assembly;
        Type[] sharedTypes = {
            typeof(OfficeConversionLossPolicy),
            typeof(OfficeConversionFileConflictPolicy),
            typeof(OfficeConversionDiagnosticCategory),
            typeof(OfficeConversionDiagnosticSeverity),
            typeof(OfficeConversionFailureReason),
            typeof(OfficeConversionLossKind),
            typeof(OfficeFeatureSupportLevel),
            typeof(OfficeOpenXmlCompatibilityLevel),
            typeof(OfficeOpenXmlMarkupCompatibilityMode),
            typeof(OfficeOpenXmlFileFormatVersion),
            typeof(OfficeOpenXmlLoadSettings),
            typeof(OfficeOpenXmlValidationErrorType),
            typeof(OfficeOpenXmlValidationError),
            typeof(OfficeImageFormat),
            typeof(OfficePageOrientation),
            typeof(OfficeSignatureMutationPolicy),
            typeof(OfficeChartDisplayUnit),
            typeof(OfficeChartDataLabelPosition),
            typeof(OfficeChartLegendPosition),
            typeof(OfficeChartMarkerShape),
            typeof(OfficeLineMarkerKind),
            typeof(OfficePresetShapeType)
        };

        Assert.All(sharedTypes, type => Assert.Same(coreAssembly, type.Assembly));

        Assert.DoesNotContain(typeof(WordDocument).Assembly.GetExportedTypes(),
            static type => type.Name == "WordPageOrientation");
        Assert.DoesNotContain(typeof(ExcelDocument).Assembly.GetExportedTypes(),
            static type => type.Name is "ExcelPageOrientation" or "ExcelSignatureMutationPolicy");
        Assert.DoesNotContain(typeof(PdfDocument).Assembly.GetExportedTypes(),
            static type => type.Name == "PdfPageOrientation");
        Assert.DoesNotContain(typeof(HtmlConversionDocument).Assembly.GetExportedTypes(),
            static type => type.Name == "HtmlConversionLossKind");
        Assert.DoesNotContain(typeof(OfficeIMO.Word.Markdown.WordMarkdownConversionReport).Assembly.GetExportedTypes(),
            static type => type.Name == "WordMarkdownConversionLossKind");
        Assert.DoesNotContain(typeof(OfficeIMO.PowerPoint.PowerPointPresentation).Assembly.GetExportedTypes(),
            static type => type.Name == "PowerPointSignatureMutationPolicy");
    }

    [Fact]
    public void GoogleWorkspaceOwnsCrossEditorImportAndDiffContracts() {
        Assembly owner = typeof(OfficeIMO.GoogleWorkspace.GoogleWorkspaceImportMode).Assembly;
        Assert.Same(owner, typeof(OfficeIMO.GoogleWorkspace.GoogleWorkspaceDiffKind).Assembly);

        Assembly[] adapters = {
            typeof(OfficeIMO.Word.GoogleDocs.GoogleDocsImportOptions).Assembly,
            typeof(OfficeIMO.Excel.GoogleSheets.GoogleSheetsImportOptions).Assembly,
            typeof(OfficeIMO.PowerPoint.GoogleSlides.GoogleSlidesImportOptions).Assembly
        };
        string[] removedNames = {
            "GoogleDocsImportMode", "GoogleSheetsImportMode", "GoogleSlidesImportMode",
            "GoogleDocsDiffKind", "GoogleSheetsDiffKind", "GoogleSlidesDiffKind"
        };

        Assert.All(adapters, assembly => Assert.DoesNotContain(
            assembly.GetExportedTypes(), type => removedNames.Contains(type.Name, StringComparer.Ordinal)));
    }

    [Fact]
    public void WordAlignmentAndTableLayoutUseOneContractPerPurpose() {
        Assembly wordAssembly = typeof(WordDocument).Assembly;
        Type[] exportedTypes = wordAssembly.GetExportedTypes();

        Assert.Contains(typeof(WordParagraphAlignment), exportedTypes);
        Assert.Contains(typeof(WordTableAlignment), exportedTypes);
        Assert.Contains(typeof(WordTextBoxHorizontalAlignment), exportedTypes);
        Assert.Contains(typeof(WordTableLayoutMode), exportedTypes);
        Assert.DoesNotContain(exportedTypes, type => type.Name is
            "HorizontalAlignment" or
            "VerticalAlignment" or
            "WordHorizontalAlignmentValues" or
            "WordTableLayoutType");

        PropertyInfo layoutMode = Assert.Single(typeof(WordTable).GetProperties(
            BindingFlags.Public | BindingFlags.Instance | BindingFlags.DeclaredOnly),
            property => property.Name == "LayoutMode");
        Assert.Equal(typeof(WordTableLayoutMode), layoutMode.PropertyType);
        Assert.Null(typeof(WordTable).GetProperty("LayoutType"));
        Assert.DoesNotContain(typeof(WordTable).GetMethods(BindingFlags.Public | BindingFlags.Instance),
            method => method.Name is "SetTableLayout" or "GetCurrentLayoutMode" or "GetCurrentLayoutType");
    }

    [Fact]
    public void OfficePackagesDoNotExportCollidingSimpleTypeNames() {
        Assembly[] assemblies = Directory
            .EnumerateFiles(AppContext.BaseDirectory, "OfficeIMO*.dll")
            .Where(static path => !Path.GetFileNameWithoutExtension(path).EndsWith(".Tests", StringComparison.Ordinal))
            .Select(static path => Assembly.Load(new AssemblyName(AssemblyName.GetAssemblyName(path).Name!)))
            .Distinct()
            .ToArray();
        IGrouping<string, Type>[] collisions = assemblies
            .SelectMany(assembly => assembly.GetExportedTypes())
            .GroupBy(type => type.Name, StringComparer.Ordinal)
            .Where(group => group.Select(type => type.Assembly).Distinct().Count() > 1)
            .ToArray();

        Assert.Empty(collisions);
    }

    [Fact]
    public void FormatOwnedBuilderAndTextRunTypesAreUnambiguous() {
        Type[] canonicalTypes = {
            typeof(WordParagraphBuilder),
            typeof(OfficeIMO.Markdown.MarkdownParagraphBuilder),
            typeof(PdfTextRun),
            typeof(OfficeIMO.Markdown.MarkdownTextRun)
        };

        Assert.Equal(4, canonicalTypes.Select(static type => type.Name).Distinct(StringComparer.Ordinal).Count());
        Assert.DoesNotContain(typeof(WordDocument).Assembly.GetExportedTypes(), static type => type.Name == "ParagraphBuilder");
        Assert.DoesNotContain(typeof(OfficeIMO.Markdown.MarkdownDoc).Assembly.GetExportedTypes(),
            static type => type.Name is "ParagraphBuilder" or "TextRun");
        Assert.DoesNotContain(typeof(PdfDocument).Assembly.GetExportedTypes(), static type => type.Name == "TextRun");
    }

    [Fact]
    public void FormatSpecificPublicTypesUseFormatSpecificNames() {
        Type[] canonicalTypes = {
            typeof(WordApplicationProperties),
            typeof(WordBuiltinDocumentProperties),
            typeof(WordCapsStyle),
            typeof(WordCompatibilityMode),
            typeof(WordCoverPageTemplate),
            typeof(WordDocumentCleanupOptions),
            typeof(WordImageFillMode),
            typeof(WordShapeType),
            typeof(WordSmartArtType),
            typeof(WordTableOfContentsStyle),
            typeof(WordHyperlinkTargetFrame),
            typeof(WordTextMatchType),
            typeof(WordImageTextWrapping),
            typeof(ExcelApplicationProperties),
            typeof(ExcelBuiltinDocumentProperties),
            typeof(ExcelExecutionMode),
            typeof(ExcelExecutionPolicy),
            typeof(ExcelHeaderFooterPosition),
            typeof(ExcelDefinedNameValidationMode),
            typeof(ExcelTableStyle),
            typeof(OfficeIMO.PowerPoint.PowerPointSlideTransition),
            typeof(OfficeIMO.PowerPoint.PowerPointSlideTransitionSpeed),
            typeof(OfficeIMO.PowerPoint.PowerPointTableCellBorders)
        };
        string[] removedNames = {
            "ApplicationProperties", "BuiltinDocumentProperties", "CapsStyle", "CompatibilityMode",
            "CoverPageTemplate", "CustomImagePartType", "WordImagePartType", "DocumentCleanupOptions", "ImageFillMode",
            "ShapeType", "SmartArtType", "TableOfContentStyle", "TargetFrame", "TextMatchType",
            "WrapTextImage", "ExecutionMode", "ExecutionPolicy", "HeaderFooterPosition",
            "NameValidationMode", "TableStyle", "ImagePartType", "PowerPointImagePartType", "SlideTransition",
            "SlideTransitionSpeed", "TableCellBorders"
        };
        Assembly[] assemblies = canonicalTypes.Select(type => type.Assembly).Distinct().ToArray();

        Assert.All(canonicalTypes, type => Assert.StartsWith(
            type.Assembly == typeof(WordDocument).Assembly ? "Word" :
            type.Assembly == typeof(ExcelDocument).Assembly ? "Excel" : "PowerPoint",
            type.Name));
        Assert.All(assemblies, assembly => Assert.DoesNotContain(
            assembly.GetExportedTypes(), type => removedNames.Contains(type.Name, StringComparer.Ordinal)));
    }

    [Fact]
    public void ExcelBatchingDoesNotExposeLockBypassScopes() {
        Assert.NotNull(typeof(ExcelSheet).GetMethod("Batch", [typeof(Action<ExcelSheet>)]));
        Assert.Null(typeof(ExcelSheet).GetMethod("BeginNoLock", BindingFlags.Public | BindingFlags.Instance));
        Assert.DoesNotContain(typeof(ExcelSheet).GetNestedTypes(BindingFlags.Public),
            type => type.Name.Contains("NoLock", StringComparison.Ordinal));
    }

    [Fact]
    public void ExcelMetadataOperationsReturnOwnedPartDescriptors() {
        string[] methodNames = {
            "AddWorkbookConnectionMetadata",
            "AddWorksheetQueryTableMetadata",
            "AddWorkbookSlicerCache",
            "AddWorkbookTimelineCache",
            "AddWorkbookMetadataPart",
            "AddWorksheetMetadataPart",
            "AddPivotSlicerCache",
            "AddPivotTimelineCache"
        };
        MethodInfo[] methods = typeof(ExcelDocument).GetMethods(
            BindingFlags.Public | BindingFlags.Instance | BindingFlags.DeclaredOnly);

        Assert.All(methodNames, name => Assert.All(
            methods.Where(method => method.Name == name),
            method => Assert.Equal(typeof(ExcelPackagePartInfo), method.ReturnType)));
        Assert.All(methodNames, name => Assert.Contains(methods, method => method.Name == name));
    }

    [Fact]
    public void OfficeImageApisUseCanonicalVerbAndCardinalityVocabulary() {
        Assembly[] assemblies = {
            typeof(WordDocument).Assembly,
            typeof(ExcelDocument).Assembly,
            typeof(OfficeIMO.PowerPoint.PowerPointPresentation).Assembly
        };
        MethodInfo[] methods = assemblies
            .SelectMany(static assembly => assembly.GetExportedTypes())
            .SelectMany(static type => type.GetMethods(BindingFlags.Public | BindingFlags.Instance | BindingFlags.Static))
            .ToArray();

        Assert.DoesNotContain(methods, static method =>
            method.Name is "SaveImage" or "SaveAsImage");

        Assert.NotNull(typeof(WordDocument).GetMethod("ToImage", Type.EmptyTypes));
        Assert.NotNull(typeof(WordDocument).GetMethods().SingleOrDefault(static method =>
            method.Name == "ExportImage" && method.GetParameters().Length > 0));
        Assert.Contains(typeof(WordDocument).GetMethods(), static method => method.Name == "SaveAsImages");

        Assert.NotNull(typeof(ExcelSheet).GetMethod("ToImage", Type.EmptyTypes));
        Assert.Contains(typeof(ExcelSheet).GetMethods(), static method => method.Name == "ExportImage");
        Assert.Contains(typeof(ExcelDocument).GetMethods(), static method => method.Name == "SaveAsImages");

        Assert.NotNull(typeof(OfficeIMO.PowerPoint.PowerPointSlide).GetMethod("ToImage", Type.EmptyTypes));
        Assert.Contains(typeof(OfficeIMO.PowerPoint.PowerPointSlide).GetMethods(), static method => method.Name == "ExportImage");
        Assert.Contains(typeof(OfficeIMO.PowerPoint.PowerPointPresentation).GetMethods(), static method => method.Name == "SaveAsImages");
    }

    [Fact]
    public void ConversionTargetNamesUseDotNetAcronymCasing() {
        Assembly[] assemblies = {
            typeof(WordDocument).Assembly,
            typeof(ExcelDocument).Assembly,
            typeof(OfficeIMO.PowerPoint.PowerPointPresentation).Assembly,
            typeof(HtmlConversionDocument).Assembly,
            typeof(PdfDocument).Assembly
        };
        string[] methodNames = assemblies
            .SelectMany(static assembly => assembly.GetExportedTypes())
            .SelectMany(static type => type.GetMethods(BindingFlags.Public | BindingFlags.Instance | BindingFlags.Static))
            .Select(static method => method.Name)
            .Distinct(StringComparer.Ordinal)
            .ToArray();

        Assert.DoesNotContain(methodNames, static name =>
            name.Contains("PDF", StringComparison.Ordinal) ||
            name.Contains("HTML", StringComparison.Ordinal) ||
            name.Contains("RTF", StringComparison.Ordinal) ||
            name.Contains("ODT", StringComparison.Ordinal) ||
            name.Contains("ODS", StringComparison.Ordinal) ||
            name.Contains("ODP", StringComparison.Ordinal));
    }

    [Fact]
    public void CanonicalSchemaAndAstNamesDoNotExposeAliases() {
        Assert.Null(typeof(OfficeDocumentReadResultSchema).GetField("Version", BindingFlags.Public | BindingFlags.Static));
        Assert.NotNull(typeof(OfficeIMO.Markdown.FootnoteDefinitionBlock).GetProperty("ChildBlocks"));
        Assert.Null(typeof(OfficeIMO.Markdown.FootnoteDefinitionBlock).GetProperty("Blocks"));
        Assert.Null(typeof(OfficeIMO.Markdown.DefinitionListBlock).GetProperty("Items"));
        Assert.Null(typeof(OfficeIMO.Markdown.OrderedListBlock).GetProperty("ListItems"));
        Assert.Null(typeof(OfficeIMO.Markdown.UnorderedListBlock).GetProperty("ListItems"));
        Assert.NotNull(typeof(OfficeIMO.Markdown.DefinitionListDefinition).GetProperty("ChildBlocks"));
        Assert.Null(typeof(OfficeIMO.Markdown.DefinitionListDefinition).GetProperty("Blocks"));
        Assert.NotNull(typeof(OfficeIMO.Markdown.QuoteBlock).GetProperty("ChildBlocks"));
        Assert.Null(typeof(OfficeIMO.Markdown.QuoteBlock).GetProperty("Children"));
        Assert.NotNull(typeof(OfficeIMO.Markdown.DetailsBlock).GetProperty("ChildBlocks"));
        Assert.Null(typeof(OfficeIMO.Markdown.DetailsBlock).GetProperty("Children"));
        Assert.NotNull(typeof(OfficeIMO.Markdown.TableCell).GetProperty("ChildBlocks"));
        Assert.Null(typeof(OfficeIMO.Markdown.TableCell).GetProperty("Blocks"));
        Assert.NotNull(typeof(OfficeIMO.Markdown.ListItem).GetProperty("NestedBlocks"));
        Assert.NotNull(typeof(OfficeIMO.Markdown.ListItem).GetProperty("ChildBlocks"));
        Assert.Null(typeof(OfficeIMO.Markdown.ListItem).GetProperty("Children"));
        Assert.Null(typeof(OfficeIMO.Excel.Fluent.SheetComposer).GetMethod("DefinitionList"));
        Assert.Null(typeof(OfficeIMO.PowerPoint.PowerPointUnits).GetMethod("Inches"));
        Assert.Null(typeof(OfficeIMO.PowerPoint.PowerPointUnits).GetMethod("Points"));
        Assert.Null(typeof(OfficeIMO.PowerPoint.PowerPointUnits).GetMethod("Cm"));
        Assert.Null(typeof(OfficeIMO.PowerPoint.PowerPointUnits).GetMethod("Mm"));
    }
}
