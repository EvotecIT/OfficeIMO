using System.Reflection;
using OfficeIMO.Drawing;
using OfficeIMO.Email;
using OfficeIMO.Excel;
using OfficeIMO.Html;
using OfficeIMO.Pdf;
using OfficeIMO.Reader;
using OfficeIMO.Word;
using OfficeIMO.Word.Fluent;
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
    public void ExcelRemoteLoadsAreAsyncOnly() {
        MethodInfo[] documentMethods = typeof(ExcelDocument).GetMethods(
            BindingFlags.Public | BindingFlags.Static | BindingFlags.DeclaredOnly);
        MethodInfo[] readerMethods = typeof(ExcelDocumentReader).GetMethods(
            BindingFlags.Public | BindingFlags.Static | BindingFlags.DeclaredOnly);

        Assert.Contains(documentMethods, static method =>
            method.Name == "LoadAsync" && method.GetParameters().FirstOrDefault()?.ParameterType == typeof(Uri));
        Assert.DoesNotContain(documentMethods, static method =>
            method.Name == "Load" && method.GetParameters().FirstOrDefault()?.ParameterType == typeof(Uri));
        Assert.Contains(readerMethods, static method =>
            method.Name == "OpenAsync" && method.GetParameters().FirstOrDefault()?.ParameterType == typeof(Uri));
        Assert.DoesNotContain(readerMethods, static method =>
            method.Name == "Open" && method.GetParameters().FirstOrDefault()?.ParameterType == typeof(Uri));
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
    public void DrawingOwnsTheSharedRemoteImageLoader() {
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
