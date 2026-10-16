using System.Collections;
using System.Globalization;
using System.Resources;
using System.Text;
using System.Text.RegularExpressions;
using OfficeIMO.Ocr.Tesseract;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Editor;
using OfficeIMO.Studio.Features.Reader;
using OfficeIMO.Studio.Features.Workspace;
using OfficeIMO.Studio.Infrastructure.Localization;
using OfficeIMO.Workflows;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioTranslationResourceTests {
    [Fact]
    public void AvailableTranslationsContainEveryUiResourceAndPreserveFormatContracts() {
        var manager = new ResourceManager("OfficeIMO.Studio.Localization.StudioStrings", typeof(StudioLocalizer).Assembly);
        ResourceSet neutral = manager.GetResourceSet(CultureInfo.InvariantCulture, true, false)!;
        var source = neutral.Cast<DictionaryEntry>().ToDictionary(e => (string)e.Key, e => (string)e.Value!);
        foreach (var descriptor in StudioCultureCatalog.Available.Where(c => c.Name is not "en" and not StudioCultureCatalog.PseudoCulture)) {
            ResourceSet? translated = manager.GetResourceSet(CultureInfo.GetCultureInfo(descriptor.Name), true, false);
            Assert.NotNull(translated);
            var values = translated.Cast<DictionaryEntry>().ToDictionary(e => (string)e.Key, e => (string)e.Value!);
            Assert.Equal(source.Keys.Order(), values.Keys.Order());
            foreach ((string key, string english) in source) {
                string value = values[key];
                Assert.False(string.IsNullOrWhiteSpace(value), descriptor.Name + ": " + key);
                Assert.Equal(Tokens(english), Tokens(value));
                if (Regex.IsMatch(english, @"\{\d+[,:}]")) {
                    var format = CompositeFormat.Parse(value);
                    Assert.Equal(CompositeFormat.Parse(english).MinimumArgumentCount, format.MinimumArgumentCount);
                    object[] arguments = Enumerable.Repeat<object>(2, format.MinimumArgumentCount).ToArray();
                    _ = string.Format(CultureInfo.GetCultureInfo(descriptor.Name), format, arguments);
                }
            }
        }
    }

    /// <summary>
    /// Keys the interface builds from a value, such as <c>Reader.Layout.{mode}.Label</c>, are looked up with an English
    /// fallback, so a missing resource would show English in every language without failing any other check.
    /// </summary>
    [Theory]
    [InlineData(typeof(ReaderLayoutMode), "Reader.Layout.{0}.Label")]
    [InlineData(typeof(ReaderLayoutMode), "Reader.Layout.{0}.Description")]
    [InlineData(typeof(PdfEditorTool), "Editor.Tool.{0}.Label")]
    [InlineData(typeof(PdfEditorTool), "Editor.Tool.{0}.Hint")]
    [InlineData(typeof(PdfFormFieldCreationKind), "Editor.FormKind.{0}")]
    [InlineData(typeof(TesseractOcrLanguage), "Ocr.Language.{0}")]
    [InlineData(typeof(PdfFormFieldValueIssueCode), "FormOcr.Constraint.{0}")]
    [InlineData(typeof(PdfWorkspaceOperationKind), "Operation.{0}")]
    [InlineData(typeof(OfficeWorkflowDiagnosticSeverity), "Workflow.Severity.{0}")]
    [InlineData(typeof(OfficeWorkflowStatus), "Workflow.Status.{0}")]
    public void EveryValueTheInterfaceNamesHasAResource(Type values, string keyFormat) {
        var manager = new ResourceManager("OfficeIMO.Studio.Localization.StudioStrings", typeof(StudioLocalizer).Assembly);
        ResourceSet neutral = manager.GetResourceSet(CultureInfo.InvariantCulture, true, false)!;
        string[] missing = Enum.GetNames(values).Select(name => string.Format(CultureInfo.InvariantCulture, keyFormat, name))
            .Where(key => neutral.GetString(key) is null).ToArray();
        Assert.True(missing.Length == 0, "Missing from StudioStrings.resx: " + string.Join(", ", missing));
    }

    private static string[] Tokens(string value) => Regex.Matches(value, @"(?<!\{)\{[^{}]+\}(?!\})")
        .Select(match => match.Value).Order().ToArray();
}
