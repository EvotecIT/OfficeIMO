using System.Globalization;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Cmdlets.PowerPoint;

namespace PSWriteOffice.Tests;

[Collection(PowerShellRunspaceCollection.Name)]
public sealed class PowerPointObjectExportWorkflowTests : IDisposable {
    private readonly string _root = Path.Combine(Path.GetTempPath(), "officeimo-ppt-object-export-" + Guid.NewGuid().ToString("N"));
    public PowerPointObjectExportWorkflowTests() => Directory.CreateDirectory(_root);
    public void Dispose() => Directory.Delete(_root, true);

    private const string Sales = """
        $rows = @(1..7 | ForEach-Object {
            [pscustomobject]@{ Name = "Item $_"; Region = $(if ($_ % 2) { 'West' } else { 'East' }); Revenue = ($_ * 10); Deals = $_ }
        })
        """;

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void PipelineAndArrayExportNativeTablesSummaryAndChartInOrder(bool pipeline) {
        string path = Path.Combine(_root, "Sales.pptx");
        using var runspace = Open(path);
        using var shell = Shell(runspace);
        string invocation = pipeline ? "$rows | Export-OfficePowerPoint" : "Export-OfficePowerPoint -InputObject $rows";
        shell.AddScript(Sales + "\n" + invocation + """
             -Path $target -Title 'Quarterly sales' -Subtitle 'Regional performance' -Property Name,Revenue,Region -Category Region -Value Revenue,Deals -ChartTitle 'Sales by region' -MaxRows 3 -Theme Boardroom -PassThru
            """);
        var output = shell.Invoke();
        Assert.False(shell.HadErrors, Errors(shell));
        Assert.Equal(path, Assert.IsType<FileInfo>(Assert.Single(output).BaseObject).FullName);
        using var saved = PowerPointDocument.Load(path);
        Assert.Empty(saved.ValidateDocument());
        Assert.Equal(6, saved.Slides.Count);
        Assert.Contains("Quarterly sales", Text(saved.Slides[0]));
        Assert.Contains("Regional performance", Text(saved.Slides[0]));
        Assert.Empty(saved.Slides[0].Tables);
        string summary = Text(saved.Slides[1]);
        Assert.Contains("At a glance", summary);
        Assert.Contains("West", summary);
        Assert.Contains("East", summary);
        Assert.Contains(saved.Slides[1].TextBoxes, box => box.Text == "4");
        Assert.Contains(saved.Slides[1].TextBoxes, box => box.Text == "3");
        var tables = saved.Slides.Skip(2).Take(3).Select(slide => Assert.Single(slide.Tables)).ToArray();
        Assert.All(tables, table => Assert.InRange(table.Rows, 2, 4));
        AssertHeaders(tables, "Name", "Revenue", "Region");
        Assert.Equal(Enumerable.Range(1, 7).Select(index => "Item " + index),
            tables.SelectMany(table => Enumerable.Range(1, table.Rows - 1).Select(row => table.GetCell(row, 0).Text)));
        Assert.Equal(new[] { "10", "20", "30", "40", "50", "60", "70" },
            tables.SelectMany(table => Enumerable.Range(1, table.Rows - 1).Select(row => table.GetCell(row, 1).Text)));
        Assert.Contains("continued", Text(saved.Slides[3]));
        Assert.Contains("continued", Text(saved.Slides[4]));
        Assert.Empty(saved.Slides.Last().Tables);
        var chart = Assert.Single(saved.Slides.Last().Charts);
        Assert.True(chart.TryGetOfficeSnapshot(out var snapshot));
        Assert.Equal(OfficeChartType.ColumnClustered, snapshot.ChartKind);
        // Composed chart labels are native slide text above the plot, aligned with the content column.
        Assert.Contains(saved.Slides.Last().TextBoxes, box => box.Text == "Sales by region");
        Assert.Equal(new[] { "West", "East", "West", "East", "West", "East", "West" }, snapshot.Data.Categories);
        Assert.Equal(new[] { "Revenue", "Deals" }, snapshot.Data.Series.Select(series => series.Name));
        Assert.Equal(Enumerable.Range(1, 7).Select(index => index * 10D), snapshot.Data.Series[0].Values);
        Assert.Equal(Enumerable.Range(1, 7).Select(index => (double)index), snapshot.Data.Series[1].Values);
        Assert.Equal(OfficeTheme.Boardroom.Palette.Accent1.ToRgbHex(), saved.GetThemeColor(PowerPointThemeColor.Accent1));
        Assert.Equal(OfficeTheme.Boardroom.Fonts.Heading.Family, saved.GetThemeLatinFonts().MajorLatin);
    }

    [Fact]
    public void AutomaticCapacityRepeatsHeadersAndRetainsEveryRowWithoutPassThruOutput() {
        string path = Path.Combine(_root, "automatic.pptx");
        using var runspace = Open(path);
        using var shell = Shell(runspace);
        shell.AddScript("""
            1..45 | ForEach-Object { [pscustomobject]@{ Item = "Row $_"; Number = $_ } } |
                Export-OfficePowerPoint -Path $target -NoSummary
            """);
        Assert.Empty(shell.Invoke());
        Assert.False(shell.HadErrors, Errors(shell));
        using var saved = PowerPointDocument.Load(path);
        Assert.Empty(saved.ValidateDocument());
        Assert.Contains("automatic", Text(saved.Slides[0]));
        var tables = saved.Slides.SelectMany(slide => slide.Tables).ToArray();
        Assert.True(tables.Length > 1);
        AssertHeaders(tables, "Item", "Number");
        Assert.Equal(Enumerable.Range(1, 45).Select(index => "Row " + index),
            tables.SelectMany(table => Enumerable.Range(1, table.Rows - 1).Select(row => table.GetCell(row, 0).Text)));
        Assert.Equal(1 + tables.Length, saved.Slides.Count);
    }

    [Theory]
    [InlineData("", true)]
    [InlineData("-GroupBy Deals", true)]
    [InlineData("-NoSummary", false)]
    public void SummaryUsesEligibleTextGroupsOrExplicitGroupByAndCanBeSuppressed(string arguments, bool summaryExpected) {
        using var runspace = Open(Path.Combine(_root, "summary.pptx"));
        using var shell = Shell(runspace);
        shell.AddScript(Sales + "\n$rows | Export-OfficePowerPoint -Path $target -MaxRows 10 " + arguments);
        Assert.Empty(shell.Invoke());
        Assert.False(shell.HadErrors, Errors(shell));
        using var saved = PowerPointDocument.Load((string)runspace.SessionStateProxy.GetVariable("target"));
        var summaries = saved.Slides.Where(slide => Text(slide).Contains("At a glance", StringComparison.Ordinal)).ToArray();
        if (!summaryExpected) { Assert.Empty(summaries); return; }
        Assert.NotEmpty(summaries);
        string text = string.Join("\n", summaries.Select(Text));
        if (arguments.Length == 0) {
            Assert.Single(summaries);
            Assert.Contains("West", text);
            Assert.Contains("East", text);
            Assert.DoesNotContain("Item 1", text);
        } else {
            // Seven explicitly requested groups flow through native KPI continuation rather than being truncated.
            foreach (int index in Enumerable.Range(1, 7))
                Assert.Contains(summaries.SelectMany(slide => slide.TextBoxes), box => box.Text == index.ToString(CultureInfo.InvariantCulture));
        }
    }

    [Fact]
    public void ChartOnlyUsesCultureForNumbersAndOmitsTablesAndSummary() {
        string path = Path.Combine(_root, "chart-only.pptx");
        using var runspace = Open(path);
        using var shell = Shell(runspace);
        shell.AddScript("""
            Export-OfficePowerPoint -InputObject @(
                [pscustomobject]@{ Region = 'West'; Revenue = '12,5' },
                [pscustomobject]@{ Region = 'East'; Revenue = '24,75' }
            ) -Path $target -Category Region -Value Revenue -Culture fr-FR -ChartType BarClustered -NoTable
            """);
        Assert.Empty(shell.Invoke());
        Assert.False(shell.HadErrors, Errors(shell));
        using var saved = PowerPointDocument.Load(path);
        Assert.Equal(2, saved.Slides.Count);
        Assert.Empty(saved.Slides.SelectMany(slide => slide.Tables));
        Assert.DoesNotContain("At a glance", string.Join("\n", saved.Slides.Select(Text)));
        var chart = Assert.Single(saved.Slides[1].Charts);
        Assert.True(chart.TryGetOfficeSnapshot(out var snapshot));
        Assert.Equal(OfficeChartType.BarClustered, snapshot.ChartKind);
        Assert.Equal(new[] { 12.5D, 24.75D }, Assert.Single(snapshot.Data.Series).Values);
    }

    [Fact]
    public void SharedNormalizerPreservesNestedValuesAndCultureInTableCells() {
        string path = Path.Combine(_root, "culture.pptx");
        using var runspace = Open(path);
        using var shell = Shell(runspace);
        shell.AddScript("""
            [pscustomobject]@{ Amount = 12.5; Tags = @('A', 'B'); Detail = [ordered]@{ State = 'Ready'; Count = 2 } } |
                Export-OfficePowerPoint -Path $target -Property Amount,Tags,Detail -Culture fr-FR -NoSummary
            """);
        Assert.Empty(shell.Invoke());
        Assert.False(shell.HadErrors, Errors(shell));
        using var saved = PowerPointDocument.Load(path);
        var table = Assert.Single(saved.Slides.SelectMany(slide => slide.Tables));
        AssertHeaders(new[] { table }, "Amount", "Tags", "Detail");
        Assert.Equal("12,5", table.GetCell(1, 0).Text);
        Assert.Equal("A, B", table.GetCell(1, 1).Text);
        Assert.Equal("State: Ready; Count: 2", table.GetCell(1, 2).Text);
    }

    [Fact]
    public void RealThemeModelAndAccentAreReadBackWithoutMutatingTheModelOrBuiltInTheme() {
        string path = Path.Combine(_root, "brand.pptx");
        OfficeTheme source = OfficeTheme.Dark.WithName("Custom dark");
        string originalAccent = source.Palette.Accent1.ToRgbHex();
        using var runspace = Open(path);
        runspace.SessionStateProxy.SetVariable("customTheme", source);
        using var shell = Shell(runspace);
        shell.AddScript(Sales + "\n$rows | Export-OfficePowerPoint -Path $target -Theme $customTheme -Accent '#A43158' -NoSummary");
        Assert.Empty(shell.Invoke());
        Assert.False(shell.HadErrors, Errors(shell));
        using var saved = PowerPointDocument.Load(path);
        Assert.Equal("A43158", saved.GetThemeColor(PowerPointThemeColor.Accent1));
        Assert.Equal(source.Palette.Background.ToRgbHex(), saved.GetThemeColor(PowerPointThemeColor.Light1));
        Assert.Equal(source.Fonts.Body.Family, saved.GetThemeLatinFonts().MinorLatin);
        Assert.Equal(originalAccent, source.Palette.Accent1.ToRgbHex());
        Assert.Equal(originalAccent, OfficeTheme.Dark.Palette.Accent1.ToRgbHex());
    }

    [Theory]
    [InlineData(".pptx", false)]
    [InlineData(".pptx", true)]
    [InlineData(".potx", false)]
    [InlineData(".potx", true)]
    public void TemplateRetainsNativeMastersFontsAndPaletteWhileReplacingItsSampleSlides(string templateExtension, bool useAccentRole) {
        string templatePath = Path.Combine(_root, "template" + templateExtension), path = Path.Combine(_root, "from-template.pptx");
        using (var template = PowerPointDocument.Create(new PowerPointCreateOptions { Theme = OfficeTheme.Classic })) {
            template.SetThemeColorForAllMasters(PowerPointThemeColor.Accent2, "987654");
            template.SetThemeLatinFontsForAllMasters("Georgia", "Verdana");
            var options = PowerPointCompositionOptions.FromDesign(PowerPointDeckDesign.FromBrand("123456", "template"));
            options.ApplyTheme = false;
            template.Compose(new PowerPointDeckPlan().AddContent("Template sample", Array.Empty<PowerPointLayoutBlock>(), PowerPointLayout.Title), options);
            template.Save(templatePath);
        }
        // Inspect the persisted packages independently of OfficeIMO's loader and its format normalization.
        using (var templatePackage = PresentationDocument.Open(templatePath, false))
            Assert.Equal(templateExtension == ".potx" ? PresentationDocumentType.Template : PresentationDocumentType.Presentation,
                templatePackage.DocumentType);
        byte[] original = File.ReadAllBytes(templatePath);
        using var runspace = Open(path);
        runspace.SessionStateProxy.SetVariable("templatePath", templatePath);
        using var shell = Shell(runspace);
        shell.AddScript(Sales + "\n$rows | Export-OfficePowerPoint -Path $target -Theme $templatePath -NoSummary " +
            (useAccentRole ? "-Accent Accent2" : ""));
        Assert.Empty(shell.Invoke());
        Assert.False(shell.HadErrors, Errors(shell));
        using (var exportedPackage = PresentationDocument.Open(path, false))
            Assert.Equal(PresentationDocumentType.Presentation, exportedPackage.DocumentType);
        using var saved = PowerPointDocument.Load(path);
        Assert.Equal("Georgia", saved.GetThemeLatinFonts().MajorLatin);
        Assert.Equal("Verdana", saved.GetThemeLatinFonts().MinorLatin);
        Assert.Equal(useAccentRole ? "987654" : OfficeTheme.Classic.Palette.Accent1.ToRgbHex(),
            saved.GetThemeColor(PowerPointThemeColor.Accent1));
        Assert.DoesNotContain("Template sample", string.Join("\n", saved.Slides.Select(Text)));
        using var source = PowerPointDocument.Load(templatePath);
        Assert.Equal(source.OpenXmlDocument.PresentationPart!.SlideMasterParts.Count(), saved.OpenXmlDocument.PresentationPart!.SlideMasterParts.Count());
        Assert.Equal(source.GetSlideLayouts().Count, saved.GetSlideLayouts().Count);
        Assert.Equal(original, File.ReadAllBytes(templatePath));
    }

    [Fact]
    public void WhatIfDoesNotCreateDestinationFoldersOrPassThruOutput() {
        string folder = Path.Combine(_root, "not-created");
        using var runspace = Open(Path.Combine(folder, "whatif.pptx"));
        using var shell = Shell(runspace);
        shell.AddScript(Sales + "\n$rows | Export-OfficePowerPoint -Path $target -WhatIf -PassThru");
        Assert.Empty(shell.Invoke());
        Assert.False(shell.HadErrors, Errors(shell));
        Assert.False(Directory.Exists(folder));
        Assert.Empty(Directory.GetFileSystemEntries(_root));
    }

    [Fact]
    public void NoClobberLeavesExistingBytesAndNoTemporaryFiles() {
        string path = Path.Combine(_root, "existing.pptx");
        byte[] original = { 1, 3, 5, 7 };
        File.WriteAllBytes(path, original);
        using var runspace = Open(path);
        using var shell = Shell(runspace);
        shell.AddScript(Sales + "\n$rows | Export-OfficePowerPoint -Path $target -NoClobber");
        AssertFailed(shell);
        Assert.Equal(original, File.ReadAllBytes(path));
        Assert.Equal(new[] { path }, Directory.GetFileSystemEntries(_root));
    }

    [Fact]
    public void DefaultOverwritePublishesACompleteDeckAndOnlyReturnsItsFile() {
        string path = Path.Combine(_root, "replace.pptx");
        File.WriteAllBytes(path, new byte[] { 2, 4, 6 });
        using var runspace = Open(path);
        using var shell = Shell(runspace);
        shell.AddScript(Sales + "\n$rows | Export-OfficePowerPoint -Path $target -NoSummary -PassThru");
        var result = shell.Invoke();
        Assert.False(shell.HadErrors, Errors(shell));
        Assert.Equal(path, Assert.IsType<FileInfo>(Assert.Single(result).BaseObject).FullName);
        using var saved = PowerPointDocument.Load(path);
        Assert.Equal(7, saved.Slides.SelectMany(slide => slide.Tables).Sum(table => table.Rows - 1));
        Assert.Empty(saved.ValidateDocument());
        Assert.Equal(new[] { path }, Directory.GetFileSystemEntries(_root));
    }

    [Theory]
    [InlineData("-Property Missing")]
    [InlineData("-Property Name,name")]
    [InlineData("-Category Missing -Value Revenue")]
    [InlineData("-Category Region -Value Missing")]
    [InlineData("-Category Region")]
    [InlineData("-Value Revenue")]
    [InlineData("-GroupBy Missing")]
    [InlineData("-ChartTitle Orphan")]
    [InlineData("-Category Region -Value Revenue -ChartType 999")]
    [InlineData("-Category Region -Value Revenue -ChartType Bubble")]
    [InlineData("-Category Region -Value Revenue,Deals -ChartType Pie")]
    [InlineData("-Category Region -Value Revenue,Deals -ChartType Doughnut")]
    [InlineData("-NoTable")]
    [InlineData("-Theme missing-template.potx")]
    public void InvalidSettingsCannotReplaceExistingBytesOrCreateDestinationFolders(string arguments) {
        string existing = Path.Combine(_root, "existing.pptx");
        byte[] original = { 9, 8, 7, 6 };
        File.WriteAllBytes(existing, original);
        foreach (string path in new[] { existing, Path.Combine(_root, "absent", "new.pptx") }) {
            using var runspace = Open(path);
            using var shell = Shell(runspace);
            shell.AddScript(Sales + "\n$rows | Export-OfficePowerPoint -Path $target " + arguments);
            AssertFailed(shell);
            Assert.Equal(original, File.ReadAllBytes(existing));
            Assert.False(Directory.Exists(Path.Combine(_root, "absent")));
            Assert.Equal(new[] { existing }, Directory.GetFileSystemEntries(_root));
        }
    }

    [Theory]
    [InlineData("$null")]
    [InlineData("$true")]
    [InlineData("'not a number'")]
    [InlineData("[double]::NaN")]
    public void InvalidChartValuesLeaveExistingDestinationUnchanged(string expression) {
        string path = Path.Combine(_root, "invalid-value.pptx");
        byte[] original = { 4, 2 };
        File.WriteAllBytes(path, original);
        using var runspace = Open(path);
        using var shell = Shell(runspace);
        shell.AddScript("[pscustomobject]@{ Region = 'West'; Revenue = " + expression +
            " } | Export-OfficePowerPoint -Path $target -Category Region -Value Revenue");
        AssertFailed(shell);
        Assert.Equal(original, File.ReadAllBytes(path));
        Assert.Equal(new[] { path }, Directory.GetFileSystemEntries(_root));
    }

    private static void AssertHeaders(IEnumerable<PowerPointTable> tables, params string[] headers) {
        foreach (var table in tables) {
            Assert.True(table.FirstRow);
            Assert.Equal(headers.Length, table.Columns);
            Assert.Equal(headers, Enumerable.Range(0, table.Columns).Select(column => table.GetCell(0, column).Text));
        }
    }

    private static void AssertFailed(PowerShell shell) {
        Exception? failure = Record.Exception(() => shell.Invoke());
        Assert.True(failure != null || shell.HadErrors, "Expected export validation or native preservation to reject the request.");
        Assert.Empty(shell.Streams.Warning);
    }

    private static string Text(PowerPointSlide slide) => string.Join("\n", slide.TextBoxes.Select(box => box.Text));
    private static string Errors(PowerShell shell) => string.Join(Environment.NewLine, shell.Streams.Error);
    private static PowerShell Shell(Runspace runspace) { var shell = PowerShell.Create(); shell.Runspace = runspace; return shell; }
    private static Runspace Open(string path) {
        var state = InitialSessionState.CreateDefault();
        state.Commands.Add(new SessionStateCmdletEntry("Export-OfficePowerPoint", typeof(ExportOfficePowerPointCommand), null));
        var runspace = RunspaceFactory.CreateRunspace(state);
        runspace.Open();
        runspace.SessionStateProxy.SetVariable("target", path);
        return runspace;
    }
}
