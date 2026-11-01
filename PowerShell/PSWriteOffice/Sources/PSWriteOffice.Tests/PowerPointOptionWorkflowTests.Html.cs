using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Html;

namespace PSWriteOffice.Tests;

public sealed partial class PowerPointOptionWorkflowTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void HtmlOptionsReachSavedAndStringConversionsWithoutChangingTheDeckOrReusableSettings(bool include) {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-ppt-html-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            string source = Path.Combine(root, "deck[1].pptx"), target = Path.Combine(root, "review.html");
            using (var deck = PowerPointDocument.Create()) {
                var slide = deck.AddSlide();
                slide.AddTextBox("Visible slide content");
                slide.Notes.Text = "Presenter talking points";
                var table = slide.AddTable(1, 1);
                table.GetCell(0, 0).Text = "Table evidence";
                var hidden = deck.AddSlide();
                hidden.Hidden = true;
                hidden.AddTextBox("Hidden slide content");
                deck.Save(source);
            }
            byte[] original = File.ReadAllBytes(source);
            using var runspace = Open();
            runspace.SessionStateProxy.SetVariable("source", source);
            runspace.SessionStateProxy.SetVariable("target", target);
            runspace.SessionStateProxy.SetVariable("customTheme", OfficeTheme.Dark);
            using var shell = PowerShell.Create();
            shell.Runspace = runspace;
            string flag = include ? "$true" : "$false", negative = include ? "$false" : "$true";
            shell.AddScript($$"""
                $settings = New-OfficePowerPointHtmlOptions -Profile SemanticSlides -Theme $customTheme -Title 'Service review' -IncludeHiddenSlides:{{flag}} -NoNotes:{{negative}} -NoTables:{{negative}} -ErrorAction Stop
                Get-Item -LiteralPath $source | Convert-OfficeDocument -To Html -OutputPath $target -Options $settings -ErrorAction Stop
                Convert-OfficeDocument -LiteralPath $source -To Html -AsString -Options $settings -ErrorAction Stop
                $settings
                """);
            var output = shell.Invoke();
            Assert.False(shell.HadErrors, Errors(shell));
            Assert.Equal(3, output.Count);
            Assert.Equal(target, Assert.IsType<OfficeConversionReport>(output[0].BaseObject).OutputPath);
            string html = Assert.IsType<string>(output[1].BaseObject);
            Assert.Equal(html, File.ReadAllText(target));
            Assert.Contains("Visible slide content", html, StringComparison.Ordinal);
            Assert.Equal(include, html.Contains("Hidden slide content", StringComparison.Ordinal));
            Assert.Equal(include, html.Contains("Presenter talking points", StringComparison.Ordinal));
            Assert.Equal(include, html.Contains("Table evidence", StringComparison.Ordinal));
            Assert.Contains("<title>Service review</title>", html, StringComparison.Ordinal);
            var options = Assert.IsType<PowerPointHtmlSaveOptions>(output[2].BaseObject);
            Assert.Same(OfficeTheme.Dark, options.DocumentOutput.DocumentTheme);
            Assert.Equal(include, options.IncludeTables);
            Assert.Equal(original, File.ReadAllBytes(source));
        } finally { Directory.Delete(root, true); }
    }
}
