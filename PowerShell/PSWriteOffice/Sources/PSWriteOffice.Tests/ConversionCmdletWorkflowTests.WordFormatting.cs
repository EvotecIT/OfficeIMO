using System.Management.Automation;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO;
using OfficeIMO.Word;

namespace PSWriteOffice.Tests;

public sealed partial class ConversionCmdletWorkflowTests {
    [Fact]
    public void WordSettersPreserveOmittedFlagsAndSaveScopedRunsAndPhysicalParagraphSpacing() {
        using var document = WordDocument.Create();
        var paragraph = document.AddParagraph("First ");
        paragraph.AddText("second");
        paragraph.KeepWithNext = true;
        paragraph.KeepLinesTogether = true;
        var runs = paragraph.GetRuns().ToArray();
        runs[0].Bold = true;
        runs[1].Bold = true;
        runs[1].Italic = true;
        using var runspace = OpenRunspace();
        runspace.SessionStateProxy.SetVariable("paragraph", paragraph);
        runspace.SessionStateProxy.SetVariable("run", runs[1]);
        using var shell = PowerShell.Create(); shell.Runspace = runspace;
        shell.AddScript("""
            $paragraph | Set-OfficeWordParagraph -Alignment Start -IndentationBefore 1cm -SpacingAfter 6pt -LineSpacing 14pt -KeepWithNext:$false -PassThru
            $run | Set-OfficeWordText -Bold:$false -Strikethrough -Font Aptos -Size 10.5pt -PassThru
            """);
        var output = shell.Invoke();
        Assert.False(shell.HadErrors, string.Join(Environment.NewLine, shell.Streams.Error));
        Assert.Equal(2, output.Count);
        Assert.Same(paragraph, output[0].BaseObject);
        Assert.Same(runs[1], output[1].BaseObject);
        using var saved = new MemoryStream(); document.Save(saved); saved.Position = 0;
        using var package = WordprocessingDocument.Open(saved, false);
        var logical = Assert.Single(package.MainDocumentPart!.Document!.Body!.Elements<Paragraph>());
        var properties = logical.ParagraphProperties!;
        Assert.Equal("start", properties.Justification!.Val!.InnerText);
        Assert.Equal("567", properties.Indentation!.Left!.Value);
        Assert.Equal("120", properties.SpacingBetweenLines!.After!.Value);
        Assert.Equal("280", properties.SpacingBetweenLines.Line!.Value);
        Assert.Equal("exact", properties.SpacingBetweenLines.LineRule!.InnerText);
        Assert.False(properties.KeepNext?.Val?.Value ?? properties.KeepNext != null);
        Assert.True(properties.KeepLines?.Val?.Value ?? properties.KeepLines != null);
        var storedRuns = logical.Elements<Run>().ToArray();
        Assert.True(storedRuns[0].RunProperties!.Bold!.Val?.Value ?? true);
        var scoped = storedRuns[1].RunProperties!;
        Assert.False(scoped.Bold!.Val!.Value);
        Assert.True(scoped.Italic!.Val?.Value ?? true);
        Assert.True(scoped.Strike!.Val?.Value ?? true);
        Assert.Equal("21", scoped.FontSize!.Val!.Value);
        Assert.Equal("Aptos", scoped.RunFonts!.Ascii!.Value);
    }

    [Theory]
    [InlineData("Set-OfficeWordText", "Size")]
    [InlineData("Set-OfficeWordText", "Color")]
    [InlineData("Set-OfficeWordParagraph", "LineSpacing")]
    public void NullableWordSettersClearOnlyTheBoundOverride(string command, string parameter) {
        using var document = WordDocument.Create();
        var paragraph = document.AddParagraph("Retained text");
        paragraph.FontSize = OfficeLength.Points(14);
        paragraph.ColorHex = "112233";
        paragraph.LineSpacing = WordLineSpacing.Exact(OfficeLength.Points(18));
        paragraph.LineSpacingBefore = OfficeLength.Points(6);
        paragraph.IndentationBefore = OfficeLength.Points(18);
        using var runspace = OpenRunspace();
        using var shell = PowerShell.Create(); shell.Runspace = runspace;
        // Omitting each override preserves it, including through the shared binder.
        shell.AddCommand(command).AddParameter("Paragraph", paragraph).AddParameter("PassThru");
        shell.Invoke();
        Assert.False(shell.HadErrors, string.Join(Environment.NewLine, shell.Streams.Error));
        Assert.Equal(14, paragraph.FontSize!.Value.ToPoints());
        Assert.Equal("112233", paragraph.ColorHex);
        Assert.Equal(18, paragraph.LineSpacing!.Value.Height!.Value.ToPoints());
        shell.Commands.Clear();
        shell.AddCommand(command).AddParameter("Paragraph", paragraph).AddParameter(parameter, null);
        if (parameter == "LineSpacing") shell.AddParameter("SpacingBefore", null).AddParameter("IndentationBefore", null);
        shell.Invoke();
        Assert.False(shell.HadErrors, string.Join(Environment.NewLine, shell.Streams.Error));
        using var bytes = new MemoryStream(); document.Save(bytes); bytes.Position = 0;
        using var package = WordprocessingDocument.Open(bytes, false);
        var stored = Assert.Single(package.MainDocumentPart!.Document!.Body!.Elements<Paragraph>());
        var run = Assert.Single(stored.Elements<Run>());
        Assert.Equal("Retained text", stored.InnerText);
        Assert.Equal("120", stored.ParagraphProperties!.SpacingBetweenLines!.Before!.Value);
        Assert.Equal("360", stored.ParagraphProperties.Indentation!.Left!.Value);
        if (parameter == "Size") Assert.Null(run.RunProperties?.FontSize);
        else Assert.Equal("28", run.RunProperties!.FontSize!.Val!.Value);
        if (parameter == "Color") Assert.Null(run.RunProperties?.Color);
        else Assert.Equal("112233", run.RunProperties!.Color!.Val!.Value);
        if (parameter == "LineSpacing") Assert.Null(stored.ParagraphProperties?.SpacingBetweenLines?.Line);
        else Assert.Equal("360", stored.ParagraphProperties!.SpacingBetweenLines!.Line!.Value);
    }

    [Fact]
    public void PhysicalLineSpacingCannotBeReinterpretedAsAnAutomaticMultipleOrPartiallyChangeTheParagraph() {
        using var document = WordDocument.Create();
        var paragraph = document.AddParagraph("Unchanged");
        paragraph.ParagraphAlignment = WordParagraphAlignment.Left;
        paragraph.LineSpacing = WordLineSpacing.Multiple(1.5);
        using var runspace = OpenRunspace();
        using var shell = PowerShell.Create(); shell.Runspace = runspace;
        shell.AddCommand("Set-OfficeWordParagraph").AddParameter("Paragraph", paragraph)
            .AddParameter("Alignment", OfficeAlignment.Right).AddParameter("LineSpacing", "14pt")
            .AddParameter("LineSpacingRule", WordLineSpacingRule.Auto);
        var exception = Record.Exception(() => shell.Invoke());
        Assert.True(exception != null || shell.HadErrors);
        string errors = exception?.ToString() ?? string.Join(Environment.NewLine, shell.Streams.Error);
        Assert.Contains("physical height", errors, StringComparison.Ordinal);
        Assert.Equal(WordParagraphAlignment.Left, paragraph.ParagraphAlignment);
        Assert.Equal(1.5, paragraph.LineSpacing!.Value.MultipleValue);
    }
}
