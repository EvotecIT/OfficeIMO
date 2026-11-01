using System.Management.Automation;
using OfficeIMO.Drawing;
using OfficeIMO.Word;
using PSWriteOffice.Services;

namespace PSWriteOffice.Tests;

public sealed partial class ConversionCmdletWorkflowTests {
    [Theory]
    [InlineData("Word", OfficeImageExportFormat.Svg)]
    [InlineData("Excel", OfficeImageExportFormat.Svg)]
    [InlineData("PowerPoint", OfficeImageExportFormat.Svg)]
    [InlineData("Pdf", OfficeImageExportFormat.Svg)]
    [InlineData("Visio", OfficeImageExportFormat.Png)]
    [InlineData("Html", OfficeImageExportFormat.Svg)]
    public void NativeImagesUseTheSameNativePayloadForLiveAndLiteralFileInputs(string kind, OfficeImageExportFormat format) {
        object document;
        if (kind == "Word") { var word = WordDocument.Create(); word.AddParagraph("Live source evidence"); document = word; }
        else document = LiveSource(kind);
        string root = Directory.CreateTempSubdirectory("officeimo-native-images-").FullName;
        try {
            var native = OfficeLiveDocumentSource.Create(document);
            string path = Path.Combine(root, "source[1]" + native.Extension);
            byte[] saved = native.Capture(default);
            File.WriteAllBytes(path, saved);
            using var runspace = OpenRunspace();
            runspace.SessionStateProxy.SetVariable("document", document);
            runspace.SessionStateProxy.SetVariable("inputFile", new FileInfo(path));
            runspace.SessionStateProxy.SetVariable("encoding", format.ToString());
            using var shell = PowerShell.Create(); shell.Runspace = runspace;
            shell.AddScript("$document | ConvertTo-OfficeDocumentImage -Format $encoding -ErrorAction Stop; $inputFile | ConvertTo-OfficeDocumentImage -Format $encoding -ErrorAction Stop");
            var output = shell.Invoke();
            Assert.False(shell.HadErrors, string.Join(Environment.NewLine, shell.Streams.Error));
            Assert.Equal(2, output.Count);
            var live = Assert.IsType<OfficeImageExportResult>(output[0].BaseObject);
            var file = Assert.IsType<OfficeImageExportResult>(output[1].BaseObject);
            Assert.Equal(format, live.Format);
            Assert.Equal(live.Bytes, file.Bytes);
            Assert.Null(live.SavedPath);
            Assert.NotNull(live.Name);
            Assert.True(live.Width > 0 && live.Height > 0);
            Assert.Equal(saved, File.ReadAllBytes(path));
            Assert.NotEmpty(native.Capture(default));
            Assert.Single(Directory.GetFiles(root));
        } finally { (document as IDisposable)?.Dispose(); Directory.Delete(root, true); }
    }

    [Fact]
    public void NativeImageSingleFileInferenceWhatIfAndNoClobberPreserveTheDestination() {
        string root = Directory.CreateTempSubdirectory("officeimo-native-image-save-").FullName;
        try {
            string outputPath = Path.Combine(root, "new", "review.svg");
            using var runspace = OpenRunspace();
            runspace.SessionStateProxy.SetVariable("outputPath", outputPath);
            using var shell = PowerShell.Create(); shell.Runspace = runspace;
            shell.AddScript("""
                $html = '<h1>Native image evidence</h1>'
                Export-OfficeDocumentImage -Text $html -From Html -OutputPath $outputPath -WhatIf -ErrorAction Stop
                if (Test-Path -LiteralPath (Split-Path $outputPath)) { throw 'WhatIf created a destination' }
                Export-OfficeDocumentImage -Text $html -From Html -OutputPath $outputPath -NoClobber -ErrorAction Stop
                $first = [IO.File]::ReadAllBytes($outputPath)
                $refused = $false
                try { Export-OfficeDocumentImage -Text '<p>Replacement</p>' -From Html -OutputPath $outputPath -NoClobber -ErrorAction Stop }
                catch { $refused = $true }
                if (-not $refused) { throw 'NoClobber replaced an image' }
                if ([Convert]::ToBase64String($first) -ne [Convert]::ToBase64String([IO.File]::ReadAllBytes($outputPath))) { throw 'Destination changed' }
                """);
            var result = shell.Invoke();
            Assert.Empty(shell.Streams.Error);
            var image = Assert.IsType<OfficeImageExportResult>(Assert.Single(result).BaseObject);
            Assert.Equal(OfficeImageExportFormat.Svg, image.Format);
            Assert.Equal(outputPath, image.SavedPath);
            Assert.Equal(image.Bytes, File.ReadAllBytes(outputPath));
        } finally { Directory.Delete(root, true); }
    }

    [Theory]
    [InlineData("-Sheet 1")]
    [InlineData("-Range A1:B2")]
    [InlineData("-Range A1:B2 -Chart Revenue")]
    [InlineData("-Password synthetic")]
    [InlineData("-Format Gif")]
    [InlineData("-Format Png")]
    [InlineData("-Page 0")]
    [InlineData("-Options @{ UnknownImageSetting = 1 }")]
    public void InvalidNativeImageArgumentsCreateNoDestination(string arguments) {
        string root = Directory.CreateTempSubdirectory("officeimo-native-image-invalid-").FullName;
        try {
            string path = Path.Combine(root, "new", "review.svg");
            using var runspace = OpenRunspace(); runspace.SessionStateProxy.SetVariable("outputPath", path);
            using var shell = PowerShell.Create(); shell.Runspace = runspace;
            shell.AddScript("Export-OfficeDocumentImage -Text '<h1>Evidence</h1>' -From Html -OutputPath $outputPath " + arguments + " -ErrorAction Stop");
            Exception? error = Record.Exception(() => shell.Invoke());
            Assert.True(error != null || shell.HadErrors);
            Assert.False(Directory.Exists(Path.GetDirectoryName(path)));
        } finally { Directory.Delete(root, true); }
    }

    [Fact]
    public void NativeImageOutputPathRejectsMultipleInputsAndMultiplePagesBeforeWriting() {
        string root = Directory.CreateTempSubdirectory("officeimo-native-image-count-").FullName;
        try {
            using var runspace = OpenRunspace();
            string destination = Path.Combine(root, "new", "review.svg");
            runspace.SessionStateProxy.SetVariable("outputPath", destination);
            using var shell = PowerShell.Create(); shell.Runspace = runspace;
            shell.AddScript("""Export-OfficeDocumentImage -Text '<p>First</p><p style="break-before:page">Second</p>' -From Html -OutputPath $outputPath -ErrorAction Stop""");
            Exception? error = Record.Exception(() => shell.Invoke());
            Assert.True(error != null || shell.HadErrors);
            Assert.False(Directory.Exists(Path.GetDirectoryName(destination)));
            shell.Commands.Clear(); shell.Streams.Error.Clear();
            shell.AddScript("'<p>First</p>','<p>Second</p>' | Export-OfficeDocumentImage -From Html -OutputPath $outputPath -ErrorAction Stop");
            error = Record.Exception(() => shell.Invoke());
            Assert.True(error != null || shell.HadErrors);
            Assert.False(Directory.Exists(Path.GetDirectoryName(destination)));
        } finally { Directory.Delete(root, true); }
    }
}
