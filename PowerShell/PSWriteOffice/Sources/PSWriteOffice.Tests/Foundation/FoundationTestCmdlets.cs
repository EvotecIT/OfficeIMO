using System.Management.Automation;
using System.Text;
using OfficeIMO;
using OfficeIMO.Drawing;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Tests.Foundation;

// Only the test assembly contains these commands; the shipped module never references it.
public sealed class FoundationOptions {
    public string Title { get; set; } = "";
    public int Count { get; set; }
    public OfficeLength Width { get; set; }
    public OfficeColor Color { get; set; }
}

[Cmdlet(VerbsDiagnostic.Test, "OfficeFoundationBinding")]
[OutputType(typeof(PSObject))]
public sealed class TestOfficeFoundationBindingCommand : OfficeCmdlet {
    [Parameter][OfficeLengthTransformation][ArgumentCompleter(typeof(OfficeLengthCompleter))]
    public OfficeLength Width { get; set; }
    [Parameter][OfficeColorTransformation][ArgumentCompleter(typeof(OfficeColorCompleter))]
    public OfficeColor Color { get; set; }
    [Parameter][OfficeThemeTransformation][ArgumentCompleter(typeof(OfficeThemeCompleter))]
    public OfficeThemeInput? Theme { get; set; }
    [Parameter][OfficePasswordTransformation]
    public System.Security.SecureString? Password { get; set; }
    [Parameter][OfficeSheetTransformation][ArgumentCompleter(typeof(OfficeSheetCompleter))]
    public OfficeSheetSelector? Sheet { get; set; }
    [Parameter][OfficeOptionsTransformation(typeof(FoundationOptions))]
    public FoundationOptions? Options { get; set; }
    [Parameter][OfficeFormatTransformation][ArgumentCompleter(typeof(OfficeFormatCompleter))]
    public OfficeDocumentFormat To { get; set; }
    [Parameter] public SwitchParameter Bold { get; set; }
    [Parameter] public SwitchParameter NoFreeze { get; set; }
    protected override void ProcessRecord() => WriteObject(new {
        Width, Color, Theme, Sheet, Options, To,
        PasswordLength = Password?.Length, Bold = BoundSwitch(nameof(Bold)), Freeze = FeatureEnabled("Freeze")
    });
}

public sealed class FoundationDocument : IDisposable {
    public string Text { get; set; } = "";
    public bool Loss { get; set; }
    public bool Disposed { get; private set; }
    public void Dispose() => Disposed = true;
}

[Cmdlet(VerbsCommon.New, "OfficeFoundationDocument", SupportsShouldProcess = true)]
[OutputType(typeof(FoundationDocument), typeof(FileInfo))]
public sealed class NewOfficeFoundationDocumentCommand : OfficeLifecycleCmdlet<FoundationDocument> {
    [Parameter(Position = 0)] public string? Path { get; set; }
    [Parameter] public SwitchParameter Force { get; set; }
    protected override void ProcessRecord() => RunDocument(Path, isNew: true, Force, path => {
        if (Content != null && path != null) File.WriteAllText(path, "partial creation");
        return new FoundationDocument();
    }, FoundationSave.Save);
}

[Cmdlet(VerbsCommon.Open, "OfficeFoundationDocument", SupportsShouldProcess = true, DefaultParameterSetName = "Path")]
[OutputType(typeof(FoundationDocument), typeof(FileInfo))]
public sealed class OpenOfficeFoundationDocumentCommand : OfficeLifecycleCmdlet<FoundationDocument> {
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = "Path")][SupportsWildcards]
    public string? Path { get; set; }
    [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true, ParameterSetName = "LiteralPath")][Alias("PSPath", "FullName")]
    public string? LiteralPath { get; set; }
    protected override void ProcessRecord() {
        foreach (string path in OfficePath.ResolveInput(SessionState, LiteralPath ?? Path!, LiteralPath != null))
            RunDocument(path, isNew: false, force: false, p => new FoundationDocument { Text = File.ReadAllText(p!) }, FoundationSave.Save);
    }
}

internal static class FoundationSave {
    internal static OfficeConversionReport Save(FoundationDocument document, string path, bool allowLoss) {
        File.WriteAllText(path, document.Text);
        return new OfficeConversionReport("native", "native", path, diagnostics: document.Loss ? new[] {
            new OfficeDiagnostic("chart-flattened", OfficeDiagnosticSeverity.Warning, "Native content would be lost.", lossKind: OfficeConversionLossKind.Approximation)
        } : null);
    }
}

[Cmdlet(VerbsCommon.Add, "OfficeFoundationText")]
[OutputType(typeof(void))]
public sealed class AddOfficeFoundationTextCommand : OfficeDocumentCmdlet<FoundationDocument> {
    [Parameter(Mandatory = true, Position = 0)] public string Text { get; set; } = "";
    [Parameter] public SwitchParameter Loss { get; set; }
    protected override void ProcessRecord() {
        FoundationDocument document = GetDocument();
        document.Text += Text;
        document.Loss |= Loss;
    }
}

[Cmdlet(VerbsDiagnostic.Test, "OfficeFoundationPath", DefaultParameterSetName = "Path")]
[OutputType(typeof(string))]
public sealed class TestOfficeFoundationPathCommand : OfficeInputCmdlet {
    [Parameter] public SwitchParameter IncludeDirectory { get; set; }
    protected override Task ProcessRecordAsync() {
        foreach (string path in ResolveInputs(IncludeDirectory)) WriteObject(path);
        return Task.CompletedTask;
    }
}

[Cmdlet(VerbsData.Export, "OfficeFoundation", SupportsShouldProcess = true, DefaultParameterSetName = "File")]
[OutputType(typeof(FileInfo))]
public sealed class ExportOfficeFoundationCommand : OfficeCmdlet {
    [Parameter(Mandatory = true, ParameterSetName = "File")] public string OutputPath { get; set; } = "";
    [Parameter(Mandatory = true, ParameterSetName = "Directory")] public string OutputDirectory { get; set; } = "";
    [Parameter] public SwitchParameter NoClobber { get; set; }
    protected override void ProcessRecord() {
        string path;
        try {
            path = ParameterSetName == "File" ? ResolveOutputPath(OutputPath) : System.IO.Path.Combine(ResolveOutputDirectory(OutputDirectory), "result.txt");
        } catch (ArgumentException error) { InvalidArgument(error, OutputPath); return; }
        if (!ShouldProcess(path, "Export")) return;
        try {
            OfficeOutput.Write(path, stream => { byte[] bytes = Encoding.UTF8.GetBytes("export"); stream.Write(bytes, 0, bytes.Length); }, NoClobber);
            WriteObject(new FileInfo(path));
        } catch (IOException error) { ItemError(error, "output-exists", path); }
    }
}

[Cmdlet(VerbsDiagnostic.Test, "OfficeFoundationBatch", DefaultParameterSetName = "Path")]
[OutputType(typeof(string))]
public sealed class TestOfficeFoundationBatchCommand : OfficeInputCmdlet {
    [Parameter]
    public Task? SlowRelease { get; set; }

    protected override Task ProcessRecordAsync() => RunBatchAsync(ResolveInputs(), async (path, token) => {
        if (System.IO.Path.GetFileName(path).StartsWith("cancel")) {
            FoundationBatchControl.Started.Set();
            try { await Task.Delay(Timeout.Infinite, token); }
            finally { FoundationBatchControl.Cancelled.Set(); }
        }
        if (SlowRelease != null && System.IO.Path.GetFileName(path).StartsWith("slow"))
            await SlowRelease.WaitAsync(TimeSpan.FromSeconds(10), token);
        string text = await File.ReadAllTextAsync(path, token);
        if (text == "fail") throw new InvalidDataException("Cannot read this item.");
        return text;
    }, "Reading foundation fixtures", "input-invalid");
}

internal static class FoundationBatchControl {
    internal static readonly ManualResetEventSlim Started = new();
    internal static readonly ManualResetEventSlim Cancelled = new();
}
