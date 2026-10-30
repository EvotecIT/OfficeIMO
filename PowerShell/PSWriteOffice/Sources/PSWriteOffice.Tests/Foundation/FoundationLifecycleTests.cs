using System.Management.Automation;
using System.Management.Automation.Runspaces;
using OfficeIMO;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Tests.Foundation;

public sealed class FoundationLifecycleTests : IDisposable {
    private readonly string _root = Path.Combine(Path.GetTempPath(), "office-foundation-" + Guid.NewGuid().ToString("N"));
    public FoundationLifecycleTests() => Directory.CreateDirectory(_root);
    public void Dispose() => Directory.Delete(_root, true);
    private string Output => Path.Combine(_root, "document.txt");

    [Theory]
    [InlineData(false)][InlineData(true)]
    public void ThrowingContentLeavesNoNewFileOrReplacement(bool existing) {
        if (existing) File.WriteAllText(Output, "original");
        FoundationDocument? document = null;
        var expected = new InvalidOperationException("authoring failure");
        Exception actual = Assert.Throws<InvalidOperationException>(() => OfficeContentLifecycle.Run(Output, true, !existing, false,
            stage => { File.WriteAllText(stage, "partial"); return document = new FoundationDocument(); },
            _ => throw expected, FoundationSave.Save));
        Assert.Same(expected, actual);
        Assert.True(document!.Disposed);
        Assert.Null(OfficeDocumentContext.Find<FoundationDocument>());
        Assert.Equal(existing ? new[] { Output } : Array.Empty<string>(), Directory.GetFiles(_root));
        if (existing) Assert.Equal("original", File.ReadAllText(Output));
    }
    [Fact]
    public void OpenBlockFailurePreservesInput() {
        File.WriteAllText(Output, "original");
        Assert.Throws<InvalidOperationException>(() => OfficeContentLifecycle.Run(Output, false, false, false,
            p => new FoundationDocument { Text = File.ReadAllText(p) }, _ => throw new InvalidOperationException("failed"), FoundationSave.Save));
        Assert.Equal("original", File.ReadAllText(Output));
    }
    [Theory]
    [InlineData(false)][InlineData(true)]
    public void NativeLossRequiresAllowLoss(bool allowLoss) {
        FoundationDocument? document = null;
        Action operation = () => OfficeContentLifecycle.Run(Output, true, true, allowLoss,
            _ => document = new FoundationDocument(), d => { d.Text = "saved"; d.Loss = true; }, FoundationSave.Save);
        if (allowLoss) { operation(); Assert.Equal("saved", File.ReadAllText(Output)); }
        else { Assert.Throws<OfficeConversionException>(operation); Assert.False(File.Exists(Output)); }
        Assert.True(document!.Disposed);
        Assert.Empty(Directory.GetFiles(_root, ".*"));
    }
    [Fact]
    public void DestinationCreatedDuringBlockIsNotClobbered() {
        Assert.Throws<IOException>(() => OfficeContentLifecycle.Run(Output, true, true, false,
            _ => new FoundationDocument(), d => { d.Text = "new"; File.WriteAllText(Output, "other writer"); }, FoundationSave.Save));
        Assert.Equal("other writer", File.ReadAllText(Output));
        Assert.Single(Directory.GetFiles(_root));
    }
    [Fact]
    public void CancellationBeforePublicationRemovesStaging() {
        using var stop = new CancellationTokenSource();
        Assert.Throws<OperationCanceledException>(() => OfficeContentLifecycle.Run(Output, true, true, false,
            stage => { File.WriteAllText(stage, "partial"); return new FoundationDocument(); }, _ => stop.Cancel(), FoundationSave.Save, stop.Token));
        Assert.Empty(Directory.GetFiles(_root));
    }
    [Fact]
    public void NestedContextsRestoreParentAndDoNotFlowToWorkers() {
        var outer = new FoundationDocument();
        using (OfficeDocumentContext.Push(outer)) {
            Assert.Same(outer, OfficeDocumentContext.Find<FoundationDocument>());
            using (OfficeDocumentContext.Push(new FoundationDocument())) Assert.NotSame(outer, OfficeDocumentContext.Find<FoundationDocument>());
            Assert.Same(outer, OfficeDocumentContext.Find<FoundationDocument>());
            FoundationDocument? fromWorker = outer;
            var thread = new Thread(() => fromWorker = OfficeDocumentContext.Find<FoundationDocument>());
            thread.Start(); thread.Join();
            Assert.Null(fromWorker);
        }
        Assert.Null(OfficeDocumentContext.Find<FoundationDocument>());
    }
    [Fact]
    public void OutputPoliciesRefuseClobberAndSupportExplicitReplacement() {
        File.WriteAllText(Output, "original");
        Assert.Throws<IOException>(() => OfficeOutput.Write(Output, _ => { }, noClobber: true));
        Assert.Throws<IOException>(() => OfficeOutput.WriteNew(Output, _ => { }, force: false));
        OfficeOutput.WriteNew(Output, stream => stream.WriteByte(42), force: true);
        Assert.Equal(new byte[] { 42 }, File.ReadAllBytes(Output));
        Assert.Single(Directory.GetFiles(_root));
    }
    [Fact]
    public void ShowUsesAnArgumentInsteadOfShellInterpolation() {
        string path = Path.Combine(_root, "result with 'quotes' and $variables.txt");
        File.WriteAllText(path, "fixture");
        var info = OfficeShow.CreateStartInfo(path);
        if (OperatingSystem.IsWindows()) { Assert.True(info.UseShellExecute); Assert.Equal(path, info.FileName); }
        else { Assert.False(info.UseShellExecute); Assert.Equal(path, Assert.Single(info.ArgumentList)); }
    }
}

[Collection(PowerShellRunspaceCollection.Name)]
public sealed class FoundationRunspaceTests {
    internal static Runspace OpenRunspace() {
        var state = InitialSessionState.CreateDefault();
        foreach (Type type in typeof(FoundationRunspaceTests).Assembly.GetTypes()) {
            var attribute = (CmdletAttribute?)Attribute.GetCustomAttribute(type, typeof(CmdletAttribute));
            if (type.Namespace == typeof(FoundationRunspaceTests).Namespace && attribute != null)
                state.Commands.Add(new SessionStateCmdletEntry(attribute.VerbName + "-" + attribute.NounName, type, null));
        }
        var runspace = RunspaceFactory.CreateRunspace(state);
        runspace.Open();
        return runspace;
    }
    [Fact]
    public void TransformationsRunInActualParameterBinding() {
        using var runspace = OpenRunspace();
        using var shell = PowerShell.Create();
        shell.Runspace = runspace;
        var result = Assert.Single(shell.AddScript("Test-OfficeFoundationBinding -Width 5cm -Color Accent1 -Theme Dark -Sheet '1' -Bold:$false -NoFreeze -Options @{Title='hello';Count=2} -To Pdf -Password 'test-only'").Invoke());
        Assert.False(shell.HadErrors, string.Join(Environment.NewLine, shell.Streams.Error));
        Assert.Equal(OfficeLength.Centimeters(5), result.Properties["Width"].Value);
        Assert.Equal(false, result.Properties["Bold"].Value);
        Assert.Equal(false, result.Properties["Freeze"].Value);
        Assert.Equal(9, result.Properties["PasswordLength"].Value);
        Assert.Single(shell.Streams.Warning);
    }
    [Fact]
    public void MissingDocumentTerminatesAndContentHasNoWhatIfParameter() {
        using var runspace = OpenRunspace();
        using var shell = PowerShell.Create();
        shell.Runspace = runspace;
        shell.AddScript("Add-OfficeFoundationText 'test' -ErrorAction Stop");
        var output = shell.Invoke();
        Assert.Empty(output);
        var error = Assert.Single(shell.Streams.Error);
        Assert.StartsWith("document-required,", error.FullyQualifiedErrorId);
        Assert.Equal(ErrorCategory.InvalidArgument, error.CategoryInfo.Category);
        shell.Commands.Clear(); shell.Streams.Error.Clear();
        var result = Assert.Single(shell.AddScript("(Get-Command Add-OfficeFoundationText).Parameters.ContainsKey('WhatIf')").Invoke());
        Assert.Equal(false, result.BaseObject);
    }
    [Fact]
    public void StopProcessingCancelsBackgroundWork() {
        string root = Path.Combine(Path.GetTempPath(), "foundation-stop-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            string path = Path.Combine(root, "cancel.txt"); File.WriteAllText(path, "fixture");
            using var runspace = OpenRunspace();
            using var shell = PowerShell.Create(); shell.Runspace = runspace;
            shell.AddCommand("Test-OfficeFoundationBatch").AddParameter("LiteralPath", path);
            FoundationBatchControl.Started.Reset(); FoundationBatchControl.Cancelled.Reset();
            var running = shell.BeginInvoke();
            Assert.True(FoundationBatchControl.Started.Wait(TimeSpan.FromSeconds(10)));
            shell.Stop();
            Assert.Throws<PipelineStoppedException>(() => shell.EndInvoke(running));
            Assert.True(FoundationBatchControl.Cancelled.Wait(TimeSpan.FromSeconds(10)));
        } finally { Directory.Delete(root, true); }
    }
}
