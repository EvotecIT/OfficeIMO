using System;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Threading.Tasks;
using PSWriteOffice;

namespace PSWriteOffice.Tests;

[Collection(PowerShellRunspaceCollection.Name)]
public sealed class AsyncPSCmdletTests
{
    [Fact]
    public void AsyncPSCmdlet_drains_worker_thread_writes_when_task_completes_synchronously()
    {
        var sessionState = InitialSessionState.CreateDefault();
        sessionState.Commands.Add(new SessionStateCmdletEntry(
            "Test-AsyncQueuedOutput",
            typeof(TestAsyncQueuedOutputCommand),
            helpFileName: null));

        using var runspace = RunspaceFactory.CreateRunspace(sessionState);
        runspace.Open();
        using var powerShell = PowerShell.Create();
        powerShell.Runspace = runspace;
        powerShell.AddCommand("Test-AsyncQueuedOutput");

        var result = powerShell.Invoke();

        Assert.False(powerShell.HadErrors, string.Join(Environment.NewLine, powerShell.Streams.Error.Select(static error => error.ToString())));
        var item = Assert.Single(result);
        Assert.Equal("queued-output", item.BaseObject);
    }

    [Fact]
    public void AsyncPSCmdlet_routes_helper_writes_through_async_pipeline_interface()
    {
        var sessionState = InitialSessionState.CreateDefault();
        sessionState.Commands.Add(new SessionStateCmdletEntry(
            "Test-AsyncHelperQueuedOutput",
            typeof(TestAsyncHelperQueuedOutputCommand),
            helpFileName: null));

        using var runspace = RunspaceFactory.CreateRunspace(sessionState);
        runspace.Open();
        using var powerShell = PowerShell.Create();
        powerShell.Runspace = runspace;
        powerShell.AddCommand("Test-AsyncHelperQueuedOutput");

        var result = powerShell.Invoke();

        Assert.False(powerShell.HadErrors, string.Join(Environment.NewLine, powerShell.Streams.Error.Select(static error => error.ToString())));
        var item = Assert.Single(result);
        Assert.Equal("helper-output", item.BaseObject);
    }

    [Fact]
    public async Task AsyncPSCmdlet_pumps_should_process_during_worker_fan_out()
    {
        var sessionState = InitialSessionState.CreateDefault();
        sessionState.Commands.Add(new SessionStateCmdletEntry(
            "Test-AsyncSynchronousFanOut",
            typeof(TestAsyncSynchronousFanOutCommand),
            helpFileName: null));

        using var runspace = RunspaceFactory.CreateRunspace(sessionState);
        runspace.Open();
        using var powerShell = PowerShell.Create();
        powerShell.Runspace = runspace;
        powerShell.AddCommand("Test-AsyncSynchronousFanOut");

        var invokeTask = Task.Run(() => powerShell.Invoke());

        var result = await invokeTask.WaitAsync(TimeSpan.FromSeconds(10));
        Assert.False(powerShell.HadErrors, string.Join(Environment.NewLine, powerShell.Streams.Error.Select(static error => error.ToString())));
        var item = Assert.Single(result);
        Assert.Equal("fan-out-output", item.BaseObject);
    }
}

[Cmdlet(VerbsDiagnostic.Test, "AsyncQueuedOutput")]
public sealed class TestAsyncQueuedOutputCommand : AsyncPSCmdlet
{
    protected override Task ProcessRecordAsync()
        => Task.Run(() => WriteObject("queued-output"));
}

[Cmdlet(VerbsDiagnostic.Test, "AsyncHelperQueuedOutput")]
public sealed class TestAsyncHelperQueuedOutputCommand : AsyncPSCmdlet
{
    protected override Task ProcessRecordAsync()
        => Task.Run(() => AsyncPipelineHelper.WriteOutput(this));
}

internal static class AsyncPipelineHelper
{
    public static void WriteOutput(IAsyncCmdletPipeline pipeline)
        => pipeline.WriteObject("helper-output");
}

[Cmdlet(VerbsDiagnostic.Test, "AsyncSynchronousFanOut", SupportsShouldProcess = true)]
public sealed class TestAsyncSynchronousFanOutCommand : AsyncPSCmdlet
{
    protected override async Task ProcessRecordAsync()
    {
        await Task.Run(() =>
        {
            if (ShouldProcess("fan-out-target", "write output"))
            {
                WriteObject("fan-out-output");
            }
        });
    }
}
