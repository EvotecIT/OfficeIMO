using System.Management.Automation;

namespace PSWriteOffice;

/// <summary>
/// Provides async-safe access to PowerShell pipeline operations.
/// </summary>
public interface IAsyncCmdletPipeline {
    /// <summary>Writes an object to the PowerShell output pipeline.</summary>
    void WriteObject(object? sendToPipeline);

    /// <summary>Writes an object to the PowerShell output pipeline.</summary>
    void WriteObject(object? sendToPipeline, bool enumerateCollection);

    /// <summary>Writes an error record to the PowerShell error stream.</summary>
    void WriteError(ErrorRecord errorRecord);

    /// <summary>Reports a terminating error on the PowerShell pipeline thread.</summary>
    void ThrowTerminatingError(ErrorRecord errorRecord);

    /// <summary>Writes a warning message to the PowerShell warning stream.</summary>
    void WriteWarning(string message);

    /// <summary>Writes a verbose message to the PowerShell verbose stream.</summary>
    void WriteVerbose(string message);

    /// <summary>Writes a debug message to the PowerShell debug stream.</summary>
    void WriteDebug(string message);

    /// <summary>Writes an information record to the PowerShell information stream.</summary>
    void WriteInformation(InformationRecord informationRecord);

    /// <summary>Writes tagged information to the PowerShell information stream.</summary>
    void WriteInformation(object messageData, string[] tags);

    /// <summary>Writes a progress record to the PowerShell progress stream.</summary>
    void WriteProgress(ProgressRecord progressRecord);

    /// <summary>Runs ShouldProcess on the PowerShell pipeline thread.</summary>
    bool ShouldProcess(string? target, string action);

    /// <summary>Runs ShouldContinue on the PowerShell pipeline thread.</summary>
    bool ShouldContinue(string query, string caption);

    /// <summary>Prompts for credentials on the PowerShell pipeline thread.</summary>
    PSCredential? PromptForCredential(string caption, string message, string userName, string targetName);
}
