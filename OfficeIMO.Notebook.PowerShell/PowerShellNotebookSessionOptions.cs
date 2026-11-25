namespace OfficeIMO.Notebook.PowerShell;

/// <summary>Configures where and how a notebook session runs PowerShell.</summary>
public sealed class PowerShellNotebookSessionOptions {
    /// <summary>Gets or sets where cells run. The default starts a separate PowerShell process.</summary>
    public PowerShellNotebookIsolation Isolation { get; set; }

    /// <summary>Gets or sets the PowerShell executable for process isolation. The default resolves <c>pwsh</c> on PATH.</summary>
    public string? ExecutablePath { get; set; }

    /// <summary>Gets or sets the initial working directory. The default inherits the caller's working directory.</summary>
    public string? WorkingDirectory { get; set; }

    /// <summary>Gets or sets the deadline for launching a PowerShell process and opening its runspace.</summary>
    public TimeSpan StartupTimeout { get; set; } = TimeSpan.FromSeconds(30);
}
