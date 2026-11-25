namespace OfficeIMO.Notebook.PowerShell;

/// <summary>Configures an installed PowerShell host for a notebook session.</summary>
public sealed class PowerShellNotebookSessionOptions {
    /// <summary>Gets or sets the PowerShell executable. The default resolves <c>pwsh</c> on PATH.</summary>
    public string? ExecutablePath { get; set; }

    /// <summary>Gets or sets the initial working directory. The default inherits the caller's working directory.</summary>
    public string? WorkingDirectory { get; set; }

    /// <summary>Gets or sets the deadline for launching PowerShell and opening the runspace.</summary>
    public TimeSpan StartupTimeout { get; set; } = TimeSpan.FromSeconds(30);
}
