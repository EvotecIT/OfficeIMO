namespace OfficeIMO.Notebook.PowerShell;

/// <summary>Chooses where a notebook session runs its PowerShell cells.</summary>
public enum PowerShellNotebookIsolation {
    /// <summary>Run cells in a separate PowerShell process. A timeout or cancellation ends that process.</summary>
    Process,

    /// <summary>Run cells in a new runspace inside the calling PowerShell process.</summary>
    /// <remarks>Requires a PowerShell host, such as a module cmdlet. Cells share the process environment, current
    /// directory and loaded assemblies with the caller. A timeout or cancellation stops the running pipeline. A cell
    /// blocked in a call that PowerShell cannot stop keeps its thread until that call returns.</remarks>
    Runspace
}
