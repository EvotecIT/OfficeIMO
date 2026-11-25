using System.Management.Automation;
using OfficeIMO.Core.Internal;

namespace PSWriteOffice.Cmdlets.Notebook;

internal static class NotebookCmdletPaths {
    internal static void RequireJupyterDestination(string path) {
        if (!string.Equals(Path.GetExtension(path), ".ipynb", StringComparison.OrdinalIgnoreCase))
            throw new PSArgumentException("The native notebook destination must end in .ipynb. Use Export-OfficeNotebook for other formats.");
    }

    internal static OfficeFileCommit.ConflictPolicy ConflictPolicy(bool noClobber) => noClobber
        ? OfficeFileCommit.ConflictPolicy.FailIfExists
        : OfficeFileCommit.ConflictPolicy.Replace;
}
