# OfficeIMO.Notebook.PowerShell

Run PowerShell cells explicitly and save their results into a Jupyter notebook.
By default a session starts an installed PowerShell executable. It defaults to
`pwsh` on PATH and accepts a full executable path, including Windows PowerShell 5.1.
Code that already runs inside PowerShell, such as a module cmdlet, can set
`Isolation = PowerShellNotebookIsolation.Runspace` to run the cells in a new
runspace of the same process. That skips starting a second PowerShell and the
message exchange per cell. Runspace cells share the process environment,
current directory and loaded assemblies with the caller. Stopping a runspace
cell stops its pipeline but not programs it started in the background, and a
cell blocked in a call that PowerShell cannot interrupt keeps its thread until
the call returns. Choose process isolation for those cells.

Only run notebooks you trust. Cells run with your account's permissions.
Importing, opening, previewing, and exporting saved notebooks remain
independent operations.

Add a reference from your application to the PowerShell execution project in
your OfficeIMO checkout.

```sh
dotnet add ./MyApp/MyApp.csproj reference ./OfficeIMO/OfficeIMO.Notebook.PowerShell/OfficeIMO.Notebook.PowerShell.csproj
```

```csharp
using OfficeIMO.Notebook;
using OfficeIMO.Notebook.PowerShell;

using var notebook = NotebookDocument.Load("lab.ipynb");
using var session = await PowerShellNotebookSession.StartAsync(
    new PowerShellNotebookSessionOptions { WorkingDirectory = Environment.CurrentDirectory });
var report = await session.ExecuteAsync(notebook, new NotebookExecutionOptions {
    CellTimeout = TimeSpan.FromSeconds(30),
    OutputReceived = update => Console.WriteLine($"Cell {update.CellIndex} produced {update.Output.Kind}")
});
notebook.Save("lab-executed.ipynb");
Console.WriteLine(report.Succeeded);
```

Set `CellIndices` to zero-based indices to select cells. Selected cells run in
notebook order. An explicit cell language takes precedence over notebook
language metadata. When no language is declared, the kernel name identifies the
notebook language. Enable
`IncludeUnspecifiedLanguage` to run code cells with no language metadata.

One session keeps variables, functions, and working location across cells and
across calls to `ExecuteAsync`. The same notebook must be edited or saved
after the run finishes. Each attempted cell replaces its previous outputs and
receives a session execution count. Pipeline objects are saved as text/plain
execution results. PowerShell host and diagnostic streams and structured
PowerShell errors are saved in Jupyter output records. The per-cell record and
character limits report truncation through `OutputTruncated`. The host also
records an interruption error when a cell is stopped.

The default cell deadline is 60 seconds. Set `CellTimeout` to `null` for an
unlimited cell. A PowerShell error stops the run by default. Enable
`ContinueOnError` to attempt later cells after those errors. A timeout ends the
session and returns a timed-out result. Cancelling an active cell interrupts the session and
throws `OperationCanceledException`. Both preserve outputs received before the
interruption and record an error on the interrupted cell. Start a new session
to execute more cells after an interruption.

PSWriteOffice exposes the same execution engine through `Invoke-OfficeNotebook`.
The cmdlet runs cells in a runspace of the current PowerShell by default;
`-Isolation Process` or `-PowerShellPath` starts a separate process.
`-OutputPath` saves a copy of the executed notebook.
