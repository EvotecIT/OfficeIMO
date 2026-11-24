using System.Diagnostics;
using System.Globalization;
using System.Text;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Notebook.Import;

/// <summary>Projects PowerShell and Markdown notebook cells into a static PowerShell script.</summary>
/// <remarks>Import the generated script with <see cref="NotebookImportFormat.PowerShellScript"/> to recover the supported cells.
/// Conversion reads saved source and never invokes PowerShell.</remarks>
public static partial class NotebookPowerShellScript {
    private const string TargetFormat = "PowerShell script";
    private const int PreambleLength = 3;
    private const int MaximumDiagnosticDetails = 4_096;
    private static readonly int MaximumScriptBytes = (int)NotebookLoadOptions.UntrustedDefaults.MaxInputBytes;
    private static readonly int MaximumScriptCells = NotebookLoadOptions.UntrustedDefaults.MaxCells;
    private static readonly Encoding ScriptEncoding = new UTF8Encoding(encoderShouldEmitUTF8Identifier: true, throwOnInvalidBytes: true);

    /// <summary>Returns the script text and any omissions or approximations.</summary>
    public static OfficeConversionResult<string> Convert(NotebookDocument notebook, bool failOnLoss = false,
        CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(notebook);
        notebook.CheckDisposed();
        cancellationToken.ThrowIfCancellationRequested();
        var timer = Stopwatch.StartNew();
        var diagnostics = new List<OfficeDiagnostic>(notebook.Diagnostics.Take(MaximumDiagnosticDetails));
        if (notebook.Diagnostics.Count > diagnostics.Count) AddDiagnosticLimit(diagnostics);
        if (notebook.HasAdditionalSourceProperties || notebook.Metadata.Any(static entry => !IsNotebookLanguageMetadata(entry))) {
            AddDiagnostic(diagnostics, "notebook-powershell-notebook-metadata-omitted",
                "Notebook metadata and extension properties outside the PowerShell script profile were omitted.", "metadata");
        }

        var cells = new List<ScriptCell>();
        var reservedBoundaries = new HashSet<int>();
        string? defaultLanguage = notebook.Language;
        int sourceCharacters = 0;
        for (int i = 0; i < notebook.Cells.Count; i++) {
            cancellationToken.ThrowIfCancellationRequested();
            NotebookCell cell = notebook.Cells[i];
            if (!CanExport(cell, defaultLanguage, i, diagnostics)) continue;

            if (cells.Count == MaximumScriptCells) {
                throw NotebookParser.Limit(nameof(NotebookLoadOptions.MaxCells), cells.Count);
            }
            ReportCellLosses(cell, i, diagnostics);
            string source = cell.GetSource(MaximumScriptBytes - PreambleLength - sourceCharacters, out bool truncated);
            if (truncated) ThrowOutputLimit();
            sourceCharacters += source.Length;
            if (cell.Kind == NotebookCellKind.Code) ReserveBoundaries(source, reservedBoundaries, cancellationToken);
            cells.Add(new ScriptCell(cell.Kind, cell.Id, source));
        }

        int boundaryNumber = 0;
        while (reservedBoundaries.Contains(boundaryNumber)) boundaryNumber++;
        string marker = CellMarker(boundaryNumber);
        var script = new StringBuilder();
        script.Append(Header).Append(boundaryNumber.ToString(CultureInfo.InvariantCulture)).Append('\n');
        foreach (ScriptCell cell in cells) {
            cancellationToken.ThrowIfCancellationRequested();
            bool markdown = cell.Kind == NotebookCellKind.Markdown;
            script.Append(marker).Append(markdown ? "markdown " : "code ").Append(cell.Id).Append('\n');
            CheckOutputLength(script.Length);
            if (markdown) {
                WriteCommentedMarkdown(script, cell.Source, cancellationToken);
            } else {
                CheckOutputLength((long)script.Length + cell.Source.Length);
                script.Append(cell.Source);
            }
            script.Append('\n');
            CheckOutputLength(script.Length);
        }

        string content = script.ToString();
        CheckOutputLength(ScriptEncoding.GetByteCount(content));
        cancellationToken.ThrowIfCancellationRequested();
        var report = new OfficeConversionReport("Jupyter Notebook", TargetFormat, itemsWritten: cells.Count,
            elapsed: timer.Elapsed, diagnostics: diagnostics);
        if (failOnLoss) report.ThrowIfLoss();
        return new OfficeConversionResult<string>(content, report);
    }

    /// <summary>Writes a UTF-8 PowerShell script with a byte-order mark and publishes it atomically.</summary>
    public static OfficeConversionReport Save(NotebookDocument notebook, string path, bool failOnLoss = false,
        CancellationToken cancellationToken = default) {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        string fullPath = Path.GetFullPath(path);
        OfficeConversionResult<string> converted = Convert(notebook, failOnLoss, cancellationToken);
        byte[] script = Encode(converted.Value);
        OfficeConversionFailures.Run(converted.Report,
            () => OfficeFileCommit.WriteAtomically(fullPath, stream => stream.Write(script, 0, script.Length), cancellationToken));
        return OfficeConversionOutput.WithPath(converted.Report, fullPath);
    }

    /// <summary>Writes a UTF-8 PowerShell script with a byte-order mark and leaves the stream open.</summary>
    /// <remarks>A seekable destination is replaced and rewound.</remarks>
    public static OfficeConversionReport Save(NotebookDocument notebook, Stream stream, bool failOnLoss = false,
        CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(stream);
        OfficeConversionResult<string> converted = Convert(notebook, failOnLoss, cancellationToken);
        byte[] script = Encode(converted.Value);
        cancellationToken.ThrowIfCancellationRequested();
        return OfficeConversionOutput.Save(converted.Report, stream, script);
    }

    /// <summary>Asynchronously writes and publishes a UTF-8 PowerShell script with a byte-order mark.</summary>
    public static async Task<OfficeConversionReport> SaveAsync(NotebookDocument notebook, string path, bool failOnLoss = false,
        CancellationToken cancellationToken = default) {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        string fullPath = Path.GetFullPath(path);
        OfficeConversionResult<string> converted = Convert(notebook, failOnLoss, cancellationToken);
        byte[] script = Encode(converted.Value);
        await OfficeConversionFailures.RunAsync(converted.Report,
            () => OfficeFileCommit.WriteAtomicallyAsync(fullPath,
                (stream, token) => stream.WriteAsync(script, 0, script.Length, token), cancellationToken)).ConfigureAwait(false);
        return OfficeConversionOutput.WithPath(converted.Report, fullPath);
    }

    /// <summary>Asynchronously writes a UTF-8 PowerShell script with a byte-order mark and leaves the stream open.</summary>
    /// <remarks>A seekable destination is replaced and rewound.</remarks>
    public static async Task<OfficeConversionReport> SaveAsync(NotebookDocument notebook, Stream stream, bool failOnLoss = false,
        CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(stream);
        OfficeConversionResult<string> converted = Convert(notebook, failOnLoss, cancellationToken);
        byte[] script = Encode(converted.Value);
        cancellationToken.ThrowIfCancellationRequested();
        return await OfficeConversionOutput.SaveAsync(converted.Report, stream, script, cancellationToken).ConfigureAwait(false);
    }

    private static byte[] Encode(string script) {
        byte[] bytes = new byte[ScriptEncoding.GetByteCount(script) + PreambleLength];
        bytes[0] = 0xEF;
        bytes[1] = 0xBB;
        bytes[2] = 0xBF;
        ScriptEncoding.GetBytes(script, 0, script.Length, bytes, PreambleLength);
        return bytes;
    }

    private static void CheckOutputLength(long count) {
        if (count > MaximumScriptBytes - PreambleLength) ThrowOutputLimit();
    }

    private static void ThrowOutputLimit() =>
        throw new InvalidDataException("The PowerShell script exceeds the 64 MiB output limit.");

    private readonly record struct ScriptCell(NotebookCellKind Kind, string Id, string Source);
}
