using System.Collections;
using System.Collections.Concurrent;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Runtime.CompilerServices;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace OfficeIMO.Notebook.PowerShell;

/// <summary>Starts runspace sessions without loading PowerShell into hosts that only use process sessions.</summary>
internal static class PowerShellNotebookRunspaceHost {
    internal static IPowerShellNotebookWorker Start(string? workingDirectory) {
        try {
            return StartRunspace(workingDirectory);
        } catch (FileNotFoundException error) when (error.FileName?.StartsWith("System.Management.Automation", StringComparison.OrdinalIgnoreCase) == true) {
            throw new PlatformNotSupportedException("A runspace notebook session needs a PowerShell host. Use process isolation from other applications.", error);
        }
    }

    [MethodImpl(MethodImplOptions.NoInlining)]
    private static IPowerShellNotebookWorker StartRunspace(string? workingDirectory) => PowerShellNotebookRunspace.Start(workingDirectory);
}

/// <summary>Runs cells in a runspace of the calling PowerShell process with the output rules of the process worker.</summary>
internal sealed class PowerShellNotebookRunspace : IPowerShellNotebookWorker {
    private static readonly Regex LineBreak = new("\r?\n", RegexOptions.CultureInvariant);
    private static readonly JsonWriterOptions TextWriterOptions = new() { Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping };
    private readonly Runspace _runspace;
    private readonly ConcurrentQueue<PowerShellNotebookFrame> _frames = new();
    private readonly SemaphoreSlim _available = new(0);
    private System.Management.Automation.PowerShell? _pipeline;
    private string? _closeReason;
    private int _closed;

    private PowerShellNotebookRunspace(Runspace runspace, string version) {
        _runspace = runspace;
        Version = version;
    }

    public bool IsOpen => Volatile.Read(ref _closed) == 0;

    public string Version { get; }

    internal static PowerShellNotebookRunspace Start(string? workingDirectory) {
        Runspace runspace = RunspaceFactory.CreateRunspace();
        try {
            runspace.ThreadOptions = PSThreadOptions.ReuseThread;
            runspace.Open();
            runspace.SessionStateProxy.SetVariable("ProgressPreference", ActionPreference.SilentlyContinue);
            if (workingDirectory != null) runspace.SessionStateProxy.Path.SetLocation(WildcardPattern.Escape(workingDirectory));
            string version = (runspace.SessionStateProxy.GetVariable("PSVersionTable") as IDictionary)?["PSVersion"]?.ToString()
                ?? throw new InvalidOperationException("The PowerShell runspace did not report its version.");
            return new PowerShellNotebookRunspace(runspace, version);
        } catch {
            runspace.Dispose();
            throw;
        }
    }

    public Task SendAsync(string source, int executionCount, int outputCharactersLimit, int outputRecordsLimit, CancellationToken cancellationToken) {
        if (!IsOpen) throw new PowerShellNotebookConnectionException("The PowerShell runspace closed its notebook session.");
        if (string.IsNullOrWhiteSpace(source)) {
            Publish(PowerShellNotebookFrame.Completed(false, false));
            return Task.CompletedTask;
        }

        var cell = new CellState(this, executionCount, outputCharactersLimit, outputRecordsLimit);
        System.Management.Automation.PowerShell pipeline = System.Management.Automation.PowerShell.Create();
        var output = new PSDataCollection<PSObject>();
        try {
            pipeline.Runspace = _runspace;
            pipeline.AddScript(source, false);
            pipeline.Commands.Commands[0].MergeMyResults(PipelineResultTypes.All, PipelineResultTypes.Output);
            // Records convert on the pipeline thread, where the cell's runspace is the default runspace.
            output.DataAdded += (_, _) => cell.Drain(output);
            pipeline.Streams.Progress.DataAdded += (_, _) => pipeline.Streams.Progress.ReadAll();
            _pipeline = pipeline;
            pipeline.BeginInvoke<PSObject, PSObject>(null, output, null, result => Complete(cell, pipeline, output, result), null);
        } catch (Exception error) {
            _pipeline = null;
            pipeline.Dispose();
            output.Dispose();
            if (error is InvalidRunspaceStateException or PSInvalidOperationException or InvalidOperationException) {
                throw new PowerShellNotebookConnectionException("The PowerShell runspace could not run the notebook cell.", error);
            }
            throw;
        }
        return Task.CompletedTask;
    }

    public async Task<PowerShellNotebookFrame> ReadAsync() {
        await _available.WaitAsync().ConfigureAwait(false);
        if (!IsOpen || !_frames.TryDequeue(out PowerShellNotebookFrame frame)) {
            throw new PowerShellNotebookConnectionException(Volatile.Read(ref _closeReason) ?? "The PowerShell runspace closed its notebook session.");
        }
        return frame;
    }

    public CancellationTokenRegistration InterruptWhenCanceled(CancellationToken token) =>
        token.Register(static worker => ((PowerShellNotebookRunspace)worker!).Dispose(), this);

    public void Dispose() {
        if (Interlocked.Exchange(ref _closed, 1) != 0) return;
        // Wake a pending read; the closed session reports the interruption.
        _available.Release();
        System.Management.Automation.PowerShell? pipeline = Interlocked.Exchange(ref _pipeline, null);
        if (pipeline == null) {
            _runspace.Dispose();
            return;
        }
        // Stopping waits for the cell to reach a stoppable point, so do not block the caller.
        _ = Task.Run(() => {
            try {
                pipeline.Stop();
            } catch (Exception error) when (error is InvalidOperationException or ObjectDisposedException or PSInvalidOperationException) {
                // The cell can finish while the session closes.
            } finally {
                _runspace.Dispose();
            }
        });
    }

    private void Complete(CellState cell, System.Management.Automation.PowerShell pipeline, PSDataCollection<PSObject> output, IAsyncResult result) {
        PowerShellNotebookFrame completion = default;
        string? failure = null;
        try {
            cell.Drain(output);
            try {
                pipeline.EndInvoke(result);
            } catch {
                Exception? reason = pipeline.InvocationStateInfo.Reason;
                if (reason == null) throw;
                object record = reason is RuntimeException { ErrorRecord: not null } runtime ? runtime.ErrorRecord : reason;
                cell.Write(record, terminatingError: true);
            }
            completion = cell.Finish(pipeline.HadErrors);
        } catch (Exception error) {
            failure = "The PowerShell runspace could not finish the notebook cell. " + error.Message;
        } finally {
            Interlocked.CompareExchange(ref _pipeline, null, pipeline);
            pipeline.Dispose();
            output.Dispose();
        }
        // Report completion only after the pipeline is released, so the next cell finds the runspace available.
        if (failure != null) {
            Volatile.Write(ref _closeReason, failure);
            Dispose();
        } else {
            Publish(completion);
        }
    }

    private void Publish(PowerShellNotebookFrame frame) {
        if (!IsOpen) return;
        _frames.Enqueue(frame);
        _available.Release();
    }

    /// <summary>Applies the record and character budgets of one cell in output order.</summary>
    private sealed class CellState {
        private readonly PowerShellNotebookRunspace _owner;
        private readonly int _executionCount;
        private readonly object _gate = new();
        private int _charactersRemaining;
        private int _recordsRemaining;
        private bool _failed;
        private bool _truncated;
        private bool _truncationReported;
        private object? _lastError;

        internal CellState(PowerShellNotebookRunspace owner, int executionCount, int characters, int records) {
            _owner = owner;
            _executionCount = executionCount;
            _charactersRemaining = characters;
            _recordsRemaining = records;
        }

        internal void Drain(PSDataCollection<PSObject> output) {
            lock (_gate) {
                // Drain even after truncation so a noisy pipeline can finish and keep its session state.
                foreach (PSObject? record in output.ReadAll()) {
                    object? value = record?.BaseObject;
                    if (_recordsRemaining == 0 || _charactersRemaining == 0) {
                        _truncated = true;
                        if (_failed) break;
                        // Merged PowerShell errors can leave HadErrors unset even when the cell failed.
                        if (value is not ErrorRecord) continue;
                    }
                    WriteRecord(record, value, terminatingError: false);
                }
                ReportTruncation();
            }
        }

        internal void Write(object record, bool terminatingError) {
            lock (_gate) {
                WriteRecord(record, record, terminatingError);
            }
        }

        internal PowerShellNotebookFrame Finish(bool hadErrors) {
            lock (_gate) {
                return PowerShellNotebookFrame.Completed(_failed || hadErrors, _truncated);
            }
        }

        private void ReportTruncation() {
            if (_truncated && !_truncationReported) {
                _truncationReported = true;
                _owner.Publish(PowerShellNotebookFrame.TruncationNotice);
            }
        }

        private void WriteRecord(object? record, object? value, bool terminatingError) {
            var errorRecord = value as ErrorRecord;
            bool nativeStandardError = errorRecord != null && !terminatingError &&
                (string.Equals(errorRecord.FullyQualifiedErrorId, "NativeCommandError", StringComparison.OrdinalIgnoreCase) ||
                 string.Equals(errorRecord.FullyQualifiedErrorId, "NativeCommandErrorMessage", StringComparison.OrdinalIgnoreCase));
            bool error = terminatingError || (errorRecord != null && !nativeStandardError);
            if (error) {
                _failed = true;
                if (terminatingError && ReferenceEquals(_lastError, value)) return;
                _lastError = value;
            }
            if (_recordsRemaining == 0 || _charactersRemaining == 0) {
                _truncated = true;
                return;
            }

            NotebookOutput output;
            if (error) {
                string text = Limit(value is Exception exception ? exception.Message : ToText(record));
                output = errorRecord != null
                    ? NotebookOutput.CreateError(errorRecord.Exception.GetType().Name, text, Traceback(errorRecord.ScriptStackTrace))
                    : NotebookOutput.CreateError(value!.GetType().Name, text, Array.Empty<string>());
            } else if (value is InformationRecord information) {
                string text = ToText(information.MessageData);
                if (information.MessageData is not HostInformationMessage { NoNewLine: true }) text += "\n";
                output = NotebookOutput.CreateStream("stdout", Limit(text));
            } else if (nativeStandardError || value is InformationalRecord) {
                string text = nativeStandardError ? ToText(record) : ((InformationalRecord)value!).Message;
                output = NotebookOutput.CreateStream("stderr", Limit(text + "\n"));
            } else {
                string text;
                try {
                    text = ToText(record);
                } catch (Exception conversion) when (conversion is not OutOfMemoryException) {
                    text = "[PowerShell could not render this result as text.]";
                    _truncated = true;
                }
                output = NotebookOutput.CreateExecuteResult(new Dictionary<string, JsonElement> { ["text/plain"] = TextElement(Limit(text)) }, _executionCount);
            }
            _recordsRemaining--;
            _owner.Publish(PowerShellNotebookFrame.ForOutput(output));
        }

        private string[] Traceback(string? stackTrace) {
            var traceback = new List<string>();
            if (stackTrace == null) return traceback.ToArray();
            foreach (string frame in LineBreak.Split(stackTrace, 17)) {
                if (frame.Length == 0) continue;
                if (traceback.Count == 16 || _charactersRemaining == 0) {
                    _truncated = true;
                    break;
                }
                traceback.Add(Limit(frame));
            }
            return traceback.ToArray();
        }

        private string Limit(string text) {
            if (text.Length <= _charactersRemaining) {
                _charactersRemaining -= text.Length;
                return text;
            }
            int length = _charactersRemaining;
            if (length > 0 && char.IsHighSurrogate(text[length - 1])) length--;
            _charactersRemaining = 0;
            _truncated = true;
            return text.Substring(0, length);
        }

        // PowerShell's [string] conversion, including ETS ToString overrides and array joining.
        private static string ToText(object? value) => value == null ? string.Empty : (string)LanguagePrimitives.ConvertTo(value, typeof(string));

        private static JsonElement TextElement(string text) {
            using var buffer = new MemoryStream(text.Length + 2);
            using (var writer = new Utf8JsonWriter(buffer, TextWriterOptions)) {
                writer.WriteStringValue(text);
            }
            using JsonDocument document = JsonDocument.Parse(buffer.ToArray());
            return document.RootElement.Clone();
        }
    }
}
