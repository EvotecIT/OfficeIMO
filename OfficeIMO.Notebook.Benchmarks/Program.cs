using System.Text;
using System.Text.Json;
using System.Runtime.InteropServices;
using OfficeIMO.Markdown;
using OfficeIMO.Notebook.Html;
using OfficeIMO.Notebook.Markdown;
using OfficeIMO.Notebook.Pdf;
#if NOTEBOOK_COMPETITORS
using Microsoft.DotNet.Interactive.Documents.Jupyter;
using JupyterNotebook = Microsoft.DotNet.Interactive.Documents.Jupyter.Notebook;
using Verso.Serializers;
#endif

namespace OfficeIMO.Notebook.Benchmarks;

internal sealed class NotebookBenchmarkWorker : IDisposable {
    private static readonly JsonSerializerOptions JsonOptions = new(JsonSerializerDefaults.Web);

    private string _engine = string.Empty;
    private string _operation = string.Empty;
    private byte[] _input = Array.Empty<byte>();
    private string _artifact = string.Empty;
    private NotebookDocument? _document;
    private object? _result;
    private OfficeConversionReport? _report;
    private long _allocatedBytes;

    internal static int Main(string[] args) {
        if (args is ["environment"]) {
            Respond(new {
                ok = true,
                runtime = RuntimeInformation.FrameworkDescription,
#if NOTEBOOK_COMPETITORS
                microsoft = "1.0.0-beta.26120.1",
                verso = "1.2.3",
#endif
            });
            return 0;
        }
        if (args is not ["worker"]) return 2;

        using var worker = new NotebookBenchmarkWorker();
        string? line;
        while ((line = Console.ReadLine()) != null) {
            try {
                using JsonDocument message = JsonDocument.Parse(line);
                JsonElement request = message.RootElement;
                switch (request.GetProperty("action").GetString()) {
                    case "prepare":
                        worker.Prepare(request);
                        Respond(new { ok = true });
                        break;
                    case "run":
                        worker.Run();
                        Respond(new { ok = true });
                        break;
                    case "finish":
                        Respond(worker.Finish());
                        break;
                    default:
                        throw new InvalidOperationException("Unknown benchmark action.");
                }
            } catch (Exception error) {
                Respond(new { ok = false, error = error.ToString() });
            }
        }
        return 0;
    }

    public void Dispose() => _document?.Dispose();

    private void Prepare(JsonElement request) {
        _document?.Dispose();
        _document = null;
        _report = null;
        _result = null;
        _engine = request.GetProperty("engine").GetString()!;
        _operation = request.GetProperty("operation").GetString()!;
        _artifact = request.GetProperty("artifact").GetString()!;
        _input = File.ReadAllBytes(request.GetProperty("input").GetString()!);
        if (_engine == "OfficeIMO" && _operation is "Html" or "Markdown") {
            _document = NotebookDocument.Load(_input);
        }
    }

    private void Run() {
        long before = GC.GetTotalAllocatedBytes(precise: true);
        _result = _engine switch {
            "OfficeIMO" => RunOfficeImo(),
#if NOTEBOOK_COMPETITORS
            "Microsoft" => RunMicrosoft(),
            "Verso" => RunVerso(),
#endif
            _ => throw new InvalidOperationException($"Requested engine '{_engine}' was not built.")
        };
        _allocatedBytes = GC.GetTotalAllocatedBytes(precise: true) - before;
    }

    private object RunOfficeImo() {
        if (_operation == "Parse") {
            _document = NotebookDocument.Load(_input);
            return _document;
        }
        if (_operation == "Html") {
            string html = _document!.ToHtml(null, out OfficeConversionReport report);
            _report = report;
            return Encoding.UTF8.GetBytes(html);
        }
        if (_operation == "Markdown") {
            using MarkdownDocument projected = _document!.ToMarkdown(null, out OfficeConversionReport report);
            _report = report;
            return Encoding.UTF8.GetBytes(projected.ToMarkdown());
        }

        using NotebookDocument notebook = NotebookDocument.Load(_input);
        switch (_operation) {
            case "SourceProjection":
                return notebook.Cells.Select(static cell => new[] { cell.OriginalKind, cell.Source }).ToArray();
            case "RoundTrip":
            case "RichRoundTrip":
                return notebook.ToBytes(new NotebookSaveOptions { AllowLoss = _operation == "RichRoundTrip" });
            case "Pdf":
                using (var result = new MemoryStream()) {
                    _report = notebook.SaveAsPdf(result);
                    return result.ToArray();
                }
            default:
                throw new InvalidOperationException($"Unknown operation: {_operation}");
        }
    }

#if NOTEBOOK_COMPETITORS
    private object RunMicrosoft() {
        var notebook = JupyterNotebook.Parse(Encoding.UTF8.GetString(_input));
        if (_operation == "SourceProjection") {
            return notebook.Elements.Select(static cell => new[] {
                cell.KernelName == "markdown" ? "markdown" : "code", cell.Contents
            }).ToArray();
        }
        return Encoding.UTF8.GetBytes(notebook.ToJupyterJson());
    }

    private object RunVerso() {
        var serializer = new JupyterSerializer();
        var notebook = serializer.DeserializeAsync(Encoding.UTF8.GetString(_input)).GetAwaiter().GetResult();
        if (_operation == "SourceProjection") {
            return notebook.Cells.Select(static cell => new[] { cell.Type, cell.Source }).ToArray();
        }
        return Encoding.UTF8.GetBytes(serializer.SerializeAsync(notebook).GetAwaiter().GetResult());
    }
#endif

    private object Finish() {
        byte[] data = _result switch {
            byte[] bytes => bytes,
            NotebookDocument notebook => notebook.ToBytes(),
            string[][] projection => JsonSerializer.SerializeToUtf8Bytes(projection),
            _ => throw new InvalidOperationException("The operation did not produce a result.")
        };
        File.WriteAllBytes(_artifact, data);
        if (_report is not null) {
            string path = Path.Combine(Path.GetDirectoryName(_artifact)!, "conversion-report.json");
            File.WriteAllBytes(path, JsonSerializer.SerializeToUtf8Bytes(
                _report.Diagnostics.Select(static diagnostic => new {
                    diagnostic.Code, diagnostic.Message, diagnostic.LossKind, diagnostic.Severity
                }), JsonOptions));
        }
        return new { ok = true, outputBytes = data.Length, assetCount = 0, allocatedBytes = _allocatedBytes };
    }

    private static void Respond<T>(T response) {
        Console.WriteLine(JsonSerializer.Serialize(response, JsonOptions));
        Console.Out.Flush();
    }
}
