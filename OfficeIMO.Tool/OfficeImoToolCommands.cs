using OfficeIMO.Tool.Commands.Agent;
using OfficeIMO.Tool.Commands.Convert;
using OfficeIMO.Tool.Commands.HiddenData;
using OfficeIMO.Tool.Commands.Html;
using OfficeIMO.Tool.Commands.Invoice;
using OfficeIMO.Tool.Commands.Markup;
using OfficeIMO.Tool.Commands.Mcp;
using OfficeIMO.Tool.Commands.Pdf;
using OfficeIMO.Tool.Commands.Provenance;
using OfficeIMO.Tool.Commands.Reader;
using OfficeIMO.Tool.Commands.Tabular;
using OfficeIMO.Tool.Commands.Workflow;

namespace OfficeIMO.Tool;

internal delegate Task<int> OfficeImoToolCommandHandler(
    string[] args,
    Stream standardInput,
    Stream standardOutput,
    TextWriter standardError,
    CancellationToken cancellationToken);

/// <summary>One top-level command area: how it is dispatched and how it is described in help and the command catalog.</summary>
internal sealed class OfficeImoToolCommandGroup {
    internal required string Id { get; init; }
    internal required string Title { get; init; }
    internal required string Summary { get; init; }
    /// <summary>Arguments shown after the command name in the top-level usage.</summary>
    internal required string Synopsis { get; init; }
    internal required OfficeImoToolCommandHandler Run { get; init; }
    internal IReadOnlyList<string> Aliases { get; init; } = [];
    /// <summary>The namespaced command this shortcut runs, such as <c>reader read</c>.</summary>
    internal string? ShortcutFor { get; init; }
    /// <summary>The help text the area prints for <c>--help</c>; usage lines and subcommands are read from it.</summary>
    internal IReadOnlyList<string> Help { get; init; } = [];
    /// <summary>One-line summaries keyed by the subcommand name as it appears in the help usage lines.</summary>
    internal IReadOnlyDictionary<string, string> Subcommands { get; init; } = new Dictionary<string, string>();
    /// <summary>Options worth calling out, keyed by option name; each must appear in the area's help text.</summary>
    internal IReadOnlyList<OfficeImoToolKeyOption> KeyOptions { get; init; } = [];
    /// <summary>File extensions, without the dot, that the area reads; computed on demand from the owning catalog where one exists.</summary>
    internal Func<IEnumerable<string>> InputFormats { get; init; } = static () => [];
    /// <summary>File extensions, without the dot, that the area writes.</summary>
    internal Func<IEnumerable<string>> OutputFormats { get; init; } = static () => [];
    /// <summary>Formats written to standard output: json, ndjson, sarif, text, markdown or path.</summary>
    internal IReadOnlyList<string> ResultFormats { get; init; } = [];
}

internal sealed record OfficeImoToolKeyOption(string Name, string Description, string? Value = null);

internal static class OfficeImoToolCommands {
    // Removal preserves the input format.
    private static readonly string[] HiddenDataFormats = ["docx", "xlsx", "pptx", "vsdx", "odt", "ods", "odp", "epub", "pdf", "jpeg", "png", "webp", "gif", "tiff"];

    internal static IReadOnlyList<OfficeImoToolCommandGroup> Groups { get; } = [
        new() {
            Id = "convert",
            Title = "Convert",
            Summary = "Convert one file; the input and output extensions select a route from the shared conversion catalog.",
            Synopsis = "<input> [output.pdf|output.docx|output.xlsx|output.pptx|output.md|output.json] [options]",
            Run = static (args, input, output, error, token) => ConvertCommand.RunAsync(args, input, output, error, token),
            Help = [ConvertCommand.Usage],
            KeyOptions = [
                new("--output", "Destination file; the default is a sibling PDF."),
                new("--force", "Replace an existing destination."),
                new("--require-no-loss", "Fail instead of publishing when the route reports source or output fidelity loss."),
                new("--max-input-bytes", "Refuse inputs larger than this many bytes."),
                new("--max-output-bytes", "Refuse to publish outputs larger than this many bytes."),
                new("--assets", "Directory for images extracted during Markdown or JSON output."),
                new("--select", "Path to the rows to convert in a structured-data source, such as data.items."),
                new("--iwork-mode", "How Pages, Numbers and Keynote documents are rebuilt: auto, editable or visual.")
            ],
            InputFormats = static () => ConvertRoutes.All().SelectMany(static route => route.SourceExtensions),
            OutputFormats = static () => ConvertRoutes.All().Select(static route => route.TargetExtension),
            ResultFormats = ["path"]
        },
        new() {
            Id = "read",
            Title = "Read",
            Summary = "Extract one document as Markdown or JSON; a shortcut for 'reader read'.",
            Synopsis = "<path|-> [options]",
            Aliases = ["extract"],
            ShortcutFor = "reader read",
            InputFormats = ReaderCommand.SupportedExtensions,
            OutputFormats = static () => ["md", "json"],
            Run = static (args, input, output, error, token) => OfficeImoToolApp.RunReaderAsync(["read", .. args], input, output, error, token),
            ResultFormats = ["markdown", "json"]
        },
        new() {
            Id = "inspect",
            Title = "Inspect",
            Summary = "Return a compact, bounded JSON summary of one document; a shortcut for 'agent inspect'.",
            Synopsis = "<path> [options]",
            ShortcutFor = "agent inspect",
            InputFormats = ReaderCommand.SupportedExtensions,
            Run = static (args, _, output, error, token) => OfficeImoToolApp.RunAgentAsync(["inspect", .. args], output, error, token),
            ResultFormats = ["json"]
        },
        new() {
            Id = "html",
            Title = "HTML",
            Summary = "Render HTML, MHTML and ZIP site bundles to PDF or to page images.",
            Synopsis = "<command> [options]",
            Run = static (args, input, output, error, token) => HtmlCommand.RunAsync(args, input, output, error, token),
            Help = [HtmlCommand.Usage],
            Subcommands = new Dictionary<string, string> {
                ["convert"] = "Render an HTML page, MHTML archive or ZIP site bundle to PDF.",
                ["render"] = "Render pages to PNG or SVG images packaged in a ZIP archive with a manifest.",
                ["capabilities"] = "List the HTML engine's supported features and profiles."
            },
            KeyOptions = [
                new("--profile", "Rendering intent, such as print-paged or screen-full-page."),
                new("--pages", "Pages to produce: all, one page, a range or stitched."),
                new("--stylesheet", "Additional CSS file applied to the document."),
                new("--base-uri", "Absolute URI used to resolve relative links."),
                new("--font-family", "Name a font family and register its faces with --font-regular and related options."),
                new("--max-pages", "Refuse documents that would produce more pages than this."),
                new("--pdf-ua-language", "Language tag for PDF/UA groundwork.")
            ],
            InputFormats = static () => ["html", "mhtml", "zip"],
            OutputFormats = static () => ["pdf", "png", "svg"],
            ResultFormats = ["text", "json"]
        },
        new() {
            Id = "reader",
            Title = "Reader",
            Summary = "Extract documents or whole folders as Markdown or JSON for search, review and language models.",
            Synopsis = "<command> [options]",
            Run = static (args, input, output, error, token) => OfficeImoToolApp.RunReaderAsync(args, input, output, error, token),
            Help = [ReaderCommand.Usage],
            Subcommands = new Dictionary<string, string> {
                ["read"] = "Extract one document, or standard input, as Markdown or JSON.",
                ["folder"] = "Extract every supported document under a folder into an output folder.",
                ["capabilities"] = "List the formats the reader handles."
            },
            KeyOptions = [
                new("--format", "Output format: markdown or json."),
                new("--output", "Write to a file, or to a directory for folder extraction."),
                new("--assets", "Directory for extracted images."),
                new("--name", "Source name used to choose a handler when reading standard input."),
                new("--concurrency", "Number of documents extracted at once in folder mode."),
                new("--max-files", "Stop folder extraction after this many files.")
            ],
            InputFormats = ReaderCommand.SupportedExtensions,
            OutputFormats = static () => ["md", "json"],
            ResultFormats = ["markdown", "json", "text"]
        },
        new() {
            Id = "markup",
            Title = "Markup",
            Summary = "Parse, validate and preview OfficeIMO Markup, emit C# or PowerShell, and export Word, Excel or PowerPoint files.",
            Synopsis = "<command> [options]",
            Run = static (args, input, output, error, token) => OfficeImoToolApp.WithWriterAsync(output, writer => MarkupCommand.RunAsync(args, input, writer, error, token)),
            Help = [MarkupCommand.Usage],
            Subcommands = new Dictionary<string, string> {
                ["parse"] = "Parse markup and print the document model with diagnostics as JSON.",
                ["validate"] = "Validate markup; exit code 1 when it has errors.",
                ["preview"] = "Print the parsed document model as JSON for previews.",
                ["emit"] = "Generate C# or PowerShell code that builds the document.",
                ["export"] = "Write the markup as a PPTX, XLSX or DOCX file."
            },
            KeyOptions = [
                new("--profile", "Markup profile: presentation, document, workbook or common."),
                new("--target", "Code or file target for emit and export."),
                new("--output", "Destination file for emit or export.")
            ],
            InputFormats = static () => ["md"],
            OutputFormats = static () => ["docx", "xlsx", "pptx", "cs", "ps1"],
            ResultFormats = ["json", "text"]
        },
        new() {
            Id = "tabular",
            Title = "Tabular",
            Summary = "List worksheets, report column schemas and convert between CSV, TSV and workbook formats.",
            Synopsis = "<command> [options]",
            Run = static (args, _, output, error, token) => OfficeImoToolApp.WithWriterAsync(output, writer => TabularCommand.RunAsync(args, writer, error, token)),
            Help = [TabularCommand.Usage],
            Subcommands = new Dictionary<string, string> {
                ["sheets"] = "List the worksheets in a workbook.",
                ["schema"] = "Report column names and inferred types for a worksheet or CSV file.",
                ["convert"] = "Convert between CSV, TSV and workbook formats."
            },
            KeyOptions = [
                new("--sheet", "Worksheet to read by name."),
                new("--sheet-index", "Worksheet to read by position."),
                new("--delimiter", "Delimiter of delimited input."),
                new("--output-delimiter", "Delimiter of delimited output."),
                new("--no-header", "Treat the first row as data.")
            ],
            InputFormats = static () => ["xlsx", "xlsm", "xltx", "xltm", "xlam", "xlsb", "xls", "csv", "tsv"],
            OutputFormats = static () => ["xlsx", "xlsb", "xls", "csv", "tsv"],
            ResultFormats = ["text"]
        },
        new() {
            Id = "workflow",
            Title = "Workflow",
            Summary = "Batch conversion with checkpoints, saved jobs, page images, assembly, comparison, print plans and printing.",
            Synopsis = "<command> [options]",
            Run = static (args, _, output, error, token) => OfficeImoToolApp.WithWriterAsync(output, writer => WorkflowCommand.RunAsync(args, writer, error, token)),
            Help = [WorkflowCommand.Usage],
            Subcommands = new Dictionary<string, string> {
                ["compare"] = "Compare two PDFs and write an HTML difference report.",
                ["optimize-images"] = "Analyze or reduce the images in Word documents.",
                ["export-pages"] = "Render PDF pages to PNG, JPEG, WebP, TIFF or SVG images.",
                ["assemble"] = "Combine documents, folders and ZIP archives into one PDF.",
                ["print-plan"] = "Plan how a PDF lays out on printed sheets without printing.",
                ["batch"] = "Convert many files with per-file results, retries and optional checkpoints.",
                ["job plan"] = "Show what a saved job would convert.",
                ["job run"] = "Run a saved job.",
                ["job report"] = "Inspect a saved run report.",
                ["printers"] = "List printers, their capabilities or their paper sources.",
                ["print"] = "Print a document to a queue or an IPP printer.",
                ["print-job status|wait|cancel"] = "Check, wait for or cancel a submitted print job."
            },
            KeyOptions = [
                new("--output", "Destination file or folder."),
                new("--input-directory", "Folder to convert in a batch."),
                new("--target", "Batch target format; PDF by default."),
                new("--checkpoint", "Folder that records verified outputs so a rerun resumes where it stopped."),
                new("--conflict", "What a batch does with an existing output: fail, rename or replace."),
                new("--retry-failed", "Retry items that failed in an earlier checkpointed run."),
                new("--concurrency", "Number of files converted at once."),
                new("--pages", "Page selection such as 1-3,last."),
                new("--save-job", "Write the batch settings to a new job file without converting."),
                new("--format", "Result format: json, ndjson for batches, or text for single operations."),
                new("--force", "Replace an existing output.")
            ],
            InputFormats = static () => ConvertRoutes.All().SelectMany(static route => route.SourceExtensions).Append(".pdf").Append(".zip"),
            OutputFormats = static () => ["pdf", "docx", "xlsx", "pptx", "html", "json", "jsonl", "yaml", "toml", "xml", "csv", "tsv", "png", "jpeg", "webp", "tiff", "svg"],
            ResultFormats = ["json", "ndjson", "text"]
        },
        new() {
            Id = "invoice",
            Title = "Invoice",
            Summary = "Inspect, validate, edit, convert and render CII and UBL electronic invoices, including Factur-X hybrid PDFs.",
            Synopsis = "<command> [options]",
            Run = static (args, _, output, error, token) => OfficeImoToolApp.WithWriterAsync(output, writer => InvoiceCommand.RunAsync(args, writer, error, token)),
            Help = [InvoiceCommand.Usage],
            Subcommands = new Dictionary<string, string> {
                ["inspect|validate"] = "Read an invoice into the model; validate returns exit code 1 when it has errors.",
                ["convert|render|hybrid"] = "Convert to another syntax or profile, render a presentation PDF, or embed the XML in a hybrid PDF.",
                ["batch inspect|validate"] = "Inspect or validate many invoices with one result per item.",
                ["batch convert|render|hybrid"] = "Convert or render many invoices into an output directory.",
                ["edit"] = "Replace selected header fields while keeping the rest of the XML.",
                ["batch edit"] = "Apply the same field replacements to many invoices."
            },
            KeyOptions = [
                new("--release", "Pinned standards release of the target, such as En16931_1_3_16."),
                new("--syntax", "Target syntax: Cii or Ubl."),
                new("--profile", "Target profile, such as En16931 or XRechnung."),
                new("--output", "Destination file for a single invoice."),
                new("--output-directory", "Destination folder for a batch."),
                new("--rule-bundle", "Pinned XRechnung rule bundle for standards validation."),
                new("--language", "Languages of the rendered presentation.")
            ],
            InputFormats = static () => ["xml"],
            OutputFormats = static () => ["xml", "pdf"],
            ResultFormats = ["json"]
        },
        new() {
            Id = "pdf",
            Title = "PDF",
            Summary = "Extract pages, split, decrypt, flatten, optimize, sanitize, OCR, convert to PDF/A and redact PDFs into separate copies.",
            Synopsis = "<command> [options]",
            Run = static (args, _, output, error, token) => OfficeImoToolApp.WithWriterAsync(output, writer => PdfCommand.RunAsync(args, writer, error, token)),
            Help = [PdfWorkflowCommand.Usage, PdfCommand.Usage],
            Subcommands = new Dictionary<string, string> {
                ["extract"] = "Copy selected pages into a new PDF.",
                ["split"] = "Split a PDF into parts of a fixed page count.",
                ["decrypt"] = "Write an unencrypted copy using the owner password from an environment variable.",
                ["flatten"] = "Replace pages with their rendered appearance.",
                ["ocr"] = "Add searchable text from an OCR provider.",
                ["optimize|sanitize"] = "Write a smaller copy, or a copy without active content.",
                ["convert-to-pdfa"] = "Write a PDF/A copy, or refuse and report every violation.",
                ["providers"] = "List the OCR providers that are loaded.",
                ["redact providers"] = "List the OCR providers available to redaction.",
                ["redact batch"] = "Run several redactions from one request file.",
                ["redact plan"] = "Find redaction candidates and write a reviewable plan.",
                ["redact apply"] = "Apply reviewed redaction decisions to a new PDF.",
                ["redact verify"] = "Check that an existing redacted PDF matches the reviewed decisions."
            },
            KeyOptions = [
                new("--output", "Destination file or folder; the source is never changed."),
                new("--pages", "Page selection such as 1-3,last, odd, even or all,!2."),
                new("--password-env", "Name of the environment variable that holds the owner password."),
                new("--recipe", "Redaction recipe JSON."),
                new("--decisions", "Reviewed redaction decisions JSON."),
                new("--evidence", "Path for the redaction evidence JSON."),
                new("--ocr-provider", "OCR provider id."),
                new("--ocr-provider-assembly", "Trusted assembly that contains an OCR provider."),
                new("--maximum-pages", "Page limit; 100 by default.", "1..10000"),
                new("--force", "Replace an existing output.")
            ],
            InputFormats = static () => ["pdf"],
            OutputFormats = static () => ["pdf", "json"],
            ResultFormats = ["json"]
        },
        new() {
            Id = "provenance",
            Title = "Provenance",
            Summary = "Inspect, assess and remove Content Credentials (C2PA) and AI-source labels, and audit files for suspicious Unicode in CI.",
            Synopsis = "<command> [options]",
            Run = static (args, _, output, error, token) => OfficeImoToolApp.WithWriterAsync(output, writer => ProvenanceCommand.RunAsync(args, writer, error, token)),
            Help = [ProvenanceCommand.Usage],
            Subcommands = new Dictionary<string, string> {
                ["doctor"] = "Check that a host-supplied c2patool executable runs.",
                ["capabilities"] = "List the formats and provenance carriers the tool understands.",
                ["inspect"] = "Report the provenance carriers in one file.",
                ["assess"] = "Inspect and verify provenance, optionally with c2patool and trust anchors.",
                ["remove"] = "Write a copy without selected provenance carriers.",
                ["audit"] = "Report provenance and dangerous Unicode across files and folders without changing them.",
                ["check"] = "Audit and return exit code 1 when the selected findings are present.",
                ["batch inspect|assess"] = "Inspect or assess many files with one result per item.",
                ["batch remove"] = "Remove provenance from many files into an output directory."
            },
            KeyOptions = [
                new("--format", "Result format; JSON by default."),
                new("--fail-on", "Findings that fail check: dangerous-text, carriers or any."),
                new("--include", "Wildcard of files to include in an audit."),
                new("--exclude", "Wildcard of files to skip in an audit."),
                new("--c2patool", "Trusted c2patool executable used for verification."),
                new("--output", "Destination file for remove."),
                new("--force", "Replace an existing output.")
            ],
            ResultFormats = ["json", "text", "ndjson", "sarif"]
        },
        new() {
            Id = "hidden-data",
            Title = "Hidden data",
            Summary = "Find and remove metadata, comments, tracked changes, hidden content and active content before sharing a file.",
            Synopsis = "<command> [options]",
            Run = static (args, _, output, error, token) => OfficeImoToolApp.WithWriterAsync(output, writer => HiddenDataCommand.RunAsync(args, writer, error, token)),
            Help = [HiddenDataCommand.Usage],
            Subcommands = new Dictionary<string, string> {
                ["inspect"] = "List the hidden data a file carries without changing it.",
                ["remove"] = "Write a cleaned copy; the input is never changed."
            },
            KeyOptions = [
                new("--kinds", "What to remove: default, all, or named kinds such as provenance or tracked-changes."),
                new("--tracked-changes", "Keep, accept or reject tracked changes."),
                new("--keep-provenance", "Provenance carriers to keep when provenance is removed."),
                new("--output", "Destination file; the default is <input>.cleaned<ext>."),
                new("--force", "Replace an existing output.")
            ],
            InputFormats = static () => HiddenDataFormats,
            OutputFormats = static () => HiddenDataFormats,
            ResultFormats = ["json", "text"]
        },
        new() {
            Id = "agent",
            Title = "Agent",
            Summary = "Inspect, search, fetch and convert documents and mailboxes with compact, bounded JSON results.",
            Synopsis = "<command> [options]",
            Run = static (args, _, output, error, token) => OfficeImoToolApp.RunAgentAsync(args, output, error, token),
            Help = [AgentCommand.Usage],
            Subcommands = new Dictionary<string, string> {
                ["inspect"] = "Summarize one document as compact JSON.",
                ["inspect-email"] = "Summarize one message or mailbox as compact JSON.",
                ["search"] = "Search the messages of a mailbox by metadata.",
                ["fetch"] = "Fetch one result returned by inspect or search.",
                ["search-email"] = "Search the content of a mailbox's messages.",
                ["convert"] = "Convert a document to Markdown, JSON or an Office or PDF file.",
                ["capabilities"] = "List the operations available for each format."
            },
            KeyOptions = [
                new("--max-output-characters", "Upper bound on the size of the JSON result."),
                new("--query", "Text to search for."),
                new("--take", "Number of results per page."),
                new("--cursor", "Continue from a previous page of results."),
                new("--source-id", "Source returned by inspect or search, used by fetch.")
            ],
            InputFormats = ReaderCommand.SupportedExtensions,
            OutputFormats = static () => ["md", "json", "docx", "xlsx", "pptx", "pdf"],
            ResultFormats = ["json"]
        },
        new() {
            Id = "mcp",
            Title = "MCP server",
            Summary = "Serve OfficeIMO tools to MCP clients over standard input and output, limited to the folders in OFFICEIMO_MCP_ALLOWED_ROOTS.",
            Synopsis = "serve --stdio",
            Run = static (args, _, _, error, token) => McpCommand.RunAsync(args, error, token),
            Help = [McpCommand.Usage],
            Subcommands = new Dictionary<string, string> {
                ["serve"] = "Start the MCP server on standard input and output."
            },
            KeyOptions = [
                new("--stdio", "Use standard input and output as the transport."),
                new("--pdf-password-env", "Admit an environment variable that PDF tools may read a password from."),
                new("--ocr-provider-assembly", "Trusted OCR provider assembly loaded at startup.")
            ],
            ResultFormats = ["json"]
        }
    ];

    internal static OfficeImoToolCommandGroup? Find(string name) {
        foreach (OfficeImoToolCommandGroup group in Groups) {
            if (string.Equals(group.Id, name, StringComparison.OrdinalIgnoreCase)) return group;
            foreach (string alias in group.Aliases) {
                if (string.Equals(alias, name, StringComparison.OrdinalIgnoreCase)) return group;
            }
        }
        return null;
    }
}
