using System.Text.Json;
using System.Text.Json.Serialization;
using OfficeIMO.AI;

internal sealed class ExampleOptions {
    public const string SyntheticInvoice = "Invoice EXAMPLE-2026-001. Seller: Example Supplies. Total: 1234.50 PLN. Due: 2026-09-30.";
    public static readonly JsonSerializerOptions RequestJsonOptions = new() {
        PropertyNameCaseInsensitive = true, UnmappedMemberHandling = JsonUnmappedMemberHandling.Disallow,
        Converters = { new JsonStringEnumConverter() }, MaxDepth = 12
    };
    public static OfficeAiRequest InvoiceRequest => new() {
        Operation = OfficeAiOperation.ExtractFields, Instruction = "Extract the invoice number, total and due date. Use exact source values.",
        Fields = new[] { new OfficeAiFieldDefinition("invoiceNumber"), new OfficeAiFieldDefinition("total", OfficeAiFieldType.Decimal),
            new OfficeAiFieldDefinition("dueDate", OfficeAiFieldType.Date, "yyyy-MM-dd") }
    };
    public const string Usage = """
        OfficeIMO.AI headless example (.NET 10)
        Without --source or --request, processes the built-in synthetic invoice.
        --source PATH         PDF, plain text, PNG, JPEG or WebP input
        --request PATH        OfficeAiRequest JSON with operation, instruction, fields, pages and limits
        --output DIRECTORY    Write report.json plus proposed Reader JSON, CSV and XLSX where applicable; files must not exist
        --allow-remote        Authorize sending the selected source evidence to the configured hosted model
        --images              Include page images for the selected vision-capable model
        --model NAME          Explicit model; default gpt-5.5
        --codex-session       Prefer the existing Codex login over IX's saved ChatGPT credential
        --endpoint URL        Use an OpenAI-compatible endpoint; optional key from OFFICEIMO_AI_API_KEY
        --copilot             Use native Copilot HTTP with a GitHub credential; requires --model
        --text-only           Evaluate native text cases and exclude image-required cases explicitly
        --local               Require an explicit loopback endpoint, with no redirects or system proxy
        --prompted-json       Provider does not enforce schemas; local validation still applies
        --help                Show this help
        --evaluate            Run the versioned synthetic evaluation corpus; requires a new --output directory
        --split NAME          all, development or heldout (default all)
        --repeat COUNT        Repeat each selected case 1-3 times to expose instability
        --request-characters N  Bound model request text, including wrappers (minimum 4096)
        --case ID             Run one named synthetic case when diagnosing an evaluation failure
        Exit: 0 completed, 1 partial/insufficient/invalid result, 2 setup/input failure, 3 cancelled/timeout.
        """;
    public string Split { get; private set; } = "all";
    public int Repeat { get; private set; } = 1;
    public int RequestCharacters { get; private set; } = 48000;
    public bool Copilot { get; private set; }
    public bool TextOnly { get; private set; }
    public bool Help { get; private set; }
    public bool Evaluate { get; private set; }
    public string? CaseId { get; private set; }
    public bool AllowRemote { get; private set; }
    public bool Images { get; private set; }
    public bool CodexSession { get; private set; }
    public bool Local { get; private set; }
    public bool PromptedJson { get; private set; }
    public string Model { get; private set; } = "gpt-5.5";
    public string? SourcePath { get; private set; }
    public string? RequestPath { get; private set; }
    public string? OutputPath { get; private set; }
    public Uri? Endpoint { get; private set; }

    public static ExampleOptions Parse(string[] args) {
        var options = new ExampleOptions();
        var seen = new HashSet<string>(StringComparer.Ordinal);
        for (int index = 0; index < args.Length; index++) {
            string key = args[index];
            if (!seen.Add(key)) throw new ArgumentException("Duplicate option.");
            string Value() => ++index < args.Length && !args[index].StartsWith("--", StringComparison.Ordinal) ? args[index] : throw new ArgumentException("Missing option value.");
            switch (key) {
                case "--help": options.Help = true; break;
                case "--evaluate": options.Evaluate = true; break;
                case "--split": options.Split = Value(); break;
                case "--repeat": options.Repeat = int.Parse(Value(), System.Globalization.CultureInfo.InvariantCulture); break;
                case "--request-characters": options.RequestCharacters = int.Parse(Value(), System.Globalization.CultureInfo.InvariantCulture); break;
                case "--case": options.CaseId = Value(); break;
                case "--allow-remote": options.AllowRemote = true; break;
                case "--images": options.Images = true; break;
                case "--codex-session": options.CodexSession = true; break;
                case "--copilot": options.Copilot = true; break;
                case "--text-only": options.TextOnly = true; break;
                case "--local": options.Local = true; break;
                case "--prompted-json": options.PromptedJson = true; break;
                case "--source": options.SourcePath = Value(); break;
                case "--request": options.RequestPath = Value(); break;
                case "--output": options.OutputPath = Value(); break;
                case "--model": options.Model = Value(); break;
                case "--endpoint": options.Endpoint = new Uri(Value(), UriKind.Absolute); break;
                default: throw new ArgumentException("Unknown option.");
            }
        }
        if (options.Copilot && !seen.Contains("--model")) throw new ArgumentException("Copilot requires an explicit --model from its available model catalog.");
        if (options.Copilot && (options.Local || options.Endpoint is not null || options.CodexSession)) throw new ArgumentException("Copilot requires its own hosted route and GitHub credential.");
        if (options.Split is not ("all" or "development" or "heldout") || options.Repeat is < 1 or > 3
            || options.RequestCharacters is < 4096 or > 2000000) throw new ArgumentException("Invalid evaluation or request bounds.");
        return options;
    }
}
