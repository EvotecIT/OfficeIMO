using OfficeIMO.Invoicing;
using OfficeIMO.Invoicing.Pdf;
using OfficeIMO.Invoicing.Validation;
using OfficeIMO.Pdf;
using OfficeIMO.Workflows;

namespace OfficeIMO.Tool.Commands.Invoice;

internal sealed class InvoiceArguments {
    internal bool Help, Batch;
    internal OfficeInvoiceWorkflowOperation Operation;
    internal readonly List<string> Inputs = new();
    internal readonly Dictionary<string, string> Options = new(StringComparer.Ordinal);
    internal readonly OfficeInvoiceWorkflowBatchOptions Limits = new();
    internal InvoiceXmlOptions? Target;
    internal InvoiceSpecificationRelease? StandardsRelease;
    internal InvoicePdfLayoutOptions Layout = new();
    internal PdfOptions? Pdf;

    internal static InvoiceArguments Parse(string[] args) {
        var result = new InvoiceArguments();
        if (args.Length == 0 || args[0] is "help" or "--help" or "-h") { result.Help = true; return result; }
        int index = 0;
        if (args[0] == "batch") { result.Batch = true; index++; }
        if (index >= args.Length) throw new ArgumentException("Choose an invoice operation.");
        result.Operation = args[index++] switch {
            "inspect" => OfficeInvoiceWorkflowOperation.Inspect,
            "validate" => OfficeInvoiceWorkflowOperation.Validate,
            "convert" => OfficeInvoiceWorkflowOperation.Convert,
            "render" => OfficeInvoiceWorkflowOperation.RenderPresentationPdf,
            "hybrid" => OfficeInvoiceWorkflowOperation.RenderHybridPdf,
            _ => throw new ArgumentException("Unknown invoice operation.")
        };
        bool projection = false;
        for (; index < args.Length; index++) {
            string argument = args[index];
            if (argument == "--allow-profile-loss") { if (projection) throw new ArgumentException("Duplicate --allow-profile-loss."); projection = true; continue; }
            if (argument == "--stop-on-failure") { result.Limits.ContinueOnFailure = false; continue; }
            if (argument is "--compact-details" or "--page-identity" or "--modern") {
                if (!result.Options.TryAdd(argument, "true")) throw new ArgumentException("Duplicate invoice option: " + argument);
                continue;
            }
            if (!argument.StartsWith("--", StringComparison.Ordinal)) { result.Inputs.Add(argument); continue; }
            if (argument is not ("--output" or "--output-directory" or "--release" or "--syntax" or "--profile" or
                "--standards-release" or "--rule-bundle" or "--peppol-rules" or "--facturx-rules" or "--saxon-jar" or "--java" or
                "--language" or "--font" or "--columns" or "--unit-display" or "--payment-display" or "--max-items" or "--max-input-bytes" or "--max-output-bytes"))
                throw new ArgumentException("Unknown invoice option: " + argument);
            if (++index >= args.Length || args[index].StartsWith("--", StringComparison.Ordinal)) throw new ArgumentException("Missing value for " + argument);
            if (!result.Options.TryAdd(argument, args[index])) throw new ArgumentException("Duplicate invoice option: " + argument);
        }
        if (result.Inputs.Count == 0 || !result.Batch && result.Inputs.Count != 1) throw new ArgumentException("Supply one input, or use batch for several inputs.");
        bool target = result.Options.ContainsKey("--release") || result.Options.ContainsKey("--syntax") || result.Options.ContainsKey("--profile");
        if (target) result.Target = new InvoiceXmlOptions(ParseEnum<InvoiceSpecificationRelease>(result.Required("--release")),
            ParseEnum<InvoiceSyntax>(result.Required("--syntax")), ParseEnum<InvoiceProfile>(result.Required("--profile")),
            projection ? InvoiceProjectionPolicy.AllowProfileDefinedDataLoss : InvoiceProjectionPolicy.RejectDataLoss);
        else if (projection) throw new ArgumentException("Profile loss requires an explicit target.");
        if (result.Options.TryGetValue("--standards-release", out string? release)) result.StandardsRelease = ParseEnum<InvoiceSpecificationRelease>(release);
        if (!result.StandardsRelease.HasValue && result.Options.Keys.Any(k => k is "--rule-bundle" or "--peppol-rules" or "--facturx-rules" or "--saxon-jar" or "--java"))
            throw new ArgumentException("Standards artifacts require --standards-release.");
        if (result.Operation is not (OfficeInvoiceWorkflowOperation.RenderHybridPdf or OfficeInvoiceWorkflowOperation.RenderPresentationPdf) &&
            result.Options.Keys.Any(k => k is "--language" or "--font" or "--columns" or "--unit-display" or "--payment-display" or "--compact-details" or "--page-identity" or "--modern"))
            throw new ArgumentException("Presentation options require render or hybrid.");
        if (result.Options.TryGetValue("--language", out string? cultures)) result.Layout = InvoicePdfLayoutOptions.ForCultures(cultures.Split(','));
        if (result.Options.TryGetValue("--columns", out string? columns)) {
            result.Layout.LineColumns.Clear();
            foreach (string column in columns.Split(',')) result.Layout.LineColumns.Add(ParseEnum<InvoicePdfLineColumn>(column));
        }
        if (result.Options.TryGetValue("--unit-display", out string? unitDisplay)) result.Layout.UnitCodeDisplay = ParseEnum<InvoicePdfCodeDisplay>(unitDisplay);
        if (result.Options.TryGetValue("--payment-display", out string? paymentDisplay)) result.Layout.PaymentCodeDisplay = ParseEnum<InvoicePdfCodeDisplay>(paymentDisplay);
        result.Layout.CompactDetails = result.Options.ContainsKey("--compact-details");
        result.Layout.IncludePageIdentity = result.Options.ContainsKey("--page-identity");
        if (result.Options.ContainsKey("--modern")) result.Layout.Theme = new InvoicePdfTheme();
        try { result.Layout = result.Layout.Clone(); }
        catch (InvalidOperationException exception) { throw new ArgumentException(exception.Message, exception); }
        if (result.Options.TryGetValue("--font", out string? fontPath)) {
            using var fontFile = File.OpenRead(fontPath);
            if (fontFile.Length is <= 0 or > 16 * 1024 * 1024) throw new ArgumentException("Embedded font must contain between 1 byte and 16 MiB.");
            byte[] font = new byte[(int)fontFile.Length]; fontFile.ReadExactly(font);
            result.Pdf = new PdfOptions().EmbedStandardFont(PdfStandardFont.Helvetica, font, "Invoice font")
                .EmbedStandardFont(PdfStandardFont.HelveticaBold, font, "Invoice font");
        }
        if (result.Options.TryGetValue("--max-items", out string? items)) result.Limits.MaximumRequests = PositiveInt(items, 10000);
        if (result.Options.TryGetValue("--max-input-bytes", out string? bytes)) result.Limits.MaximumInputBytes = PositiveLong(bytes);
        if (result.Options.TryGetValue("--max-output-bytes", out bytes)) result.Limits.MaximumOutputBytes = PositiveLong(bytes);
        if (result.Inputs.Count > result.Limits.MaximumRequests) throw new ArgumentException("Invoice inputs exceed --max-items.");
        if (result.Batch && result.Options.ContainsKey("--output") || !result.Batch && result.Options.ContainsKey("--output-directory"))
            throw new ArgumentException("Use --output for one input or --output-directory for a batch.");
        // Validate request contracts before loading standards artifacts or executing any input.
        _ = result.Requests().ToArray();
        return result;
    }
    internal IEnumerable<OfficeInvoiceFileWorkflowRequest> Requests() {
        bool writes = Operation is OfficeInvoiceWorkflowOperation.Convert or OfficeInvoiceWorkflowOperation.RenderHybridPdf or OfficeInvoiceWorkflowOperation.RenderPresentationPdf;
        if (!writes && Options.Keys.Any(k => k is "--output" or "--output-directory")) throw new ArgumentException("Inspection and validation do not create files.");
        foreach (string input in Inputs) {
            string? output = writes ? Batch
                ? Path.Combine(Required("--output-directory"), Path.GetFileNameWithoutExtension(input) + ".invoice" + (Operation == OfficeInvoiceWorkflowOperation.Convert ? ".xml" : ".pdf"))
                : Required("--output") : null;
            yield return new(input, Operation, output, Target, StandardsRelease, Layout, Pdf);
        }
    }
    internal InvoiceValidator? Validator() {
        if (!StandardsRelease.HasValue) return null;
        var bundle = InvoiceRuleBundle.Load(Required("--rule-bundle"), Get("--peppol-rules"), Get("--facturx-rules"));
        var runner = Get("--saxon-jar") is string jar ? new SaxonInvoiceRulesRunner(jar, Get("--java") ?? "java") : null;
        return new InvoiceValidator(bundle, runner);
    }
    private string Required(string key) => Get(key) ?? throw new ArgumentException("Required invoice option: " + key);
    private string? Get(string key) => Options.GetValueOrDefault(key);
    private static T ParseEnum<T>(string value) where T : struct, Enum {
        value = value.Trim();
        return value.Length > 0 && !value.Contains(',') && Enum.TryParse<T>(value, true, out T parsed) &&
            Enum.IsDefined(parsed) && !char.IsDigit(value[0]) && value[0] is not ('-' or '+') ? parsed : throw new ArgumentException("Invalid " + typeof(T).Name + ": " + value);
    }
    private static int PositiveInt(string value, int maximum) => int.TryParse(value, out int result) && result > 0 && result <= maximum ? result : throw new ArgumentException("Invalid positive item limit.");
    private static long PositiveLong(string value) => long.TryParse(value, out long result) && result > 0 ? result : throw new ArgumentException("Invalid positive byte limit.");
}
