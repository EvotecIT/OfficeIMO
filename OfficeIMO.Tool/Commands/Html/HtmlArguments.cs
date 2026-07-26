using System.Globalization;

namespace OfficeIMO.Tool.Commands.Html;

internal enum HtmlCommandKind {
    Help,
    Convert,
    Capabilities
}

internal enum HtmlInputFormat {
    Auto,
    Html,
    Mhtml
}

internal sealed class HtmlArguments {
    internal const long DefaultMaxInputBytes = 64L * 1024L * 1024L;
    internal const long MaxStylesheetBytes = 4L * 1024L * 1024L;
    internal const long MaxFontBytes = 32L * 1024L * 1024L;
    internal const int MaxStylesheetCount = 16;

    internal HtmlCommandKind Command { get; private set; }
    internal HtmlInputFormat InputFormat { get; private set; }
    internal string? InputPath { get; private set; }
    internal string? OutputPath { get; private set; }
    internal string? BaseUri { get; private set; }
    internal string? PdfUaLanguage { get; private set; }
    internal string? FontFamilyName { get; private set; }
    internal string? RegularFontPath { get; private set; }
    internal string? BoldFontPath { get; private set; }
    internal string? ItalicFontPath { get; private set; }
    internal string? BoldItalicFontPath { get; private set; }
    internal long MaxInputBytes { get; private set; } = DefaultMaxInputBytes;
    internal int MaxPages { get; private set; } = 10_000;
    internal bool Force { get; private set; }
    internal bool JsonCapabilities { get; private set; }
    internal List<string> StylesheetPaths { get; } = new List<string>();

    internal static HtmlArguments Parse(string[] args) {
        if (args == null) throw new ArgumentNullException(nameof(args));
        if (args.Length == 0 || IsHelp(args[0])) return new HtmlArguments { Command = HtmlCommandKind.Help };

        var parsed = new HtmlArguments {
            Command = args[0].ToLowerInvariant() switch {
                "convert" => HtmlCommandKind.Convert,
                "capabilities" => HtmlCommandKind.Capabilities,
                _ => throw new HtmlUsageException("Unknown command '" + args[0] + "'.")
            }
        };

        for (int index = 1; index < args.Length; index++) {
            string token = args[index];
            if (IsHelp(token)) return new HtmlArguments { Command = HtmlCommandKind.Help };
            if (parsed.Command == HtmlCommandKind.Capabilities && token != "--format") {
                throw new HtmlUsageException("The capabilities command accepts only --format text|json.");
            }
            if (parsed.Command == HtmlCommandKind.Convert && token == "--format") {
                throw new HtmlUsageException("The convert command does not accept --format.");
            }
            switch (token) {
                case "--output":
                case "-o":
                    parsed.OutputPath = NextValue(args, ref index, token);
                    break;
                case "--input-format":
                    parsed.InputFormat = ParseInputFormat(NextValue(args, ref index, token));
                    break;
                case "--stylesheet":
                    if (parsed.StylesheetPaths.Count >= MaxStylesheetCount) {
                        throw new HtmlUsageException("--stylesheet may be specified at most " + MaxStylesheetCount + " times.");
                    }
                    parsed.StylesheetPaths.Add(NextValue(args, ref index, token));
                    break;
                case "--base-uri":
                    parsed.BaseUri = NextValue(args, ref index, token);
                    break;
                case "--font-family":
                    parsed.FontFamilyName = NextValue(args, ref index, token);
                    break;
                case "--font-regular":
                    parsed.RegularFontPath = NextValue(args, ref index, token);
                    break;
                case "--font-bold":
                    parsed.BoldFontPath = NextValue(args, ref index, token);
                    break;
                case "--font-italic":
                    parsed.ItalicFontPath = NextValue(args, ref index, token);
                    break;
                case "--font-bold-italic":
                    parsed.BoldItalicFontPath = NextValue(args, ref index, token);
                    break;
                case "--max-input-bytes":
                    parsed.MaxInputBytes = ParsePositiveLong(NextValue(args, ref index, token), token);
                    break;
                case "--max-pages":
                    parsed.MaxPages = ParseBoundedInt(NextValue(args, ref index, token), token, 1, 100_000);
                    break;
                case "--pdf-ua-language":
                    parsed.PdfUaLanguage = NextValue(args, ref index, token);
                    break;
                case "--force":
                    parsed.Force = true;
                    break;
                case "--format":
                    string capabilityFormat = NextValue(args, ref index, token);
                    if (string.Equals(capabilityFormat, "json", StringComparison.OrdinalIgnoreCase)) {
                        parsed.JsonCapabilities = true;
                    } else if (string.Equals(capabilityFormat, "text", StringComparison.OrdinalIgnoreCase)) {
                        parsed.JsonCapabilities = false;
                    } else {
                        throw new HtmlUsageException("Capabilities format must be 'text' or 'json'.");
                    }
                    break;
                default:
                    if (token.StartsWith("-", StringComparison.Ordinal) && token != "-") {
                        throw new HtmlUsageException("Unknown option '" + token + "'.");
                    }
                    if (parsed.InputPath != null) throw new HtmlUsageException("Only one input path may be specified.");
                    parsed.InputPath = token;
                    break;
            }
        }

        parsed.Validate();
        return parsed;
    }

    private void Validate() {
        if (Command == HtmlCommandKind.Capabilities) {
            return;
        }

        if (string.IsNullOrWhiteSpace(InputPath)) throw new HtmlUsageException("The convert command requires <input.html|input.mhtml|->.");
        if (InputPath == "-" && InputFormat == HtmlInputFormat.Auto) {
            throw new HtmlUsageException("Standard input requires --input-format html|mhtml.");
        }
        if (string.IsNullOrWhiteSpace(OutputPath)) {
            OutputPath = InputPath == "-" ? "-" : Path.ChangeExtension(InputPath, ".pdf");
        }
        if (InputPath != "-" && OutputPath != "-" && OfficeImoToolPathSafety.PathsEqual(InputPath!, OutputPath!)) {
            throw new HtmlUsageException("Input and output paths must be different.");
        }
        if (BaseUri != null && (!Uri.TryCreate(BaseUri, UriKind.Absolute, out Uri? uri)
            || (uri.Scheme != Uri.UriSchemeHttp && uri.Scheme != Uri.UriSchemeHttps && uri.Scheme != Uri.UriSchemeFile))) {
            throw new HtmlUsageException("--base-uri must be an absolute http, https, or file URI.");
        }
        if (PdfUaLanguage != null && string.IsNullOrWhiteSpace(PdfUaLanguage)) {
            throw new HtmlUsageException("--pdf-ua-language requires a non-empty language tag.");
        }
        bool hasOptionalFontFace = BoldFontPath != null || ItalicFontPath != null || BoldItalicFontPath != null;
        if (RegularFontPath == null && (FontFamilyName != null || hasOptionalFontFace)) {
            throw new HtmlUsageException("--font-regular is required when configuring an embedded font family.");
        }
        if (RegularFontPath != null && string.IsNullOrWhiteSpace(FontFamilyName)) {
            throw new HtmlUsageException("--font-family is required with --font-regular.");
        }
    }

    internal HtmlInputFormat ResolveInputFormat() {
        if (InputFormat != HtmlInputFormat.Auto) return InputFormat;
        string extension = Path.GetExtension(InputPath!).ToLowerInvariant();
        return extension is ".mhtml" or ".mht" ? HtmlInputFormat.Mhtml : HtmlInputFormat.Html;
    }

    private static HtmlInputFormat ParseInputFormat(string value) => value.ToLowerInvariant() switch {
        "html" or "htm" => HtmlInputFormat.Html,
        "mhtml" or "mht" => HtmlInputFormat.Mhtml,
        _ => throw new HtmlUsageException("--input-format must be 'html' or 'mhtml'.")
    };

    private static string NextValue(string[] args, ref int index, string option) {
        if (++index >= args.Length || string.IsNullOrWhiteSpace(args[index])) {
            throw new HtmlUsageException(option + " requires a value.");
        }
        return args[index];
    }

    private static long ParsePositiveLong(string value, string option) {
        if (!long.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out long parsed) || parsed < 1) {
            throw new HtmlUsageException(option + " must be a positive integer.");
        }
        return parsed;
    }

    private static int ParseBoundedInt(string value, string option, int minimum, int maximum) {
        if (!int.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out int parsed)
            || parsed < minimum || parsed > maximum) {
            throw new HtmlUsageException(option + " must be between " + minimum + " and " + maximum + ".");
        }
        return parsed;
    }

    private static bool IsHelp(string value) => value is "help" or "--help" or "-h";
}

internal sealed class HtmlUsageException : Exception {
    internal HtmlUsageException(string message) : base(message) { }
}
