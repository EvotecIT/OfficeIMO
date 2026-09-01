using System.Globalization;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Workflows;

namespace OfficeIMO.Tool.Commands.Workflow;

internal enum WorkflowCommandKind {
    Help,
    ExportPages,
    Assemble,
    PrintPlan
}

internal sealed class WorkflowArguments {
    internal WorkflowCommandKind Command { get; private set; }
    internal IReadOnlyList<string> Inputs { get; private set; } = [];
    internal string? OutputPath { get; private set; }
    internal string? Pages { get; private set; }
    internal OfficeImageExportFormat ImageFormat { get; private set; } = OfficeImageExportFormat.Png;
    internal double TargetDpi { get; private set; } = 144D;
    internal int? MaximumDimension { get; private set; }
    internal bool Force { get; private set; }
    internal bool IncludeSubdirectories { get; private set; } = true;
    internal PageSize PaperSize { get; private set; } = PageSizes.A4;
    internal PdfPrintOrientation Orientation { get; private set; } = PdfPrintOrientation.Automatic;
    internal int PagesPerSheet { get; private set; } = 1;
    internal PdfPrintScaleMode ScaleMode { get; private set; } = PdfPrintScaleMode.Fit;
    internal double Margin { get; private set; } = 18D;

    internal static WorkflowArguments Parse(string[] args) {
        if (args.Length == 0 || IsHelp(args[0])) return new WorkflowArguments { Command = WorkflowCommandKind.Help };

        var parsed = new WorkflowArguments {
            Command = args[0].ToLowerInvariant() switch {
                "export-pages" => WorkflowCommandKind.ExportPages,
                "assemble" => WorkflowCommandKind.Assemble,
                "print-plan" => WorkflowCommandKind.PrintPlan,
                _ => throw new WorkflowUsageException("Unknown workflow command '" + args[0] + "'.")
            }
        };
        var inputs = new List<string>();
        for (int index = 1; index < args.Length; index++) {
            string token = args[index];
            if (IsHelp(token)) return new WorkflowArguments { Command = WorkflowCommandKind.Help };
            switch (token) {
                case "--output":
                    EnsureCommand(parsed.Command, token, WorkflowCommandKind.ExportPages, WorkflowCommandKind.Assemble);
                    parsed.OutputPath = ReadValue(args, ref index, token);
                    break;
                case "--pages":
                    EnsureCommand(parsed.Command, token, WorkflowCommandKind.ExportPages, WorkflowCommandKind.PrintPlan);
                    parsed.Pages = ReadValue(args, ref index, token);
                    break;
                case "--format":
                    EnsureCommand(parsed.Command, token, WorkflowCommandKind.ExportPages);
                    parsed.ImageFormat = ParseImageFormat(ReadValue(args, ref index, token));
                    break;
                case "--dpi":
                    EnsureCommand(parsed.Command, token, WorkflowCommandKind.ExportPages);
                    parsed.TargetDpi = ParseDouble(ReadValue(args, ref index, token), token, 36D, 600D);
                    break;
                case "--max-dimension":
                    EnsureCommand(parsed.Command, token, WorkflowCommandKind.ExportPages);
                    parsed.MaximumDimension = ParseInt(ReadValue(args, ref index, token), token, 1, 20_000);
                    break;
                case "--force":
                    EnsureCommand(parsed.Command, token, WorkflowCommandKind.ExportPages, WorkflowCommandKind.Assemble);
                    parsed.Force = true;
                    break;
                case "--no-recursive":
                    EnsureCommand(parsed.Command, token, WorkflowCommandKind.Assemble);
                    parsed.IncludeSubdirectories = false;
                    break;
                case "--paper":
                    EnsureCommand(parsed.Command, token, WorkflowCommandKind.PrintPlan);
                    parsed.PaperSize = ParsePaper(ReadValue(args, ref index, token));
                    break;
                case "--orientation":
                    EnsureCommand(parsed.Command, token, WorkflowCommandKind.PrintPlan);
                    parsed.Orientation = ParseOrientation(ReadValue(args, ref index, token));
                    break;
                case "--pages-per-sheet":
                    EnsureCommand(parsed.Command, token, WorkflowCommandKind.PrintPlan);
                    parsed.PagesPerSheet = ParsePagesPerSheet(ReadValue(args, ref index, token));
                    break;
                case "--scale":
                    EnsureCommand(parsed.Command, token, WorkflowCommandKind.PrintPlan);
                    parsed.ScaleMode = ParseScale(ReadValue(args, ref index, token));
                    break;
                case "--margin":
                    EnsureCommand(parsed.Command, token, WorkflowCommandKind.PrintPlan);
                    parsed.Margin = ParseDouble(ReadValue(args, ref index, token), token, 0D, 200D);
                    break;
                default:
                    if (token.StartsWith("-", StringComparison.Ordinal)) {
                        throw new WorkflowUsageException("Unknown option '" + token + "'.");
                    }
                    inputs.Add(token);
                    break;
            }
        }
        parsed.Inputs = inputs;
        parsed.Validate();
        return parsed;
    }

    private void Validate() {
        switch (Command) {
            case WorkflowCommandKind.ExportPages:
                RequireInputs(exactCount: 1);
                RequireOutput("export-pages");
                break;
            case WorkflowCommandKind.Assemble:
                if (Inputs.Count == 0) throw new WorkflowUsageException("assemble requires at least one source path.");
                RequireOutput("assemble");
                break;
            case WorkflowCommandKind.PrintPlan:
                RequireInputs(exactCount: 1);
                break;
        }
    }

    private void RequireInputs(int exactCount) {
        if (Inputs.Count != exactCount) {
            throw new WorkflowUsageException(CommandName(Command) + " requires exactly one input PDF.");
        }
    }

    private void RequireOutput(string command) {
        if (string.IsNullOrWhiteSpace(OutputPath)) throw new WorkflowUsageException(command + " requires --output <path>.");
    }

    private static string CommandName(WorkflowCommandKind command) => command switch {
        WorkflowCommandKind.ExportPages => "export-pages",
        WorkflowCommandKind.PrintPlan => "print-plan",
        _ => command.ToString().ToLowerInvariant()
    };

    private static string ReadValue(string[] args, ref int index, string option) {
        if (++index >= args.Length || string.IsNullOrWhiteSpace(args[index]) || args[index].StartsWith("-", StringComparison.Ordinal)) {
            throw new WorkflowUsageException(option + " requires a value.");
        }
        return args[index];
    }

    private static void EnsureCommand(WorkflowCommandKind command, string option, params WorkflowCommandKind[] allowed) {
        if (!allowed.Contains(command)) {
            throw new WorkflowUsageException(option + " is not valid with " + CommandName(command) + ".");
        }
    }

    private static int ParseInt(string value, string option, int minimum, int maximum) {
        if (!int.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out int parsed) || parsed < minimum || parsed > maximum) {
            throw new WorkflowUsageException(option + " must be between " + minimum + " and " + maximum + ".");
        }
        return parsed;
    }

    private static double ParseDouble(string value, string option, double minimum, double maximum) {
        if (!double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed) ||
            !double.IsFinite(parsed) || parsed < minimum || parsed > maximum) {
            throw new WorkflowUsageException(option + " must be between " + minimum.ToString(CultureInfo.InvariantCulture) +
                                             " and " + maximum.ToString(CultureInfo.InvariantCulture) + ".");
        }
        return parsed;
    }

    private static OfficeImageExportFormat ParseImageFormat(string value) => value.ToLowerInvariant() switch {
        "png" => OfficeImageExportFormat.Png,
        "jpg" or "jpeg" => OfficeImageExportFormat.Jpeg,
        "webp" => OfficeImageExportFormat.Webp,
        "tif" or "tiff" => OfficeImageExportFormat.Tiff,
        "svg" => OfficeImageExportFormat.Svg,
        _ => throw new WorkflowUsageException("Unknown image format '" + value + "'.")
    };

    private static PageSize ParsePaper(string value) => value.ToLowerInvariant() switch {
        "a4" => PageSizes.A4,
        "letter" => PageSizes.Letter,
        "legal" => PageSizes.Legal,
        "a3" => PageSizes.A3,
        _ => throw new WorkflowUsageException("Paper must be A4, Letter, Legal, or A3.")
    };

    private static PdfPrintOrientation ParseOrientation(string value) => value.ToLowerInvariant() switch {
        "auto" or "automatic" => PdfPrintOrientation.Automatic,
        "portrait" => PdfPrintOrientation.Portrait,
        "landscape" => PdfPrintOrientation.Landscape,
        _ => throw new WorkflowUsageException("Orientation must be auto, portrait, or landscape.")
    };

    private static int ParsePagesPerSheet(string value) {
        int parsed = ParseInt(value, "--pages-per-sheet", 1, 4);
        return parsed is 1 or 2 or 4
            ? parsed
            : throw new WorkflowUsageException("--pages-per-sheet must be 1, 2, or 4.");
    }

    private static PdfPrintScaleMode ParseScale(string value) => value.ToLowerInvariant() switch {
        "fit" => PdfPrintScaleMode.Fit,
        "actual" or "actual-size" => PdfPrintScaleMode.ActualSize,
        "fill" => PdfPrintScaleMode.Fill,
        _ => throw new WorkflowUsageException("Scale must be fit, actual, or fill.")
    };

    private static bool IsHelp(string value) => value is "help" or "--help" or "-h";
}

internal sealed class WorkflowUsageException : Exception {
    internal WorkflowUsageException(string message) : base(message) { }
}
