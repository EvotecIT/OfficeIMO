using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Text.Json;
using System.Text.Json.Serialization;
using OfficeIMO.Markup;
using OfficeIMO.Markup.Excel;
using OfficeIMO.Markup.PowerPoint;
using OfficeIMO.Markup.Word;

internal static class Program {
    private static readonly JsonSerializerOptions JsonOptions = new JsonSerializerOptions {
        WriteIndented = true,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
    };

    public static async Task<int> Main(string[] args) {
        try {
            var options = CliOptions.Parse(args);
            if (string.IsNullOrWhiteSpace(options.Command) || options.ShowHelp) {
                WriteHelp();
                return string.IsNullOrWhiteSpace(options.Command) ? 1 : 0;
            }

            var markup = await ReadMarkupAsync(options).ConfigureAwait(false);
            var result = OfficeMarkupParser.Parse(markup, new OfficeMarkupParserOptions {
                Profile = options.Profile
            });

            switch (options.Command.ToLowerInvariant()) {
                case "parse":
                case "preview":
                    WriteJson(new MarkupEnvelope(ToDocumentDto(result.Document), result.Diagnostics.Select(ToDiagnosticDto).ToList()));
                    return 0;
                case "validate":
                    WriteJson(new ValidationEnvelope(result.Diagnostics.Select(ToDiagnosticDto).ToList(), result.HasErrors));
                    return result.HasErrors ? 1 : 0;
                case "emit":
                    return await EmitAsync(result, options).ConfigureAwait(false);
                case "export":
                    return Export(result, options);
                default:
                    Console.Error.WriteLine($"Unknown command '{options.Command}'.");
                    WriteHelp();
                    return 1;
            }
        } catch (IOException ex) {
            Console.Error.WriteLine(ex.Message);
            return 1;
        } catch (UnauthorizedAccessException ex) {
            Console.Error.WriteLine(ex.Message);
            return 1;
        } catch (JsonException ex) {
            Console.Error.WriteLine(ex.Message);
            return 1;
        } catch (InvalidOperationException ex) {
            Console.Error.WriteLine(ex.Message);
            return 1;
        } catch (ArgumentException ex) {
            Console.Error.WriteLine(ex.Message);
            return 1;
        } catch (Exception ex) {
            Console.Error.WriteLine(ex.ToString());
            return 1;
        }
    }

    private static async Task<int> EmitAsync(OfficeMarkupParseResult result, CliOptions options) {
        if (result.HasErrors) {
            WriteJson(new ValidationEnvelope(result.Diagnostics.Select(ToDiagnosticDto).ToList(), true), Console.Error);
            return 1;
        }

        var target = (options.Target ?? "csharp").ToLowerInvariant();
        var text = target switch {
            "csharp" or "cs" => new OfficeMarkupCSharpEmitter().Emit(result.Document),
            "powershell" or "ps" or "ps1" => new OfficeMarkupPowerShellEmitter().Emit(result.Document),
            _ => throw new InvalidOperationException($"Unsupported emit target '{options.Target}'.")
        };

        if (!string.IsNullOrWhiteSpace(options.OutputPath)) {
            var outputPath = NormalizeWritableFilePath(options.OutputPath!);
            await File.WriteAllTextAsync(outputPath, text).ConfigureAwait(false);
        } else {
            Console.WriteLine(text);
        }

        return 0;
    }

    private static int Export(OfficeMarkupParseResult result, CliOptions options) {
        if (result.HasErrors) {
            WriteJson(new ValidationEnvelope(result.Diagnostics.Select(ToDiagnosticDto).ToList(), true), Console.Error);
            return 1;
        }

        var target = (options.Target ?? "pptx").ToLowerInvariant();
        switch (target) {
            case "pptx":
            case "powerpoint":
            case "presentation":
                var inputPath = ResolveInputFilePath(options);
                var outputPath = options.OutputPath;
                if (string.IsNullOrWhiteSpace(outputPath)) {
                    throw new InvalidOperationException("Export target 'pptx' requires --output <file.pptx>.");
                }

                outputPath = NormalizeWritableFilePath(outputPath);

                new OfficeMarkupPowerPointExporter().Export(result.Document, new OfficeMarkupPowerPointExportOptions {
                    OutputPath = outputPath!,
                    BaseDirectory = inputPath == null
                        ? Directory.GetCurrentDirectory()
                        : Path.GetDirectoryName(inputPath),
                    MermaidRendererPath = options.MermaidRendererPath,
                    RenderMermaidDiagrams = options.RenderMermaidDiagrams
                });
                WriteJson(new ExportEnvelope(outputPath!, target));
                return 0;
            case "xlsx":
            case "excel":
            case "workbook":
                var workbookOutputPath = options.OutputPath;
                if (string.IsNullOrWhiteSpace(workbookOutputPath)) {
                    throw new InvalidOperationException("Export target 'xlsx' requires --output <file.xlsx>.");
                }

                workbookOutputPath = NormalizeWritableFilePath(workbookOutputPath);

                new OfficeMarkupExcelExporter().Export(result.Document, new OfficeMarkupExcelExportOptions {
                    OutputPath = workbookOutputPath!,
                    SafePreflight = options.WorkbookSafePreflight,
                    ValidateOpenXml = options.WorkbookValidateOpenXml,
                    SafeRepairDefinedNames = options.WorkbookRepairDefinedNames
                });
                WriteJson(new ExportEnvelope(workbookOutputPath!, target));
                return 0;
            case "docx":
            case "word":
            case "document":
                var documentInputPath = ResolveInputFilePath(options);
                var documentOutputPath = options.OutputPath;
                if (string.IsNullOrWhiteSpace(documentOutputPath)) {
                    throw new InvalidOperationException("Export target 'docx' requires --output <file.docx>.");
                }

                documentOutputPath = NormalizeWritableFilePath(documentOutputPath);

                new OfficeMarkupWordExporter().Export(result.Document, new OfficeMarkupWordExportOptions {
                    OutputPath = documentOutputPath!,
                    BaseDirectory = documentInputPath == null
                        ? Environment.CurrentDirectory
                        : Path.GetDirectoryName(documentInputPath)
                });
                WriteJson(new ExportEnvelope(documentOutputPath!, target));
                return 0;
            default:
                throw new InvalidOperationException($"Unsupported export target '{options.Target}'.");
        }
    }

    private static async Task<string> ReadMarkupAsync(CliOptions options) {
        if (options.UseStdin || string.Equals(options.InputPath, "-", StringComparison.Ordinal)) {
            return await Console.In.ReadToEndAsync().ConfigureAwait(false);
        }

        if (!string.IsNullOrWhiteSpace(options.InputPath)) {
            var inputPath = NormalizeExistingFilePath(options.InputPath!);
            return await File.ReadAllTextAsync(inputPath).ConfigureAwait(false);
        }

        if (Console.IsInputRedirected) {
            return await Console.In.ReadToEndAsync().ConfigureAwait(false);
        }

        throw new InvalidOperationException("Input path is required. Use '-' or --stdin to read from standard input.");
    }

    private static string? ResolveInputFilePath(CliOptions options) {
        if (options.UseStdin || string.Equals(options.InputPath, "-", StringComparison.Ordinal)) {
            return null;
        }

        if (string.IsNullOrWhiteSpace(options.InputPath)) {
            return null;
        }

        return NormalizeExistingFilePath(options.InputPath!);
    }

    private static string NormalizeExistingFilePath(string path) {
        var fullPath = Path.GetFullPath(path);
        if (!File.Exists(fullPath)) {
            throw new FileNotFoundException($"The provided file path does not exist: {fullPath}", fullPath);
        }

        return fullPath;
    }

    private static string NormalizeWritableFilePath(string path) {
        var fullPath = Path.GetFullPath(path);
        var directory = Path.GetDirectoryName(fullPath);
        if (string.IsNullOrWhiteSpace(directory)) {
            throw new InvalidOperationException($"Unable to resolve an output directory for path '{path}'.");
        }

        Directory.CreateDirectory(directory);
        return fullPath;
    }

    private static void WriteJson<T>(T value, TextWriter? writer = null) {
        writer ??= Console.Out;
        writer.WriteLine(JsonSerializer.Serialize(value, JsonOptions));
    }

    private static OfficeMarkupDocumentDto ToDocumentDto(OfficeMarkupDocument document) {
        var styleResolver = OfficeMarkupStyleResolver.Create(document);
        return new OfficeMarkupDocumentDto {
            Profile = document.Profile.ToString(),
            Metadata = new Dictionary<string, string>(document.Metadata, StringComparer.OrdinalIgnoreCase),
            Blocks = document.Blocks.Select(block => ToBlockDto(block, styleResolver)).ToList()
        };
    }

    private static OfficeMarkupBlockDto ToBlockDto(OfficeMarkupBlock block, OfficeMarkupStyleResolver styleResolver) {
        var dto = new OfficeMarkupBlockDto {
            Kind = block.Kind.ToString(),
            Attributes = new Dictionary<string, string>(block.Attributes, StringComparer.OrdinalIgnoreCase),
            SourceText = block.SourceText,
            ResolvedStyle = ToStyleDto(styleResolver.Resolve(block))
        };

        switch (block) {
            case OfficeMarkupHeadingBlock heading:
                dto.Level = heading.Level;
                dto.Text = heading.Text;
                break;
            case OfficeMarkupParagraphBlock paragraph:
                dto.Text = paragraph.Text;
                break;
            case OfficeMarkupListBlock list:
                dto.Ordered = list.Ordered;
                dto.Start = list.Start;
                dto.Items = list.Items.Select(item => new OfficeMarkupListItemDto {
                    Text = item.Text,
                    IsTask = item.IsTask,
                    IsChecked = item.IsChecked,
                    Blocks = item.Blocks.Select(child => ToBlockDto(child, styleResolver)).ToList()
                }).ToList();
                break;
            case OfficeMarkupCodeBlock code:
                dto.Language = code.Language;
                dto.Content = code.Content;
                break;
            case OfficeMarkupImageBlock image:
                dto.Source = image.Source;
                dto.Alt = image.Alt;
                dto.Title = image.Title;
                dto.Width = image.Width;
                dto.Height = image.Height;
                dto.Position = ToPlacementDto(image.Placement);
                break;
            case OfficeMarkupTableBlock table:
                dto.Headers = table.Headers.ToList();
                dto.Rows = table.Rows.Select(row => row.ToList()).ToList();
                break;
            case OfficeMarkupDiagramBlock diagram:
                dto.Language = diagram.Language;
                dto.Content = diagram.Content;
                dto.RenderAsImage = diagram.RenderAsImage;
                dto.Position = ToPlacementDto(diagram.Placement);
                break;
            case OfficeMarkupSlideBlock slide:
                dto.Title = slide.Title;
                dto.Layout = slide.Layout;
                dto.Section = slide.Section;
                dto.Transition = slide.Transition;
                dto.TransitionDetails = ToTransitionDto(slide.Transition);
                dto.Background = slide.Background;
                dto.Notes = slide.Notes;
                dto.Placement = slide.Placement;
                dto.Columns = slide.Columns;
                dto.Blocks = slide.Blocks.Select(child => ToBlockDto(child, styleResolver)).ToList();
                break;
            case OfficeMarkupSectionBlock section:
                dto.Name = section.Name;
                dto.PageSize = section.PageSize;
                dto.Orientation = section.Orientation;
                dto.Blocks = section.Blocks.Select(child => ToBlockDto(child, styleResolver)).ToList();
                break;
            case OfficeMarkupHeaderFooterBlock headerFooter:
                dto.Name = headerFooter.HeaderFooterKind;
                dto.Text = headerFooter.Text;
                break;
            case OfficeMarkupTableOfContentsBlock toc:
                dto.Title = toc.Title;
                dto.MinLevel = toc.MinLevel;
                dto.MaxLevel = toc.MaxLevel;
                break;
            case OfficeMarkupSheetBlock sheet:
                dto.Name = sheet.Name;
                break;
            case OfficeMarkupRangeBlock range:
                dto.Address = range.Address;
                dto.Sheet = range.Sheet;
                dto.Rows = range.Values.Select(row => row.ToList()).ToList();
                break;
            case OfficeMarkupFormulaBlock formula:
                dto.Cell = formula.Cell;
                dto.Expression = formula.Expression;
                dto.Sheet = formula.Sheet;
                break;
            case OfficeMarkupNamedTableBlock namedTable:
                dto.Name = namedTable.Name;
                dto.Range = namedTable.Range;
                dto.HasHeader = namedTable.HasHeader;
                break;
            case OfficeMarkupChartBlock chart:
                dto.ChartType = chart.ChartType;
                dto.Title = chart.Title;
                dto.Source = chart.Source;
                dto.Sheet = chart.Sheet;
                dto.Rows = chart.Data.Select(row => row.ToList()).ToList();
                dto.Position = ToPlacementDto(chart.Placement);
                break;
            case OfficeMarkupTextBoxBlock textBox:
                dto.Text = textBox.Text;
                dto.Style = textBox.Style;
                dto.Position = ToPlacementDto(textBox.Placement);
                break;
            case OfficeMarkupColumnsBlock columns:
                dto.Gap = columns.Gap;
                dto.Position = ToPlacementDto(columns.Placement);
                break;
            case OfficeMarkupColumnBlock column:
                dto.ColumnKind = column.ColumnKind;
                dto.Body = column.Body;
                dto.WidthText = column.Width;
                break;
            case OfficeMarkupCardBlock card:
                dto.Title = card.Title;
                dto.Body = card.Body;
                dto.Style = card.Style;
                dto.Position = ToPlacementDto(card.Placement);
                break;
            case OfficeMarkupFormattingBlock formatting:
                dto.Target = formatting.Target;
                dto.Style = formatting.Style;
                dto.NumberFormat = formatting.NumberFormat;
                break;
            case OfficeMarkupExtensionBlock extension:
                dto.Command = extension.Command;
                dto.Body = extension.Body;
                break;
            case OfficeMarkupRawMarkdownBlock raw:
                dto.Markdown = raw.Markdown;
                break;
        }

        return dto;
    }

    private static OfficeMarkupTransitionDto? ToTransitionDto(string? transition) {
        if (string.IsNullOrWhiteSpace(transition)) {
            return null;
        }

        var resolved = OfficeMarkupTransitionResolver.Parse(transition);
        return new OfficeMarkupTransitionDto {
            RawText = resolved.RawText,
            Effect = resolved.Effect,
            ResolvedIdentifier = resolved.ResolvedIdentifier,
            Attributes = resolved.Attributes.ToDictionary(pair => pair.Key, pair => pair.Value, StringComparer.OrdinalIgnoreCase)
        };
    }

    private static OfficeMarkupDiagnosticDto ToDiagnosticDto(OfficeMarkupDiagnostic diagnostic) =>
        new OfficeMarkupDiagnosticDto {
            Severity = diagnostic.Severity.ToString(),
            Message = diagnostic.Message,
            NodeKind = diagnostic.Node?.Kind.ToString(),
            NodeSourceText = diagnostic.Node?.SourceText
        };

    private static OfficeMarkupPlacementDto? ToPlacementDto(OfficeMarkupPlacement? placement) =>
        placement == null || !placement.HasValue
            ? null
            : new OfficeMarkupPlacementDto {
                X = placement.X,
                Y = placement.Y,
                Width = placement.Width,
                Height = placement.Height
            };

    private static OfficeMarkupResolvedStyleDto? ToStyleDto(OfficeMarkupResolvedStyle? style) =>
        style == null
            ? null
            : new OfficeMarkupResolvedStyleDto {
                Name = style.Name,
                FontName = style.FontName,
                FontSize = style.FontSize,
                Bold = style.Bold,
                Italic = style.Italic,
                TextColor = style.TextColor,
                FillColor = style.FillColor,
                BorderColor = style.BorderColor,
                TextAlign = style.TextAlign
            };

    private static void WriteHelp() {
        Console.WriteLine("OfficeIMO Markup CLI");
        Console.WriteLine();
        Console.WriteLine("Usage:");
        Console.WriteLine("  officeimo-markup parse <file> [--profile presentation|document|workbook|common]");
        Console.WriteLine("  officeimo-markup validate <file> [--profile presentation|document|workbook|common]");
        Console.WriteLine("  officeimo-markup preview <file> [--profile presentation|document|workbook|common]");
        Console.WriteLine("  officeimo-markup emit <file> --target csharp|powershell [--output <file>]");
        Console.WriteLine("  officeimo-markup export <file> --target pptx --output <file.pptx>");
        Console.WriteLine("  officeimo-markup export <file> --target xlsx --output <file.xlsx>");
        Console.WriteLine("  officeimo-markup export <file> --target xlsx --output <file.xlsx> [--no-safe-preflight] [--no-defined-name-repair] [--no-openxml-validation]");
        Console.WriteLine("  officeimo-markup export <file> --target docx --output <file.docx>");
        Console.WriteLine("  officeimo-markup export <file> --target pptx --output <file.pptx> [--mermaid-renderer <mmdc>] [--no-mermaid]");
        Console.WriteLine("  officeimo-markup preview --stdin --profile presentation");
    }
}

internal sealed class CliOptions {
    public string Command { get; private set; } = string.Empty;
    public string? InputPath { get; private set; }
    public string? OutputPath { get; private set; }
    public string? Target { get; private set; }
    public OfficeMarkupProfile Profile { get; private set; } = OfficeMarkupProfile.Document;
    public bool UseStdin { get; private set; }
    public bool ShowHelp { get; private set; }
    public string? MermaidRendererPath { get; private set; }
    public bool RenderMermaidDiagrams { get; private set; } = true;
    public bool WorkbookSafePreflight { get; private set; } = true;
    public bool WorkbookValidateOpenXml { get; private set; } = true;
    public bool WorkbookRepairDefinedNames { get; private set; } = true;

    public static CliOptions Parse(string[] args) {
        var options = new CliOptions();
        var positionals = new List<string>();
        for (int i = 0; i < args.Length; i++) {
            var arg = args[i];
            switch (arg) {
                case "-h":
                case "--help":
                    options.ShowHelp = true;
                    break;
                case "--stdin":
                    options.UseStdin = true;
                    break;
                case "--profile":
                    options.Profile = ParseProfile(ReadValue(args, ref i, arg));
                    break;
                case "--target":
                    options.Target = ReadValue(args, ref i, arg);
                    break;
                case "--output":
                case "-o":
                    options.OutputPath = ReadValue(args, ref i, arg);
                    break;
                case "--mermaid-renderer":
                    options.MermaidRendererPath = ReadValue(args, ref i, arg);
                    break;
                case "--no-mermaid":
                    options.RenderMermaidDiagrams = false;
                    break;
                case "--no-safe-preflight":
                    options.WorkbookSafePreflight = false;
                    break;
                case "--no-openxml-validation":
                    options.WorkbookValidateOpenXml = false;
                    break;
                case "--no-defined-name-repair":
                    options.WorkbookRepairDefinedNames = false;
                    break;
                case "--format":
                    _ = ReadValue(args, ref i, arg);
                    break;
                default:
                    positionals.Add(arg);
                    break;
            }
        }

        if (positionals.Count > 0) {
            options.Command = positionals[0];
        }

        if (positionals.Count > 1) {
            options.InputPath = positionals[1];
        }

        return options;
    }

    private static string ReadValue(string[] args, ref int index, string option) {
        if (index + 1 >= args.Length) {
            throw new InvalidOperationException($"Option '{option}' requires a value.");
        }

        index++;
        return args[index];
    }

    private static OfficeMarkupProfile ParseProfile(string value) {
        if (Enum.TryParse<OfficeMarkupProfile>(value, true, out var profile)) {
            return profile;
        }

        throw new InvalidOperationException($"Unsupported profile '{value}'.");
    }
}

internal sealed class MarkupEnvelope {
    public MarkupEnvelope(OfficeMarkupDocumentDto document, IReadOnlyList<OfficeMarkupDiagnosticDto> diagnostics) {
        Document = document;
        Diagnostics = diagnostics;
    }

    public OfficeMarkupDocumentDto Document { get; }
    public IReadOnlyList<OfficeMarkupDiagnosticDto> Diagnostics { get; }
}

internal sealed class ValidationEnvelope {
    public ValidationEnvelope(IReadOnlyList<OfficeMarkupDiagnosticDto> diagnostics, bool hasErrors) {
        Diagnostics = diagnostics;
        HasErrors = hasErrors;
    }

    public IReadOnlyList<OfficeMarkupDiagnosticDto> Diagnostics { get; }
    public bool HasErrors { get; }
}

internal sealed class ExportEnvelope {
    public ExportEnvelope(string outputPath, string target) {
        OutputPath = outputPath;
        Target = target;
    }

    public string OutputPath { get; }
    public string Target { get; }
}

internal sealed class OfficeMarkupDocumentDto {
    public string Profile { get; set; } = string.Empty;
    public Dictionary<string, string> Metadata { get; set; } = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
    public List<OfficeMarkupBlockDto> Blocks { get; set; } = new List<OfficeMarkupBlockDto>();
}

internal sealed class OfficeMarkupBlockDto {
    public string Kind { get; set; } = string.Empty;
    public Dictionary<string, string> Attributes { get; set; } = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
    public string? SourceText { get; set; }
    public string? Text { get; set; }
    public int? Level { get; set; }
    public bool? Ordered { get; set; }
    public int? Start { get; set; }
    public List<OfficeMarkupListItemDto>? Items { get; set; }
    public string? Language { get; set; }
    public string? Content { get; set; }
    public string? Source { get; set; }
    public string? Alt { get; set; }
    public string? Title { get; set; }
    public double? Width { get; set; }
    public double? Height { get; set; }
    public List<string>? Headers { get; set; }
    public List<List<string>>? Rows { get; set; }
    public bool? RenderAsImage { get; set; }
    public string? Layout { get; set; }
    public string? Section { get; set; }
    public string? Transition { get; set; }
    public string? Background { get; set; }
    public string? Notes { get; set; }
    public string? Placement { get; set; }
    public int? Columns { get; set; }
    public List<OfficeMarkupBlockDto>? Blocks { get; set; }
    public string? Name { get; set; }
    public string? PageSize { get; set; }
    public string? Orientation { get; set; }
    public int? MinLevel { get; set; }
    public int? MaxLevel { get; set; }
    public string? Address { get; set; }
    public string? Sheet { get; set; }
    public string? Cell { get; set; }
    public string? Expression { get; set; }
    public string? Range { get; set; }
    public bool? HasHeader { get; set; }
    public string? ChartType { get; set; }
    public string? Target { get; set; }
    public string? Style { get; set; }
    public string? NumberFormat { get; set; }
    public string? Gap { get; set; }
    public string? ColumnKind { get; set; }
    public string? WidthText { get; set; }
    public OfficeMarkupPlacementDto? Position { get; set; }
    public OfficeMarkupResolvedStyleDto? ResolvedStyle { get; set; }
    public OfficeMarkupTransitionDto? TransitionDetails { get; set; }
    public string? Command { get; set; }
    public string? Body { get; set; }
    public string? Markdown { get; set; }
}

internal sealed class OfficeMarkupTransitionDto {
    public string? RawText { get; set; }
    public string? Effect { get; set; }
    public string? ResolvedIdentifier { get; set; }
    public Dictionary<string, string> Attributes { get; set; } = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
}

internal sealed class OfficeMarkupResolvedStyleDto {
    public string? Name { get; set; }
    public string? FontName { get; set; }
    public int? FontSize { get; set; }
    public bool? Bold { get; set; }
    public bool? Italic { get; set; }
    public string? TextColor { get; set; }
    public string? FillColor { get; set; }
    public string? BorderColor { get; set; }
    public string? TextAlign { get; set; }
}

internal sealed class OfficeMarkupPlacementDto {
    public string? X { get; set; }
    public string? Y { get; set; }
    public string? Width { get; set; }
    public string? Height { get; set; }
}

internal sealed class OfficeMarkupListItemDto {
    public string Text { get; set; } = string.Empty;
    public bool IsTask { get; set; }
    public bool IsChecked { get; set; }
    public List<OfficeMarkupBlockDto> Blocks { get; set; } = new List<OfficeMarkupBlockDto>();
}

internal sealed class OfficeMarkupDiagnosticDto {
    public string Severity { get; set; } = string.Empty;
    public string Message { get; set; } = string.Empty;
    public string? NodeKind { get; set; }
    public string? NodeSourceText { get; set; }
}
