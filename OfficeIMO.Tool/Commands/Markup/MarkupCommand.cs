using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.Json.Serialization.Metadata;
using OfficeIMO.Markup;
using OfficeIMO.Markup.Excel;
using OfficeIMO.Markup.PowerPoint;
using OfficeIMO.Markup.Word;
using OfficeIMO.Excel;

namespace OfficeIMO.Tool.Commands.Markup;

internal static class MarkupCommand {
    internal static async Task<int> RunAsync(
        string[] args,
        Stream standardInput,
        TextWriter standardOutput,
        TextWriter standardError,
        CancellationToken cancellationToken = default) {
        try {
            ArgumentNullException.ThrowIfNull(args);
            ArgumentNullException.ThrowIfNull(standardInput);
            ArgumentNullException.ThrowIfNull(standardOutput);
            ArgumentNullException.ThrowIfNull(standardError);

            var options = MarkupArguments.Parse(args);
            if (options.ShowHelp) {
                await WriteHelpAsync(standardOutput).ConfigureAwait(false);
                return (int)OfficeImoToolExitCode.Success;
            }

            if (string.IsNullOrWhiteSpace(options.Command)) {
                await WriteHelpAsync(standardOutput).ConfigureAwait(false);
                return (int)OfficeImoToolExitCode.Usage;
            }

            var markup = await ReadMarkupAsync(options, standardInput, cancellationToken).ConfigureAwait(false);
            var result = OfficeMarkupParser.Parse(markup, new OfficeMarkupParserOptions {
                Profile = options.Profile
            });

            switch (options.Command.ToLowerInvariant()) {
                case "parse":
                case "preview":
                    await WriteJsonAsync(new MarkupEnvelope(ToDocumentDto(result.Document), result.Diagnostics.Select(ToDiagnosticDto).ToList()), MarkupJsonSerializerContext.Default.MarkupEnvelope, standardOutput).ConfigureAwait(false);
                    return (int)OfficeImoToolExitCode.Success;
                case "validate":
                    await WriteJsonAsync(new ValidationEnvelope(result.Diagnostics.Select(ToDiagnosticDto).ToList(), result.HasErrors), MarkupJsonSerializerContext.Default.ValidationEnvelope, standardOutput).ConfigureAwait(false);
                    return result.HasErrors
                        ? (int)OfficeImoToolExitCode.ValidationFailed
                        : (int)OfficeImoToolExitCode.Success;
                case "emit":
                    return await EmitAsync(result, options, standardOutput, standardError, cancellationToken).ConfigureAwait(false);
                case "export":
                    return await ExportAsync(result, options, standardOutput, standardError, cancellationToken).ConfigureAwait(false);
                default:
                    await standardError.WriteLineAsync($"Unknown command '{options.Command}'.").ConfigureAwait(false);
                    await WriteHelpAsync(standardError).ConfigureAwait(false);
                    return (int)OfficeImoToolExitCode.Usage;
            }
        } catch (OperationCanceledException) {
            await standardError.WriteLineAsync("Operation cancelled.").ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.Cancelled;
        } catch (FileNotFoundException ex) {
            await standardError.WriteLineAsync(ex.Message).ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.InputNotFound;
        } catch (MarkupInputException ex) {
            await standardError.WriteLineAsync(ex.Message).ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.UnsupportedInput;
        } catch (IOException ex) {
            await standardError.WriteLineAsync(ex.Message).ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.OutputFailed;
        } catch (UnauthorizedAccessException ex) {
            await standardError.WriteLineAsync(ex.Message).ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.OutputFailed;
        } catch (JsonException ex) {
            await standardError.WriteLineAsync(ex.Message).ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.OperationFailed;
        } catch (InvalidOperationException ex) {
            await standardError.WriteLineAsync(ex.Message).ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.OperationFailed;
        } catch (ArgumentException ex) {
            await standardError.WriteLineAsync(ex.Message).ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.Usage;
        } catch (Exception ex) when (ex is not OutOfMemoryException
                                     and not StackOverflowException
                                     and not AccessViolationException
                                     and not AppDomainUnloadedException
                                     and not BadImageFormatException
                                     and not CannotUnloadAppDomainException
                                     and not InvalidProgramException) {
            await standardError.WriteLineAsync(ex.ToString()).ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.OperationFailed;
        }
    }

    private static async Task<int> EmitAsync(
        OfficeMarkupParseResult result,
        MarkupArguments options,
        TextWriter standardOutput,
        TextWriter standardError,
        CancellationToken cancellationToken) {
        if (result.HasErrors) {
            await WriteJsonAsync(
                new ValidationEnvelope(result.Diagnostics.Select(ToDiagnosticDto).ToList(), true),
                MarkupJsonSerializerContext.Default.ValidationEnvelope,
                standardError).ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.ValidationFailed;
        }

        var target = (options.Target ?? "csharp").ToLowerInvariant();
        var text = target switch {
            "csharp" or "cs" => new OfficeMarkupCSharpEmitter().Emit(result.Document),
            "powershell" or "ps" or "ps1" => new OfficeMarkupPowerShellEmitter().Emit(result.Document),
            _ => throw new InvalidOperationException($"Unsupported emit target '{options.Target}'.")
        };

        if (!string.IsNullOrWhiteSpace(options.OutputPath)) {
            var outputPath = NormalizeWritableFilePath(options.OutputPath!);
            await File.WriteAllTextAsync(outputPath, text, cancellationToken).ConfigureAwait(false);
        } else {
            await standardOutput.WriteLineAsync(text).ConfigureAwait(false);
        }

        return (int)OfficeImoToolExitCode.Success;
    }

    private static async Task<int> ExportAsync(
        OfficeMarkupParseResult result,
        MarkupArguments options,
        TextWriter standardOutput,
        TextWriter standardError,
        CancellationToken cancellationToken) {
        if (result.HasErrors) {
            await WriteJsonAsync(
                new ValidationEnvelope(result.Diagnostics.Select(ToDiagnosticDto).ToList(), true),
                MarkupJsonSerializerContext.Default.ValidationEnvelope,
                standardError).ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.ValidationFailed;
        }

        cancellationToken.ThrowIfCancellationRequested();
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

                result.Document.SaveAsPowerPoint(outputPath!, new MarkupToPowerPointOptions {
                    BaseDirectory = inputPath == null
                        ? Directory.GetCurrentDirectory()
                        : Path.GetDirectoryName(inputPath),
                    MermaidRendererPath = options.MermaidRendererPath,
                    RenderMermaidDiagrams = options.RenderMermaidDiagrams
                });
                await WriteJsonAsync(
                    new ExportEnvelope(outputPath!, target),
                    MarkupJsonSerializerContext.Default.ExportEnvelope,
                    standardOutput).ConfigureAwait(false);
                return (int)OfficeImoToolExitCode.Success;
            case "xlsx":
            case "excel":
            case "workbook":
                var workbookOutputPath = options.OutputPath;
                if (string.IsNullOrWhiteSpace(workbookOutputPath)) {
                    throw new InvalidOperationException("Export target 'xlsx' requires --output <file.xlsx>.");
                }

                workbookOutputPath = NormalizeWritableFilePath(workbookOutputPath);

                result.Document.SaveAsExcel(workbookOutputPath!, saveOptions: new ExcelSaveOptions {
                    SafePreflight = options.WorkbookSafePreflight,
                    ValidateOpenXml = options.WorkbookValidateOpenXml,
                    SafeRepairDefinedNames = options.WorkbookRepairDefinedNames
                });
                await WriteJsonAsync(
                    new ExportEnvelope(workbookOutputPath!, target),
                    MarkupJsonSerializerContext.Default.ExportEnvelope,
                    standardOutput).ConfigureAwait(false);
                return (int)OfficeImoToolExitCode.Success;
            case "docx":
            case "word":
            case "document":
                var documentInputPath = ResolveInputFilePath(options);
                var documentOutputPath = options.OutputPath;
                if (string.IsNullOrWhiteSpace(documentOutputPath)) {
                    throw new InvalidOperationException("Export target 'docx' requires --output <file.docx>.");
                }

                documentOutputPath = NormalizeWritableFilePath(documentOutputPath);

                result.Document.SaveAsWord(documentOutputPath!, new MarkupToWordOptions {
                    BaseDirectory = documentInputPath == null
                        ? Environment.CurrentDirectory
                        : Path.GetDirectoryName(documentInputPath)
                });
                await WriteJsonAsync(
                    new ExportEnvelope(documentOutputPath!, target),
                    MarkupJsonSerializerContext.Default.ExportEnvelope,
                    standardOutput).ConfigureAwait(false);
                return (int)OfficeImoToolExitCode.Success;
            default:
                throw new InvalidOperationException($"Unsupported export target '{options.Target}'.");
        }
    }

    private static async Task<string> ReadMarkupAsync(
        MarkupArguments options,
        Stream standardInput,
        CancellationToken cancellationToken) {
        if (options.UseStdin || string.Equals(options.InputPath, "-", StringComparison.Ordinal)) {
            try {
                byte[] bytes = await ReadBoundedAsync(standardInput, options.MaxInputBytes, cancellationToken).ConfigureAwait(false);
                return DecodeUtf8(bytes);
            } catch (IOException ex) {
                throw new MarkupInputException(ex.Message, ex);
            } catch (DecoderFallbackException ex) {
                throw new MarkupInputException("Input is not valid UTF-8.", ex);
            }
        }

        if (!string.IsNullOrWhiteSpace(options.InputPath)) {
            try {
                var inputPath = NormalizeExistingFilePath(options.InputPath!);
                var info = new FileInfo(inputPath);
                if (info.Length > options.MaxInputBytes) {
                    throw new MarkupInputException("Input exceeds the configured byte limit.");
                }
                byte[] bytes = await File.ReadAllBytesAsync(inputPath, cancellationToken).ConfigureAwait(false);
                return DecodeUtf8(bytes);
            } catch (MarkupInputException) {
                throw;
            } catch (FileNotFoundException) {
                throw;
            } catch (IOException ex) {
                throw new MarkupInputException(ex.Message, ex);
            } catch (UnauthorizedAccessException ex) {
                throw new MarkupInputException(ex.Message, ex);
            } catch (DecoderFallbackException ex) {
                throw new MarkupInputException("Input is not valid UTF-8.", ex);
            }
        }

        throw new InvalidOperationException("Input path is required. Use '-' or --stdin to read from standard input.");
    }

    private static string? ResolveInputFilePath(MarkupArguments options) {
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

    private static async Task<byte[]> ReadBoundedAsync(
        Stream input,
        long maximumBytes,
        CancellationToken cancellationToken) {
        var buffer = new byte[8192];
        using var output = new MemoryStream((int)Math.Min(maximumBytes, 64 * 1024));
        while (true) {
            int read = await input.ReadAsync(buffer.AsMemory(), cancellationToken).ConfigureAwait(false);
            if (read == 0) return output.ToArray();
            if (output.Length > maximumBytes - read) {
                throw new IOException("Input exceeds the configured byte limit.");
            }
            await output.WriteAsync(buffer.AsMemory(0, read), cancellationToken).ConfigureAwait(false);
        }
    }

    private static string DecodeUtf8(byte[] bytes) {
        var encoding = new UTF8Encoding(
            encoderShouldEmitUTF8Identifier: false,
            throwOnInvalidBytes: true);
        string text = encoding.GetString(bytes);
        return text.Length > 0 && text[0] == '\uFEFF' ? text.Substring(1) : text;
    }

    private sealed class MarkupInputException : IOException {
        internal MarkupInputException(string message)
            : base(message) {
        }

        internal MarkupInputException(string message, Exception innerException)
            : base(message, innerException) {
        }
    }

    private static Task WriteJsonAsync<T>(T value, JsonTypeInfo<T> typeInfo, TextWriter writer) {
        return writer.WriteLineAsync(JsonSerializer.Serialize(value, typeInfo));
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

    private const string Usage = """
OfficeIMO.Tool - Markup

Usage:
  officeimo markup parse <file|-> [--profile presentation|document|workbook|common]
  officeimo markup validate <file|-> [--profile presentation|document|workbook|common]
  officeimo markup preview <file|-> [--profile presentation|document|workbook|common]
  officeimo markup emit <file|-> --target csharp|powershell [--output <file>]
  officeimo markup export <file|-> --target pptx --output <file.pptx>
  officeimo markup export <file|-> --target xlsx --output <file.xlsx>
  officeimo markup export <file|-> --target docx --output <file.docx>

Options:
  --stdin                       Read markup from standard input.
  --max-input-bytes <bytes>     Bound input size (default: 67108864).
  --no-safe-preflight           Disable Excel safe preflight.
  --no-defined-name-repair      Disable Excel defined-name repair.
  --no-openxml-validation       Disable Excel Open XML validation.
  --mermaid-renderer <mmdc>     Configure the Mermaid renderer for PowerPoint.
  --no-mermaid                  Disable Mermaid rendering.
""";

    private static Task WriteHelpAsync(TextWriter writer) => writer.WriteLineAsync(Usage);
}
