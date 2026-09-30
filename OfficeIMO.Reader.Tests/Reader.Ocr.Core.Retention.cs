using OfficeIMO.Ocr;
using OfficeIMO.Reader;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class ReaderOcrCoreTests {
    [Fact]
    public async Task ApplyOcrAsync_BoundsProviderTextSpansAndConfidenceDiagnostics() {
        OfficeDocumentReadResult source = CreateDocument(1);
        string oversizedHierarchyId = new string('x', 257);
        var engine = new DelegateOcrEngine("bounded-fixture", (request, cancellationToken) => Task.FromResult(new OcrResult {
            Text = "1234567890",
            Confidence = 1.5,
            Spans = new[] {
                new OcrTextSpan {
                    Sequence = 0,
                    Level = OcrTextSpanLevel.Line,
                    Text = "1234567890",
                    Confidence = -0.5,
                    BlockId = oversizedHierarchyId,
                    ParagraphId = oversizedHierarchyId,
                    LineId = oversizedHierarchyId
                },
                new OcrTextSpan { Sequence = 1, Level = OcrTextSpanLevel.Word, Text = "12345" },
                new OcrTextSpan { Sequence = 2, Level = OcrTextSpanLevel.Character, Text = "1" }
            }
        }));

        OfficeDocumentOcrExecutionResult execution = await source.ApplyOcrAsync(engine, new OfficeDocumentOcrExecutionOptions {
            MaxRecognizedCharactersPerCandidate = 5,
            MaxSpansPerCandidate = 2
        });

        Assert.Equal("12345", Assert.Single(execution.Document.Blocks, block => block.Kind == "ocr-text").Text);
        OcrResult result = Assert.Single(execution.Recognitions).Result;
        Assert.Null(result.Confidence);
        Assert.Null(result.Spans[0].Confidence);
        Assert.Null(result.Spans[0].BlockId);
        Assert.Null(result.Spans[0].ParagraphId);
        Assert.Null(result.Spans[0].LineId);
        Assert.Equal(2, result.Spans.Count);
        Assert.Contains(execution.Diagnostics, diagnostic => diagnostic.Code == "ocr-text-limit");
        Assert.Contains(execution.Diagnostics, diagnostic => diagnostic.Code == "ocr-span-limit");
        Assert.Single(execution.Diagnostics, diagnostic => diagnostic.Code == "ocr-confidence-out-of-range");
        Assert.Single(execution.Diagnostics, diagnostic => diagnostic.Code == "ocr-hierarchy-id-limit");
    }

    [Fact]
    public async Task ApplyOcrAsync_BoundsAllRetainedProviderControlledTextAndDiagnostics() {
        OfficeDocumentReadResult source = CreateDocument(1);
        var engine = new DelegateOcrEngine("bounded-output-fixture", (_, _) => Task.FromResult(new OcrResult {
            Text = "recognized",
            Provider = "provider-name",
            Model = "provider-model",
            Language = "provider-language",
            Spans = new[] {
                new OcrTextSpan { Sequence = 0, Level = OcrTextSpanLevel.Word, Text = "abcdef" },
                new OcrTextSpan { Sequence = 1, Level = OcrTextSpanLevel.Word, Text = "ghijkl" }
            },
            Diagnostics = new[] {
                new OcrDiagnostic {
                    Code = "warning",
                    Message = "provider message",
                    Source = "provider source",
                    Attributes = new Dictionary<string, string> {
                        ["key"] = "value",
                        ["second"] = "attribute"
                    }
                },
                new OcrDiagnostic { Code = "second", Message = "discarded" }
            }
        }));

        OfficeDocumentOcrExecutionResult execution = await source.ApplyOcrAsync(engine, new OfficeDocumentOcrExecutionOptions {
            MaxSpanCharactersPerCandidate = 6,
            MaxResultMetadataCharactersPerCandidate = 5,
            MaxProviderDiagnosticsPerCandidate = 1,
            MaxProviderDiagnosticCharactersPerCandidate = 8,
            MaxProviderDiagnosticAttributesPerCandidate = 1,
            MaxProviderDiagnosticAttributeCharactersPerCandidate = 4
        });

        OcrResult result = Assert.Single(execution.Recognitions).Result;
        Assert.Equal("provi", result.Provider);
        Assert.Null(result.Model);
        Assert.Null(result.Language);
        Assert.Equal("abcdef", result.Spans[0].Text);
        Assert.Equal(string.Empty, result.Spans[1].Text);
        OcrDiagnostic diagnostic = Assert.Single(result.Diagnostics);
        Assert.True((diagnostic.Code.Length + diagnostic.Message.Length + (diagnostic.Source?.Length ?? 0)) <= 8);
        KeyValuePair<string, string> attribute = Assert.Single(diagnostic.Attributes);
        Assert.True(attribute.Key.Length + attribute.Value.Length <= 4);
        Assert.Contains(execution.Diagnostics, item => item.Code == "ocr-result-metadata-limit");
        Assert.Contains(execution.Diagnostics, item => item.Code == "ocr-span-text-limit");
        Assert.Contains(execution.Diagnostics, item => item.Code == "ocr-provider-diagnostic-limit");
        Assert.Contains(execution.Diagnostics, item => item.Code == "ocr-provider-diagnostic-text-limit");
        Assert.Contains(execution.Diagnostics, item => item.Code == "ocr-provider-diagnostic-attribute-limit");
        Assert.Contains(execution.Diagnostics, item => item.Code == "ocr-provider-diagnostic-attribute-text-limit");
    }

    [Fact]
    public async Task ApplyOcrAsync_BoundsRawProviderStringsBeforeTrimmingThem() {
        OfficeDocumentReadResult source = CreateDocument(1);
        string padded = new string(' ', 1024) + "unbounded-tail";
        var engine = new DelegateOcrEngine("raw-bounds-fixture", (_, _) => Task.FromResult(new OcrResult {
            Text = padded,
            Provider = padded,
            Spans = new[] {
                new OcrTextSpan {
                    Sequence = 0,
                    Level = OcrTextSpanLevel.Word,
                    Text = padded,
                    BlockId = padded
                }
            },
            Diagnostics = new[] {
                new OcrDiagnostic {
                    Code = padded,
                    Message = padded,
                    Source = padded,
                    Attributes = new Dictionary<string, string> { [padded] = padded }
                }
            }
        }));

        OfficeDocumentOcrExecutionResult execution = await source.ApplyOcrAsync(engine, new OfficeDocumentOcrExecutionOptions {
            MaxRecognizedCharactersPerCandidate = 16,
            MaxSpanCharactersPerCandidate = 16,
            MaxResultMetadataCharactersPerCandidate = 16,
            MaxProviderDiagnosticCharactersPerCandidate = 16,
            MaxProviderDiagnosticAttributeCharactersPerCandidate = 16
        });

        OcrResult result = Assert.Single(execution.Recognitions).Result;
        Assert.Equal(string.Empty, result.Text);
        Assert.Null(result.Provider);
        Assert.Equal(string.Empty, Assert.Single(result.Spans).Text);
        Assert.Null(result.Spans[0].BlockId);
        OcrDiagnostic diagnostic = Assert.Single(result.Diagnostics);
        Assert.Equal(string.Empty, diagnostic.Code);
        Assert.Equal(string.Empty, diagnostic.Message);
        Assert.Null(diagnostic.Source);
        Assert.Contains(execution.Diagnostics, item => item.Code == "ocr-text-limit");
        Assert.Contains(execution.Diagnostics, item => item.Code == "ocr-result-metadata-limit");
        Assert.Contains(execution.Diagnostics, item => item.Code == "ocr-span-text-limit");
        Assert.Contains(execution.Diagnostics, item => item.Code == "ocr-hierarchy-id-limit");
        Assert.Contains(execution.Diagnostics, item => item.Code == "ocr-provider-diagnostic-text-limit");
        Assert.Contains(execution.Diagnostics, item => item.Code == "ocr-provider-diagnostic-attribute-text-limit");
    }

    [Fact]
    public async Task ApplyOcrAsync_DoesNotReportTruncationWhenDiagnosticAttributeBudgetIsExactlyFilled() {
        OfficeDocumentReadResult source = CreateDocument(1);
        var engine = new DelegateOcrEngine("exact-attribute-budget-fixture", (_, _) => Task.FromResult(new OcrResult {
            Text = "recognized",
            Diagnostics = new[] {
                new OcrDiagnostic {
                    Code = "notice",
                    Message = "message",
                    Attributes = new Dictionary<string, string> { ["key"] = "v" }
                }
            }
        }));

        OfficeDocumentOcrExecutionResult execution = await source.ApplyOcrAsync(engine, new OfficeDocumentOcrExecutionOptions {
            MaxProviderDiagnosticAttributesPerCandidate = 1,
            MaxProviderDiagnosticAttributeCharactersPerCandidate = 4
        });

        KeyValuePair<string, string> attribute = Assert.Single(Assert.Single(execution.Recognitions).Result.Diagnostics[0].Attributes);
        Assert.Equal("key", attribute.Key);
        Assert.Equal("v", attribute.Value);
        Assert.DoesNotContain(execution.Diagnostics, item => item.Code == "ocr-provider-diagnostic-attribute-limit");
        Assert.DoesNotContain(execution.Diagnostics, item => item.Code == "ocr-provider-diagnostic-attribute-text-limit");
    }

    [Fact]
    public async Task AttributeCharacterBudgetRetainsExactFitsAcrossDiagnosticDictionaries() {
        var engine = new DelegateOcrEngine("attribute-exact-fit", (_, _) => Task.FromResult(new OcrResult {
            Text = "recognized", Diagnostics = new[] { new OcrDiagnostic {
                Attributes = new Dictionary<string, string> { [""] = "" }
            }, new OcrDiagnostic { Attributes = new Dictionary<string, string> { [""] = "" } },
                new OcrDiagnostic { Attributes = new Dictionary<string, string> { ["a"] = "" } } }
        }));
        var execution = await CreateDocument(1).ApplyOcrAsync(engine, new OfficeDocumentOcrExecutionOptions {
            MaxProviderDiagnosticsPerCandidate = 3, MaxProviderDiagnosticAttributesPerCandidate = 3,
            MaxProviderDiagnosticAttributeCharactersPerCandidate = 1
        });
        OcrResult result = Assert.Single(execution.Recognitions).Result;
        Assert.Equal(3, result.Diagnostics.Sum(item => item.Attributes.Count));
        Assert.Equal("", result.Diagnostics[2].Attributes["a"]);
        Assert.DoesNotContain(execution.Diagnostics, item => item.Code == "ocr-provider-diagnostic-attribute-limit");
        Assert.DoesNotContain(execution.Diagnostics, item => item.Code == "ocr-provider-diagnostic-attribute-text-limit");
    }

}
