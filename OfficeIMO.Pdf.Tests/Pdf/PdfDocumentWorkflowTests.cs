using System.Globalization;
using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfDocumentWorkflowTests {
    [Fact]
    public void PageSelection_ParsesAndSnapshotsCallerRanges() {
        PdfPageSelection parsed = PdfPageSelection.Parse("3,1-2;2..3");

        Assert.Equal(5, parsed.PageCount);
        Assert.Equal("3,1-2,2-3", parsed.ToString());
        Assert.Equal(new[] {
            PdfPageRange.From(3, 3),
            PdfPageRange.From(1, 2),
            PdfPageRange.From(2, 3)
        }, parsed.Ranges);

        var ranges = new[] { PdfPageRange.From(1, 1) };
        PdfPageSelection selection = PdfPageSelection.FromRanges(ranges);
        ranges[0] = PdfPageRange.From(2, 2);

        Assert.Equal("1", selection.ToString());
        Assert.True(PdfPageSelection.TryParse("1,3", out PdfPageSelection? tryParsed));
        Assert.Equal(PdfPageSelection.FromRanges(PdfPageRange.From(1, 1), PdfPageRange.From(3, 3)), tryParsed);
        Assert.False(PdfPageSelection.TryParse(" ", out _));
    }

    [Fact]
    public void KeyValueTable_RendersRichDocumentFactsAndClonesCallerStyle() {
        PdfTableStyle style = TableStyles.Minimal();
        style.HeaderRowCount = 4;

        byte[] bytes = PdfDocument.Create(new PdfOptions {
                CompressContentStreams = false
            })
            .KeyValueTable(new[] {
                PdfKeyValueRow.Rich(
                    new[] { TextRun.Bolded("Invoice") },
                    new[] { TextRun.Normal("FV/2026/001"), TextRun.Bolded(" paid") }),
                PdfKeyValueRow.Text("Customer", "Evotec")
            }, style: style, includeHeader: true, keyHeader: "Field", valueHeader: "Value")
            .ToBytes();

        string raw = Encoding.ASCII.GetString(bytes);
        string text = PdfReadDocument.Open(bytes).ExtractText();

        Assert.Equal(4, style.HeaderRowCount);
        Assert.Contains("Field", text, StringComparison.Ordinal);
        Assert.Contains("Value", text, StringComparison.Ordinal);
        Assert.Contains("Invoice", text, StringComparison.Ordinal);
        Assert.Contains("FV/2026/001 paid", text, StringComparison.Ordinal);
        Assert.Contains("Customer", text, StringComparison.Ordinal);
        Assert.Contains("Evotec", text, StringComparison.Ordinal);
        Assert.Contains("/BaseFont /Helvetica-Bold", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void TableStyle_CanShrinkTextToFitResolvedCellWidth() {
        const string longValue = "ThisIdentifierShouldShrinkToFit";
        byte[] bytes = PdfDocument.Create()
            .Table(new[] {
                new[] { "Name", "Value" },
                new[] { "Alpha", longValue }
            }, style: new PdfTableStyle {
                FontSize = 18,
                HeaderFontSize = 18,
                MinimumShrinkFontSize = 7,
                ShrinkTextToFit = true,
                ColumnWidthPoints = new List<double?> { 54, 108 },
                HeaderRowCount = 1
            })
            .ToBytes();

        IReadOnlyList<PdfLogicalTextBlock> blocks = PdfDocument.Open(bytes).Read.TextBlocks();
        PdfLogicalTextBlock valueBlock = Assert.Single(blocks, block => block.Text.Contains(longValue, StringComparison.Ordinal));

        Assert.True(valueBlock.FontSize < 18D);
        Assert.True(valueBlock.FontSize >= 7D);
    }

    [Fact]
    public void TableStyle_ShrinksExplicitRichRunFontSizesToFitResolvedCellWidth() {
        const string longValue = "ThisIdentifierShouldShrinkToFit";
        byte[] bytes = PdfDocument.Create()
            .Table(new[] {
                new[] { new PdfTableCell("Name"), new PdfTableCell("Value") },
                new[] {
                    new PdfTableCell("Alpha"),
                    PdfTableCell.RichTextCell(new[] {
                        TextRun.Normal(longValue, fontSize: 30, font: PdfStandardFont.Courier)
                    })
                }
            }, style: new PdfTableStyle {
                FontSize = 18,
                HeaderFontSize = 18,
                MinimumShrinkFontSize = 7,
                ShrinkTextToFit = true,
                ColumnWidthPoints = new List<double?> { 54, 108 },
                HeaderRowCount = 1
            })
            .ToBytes();

        PdfDocument document = PdfDocument.Open(bytes);
        IReadOnlyList<PdfLogicalTextBlock> blocks = document.Read.TextBlocks();
        string compactText = document.Read.Text()
            .Replace("\r", string.Empty)
            .Replace("\n", string.Empty)
            .Replace(" ", string.Empty);

        Assert.Contains(longValue, compactText, StringComparison.Ordinal);
        Assert.Contains(blocks, block => block.FontSize < 30D && block.FontSize >= 7D);
    }

    [Fact]
    public void TableStyle_ClampsShrunkExplicitRichRunFontSizesToMinimum() {
        const string longValue = "ThisIdentifierShouldShrinkToFit";
        byte[] bytes = PdfDocument.Create()
            .Table(new[] {
                new[] { new PdfTableCell("Name"), new PdfTableCell("Value") },
                new[] {
                    new PdfTableCell("Alpha"),
                    PdfTableCell.RichTextCell(new[] {
                        TextRun.Normal(longValue, fontSize: 8, font: PdfStandardFont.Courier)
                    })
                }
            }, style: new PdfTableStyle {
                FontSize = 18,
                HeaderFontSize = 18,
                MinimumShrinkFontSize = 7,
                ShrinkTextToFit = true,
                ColumnWidthPoints = new List<double?> { 54, 108 },
                HeaderRowCount = 1
            })
            .ToBytes();

        PdfDocument document = PdfDocument.Open(bytes);
        IReadOnlyList<PdfLogicalTextBlock> blocks = document.Read.TextBlocks();
        string compactText = document.Read.Text()
            .Replace("\r", string.Empty)
            .Replace("\n", string.Empty)
            .Replace(" ", string.Empty);

        Assert.Contains(longValue, compactText, StringComparison.Ordinal);
        Assert.Contains(blocks, block => Math.Abs(block.FontSize - 7D) < 0.001D);
    }

    [Fact]
    public void TableStyle_ShrinksExplicitRichRunsWhenRowFontIsAtMinimum() {
        const string longValue = "ThisIdentifierShouldShrinkToFit";
        byte[] bytes = PdfDocument.Create()
            .Table(new[] {
                new[] { new PdfTableCell("Name"), new PdfTableCell("Value") },
                new[] {
                    new PdfTableCell("Alpha"),
                    PdfTableCell.RichTextCell(new[] {
                        TextRun.Normal(longValue, fontSize: 30, font: PdfStandardFont.Courier)
                    })
                }
            }, style: new PdfTableStyle {
                FontSize = 8,
                HeaderFontSize = 8,
                MinimumShrinkFontSize = 8,
                ShrinkTextToFit = true,
                ColumnWidthPoints = new List<double?> { 54, 108 },
                HeaderRowCount = 1
            })
            .ToBytes();

        PdfDocument document = PdfDocument.Open(bytes);
        IReadOnlyList<PdfLogicalTextBlock> blocks = document.Read.TextBlocks();
        string compactText = document.Read.Text()
            .Replace("\r", string.Empty)
            .Replace("\n", string.Empty)
            .Replace(" ", string.Empty);

        Assert.Contains(longValue, compactText, StringComparison.Ordinal);
        Assert.Contains(blocks, block => Math.Abs(block.FontSize - 8D) < 0.001D);
        Assert.DoesNotContain(blocks, block => block.FontSize > 8.001D && block.Text.Contains(longValue, StringComparison.Ordinal));
    }

    [Fact]
    public void TableStyle_DoesNotEnlargeExplicitRichRunsBelowShrinkMinimum() {
        const string tinyValue = "Tiny";
        const string longValue = "ThisIdentifierShouldShrinkToFit";
        byte[] bytes = PdfDocument.Create()
            .Table(new[] {
                new[] { new PdfTableCell("Name"), new PdfTableCell("Value") },
                new[] {
                    new PdfTableCell("Alpha"),
                    PdfTableCell.RichTextCell(new[] {
                        TextRun.Normal(tinyValue, fontSize: 4, font: PdfStandardFont.Courier),
                        TextRun.Normal(longValue)
                    })
                }
            }, style: new PdfTableStyle {
                FontSize = 18,
                HeaderFontSize = 18,
                MinimumShrinkFontSize = 7,
                ShrinkTextToFit = true,
                ColumnWidthPoints = new List<double?> { 54, 108 },
                HeaderRowCount = 1
            })
            .ToBytes();

        string raw = PdfEncoding.Latin1GetString(bytes);

        Assert.Matches(@"(?s)/F\d+\s+4\s+Tf.*<54696E79>\s+Tj", raw);
    }

    [Fact]
    public void TableStyle_ScalesExplicitRichRunsAgainstRemainingCellWidth() {
        const string prefix = "PrefixConsumesWidth";
        const string largeValue = "WideValue";
        byte[] bytes = PdfDocument.Create()
            .Table(new[] {
                new[] { new PdfTableCell("Name"), new PdfTableCell("Value") },
                new[] {
                    new PdfTableCell("Alpha"),
                    PdfTableCell.RichTextCell(new[] {
                        TextRun.Normal(prefix),
                        TextRun.Normal(largeValue, fontSize: 30, font: PdfStandardFont.Courier)
                    })
                }
            }, style: new PdfTableStyle {
                FontSize = 8,
                HeaderFontSize = 8,
                MinimumShrinkFontSize = 8,
                ShrinkTextToFit = true,
                ColumnWidthPoints = new List<double?> { 54, 108 },
                HeaderRowCount = 1
            })
            .ToBytes();

        PdfDocument document = PdfDocument.Open(bytes);
        IReadOnlyList<PdfLogicalTextBlock> blocks = document.Read.TextBlocks();
        string compactText = document.Read.Text()
            .Replace("\r", string.Empty)
            .Replace("\n", string.Empty)
            .Replace(" ", string.Empty);

        Assert.Contains(prefix + largeValue, compactText, StringComparison.Ordinal);
        Assert.Contains(blocks, block => block.Text.Contains(largeValue, StringComparison.Ordinal) && Math.Abs(block.FontSize - 8D) < 0.001D);
        Assert.DoesNotContain(blocks, block => block.Text.Contains(largeValue, StringComparison.Ordinal) && block.FontSize > 8.001D);
    }

    [Fact]
    public void TableStyle_ChoosesLargestFittingExplicitRichRunScale() {
        const string prefix = "Short";
        const string largeValue = "WideValue";
        byte[] bytes = PdfDocument.Create()
            .Table(new[] {
                new[] { new PdfTableCell("Name"), new PdfTableCell("Value") },
                new[] {
                    new PdfTableCell("Alpha"),
                    PdfTableCell.RichTextCell(new[] {
                        TextRun.Normal(prefix),
                        TextRun.Normal(largeValue, fontSize: 20, font: PdfStandardFont.Courier)
                    })
                }
            }, style: new PdfTableStyle {
                FontSize = 10,
                HeaderFontSize = 10,
                MinimumShrinkFontSize = 6,
                ShrinkTextToFit = true,
                ColumnWidthPoints = new List<double?> { 54, 108 },
                HeaderRowCount = 1
            })
            .ToBytes();

        IReadOnlyList<PdfLogicalTextBlock> blocks = PdfDocument.Open(bytes).Read.TextBlocks();

        Assert.Contains(blocks, block => block.Text.Contains(largeValue, StringComparison.Ordinal) && block.FontSize > 6.001D && block.FontSize < 20D);
    }

    [Fact]
    public void TableStyle_CanShrinkTextToFitInsideRowColumnLayout() {
        const string longValue = "ThisIdentifierShouldShrinkToFit";
        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 240,
                PageHeight = 180
            })
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column =>
                                column.Table(new[] {
                                    new[] { "Name", "Value" },
                                    new[] { "Alpha", longValue }
                                }, style: new PdfTableStyle {
                                    FontSize = 18,
                                    HeaderFontSize = 18,
                                    MinimumShrinkFontSize = 7,
                                    ShrinkTextToFit = true,
                                    ColumnWidthPoints = new List<double?> { 54, 108 },
                                    HeaderRowCount = 1
                                }))))))
            .ToBytes();

        PdfDocument document = PdfDocument.Open(bytes);
        IReadOnlyList<PdfLogicalTextBlock> blocks = document.Read.TextBlocks();
        string compactText = document.Read.Text()
            .Replace("\r", string.Empty)
            .Replace("\n", string.Empty)
            .Replace(" ", string.Empty);

        Assert.Contains(longValue, compactText, StringComparison.Ordinal);
        Assert.Contains(blocks, block => block.FontSize < 18D && block.FontSize >= 7D);
    }

    [Fact]
    public void TableStyle_CanShrinkTextToFitInsideCanvasTable() {
        const string longValue = "ThisIdentifierShouldShrinkToFit";
        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 240,
                PageHeight = 180
            })
            .Canvas(canvas => canvas.Table(new[] {
                new[] { "Name", "Value" },
                new[] { "Alpha", longValue }
            }, 20, 20, 162, 80, new PdfTableStyle {
                FontSize = 18,
                HeaderFontSize = 18,
                MinimumShrinkFontSize = 7,
                ShrinkTextToFit = true,
                ColumnWidthPoints = new List<double?> { 54, 108 },
                HeaderRowCount = 1
            }))
            .ToBytes();

        IReadOnlyList<PdfLogicalTextBlock> blocks = PdfDocument.Open(bytes).Read.TextBlocks();
        PdfLogicalTextBlock valueBlock = Assert.Single(blocks, block => block.Text.Contains(longValue, StringComparison.Ordinal));

        Assert.True(valueBlock.FontSize < 18D);
        Assert.True(valueBlock.FontSize >= 7D);
    }

    [Fact]
    public void PageSizes_ResolveExpandedStandardNames() {
        Assert.True(PageSizes.TryGet("a0", out PageSize a0));
        Assert.Equal(2384, a0.Width);
        Assert.Equal(3370, a0.Height);

        Assert.True(PageSizes.TryGet("B10", out PageSize b10));
        Assert.Equal(88, b10.Width);
        Assert.Equal(124, b10.Height);

        Assert.Equal(PageSizes.Executive, PageSizes.Get("Executive"));
        Assert.Equal(PageSizes.Ledger, PageSizes.Get("LedgerOrTabloid"));
        Assert.Contains("Tabloid", PageSizes.Names);
        Assert.False(PageSizes.TryGet("Unknown", out _));
    }

    [Fact]
    public void FileVersion_CanEmitPdf20Header() {
        byte[] bytes = PdfDocument.Create(new PdfOptions {
                FileVersion = PdfFileVersion.Pdf20
            })
            .Paragraph(paragraph => paragraph.Text("PDF 2.0"))
            .ToBytes();

        Assert.StartsWith("%PDF-2.0", Encoding.ASCII.GetString(bytes), StringComparison.Ordinal);
        Assert.Equal("2.0", PdfDocument.Open(bytes).Inspect().HeaderVersion);
    }

    [Fact]
    public void Open_SnapshotsInputBytesAndExposesReadInspectAndPreflight() {
        byte[] source = BuildThreePagePdf();
        byte[] callerBuffer = (byte[])source.Clone();

        PdfDocument document = PdfDocument.Open(callerBuffer);
        callerBuffer[20] ^= 0x10;

        Assert.Equal(3, document.Inspect().PageCount);
        Assert.Equal("Workflow source", document.Inspect().Metadata.Title);
        Assert.Equal("Workflow source", document.Read.DocumentInfo().Metadata.Title);
        Assert.Equal("Workflow source", document.Read.Metadata().Title);
        Assert.Equal(3, document.Read.Pages().Count);
        Assert.Equal(2, document.Read.Page(2)?.PageNumber);
        Assert.Null(document.Read.Page(4));
        Assert.False(document.Read.Security().HasEncryption);
        Assert.False(document.Read.Security().HasSignatures);
        Assert.Equal(document.Inspect().HeaderVersion, document.Read.HeaderVersion());
        Assert.Equal(document.Inspect().EffectiveVersion, document.Read.EffectiveVersion());
        Assert.False(document.Read.IsPdf20OrLater());
        Assert.True(document.Read.TryDocumentInfo().Succeeded);
        Assert.True(document.Read.TryMetadata().Succeeded);
        Assert.True(document.Read.TryPages().Succeeded);
        Assert.True(document.Read.TrySecurity().Succeeded);
        Assert.Equal(PdfTextExtractor.ExtractAllText(source), document.Read.Text());
        Assert.Equal(PdfTextExtractor.ExtractTextByPage(source), document.Read.TextByPage());
        Assert.True(document.Preflight().CanRead);
        Assert.True(document.Preflight().CanRewrite);

        PdfSignatureValidationReport unsignedSignatures = document.ValidateSignatures();
        PdfAppendOnlyMutationReport unsignedMutation = document.AnalyzeAppendOnlyMutation();

        Assert.False(unsignedSignatures.HasSignatures);
        Assert.Equal("Unsigned", unsignedSignatures.ProofStatus);
        Assert.True(unsignedMutation.CanAppendMetadata);
        Assert.True(unsignedMutation.CanPrepareExternalSignature);
    }

    [Fact]
    public void Open_SeekableStreamReadsCompletePdfAndRestoresPosition() {
        using var stream = new MemoryStream(BuildThreePagePdf());
        long originalPosition = stream.Length;
        stream.Position = originalPosition;

        PdfDocument document = PdfDocument.Open(stream);

        Assert.Equal(3, document.Inspect().PageCount);
        Assert.Equal(originalPosition, stream.Position);
    }

    [Fact]
    public void Open_ExposesCatalogMetadataThroughFluentReader() {
        byte[] source = PdfDocument.Create(new PdfOptions()
                .SetSrgbOutputIntent()
                .SetPdfAIdentification(3, "B")
                .SetPdfUaIdentification())
            .TaggedPdfCatalogMarkers()
            .Language("en-US")
            .Meta(title: "Reader catalog metadata", author: "OfficeIMO", subject: "Reader inspection", keywords: "xmp, intent")
            .Paragraph(paragraph => paragraph.Text("Reader catalog metadata workflow"))
            .ToBytes();

        PdfDocument document = PdfDocument.Open(source);

        PdfXmpMetadataInfo xmp = Assert.IsType<PdfXmpMetadataInfo>(document.Read.XmpMetadata());
        Assert.True(xmp.IsWellFormedXml);
        Assert.Equal("Reader catalog metadata", xmp.Title);
        Assert.Equal("OfficeIMO", xmp.Creator);
        Assert.Equal(3, xmp.PdfAPart);
        Assert.Equal("B", xmp.PdfAConformance);
        Assert.Equal(1, xmp.PdfUaPart);
        Assert.True(document.Read.TryXmpMetadata().Succeeded);

        PdfOutputIntentInfo outputIntent = Assert.Single(document.Read.OutputIntents());
        Assert.Equal("GTS_PDFA1", outputIntent.Subtype);
        Assert.Equal(PdfIccProfiles.SrgbIec6196621OutputConditionIdentifier, outputIntent.OutputConditionIdentifier);
        Assert.True(outputIntent.HasDestinationOutputProfile);
        Assert.Single(document.Read.OutputIntentsBySubtype("GTS_PDFA1"));
        Assert.Single(document.Read.OutputIntentsByOutputConditionIdentifier(PdfIccProfiles.SrgbIec6196621OutputConditionIdentifier));
        Assert.Empty(document.Read.OutputIntentsBySubtype("GTS_PDFX"));
        Assert.True(document.Read.TryOutputIntents().Succeeded);
        Assert.True(document.Read.TryOutputIntentsBySubtype("GTS_PDFA1").Succeeded);
        Assert.True(document.Read.TryOutputIntentsByOutputConditionIdentifier(PdfIccProfiles.SrgbIec6196621OutputConditionIdentifier).Succeeded);

        PdfTaggedContentInfo tagged = Assert.IsType<PdfTaggedContentInfo>(document.Read.TaggedContent());
        Assert.True(tagged.Marked);
        Assert.Contains("Document", tagged.StructureTypes);
        Assert.Contains("P", tagged.StructureTypes);
        Assert.True(document.Read.TryTaggedContent().Succeeded);
    }

    [Fact]
    public void Open_ExposesOptionalContentAndAttachmentMetadataThroughFluentReader() {
        PdfDocument optionalContentDocument = PdfDocument.Open(PdfOptionalContentSupport.BuildOptionalContentMetadataPdf());

        PdfOptionalContentProperties optionalContent = Assert.IsType<PdfOptionalContentProperties>(optionalContentDocument.Read.OptionalContent());
        Assert.Equal("Default layers", optionalContent.DefaultConfigurationName);
        Assert.Equal(2, optionalContent.GroupCount);

        IReadOnlyList<PdfOptionalContentGroup> groups = optionalContentDocument.Read.OptionalContentGroups();
        Assert.Equal(new[] { "Print layer", "Hidden layer" }, groups.Select(group => group.Name).ToArray());
        PdfOptionalContentGroup printLayer = Assert.Single(optionalContentDocument.Read.OptionalContentGroupsByName("Print layer"));
        Assert.Equal("Print layer", printLayer.Name);
        Assert.Empty(optionalContentDocument.Read.OptionalContentGroupsByName("Missing"));
        Assert.True(optionalContentDocument.Read.TryOptionalContent().Succeeded);
        Assert.True(optionalContentDocument.Read.TryOptionalContentGroups().Succeeded);
        Assert.True(optionalContentDocument.Read.TryOptionalContentGroupsByName("Hidden layer").Succeeded);

        byte[] payload = Encoding.UTF8.GetBytes("attachment metadata payload");
        PdfDocument attachmentDocument = PdfDocument.Open(PdfDocument.Create()
            .AttachFile("payload.txt", payload, "text/plain", PdfAssociatedFileRelationship.Data, "Workflow attachment")
            .Paragraph(paragraph => paragraph.Text("Attachment metadata workflow"))
            .ToBytes());

        PdfAttachmentInfo attachment = Assert.Single(attachmentDocument.Read.AttachmentMetadata());
        Assert.Equal("payload.txt", attachment.Name);
        Assert.Equal("payload.txt", attachment.FileName);
        Assert.Equal("text/plain", attachment.MimeType);
        Assert.Equal(PdfAssociatedFileRelationship.Data, attachment.Relationship);
        Assert.False(attachment.IsAssociatedFile);
        Assert.Single(attachmentDocument.Read.AttachmentMetadataByName("payload.txt"));
        Assert.Single(attachmentDocument.Read.AttachmentMetadataByFileName("payload.txt"));
        Assert.Single(attachmentDocument.Read.AttachmentMetadataBySource("Names/EmbeddedFiles"));
        Assert.Single(attachmentDocument.Read.AttachmentMetadataByRelationship(PdfAssociatedFileRelationship.Data));
        Assert.Empty(attachmentDocument.Read.AttachmentMetadataBySource("AF"));
        Assert.True(attachmentDocument.Read.TryAttachmentMetadata().Succeeded);
        Assert.True(attachmentDocument.Read.TryAttachmentMetadataByName("payload.txt").Succeeded);
        Assert.True(attachmentDocument.Read.TryAttachmentMetadataByFileName("payload.txt").Succeeded);
        Assert.True(attachmentDocument.Read.TryAttachmentMetadataBySource("Names/EmbeddedFiles").Succeeded);
        Assert.True(attachmentDocument.Read.TryAttachmentMetadataByRelationship(PdfAssociatedFileRelationship.Data).Succeeded);
    }

    [Fact]
    public void Open_ExposesSignatureValidationAndAppendOnlyMutationPolicy() {
        byte[] source = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("External signing workflow"))
            .ToBytes();
        PdfExternalSignaturePreparation preparation = PdfIncrementalUpdater.PrepareExternalSignature(
            source,
            new PdfExternalSignatureOptions {
                FieldName = "Approval",
                Name = "Alice",
                Reason = "Approval",
                ReservedSignatureContentsBytes = 256,
                SigningTime = new DateTimeOffset(2026, 6, 22, 12, 0, 0, TimeSpan.Zero)
            });

        PdfDocument document = PdfDocument.Open(preparation.PreparedPdf);

        PdfSignatureValidationReport signatures = document.ValidateSignatures();
        PdfAppendOnlyMutationReport mutation = document.AnalyzeAppendOnlyMutation();

        Assert.True(signatures.HasSignatures);
        Assert.True(signatures.IsStructurallyValid);
        Assert.True(signatures.RequiresAppendOnlyMutation);
        Assert.False(signatures.CryptographicTrustVerified);
        Assert.Equal("ExternalCryptoValidationRequired", signatures.ProofStatus);
        Assert.Contains(signatures.Findings, finding => finding.Code == "SignatureDetachedCmsSubFilter");
        Assert.Contains(signatures.Findings, finding => finding.Code == "AcroFormAppendOnly");

        Assert.True(mutation.RequiresAppendOnlyMutation);
        Assert.False(mutation.CanAppendMetadata);
        Assert.False(mutation.CanPrepareExternalSignature);
        Assert.Contains("Signed", mutation.Blockers);
        Assert.Contains("SignaturePrepare", mutation.BlockedActions);
        Assert.Contains("AcroFormAppendOnly", mutation.Warnings);
    }

    [Fact]
    public void PageSelection_DrivesPageAndReadWorkflows() {
        byte[] source = BuildThreePagePdf();
        PdfPageSelection selection = PdfPageSelection.Parse("3,1-2");

        PdfDocument extracted = PdfDocument.Open(source).Pages.Extract(selection);
        Assert.Equal(PdfDocument.Open(source).Pages.Extract("3,1-2").ToBytes(), extracted.ToBytes());
        Assert.Equal(3, extracted.Inspect().PageCount);
        Assert.Contains("Page C", extracted.Read.Text(), StringComparison.Ordinal);

        Assert.Equal(
            PdfDocument.Open(source).Pages.Delete("2").ToBytes(),
            PdfDocument.Open(source).Pages.Delete(PdfPageSelection.From(2)).ToBytes());

        Assert.Equal(
            PdfDocument.Open(source).Pages.Reorder("2,3,1").ToBytes(),
            PdfDocument.Open(source).Pages.Reorder(PdfPageSelection.Parse("2,3,1")).ToBytes());

        Assert.Equal(
            PdfDocument.Open(source).Pages.Duplicate("2").ToBytes(),
            PdfDocument.Open(source).Pages.Duplicate(PdfPageSelection.From(PdfPageRange.From(2, 2))).ToBytes());

        Assert.Equal(
            PdfDocument.Open(source).Pages.Move(1, "3").ToBytes(),
            PdfDocument.Open(source).Pages.Move(1, PdfPageSelection.From(3)).ToBytes());

        Assert.Equal(
            PdfDocument.Open(source).Pages.Rotate(90, "2").ToBytes(),
            PdfDocument.Open(source).Pages.Rotate(90, PdfPageSelection.Parse("2")).ToBytes());

        IReadOnlyList<PdfDocument> splitRanges = PdfDocument.Open(source).Pages.Split("1,3");
        Assert.Equal(2, splitRanges.Count);
        Assert.Contains("Page A", splitRanges[0].Read.Text(), StringComparison.Ordinal);
        Assert.Contains("Page C", splitRanges[1].Read.Text(), StringComparison.Ordinal);
        Assert.Equal(
            PdfDocument.Open(source).Pages.Split(new[] { PdfPageRange.From(1, 1), PdfPageRange.From(3, 3) })[1].ToBytes(),
            splitRanges[1].ToBytes());

        PdfDocument opened = PdfDocument.Open(source);
        Assert.Equal(PdfTextExtractor.ExtractAllTextByPageRanges(source, PdfPageRange.ParseMany("2,1")), opened.Read.Text(PdfPageSelection.Parse("2,1")));
        Assert.Equal(PdfTextExtractor.ExtractTextByPageRanges(source, PdfPageRange.ParseMany("2,1")), opened.Read.TextByPage(PdfPageSelection.Parse("2,1")));
        Assert.Equal(2, opened.Read.Logical(PdfPageSelection.Parse("2,1")).Pages.Count);
        Assert.Contains("Second page body", opened.Read.Markdown(PdfPageSelection.Parse("2")), StringComparison.Ordinal);
    }

    [Fact]
    public void OperationResult_PreflightsPageOperationsAndCarriesDiagnostics() {
        byte[] source = BuildThreePagePdf();

        PdfOperationResult<PdfDocument> extracted = PdfDocument.Open(source).Pages.TryExtract(PdfPageSelection.Parse("2"));
        Assert.True(extracted.CanAttempt);
        Assert.True(extracted.Succeeded);
        Assert.Empty(extracted.Diagnostics);
        Assert.Equal(PdfPreflightCapability.ManipulatePages, extracted.Capability);
        Assert.Contains("Page B", extracted.RequireValue().Read.Text(), StringComparison.Ordinal);

        PdfOperationResult<PdfDocument> stringExtract = PdfDocument.Open(source).Pages.TryExtract("2");
        Assert.True(stringExtract.Succeeded);
        Assert.Contains("Page B", stringExtract.RequireValue().Read.Text(), StringComparison.Ordinal);

        PdfOperationResult<PdfDocument> stringDelete = PdfDocument.Open(source).Pages.TryDelete("2");
        Assert.True(stringDelete.Succeeded);
        Assert.DoesNotContain("Page B", stringDelete.RequireValue().Read.Text(), StringComparison.Ordinal);

        PdfOperationResult<PdfDocument> stringReorder = PdfDocument.Open(source).Pages.TryReorder("2,3,1");
        Assert.True(stringReorder.Succeeded);
        Assert.Contains("Page A", stringReorder.RequireValue().Read.Text(), StringComparison.Ordinal);

        PdfOperationResult<PdfDocument> stringDuplicate = PdfDocument.Open(source).Pages.TryDuplicate("2");
        Assert.True(stringDuplicate.Succeeded);
        Assert.Equal(4, stringDuplicate.RequireValue().Inspect().PageCount);

        PdfOperationResult<PdfDocument> stringMove = PdfDocument.Open(source).Pages.TryMove(1, "3");
        Assert.True(stringMove.Succeeded);
        Assert.Equal(3, stringMove.RequireValue().Inspect().PageCount);

        PdfOperationResult<PdfDocument> stringRotate = PdfDocument.Open(source).Pages.TryRotate(90, "2");
        Assert.True(stringRotate.Succeeded);
        Assert.Equal(3, stringRotate.RequireValue().Inspect().PageCount);

        PdfOperationResult<PdfDocument> malformedRange = PdfDocument.Open(source).Pages.TryExtract("2-");
        Assert.True(malformedRange.CanAttempt);
        Assert.False(malformedRange.Succeeded);
        Assert.NotNull(malformedRange.Exception);
        Assert.Contains("Page number cannot be empty", string.Join(" ", malformedRange.Diagnostics), StringComparison.Ordinal);

        PdfOperationResult<IReadOnlyList<PdfDocument>> split = PdfDocument.Open(source).Pages.TrySplit();
        Assert.True(split.Succeeded);
        Assert.Equal(3, split.RequireValue().Count);

        PdfOperationResult<IReadOnlyList<PdfDocument>> stringSplit = PdfDocument.Open(source).Pages.TrySplit("1,3");
        Assert.True(stringSplit.CanAttempt);
        Assert.True(stringSplit.Succeeded);
        Assert.Equal(2, stringSplit.RequireValue().Count);
        Assert.Contains("Page C", stringSplit.RequireValue()[1].Read.Text(), StringComparison.Ordinal);

        PdfOperationResult<IReadOnlyList<PdfDocument>> emptySelectionSplit = PdfDocument.Open(source).Pages.TrySplit(Array.Empty<PdfPageSelection>());
        Assert.False(emptySelectionSplit.Succeeded);
        Assert.Null(emptySelectionSplit.Value);
        Assert.Contains("At least one page selection", string.Join(" ", emptySelectionSplit.Diagnostics), StringComparison.Ordinal);

        PdfOperationResult<IReadOnlyList<PdfDocument>> emptyRangeSplit = PdfDocument.Open(source).Pages.TrySplit(Array.Empty<PdfPageRange>());
        Assert.False(emptyRangeSplit.Succeeded);
        Assert.Null(emptyRangeSplit.Value);
        Assert.Contains("At least one page range", string.Join(" ", emptyRangeSplit.Diagnostics), StringComparison.Ordinal);

        PdfOperationResult<IReadOnlyList<PdfDocument>> malformedSplit = PdfDocument.Open(source).Pages.TrySplit("2-");
        Assert.True(malformedSplit.CanAttempt);
        Assert.False(malformedSplit.Succeeded);
        Assert.NotNull(malformedSplit.Exception);
        Assert.Contains("Page number cannot be empty", string.Join(" ", malformedSplit.Diagnostics), StringComparison.Ordinal);

        PdfDocument invalid = PdfDocument.Open(Encoding.ASCII.GetBytes("not a pdf"));
        PdfOperationResult<PdfDocument> blocked = invalid.Pages.TryExtract(PdfPageSelection.From(1));

        Assert.False(blocked.CanAttempt);
        Assert.False(blocked.Succeeded);
        Assert.Null(blocked.Value);
        Assert.NotEmpty(blocked.Diagnostics);
        Assert.NotNull(blocked.Preflight);
        Assert.Throws<InvalidOperationException>(() => blocked.RequireValue());
    }

    [Fact]
    public void OperationResult_ExtendsAcrossMergeReadStampAndForms() {
        byte[] source = BuildThreePagePdf();
        byte[] appendix = BuildPdf("Appendix", "Appendix body");

        PdfOperationResult<PdfDocument> merged = PdfDocument.Open(source).TryMergeWith(PdfDocument.Open(appendix));
        Assert.True(merged.Succeeded);
        Assert.Equal(4, merged.RequireValue().Inspect().PageCount);

        PdfDocument opened = PdfDocument.Open(source);
        PdfOperationResult<string> text = opened.Read.TryText(PdfPageSelection.Parse("2"));
        Assert.True(text.Succeeded);
        Assert.Contains("Second page body", text.RequireValue(), StringComparison.Ordinal);

        PdfOperationResult<PdfLogicalDocument> logical = opened.Read.TryLogical(PdfPageSelection.Parse("1,3"));
        Assert.True(logical.Succeeded);
        Assert.Equal(2, logical.RequireValue().Pages.Count);

        PdfOperationResult<string> markdown = opened.Read.TryMarkdown(PdfPageSelection.Parse("1"));
        Assert.True(markdown.Succeeded);
        Assert.Contains("First page body", markdown.RequireValue(), StringComparison.Ordinal);

        PdfOperationResult<string> stringText = opened.Read.TryText("2");
        Assert.True(stringText.Succeeded);
        Assert.Contains("Second page body", stringText.RequireValue(), StringComparison.Ordinal);

        PdfOperationResult<IReadOnlyList<string>> stringTextByPage = opened.Read.TryTextByPage("3,1");
        Assert.True(stringTextByPage.Succeeded);
        Assert.Equal(2, stringTextByPage.RequireValue().Count);

        PdfOperationResult<string> stringMarkdown = opened.Read.TryMarkdown("1");
        Assert.True(stringMarkdown.Succeeded);
        Assert.Contains("First page body", stringMarkdown.RequireValue(), StringComparison.Ordinal);

        PdfOperationResult<PdfLogicalDocument> stringLogical = opened.Read.TryLogical("1,3");
        Assert.True(stringLogical.Succeeded);
        Assert.Equal(2, stringLogical.RequireValue().Pages.Count);

        PdfOperationResult<IReadOnlyList<PdfLogicalTextBlock>> stringTextBlocks = opened.Read.TryTextBlocks("2");
        Assert.True(stringTextBlocks.Succeeded);
        Assert.Contains(stringTextBlocks.RequireValue(), block => block.Text.Contains("Second page body", StringComparison.Ordinal));

        byte[] imagePdf = PdfStamper.StampImage(source, PdfPngTestImages.CreateRgbPng(255, 0, 0), new PdfImageStampOptions {
            PageNumbers = new[] { 1, 3 },
            Width = 24,
            Height = 24
        });
        PdfDocument imageDocument = PdfDocument.Open(imagePdf);
        IReadOnlyList<PdfExtractedImage> selectedImages = imageDocument.Read.Images("3,1-2");
        Assert.Equal(2, selectedImages.Count);
        Assert.Equal(3, selectedImages[0].PageNumber);
        Assert.Equal(1, selectedImages[1].PageNumber);

        PdfOperationResult<IReadOnlyList<PdfExtractedImage>> stringImages = imageDocument.Read.TryImages("3,1-2");
        Assert.True(stringImages.Succeeded);
        Assert.Equal(2, stringImages.RequireValue().Count);

        IReadOnlyList<PdfImagePlacement> selectedPlacements = imageDocument.Read.ImagePlacements("3,1-2");
        Assert.Equal(2, selectedPlacements.Count);
        Assert.Equal(3, selectedPlacements[0].PageNumber);
        Assert.Equal(1, selectedPlacements[1].PageNumber);
        Assert.True(selectedPlacements[0].Width > 0);
        Assert.True(selectedPlacements[0].Height > 0);

        PdfOperationResult<IReadOnlyList<PdfImagePlacement>> stringPlacements = imageDocument.Read.TryImagePlacements("3,1-2");
        Assert.True(stringPlacements.Succeeded);
        Assert.Equal(2, stringPlacements.RequireValue().Count);

        PdfOperationResult<IReadOnlyList<PdfImagePlacement>> malformedPlacementRange = imageDocument.Read.TryImagePlacements("2-");
        Assert.True(malformedPlacementRange.CanAttempt);
        Assert.False(malformedPlacementRange.Succeeded);
        Assert.NotNull(malformedPlacementRange.Exception);
        Assert.Contains("Page number cannot be empty", string.Join(" ", malformedPlacementRange.Diagnostics), StringComparison.Ordinal);

        PdfOperationResult<IReadOnlyList<PdfExtractedImage>> malformedImageRange = imageDocument.Read.TryImages("2-");
        Assert.True(malformedImageRange.CanAttempt);
        Assert.False(malformedImageRange.Succeeded);
        Assert.NotNull(malformedImageRange.Exception);
        Assert.Contains("Page number cannot be empty", string.Join(" ", malformedImageRange.Diagnostics), StringComparison.Ordinal);

        PdfOperationResult<string> malformedRange = opened.Read.TryText("2-");
        Assert.True(malformedRange.CanAttempt);
        Assert.False(malformedRange.Succeeded);
        Assert.NotNull(malformedRange.Exception);
        Assert.Contains("Page number cannot be empty", string.Join(" ", malformedRange.Diagnostics), StringComparison.Ordinal);

        PdfOperationResult<PdfDocument> stamped = opened.Stamp.TryText("Reviewed", new PdfTextStampOptions { X = 72, Y = 72 });
        Assert.True(stamped.Succeeded);
        Assert.Equal(3, stamped.RequireValue().Inspect().PageCount);

        byte[] formPdf = BuildSimpleFormPdf();
        PdfDocument formDocument = PdfDocument.Open(formPdf);
        PdfFormField formField = Assert.Single(formDocument.Read.FormFields());
        Assert.Equal("Person.Name", formField.Name);
        Assert.Equal(PdfFormFieldKind.Text, formField.Kind);
        Assert.Equal("Original", formField.Value);
        Assert.Equal("Person.Name", formDocument.Read.FormField("Person.Name")?.Name);
        Assert.Equal("Person.Name", Assert.Single(formDocument.Read.FormFields("Person.Name")).Name);
        Assert.Equal("Person.Name", Assert.Single(formDocument.Read.FormFields(PdfFormFieldKind.Text)).Name);
        Assert.Equal("Person.Name", Assert.Single(formDocument.Read.FormFields(1)).Name);

        PdfLogicalFormWidget widget = Assert.Single(formDocument.Read.FormWidgets());
        Assert.Equal("Person.Name", widget.FieldName);
        Assert.Equal(1, widget.PageNumber);
        Assert.True(widget.Width > 0);
        Assert.True(widget.Height > 0);
        Assert.Equal("Person.Name", Assert.Single(formDocument.Read.FormWidgets("Person.Name")).FieldName);
        Assert.Equal("Person.Name", Assert.Single(formDocument.Read.FormWidgets(1)).FieldName);

        PdfOperationResult<IReadOnlyList<PdfFormField>> safeFormFields = formDocument.Read.TryFormFields("Person.Name");
        Assert.True(safeFormFields.Succeeded);
        Assert.Equal("Person.Name", Assert.Single(safeFormFields.RequireValue()).Name);

        PdfOperationResult<IReadOnlyList<PdfLogicalFormWidget>> safeFormWidgets = formDocument.Read.TryFormWidgets("Person.Name");
        Assert.True(safeFormWidgets.Succeeded);
        Assert.Equal("Person.Name", Assert.Single(safeFormWidgets.RequireValue()).FieldName);

        PdfOperationResult<PdfDocument> filled = PdfDocument.Open(formPdf).Forms.TryFill(new Dictionary<string, string> {
            ["Person.Name"] = "Ada Lovelace"
        });
        Assert.True(filled.Succeeded);
        Assert.Equal("Ada Lovelace", Assert.Single(filled.RequireValue().Inspect().FormFields).Value);

        PdfOperationResult<PdfDocument> flattened = PdfDocument.Open(formPdf).Forms.TryFillAndFlatten(new Dictionary<string, string> {
            ["Person.Name"] = "Ada Lovelace"
        });
        Assert.True(flattened.Succeeded);
        Assert.Empty(flattened.RequireValue().Inspect().FormFields);

        PdfDocument invalid = PdfDocument.Open(Encoding.ASCII.GetBytes("not a pdf"));
        PdfOperationResult<string> blockedText = invalid.Read.TryText();
        Assert.False(blockedText.CanAttempt);
        Assert.NotEmpty(blockedText.Diagnostics);

        PdfOperationResult<PdfDocument> blockedStamp = invalid.Stamp.TryText("Reviewed");
        Assert.False(blockedStamp.CanAttempt);
        Assert.NotEmpty(blockedStamp.Diagnostics);

        PdfOperationResult<IReadOnlyList<PdfFormField>> blockedFormFields = invalid.Read.TryFormFields();
        Assert.False(blockedFormFields.CanAttempt);
        Assert.NotEmpty(blockedFormFields.Diagnostics);

        PdfOperationResult<IReadOnlyList<PdfLogicalFormWidget>> blockedFormWidgets = invalid.Read.TryFormWidgets();
        Assert.False(blockedFormWidgets.CanAttempt);
        Assert.NotEmpty(blockedFormWidgets.Diagnostics);

        PdfOperationResult<PdfDocumentInfo> blockedDocumentInfo = invalid.Read.TryDocumentInfo();
        Assert.False(blockedDocumentInfo.CanAttempt);
        Assert.NotEmpty(blockedDocumentInfo.Diagnostics);

        PdfOperationResult<PdfDocumentSecurityInfo> blockedSecurity = invalid.Read.TrySecurity();
        Assert.False(blockedSecurity.CanAttempt);
        Assert.NotEmpty(blockedSecurity.Diagnostics);

        PdfOperationResult<PdfXmpMetadataInfo> blockedXmp = invalid.Read.TryXmpMetadata();
        Assert.False(blockedXmp.CanAttempt);
        Assert.NotEmpty(blockedXmp.Diagnostics);

        PdfOperationResult<IReadOnlyList<PdfOutputIntentInfo>> blockedOutputIntents = invalid.Read.TryOutputIntents();
        Assert.False(blockedOutputIntents.CanAttempt);
        Assert.NotEmpty(blockedOutputIntents.Diagnostics);

        PdfOperationResult<IReadOnlyList<PdfAttachmentInfo>> blockedAttachmentMetadata = invalid.Read.TryAttachmentMetadata();
        Assert.False(blockedAttachmentMetadata.CanAttempt);
        Assert.NotEmpty(blockedAttachmentMetadata.Diagnostics);
    }

    [Fact]
    public void OperationResult_ExposesAttachmentExtractionThroughFluentReader() {
        byte[] payload = Encoding.UTF8.GetBytes("attachment payload");
        byte[] source = PdfDocument.Create()
            .AttachFile("payload.txt", payload, "text/plain", PdfAssociatedFileRelationship.Data, "Workflow attachment")
            .Paragraph(paragraph => paragraph.Text("Attachment workflow"))
            .ToBytes();

        PdfDocument opened = PdfDocument.Open(source);
        IReadOnlyList<PdfExtractedAttachment> attachments = opened.Read.Attachments();
        PdfOperationResult<IReadOnlyList<PdfExtractedAttachment>> result = opened.Read.TryAttachments();

        PdfExtractedAttachment attachment = Assert.Single(attachments);
        Assert.Equal("payload.txt", attachment.FileName);
        Assert.Equal(payload, attachment.Bytes);
        Assert.True(result.CanAttempt);
        Assert.True(result.Succeeded);
        Assert.Equal(PdfPreflightCapability.ExtractAttachments, result.Capability);
        Assert.Equal(payload, Assert.Single(result.RequireValue()).Bytes);
        Assert.Empty(result.Diagnostics);

        PdfOperationResult<IReadOnlyList<PdfExtractedAttachment>> blocked = PdfDocument
            .Open(Encoding.ASCII.GetBytes("not a pdf"))
            .Read
            .TryAttachments();

        Assert.False(blocked.CanAttempt);
        Assert.False(blocked.Succeeded);
        Assert.NotEmpty(blocked.Diagnostics);
    }

    [Fact]
    public void FluentReader_ExposesGenericAnnotationReadback() {
        byte[] source = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Annotation workflow"))
            .TextAnnotation("Review note", width: 20, height: 20, icon: PdfTextAnnotationIcon.Note, open: true)
            .FreeTextAnnotation("Visible reviewer note", width: 140, height: 34, fontSize: 11, textAlign: PdfAlign.Center)
            .ToBytes();

        PdfDocument document = PdfDocument.Open(source);
        IReadOnlyList<PdfAnnotation> annotations = document.Read.Annotations();
        Assert.Equal(2, annotations.Count);
        Assert.All(annotations, annotation => Assert.Equal(1, annotation.PageNumber));

        PdfAnnotation text = Assert.Single(document.Read.AnnotationsBySubtype("Text"));
        Assert.Equal("Review note", text.Contents);
        Assert.True(text.Width > 0);
        Assert.True(text.Height > 0);

        PdfAnnotation freeText = Assert.Single(document.Read.AnnotationsBySubtype("FreeText"));
        Assert.Equal("Visible reviewer note", freeText.Contents);
        Assert.True(freeText.HasFreeTextAppearanceMetadata);
        Assert.Equal(11D, freeText.EffectiveFontSize);
        Assert.Empty(document.Read.AnnotationsByActionType("JavaScript"));

        PdfOperationResult<IReadOnlyList<PdfAnnotation>> safeAnnotations = document.Read.TryAnnotations();
        Assert.True(safeAnnotations.Succeeded);
        Assert.Equal(2, safeAnnotations.RequireValue().Count);

        PdfOperationResult<IReadOnlyList<PdfAnnotation>> safeFreeText = document.Read.TryAnnotationsBySubtype("FreeText");
        Assert.True(safeFreeText.Succeeded);
        Assert.Equal("Visible reviewer note", Assert.Single(safeFreeText.RequireValue()).Contents);

        PdfOperationResult<IReadOnlyList<PdfAnnotation>> safeActions = document.Read.TryAnnotationsByActionType("JavaScript");
        Assert.True(safeActions.Succeeded);
        Assert.Empty(safeActions.RequireValue());

        PdfOperationResult<IReadOnlyList<PdfAnnotation>> blocked = PdfDocument
            .Open(Encoding.ASCII.GetBytes("not a pdf"))
            .Read
            .TryAnnotations();

        Assert.False(blocked.CanAttempt);
        Assert.False(blocked.Succeeded);
        Assert.NotEmpty(blocked.Diagnostics);
    }

    [Fact]
    public void DiagnosticsOptimizationTextBlocksFlatteningAndRedactionPlanning_AreAvailableFromDocumentFacade() {
        byte[] source = PdfDocument.Create(new PdfOptions {
                IncludePageLabels = true,
                OpenAction = new PdfOpenActionOptions(1, destinationMode: PdfOpenActionDestinationMode.Fit),
                ViewerPreferences = new PdfViewerPreferencesOptions {
                    DisplayDocTitle = true,
                    FitWindow = true
                }
            })
            .Meta(title: "Diagnostic Report")
            .H1("Diagnostic Report")
            .Paragraph(paragraph => paragraph.Text("This paragraph should appear in structured text and redaction planning."))
            .ToBytes();

        PdfDocument document = PdfDocument.Open(source);
        PdfDocumentInfo info = document.Inspect();
        Assert.True(info.HasReadablePageLabels);
        Assert.True(info.HasReadableOpenAction);
        Assert.True(info.HasReadableViewerPreferences);

        PdfDiagnosticReport diagnostics = document.Diagnostics();
        Assert.True(diagnostics.CanRead);
        Assert.True(diagnostics.ObjectGraphParsed);
        Assert.True(diagnostics.ObjectCount > 0);
        Assert.True(diagnostics.StreamCount > 0);
        Assert.True(diagnostics.StreamTypeCounts.ContainsKey("Stream"));

        PdfOptimizationReport optimization = document.AnalyzeOptimization();
        Assert.Equal(diagnostics.StreamCount, optimization.StreamCount);
        Assert.NotEmpty(optimization.LargestStreams);

        IReadOnlyList<PdfLogicalTextBlock> blocks = document.Read.TextBlocks();
        Assert.NotEmpty(blocks);
        Assert.Contains(blocks, block => block.Text.Contains("Diagnostic Report", StringComparison.Ordinal));
        Assert.True(document.Read.TryTextBlocks().Succeeded);

        PdfRedactionPlan plan = document.PlanRedactions(new[] {
            new PdfRedactionArea(1, 0, 0, 1000, 1000, "full-page")
        });
        Assert.True(plan.HasMatches);
        Assert.Contains(plan.Matches, match => match.Text != null && match.Text.Contains("structured text", StringComparison.Ordinal));

        PdfDocument flattened = document.FlattenVisualAnnotations();
        Assert.True(flattened.Preflight().CanRead);
        Assert.True(document.TryFlattenVisualAnnotations().Succeeded);
    }

    [Fact]
    public void ProofReports_StayFluentForRewritePreservationAndRedactionVerification() {
        byte[] source = PdfRewritePreservationTestSupport.BuildPreservationProofPdf();
        PdfDocument updated = PdfDocument.Open(source).UpdateMetadata(title: "Updated preservation title");
        var preservationOptions = new PdfRewritePreservationOptions()
            .AllowMetadataChanges("Title")
            .RequireTextMarkers("PreservationMarker", "SecondPageMarker");

        PdfRewritePreservationReport preserved = PdfDocument.Open(source).AssessRewritePreservation(updated, preservationOptions);
        Assert.True(preserved.IsPreserved);
        Assert.Empty(preserved.Issues);

        using var rewrittenStream = new MemoryStream(updated.ToBytes());
        Assert.True(PdfDocument.Open(source).AssertRewritePreserved(rewrittenStream, preservationOptions).IsPreserved);

        PdfDocument deleted = PdfDocument.Open(source).Pages.Delete(2);
        PdfRewritePreservationReport loss = PdfDocument.Open(source).AssessRewritePreservation(
            deleted,
            new PdfRewritePreservationOptions().RequireTextMarkers("SecondPageMarker"));
        Assert.False(loss.IsPreserved);
        Assert.Contains(loss.Issues, issue => issue.Feature == "PageCount");
        Assert.Contains(loss.Issues, issue => issue.Feature == "TextMarker" && issue.Expected == "SecondPageMarker");

        PdfRewritePreservationMatrixReport matrix = PdfDocument.Open(source).AssertRewritePreservationMatrix(
            "fluent-metadata-update",
            "MetadataUpdate",
            document => document.UpdateMetadata(title: "Updated preservation title"),
            options: preservationOptions,
            sourceFeatures: new[] { "metadata", "xmp", "attachments" });
        Assert.True(matrix.Passed);
        PdfRewritePreservationMatrixEntry matrixEntry = Assert.Single(matrix.Entries);
        Assert.Equal(PdfRewritePreservationMatrixClassification.RewriteSafe, matrixEntry.ActualClassification);
        Assert.NotNull(matrixEntry.PreservationReport);
        Assert.True(matrixEntry.PreservationReport!.IsPreserved);
        Assert.Contains("attachments", matrixEntry.SourceFeatures);

        byte[] signedSource = PdfRewritePreservationTestSupport.BuildSignedIncrementalProofPdf();
        PdfRewritePreservationMatrixReport blockedMatrix = PdfDocument.Open(signedSource).AssertRewritePreservationMatrix(
            "signed-rewrite-blocked",
            "MetadataUpdate",
            document => document.UpdateMetadata(title: "Blocked"),
            PdfRewritePreservationMatrixClassification.Blocked,
            sourceFeatures: new[] { "signature", "incremental" });
        PdfRewritePreservationMatrixEntry blockedEntry = Assert.Single(blockedMatrix.Entries);
        Assert.Equal(PdfRewritePreservationMatrixClassification.Blocked, blockedEntry.ActualClassification);
        Assert.Contains("Signed PDF files are not supported for rewriting", blockedEntry.FailureMessage, StringComparison.Ordinal);

        PdfRedactionProofResult redactionProof = PdfRedactionProofTestSupport.BuildAndVerifyRedactionRemovalProof();
        PdfRedactionVerificationOptions redactionOptions = PdfRedactionProofTestSupport.CreateVerificationOptions();

        PdfRedactionVerificationReport verified = PdfDocument.Open(redactionProof.Redacted).VerifyRedactions(redactionOptions);
        Assert.True(verified.IsVerified);
        Assert.Empty(verified.Issues);
        Assert.DoesNotContain("PAY-SECRET-2026", verified.ExtractedText, StringComparison.Ordinal);
        Assert.Contains("Visible compliance marker", verified.ExtractedText, StringComparison.Ordinal);

        PdfRedactionVerificationReport unredacted = PdfDocument.Open(redactionProof.Source).VerifyRedactions(redactionOptions);
        Assert.False(unredacted.IsVerified);
        Assert.Contains(unredacted.Issues, issue => issue.Feature == "RemovedTextMarker" && issue.Marker == "PAY-SECRET-2026");
        Assert.Throws<InvalidOperationException>(() => PdfDocument.Open(redactionProof.Source).AssertRedactionsVerified(redactionOptions));
    }

    [Fact]
    public void Forms_TryFillNullOptionsCallsRemainSourceCompatible() {
        byte[] formPdf = BuildSimpleFormPdf();
        var textValues = new Dictionary<string, string> {
            ["Person.Name"] = "Ada Lovelace"
        };
        var fieldValues = new Dictionary<string, PdfFormFieldValue> {
            ["Person.Name"] = PdfFormFieldValue.From("Ada Lovelace")
        };

        PdfDocumentForms forms = PdfDocument.Open(formPdf).Forms;

        Assert.True(forms.TryFill(textValues, null).Succeeded);
        Assert.True(forms.TryFill(fieldValues, null).Succeeded);
        Assert.True(forms.TryFillAndFlatten(textValues, null).Succeeded);
        Assert.True(forms.TryFillAndFlatten(fieldValues, null).Succeeded);
        Assert.True(forms.TryFlatten(null).Succeeded);
    }

    [Fact]
    public void PageOperations_ReturnNewDocumentsAndMatchExistingHelpers() {
        byte[] source = BuildThreePagePdf();

        Assert.Equal(
            PdfPageExtractor.ExtractPageRanges(source, PdfPageRange.ParseMany("3,1-2")),
            PdfDocument.Open(source).Pages.Extract("3,1-2").ToBytes());

        Assert.Equal(
            PdfPageEditor.DeletePageRanges(source, PdfPageRange.ParseMany("2")),
            PdfDocument.Open(source).Pages.Delete("2").ToBytes());

        Assert.Equal(
            PdfPageEditor.ReorderPageRanges(source, PdfPageRange.ParseMany("2,3,1")),
            PdfDocument.Open(source).Pages.Reorder("2,3,1").ToBytes());

        Assert.Equal(
            PdfPageEditor.RotatePageRanges(source, 90, PdfPageRange.ParseMany("2")),
            PdfDocument.Open(source).Pages.Rotate(90, "2").ToBytes());

        IReadOnlyList<PdfDocument> split = PdfDocument.Open(source).Pages.Split();
        Assert.Equal(3, split.Count);
        Assert.All(split, part => Assert.Equal(1, part.Inspect().PageCount));
        Assert.Contains("Page A", split[0].Read.Text(), StringComparison.Ordinal);
        Assert.Contains("Page B", split[1].Read.Text(), StringComparison.Ordinal);
        Assert.Contains("Page C", split[2].Read.Text(), StringComparison.Ordinal);
    }

    [Fact]
    public void PageImports_StayFluentAndDelegateToCurrentEngine() {
        byte[] target = BuildPdf("Target", "Target body");
        byte[] source = BuildThreePagePdf();
        PdfPageSelection selection = PdfPageSelection.Parse("3,1");

        PdfDocument appended = PdfDocument.Open(target).Pages.Append(PdfDocument.Open(source), selection);
        Assert.Equal(PdfPageImporter.AppendPages(target, source, 3, 1), appended.ToBytes());
        Assert.Equal(3, appended.Inspect().PageCount);
        Assert.Contains("Page C", appended.Read.Text(), StringComparison.Ordinal);
        Assert.DoesNotContain("Page B", appended.Read.Text(), StringComparison.Ordinal);

        PdfDocument prepended = PdfDocument.Open(target).Pages.Prepend(source, PdfPageSelection.From(2));
        Assert.Equal(PdfPageImporter.PrependPages(target, source, 2), prepended.ToBytes());
        Assert.Contains("Second page body", prepended.Read.Text(), StringComparison.Ordinal);

        var importOptions = new PdfPageImportOptions {
            FlattenVisualAnnotations = true
        };
        PdfDocument inserted = PdfDocument.Open(target).Pages.Insert(1, source, PdfPageSelection.From(2), importOptions);
        Assert.Equal(PdfPageImporter.InsertPages(importOptions, target, source, 1, 2), inserted.ToBytes());

        PdfOperationResult<PdfDocument> imported = PdfDocument.Open(target).Pages.TryAppend(PdfDocument.Open(source), PdfPageSelection.From(1));
        Assert.True(imported.Succeeded);
        Assert.Equal(PdfPreflightCapability.ManipulatePages, imported.Capability);
        Assert.Contains("First page body", imported.RequireValue().Read.Text(), StringComparison.Ordinal);

        PdfDocument invalid = PdfDocument.Open(Encoding.ASCII.GetBytes("not a pdf"));
        PdfOperationResult<PdfDocument> blocked = invalid.Pages.TryAppend(PdfDocument.Open(source));
        Assert.False(blocked.CanAttempt);
        Assert.False(blocked.Succeeded);
        Assert.NotEmpty(blocked.Diagnostics);

        string sourcePath = Path.Combine(Path.GetTempPath(), "officeimo-page-import-" + Guid.NewGuid().ToString("N") + ".pdf");
        try {
            File.WriteAllBytes(sourcePath, source);

            PdfOperationResult<PdfDocument> byteImport = PdfDocument.Open(target).Pages.TryAppend(source, PdfPageSelection.From(1));
            Assert.True(byteImport.Succeeded);
            Assert.Equal(2, byteImport.RequireValue().Inspect().PageCount);

            using var sourceStream = new MemoryStream(source);
            PdfOperationResult<PdfDocument> streamImport = PdfDocument.Open(target).Pages.TryPrepend(sourceStream, PdfPageSelection.From(2));
            Assert.True(streamImport.Succeeded);
            Assert.Contains("Second page body", streamImport.RequireValue().Read.Text(), StringComparison.Ordinal);

            PdfOperationResult<PdfDocument> pathImport = PdfDocument.Open(target).Pages.TryInsert(1, sourcePath, PdfPageSelection.From(3));
            Assert.True(pathImport.Succeeded);
            Assert.Contains("Third page body", pathImport.RequireValue().Read.Text(), StringComparison.Ordinal);

            PdfOperationResult<PdfDocument> failedSource = PdfDocument.Open(target).Pages.TryAppend(Encoding.ASCII.GetBytes("not a pdf"));
            Assert.True(failedSource.CanAttempt);
            Assert.False(failedSource.Succeeded);
            Assert.NotNull(failedSource.Exception);
            Assert.NotEmpty(failedSource.Diagnostics);
        } finally {
            if (File.Exists(sourcePath)) {
                File.Delete(sourcePath);
            }
        }
    }

    [Fact]
    public void PageOperations_SplitByPageCountSelectionsAndBookmarks() {
        byte[] source = BuildThreePagePdf();

        IReadOnlyList<PdfDocument> pageGroups = PdfDocument.Open(source).Pages.Split(2);
        Assert.Equal(2, pageGroups.Count);
        Assert.Equal(2, pageGroups[0].Inspect().PageCount);
        Assert.Equal(1, pageGroups[1].Inspect().PageCount);
        Assert.Contains("Page A", pageGroups[0].Read.Text(), StringComparison.Ordinal);
        Assert.Contains("Page C", pageGroups[1].Read.Text(), StringComparison.Ordinal);

        IReadOnlyList<PdfDocument> selections = PdfDocument.Open(source).Pages.Split(
            PdfPageSelection.Parse("1-2"),
            PdfPageSelection.Parse("3"));
        Assert.Equal(2, selections.Count);
        Assert.Contains("Second page body", selections[0].Read.Text(), StringComparison.Ordinal);
        Assert.Contains("Third page body", selections[1].Read.Text(), StringComparison.Ordinal);

        byte[] bookmarked = BuildThreeBookmarkPdf();
        PdfDocument bookmarkDocument = PdfDocument.Open(bookmarked);
        IReadOnlyList<PdfBookmarkPageRange> ranges = bookmarkDocument.Pages.BookmarkPageRanges();
        Assert.Equal(3, ranges.Count);
        Assert.Equal("Chapter One", ranges[0].Title);
        Assert.Equal(PdfPageRange.From(1, 1), ranges[0].PageRange);
        Assert.Equal(PdfPageRange.From(3, 3), ranges[2].PageRange);

        IReadOnlyList<PdfBookmarkPageRange> selectedRange = bookmarkDocument.Pages.BookmarkPageRanges("Chapter Two");
        Assert.Equal(PdfPageRange.From(2, 2), Assert.Single(selectedRange).PageRange);

        IReadOnlyList<PdfDocument> bookmarkSplit = bookmarkDocument.Pages.SplitByBookmarks("Chapter Two");
        PdfDocument chapterTwo = Assert.Single(bookmarkSplit);
        Assert.Contains("Chapter Two", chapterTwo.Read.Text(), StringComparison.Ordinal);
        Assert.DoesNotContain("Chapter Three", chapterTwo.Read.Text(), StringComparison.Ordinal);
    }

    [Fact]
    public void PageOperations_BookmarkPageRangesUsePageOrderWhenOutlineOrderDiffers() {
        PdfDocument document = PdfDocument.Open(BuildOutOfOrderBookmarkPdf());

        IReadOnlyList<PdfBookmarkPageRange> ranges = document.Pages.BookmarkPageRanges();

        Assert.Equal(new[] { "Chapter One", "Chapter Two", "Chapter Three" }, ranges.Select(range => range.Title).ToArray());
        Assert.Equal(PdfPageRange.From(1, 1), ranges[0].PageRange);
        Assert.Equal(PdfPageRange.From(2, 2), ranges[1].PageRange);
        Assert.Equal(PdfPageRange.From(3, 3), ranges[2].PageRange);
    }

    [Fact]
    public void MergeMetadataAndStamping_StayFluentAndDelegateToCurrentEngine() {
        byte[] source = BuildThreePagePdf();
        byte[] appendix = BuildPdf("Appendix", "Appendix body");

        PdfDocument merged = PdfDocument.Open(source).MergeWith(PdfDocument.Open(appendix));
        Assert.Equal(PdfMerger.Merge(source, appendix), merged.ToBytes());
        Assert.Equal(4, merged.Inspect().PageCount);

        PdfDocument bulkMerged = PdfDocument.Merge(PdfDocument.Open(source), PdfDocument.Open(appendix));
        Assert.Equal(PdfMerger.Merge(source, appendix), bulkMerged.ToBytes());
        Assert.Equal(4, bulkMerged.Inspect().PageCount);

        PdfDocument metadata = merged.UpdateMetadata(title: "Workflow updated", author: "OfficeIMO Tests");
        Assert.Equal(
            PdfMetadataEditor.UpdateMetadata(merged.ToBytes(), title: "Workflow updated", author: "OfficeIMO Tests"),
            metadata.ToBytes());
        Assert.Equal("Workflow updated", metadata.Inspect().Metadata.Title);
        Assert.Equal("OfficeIMO Tests", metadata.Inspect().Metadata.Author);

        var stampOptions = new PdfTextStampOptions {
            X = 72,
            Y = 72,
            FontSize = 12
        };

        Assert.Equal(
            PdfStamper.StampText(metadata.ToBytes(), "Reviewed", stampOptions),
            metadata.Stamp.Text("Reviewed", stampOptions).ToBytes());
    }

    [Fact]
    public void AppendOnlyRevisionWorkflows_StayFluentAndDelegateToIncrementalEngine() {
        byte[] source = BuildThreePagePdf();

        PdfDocument metadata = PdfDocument.Open(source).AppendMetadataRevision(title: "Append-only workflow");
        Assert.Equal(
            PdfIncrementalUpdater.UpdateMetadata(source, title: "Append-only workflow"),
            metadata.ToBytes());
        Assert.Equal("Append-only workflow", metadata.Inspect().Metadata.Title);
        Assert.True(metadata.Inspect().Security.HasIncrementalUpdates);

        PdfOperationResult<PdfDocument> metadataResult = PdfDocument.Open(source).TryAppendMetadataRevision(author: "OfficeIMO Incremental");
        Assert.True(metadataResult.Succeeded);
        Assert.Equal(PdfPreflightCapability.AppendMetadataRevision, metadataResult.Capability);
        Assert.Equal("OfficeIMO Incremental", metadataResult.RequireValue().Inspect().Metadata.Author);

        byte[] formSource = BuildSimpleFormPdf();
        var fieldValues = new Dictionary<string, string> {
            ["Person.Name"] = "Ada"
        };
        var formOptions = new PdfIncrementalFormFieldUpdateOptions {
            GenerateAppearanceStreams = true,
            KeepNeedAppearances = false
        };

        PdfDocument form = PdfDocument.Open(formSource).Forms.AppendRevision(fieldValues, formOptions);
        Assert.Equal(
            PdfIncrementalUpdater.UpdateFormFields(formSource, fieldValues, formOptions),
            form.ToBytes());
        Assert.Equal("Ada", Assert.Single(form.Inspect().FormFields).Value);
        Assert.True(form.Inspect().Security.HasIncrementalUpdates);

        PdfOperationResult<PdfDocument> formResult = PdfDocument.Open(formSource).Forms.TryAppendRevision(fieldValues, formOptions, readOptions: null);
        Assert.True(formResult.Succeeded);
        Assert.Equal(PdfPreflightCapability.AppendFormFieldRevision, formResult.Capability);
        Assert.Equal("Ada", Assert.Single(formResult.RequireValue().Inspect().FormFields).Value);

        var signatureOptions = new PdfExternalSignatureOptions {
            FieldName = "Approval",
            Name = "Alice",
            Reason = "Approval",
            ReservedSignatureContentsBytes = 256,
            SigningTime = new DateTimeOffset(2026, 6, 22, 12, 0, 0, TimeSpan.Zero)
        };

        PdfExternalSignaturePreparation preparation = PdfDocument.Open(source).PrepareExternalSignature(signatureOptions);
        Assert.Equal(
            PdfIncrementalUpdater.PrepareExternalSignature(source, signatureOptions).PreparedPdf,
            preparation.PreparedPdf);
        Assert.Equal("Approval", preparation.FieldName);
        Assert.True(PdfDocument.Open(preparation.PreparedPdf).ValidateSignatures().HasSignatures);

        PdfOperationResult<PdfExternalSignaturePreparation> signatureResult = PdfDocument.Open(source).TryPrepareExternalSignature(signatureOptions);
        Assert.True(signatureResult.Succeeded);
        Assert.Equal(PdfPreflightCapability.PrepareExternalSignatureRevision, signatureResult.Capability);
        Assert.Equal("Approval", signatureResult.RequireValue().FieldName);
    }

    [Fact]
    public void ExistingDocumentFacade_PreservesPerDocumentReadContractsAcrossReparsingOperations() {
        byte[] onePage = BuildPdf("One page", "First source");
        byte[] twoPages = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Page one"))
            .PageBreak()
            .Paragraph(paragraph => paragraph.Text("Page two"))
            .ToBytes();
        var onePageLimit = new PdfReadOptions {
            Limits = new PdfReadLimits { MaxPages = 1 }
        };
        PdfDocument constrained = PdfDocument.Open(twoPages, onePageLimit);

        PdfMutationBlockedException mergeException = Assert.Throws<PdfMutationBlockedException>(
            () => PdfDocument.Merge(PdfDocument.Open(onePage), constrained));
        PdfMutationBlockedException resizeException = Assert.Throws<PdfMutationBlockedException>(
            () => constrained.Pages.Resize(new PageSize(612, 792)));

        PdfMutationBlockedException optimizeException = Assert.Throws<PdfMutationBlockedException>(
            () => constrained.Optimize());

        Assert.Contains("Read.ParserUnsupported", mergeException.Plan.BlockerCodes);
        Assert.Contains("Read.ParserUnsupported", resizeException.Plan.BlockerCodes);
        Assert.Contains("Read.ParserUnsupported", optimizeException.Plan.BlockerCodes);
    }

    [Fact]
    public void AppendMetadataRevision_PreservesTaggedStructureWhenRewriteIsBlocked() {
        byte[] source = PdfDocument.Create()
            .TaggedPdfCatalogMarkers()
            .Language("en-US")
            .H1("Tagged append-only heading")
            .Paragraph(paragraph => paragraph.Text("Tagged append-only body."))
            .ToBytes();

        PdfDocument tagged = PdfDocument.Open(source);
        PdfDocumentPreflight preflight = tagged.Preflight();
        Assert.True(preflight.CanRead);
        Assert.False(preflight.CanRewrite);
        Assert.True(preflight.CanAppendMetadataRevision);
        Assert.True(preflight.Can(PdfPreflightCapability.AppendMetadataRevision));
        Assert.Empty(preflight.GetCapabilityDiagnostics(PdfPreflightCapability.AppendMetadataRevision));
        Assert.Contains(preflight.RewriteBlockers, blocker => blocker.Kind == PdfRewriteBlockerKind.TaggedContent);

        PdfOperationResult<PdfDocument> result = tagged.TryAppendMetadataRevision(title: "Tagged append-only update");

        Assert.True(result.Succeeded);
        PdfDocument updated = result.RequireValue();
        PdfDocumentInfo updatedInfo = updated.Inspect();
        Assert.Equal("Tagged append-only update", updatedInfo.Metadata.Title);
        Assert.True(updatedInfo.Security.HasIncrementalUpdates);
        Assert.True(updatedInfo.HasReadableTaggedContent);
        Assert.NotNull(updatedInfo.TaggedContent);
        Assert.Contains("Document", updatedInfo.TaggedContent!.StructureTypes);
        Assert.Contains("H1", updatedInfo.TaggedContent.StructureTypes);
        Assert.Contains("P", updatedInfo.TaggedContent.StructureTypes);

        PdfRewritePreservationReport report = PdfRewritePreservation.AssertPreserved(
            source,
            updated.ToBytes(),
            new PdfRewritePreservationOptions().AllowMetadataChanges("Title"));
        Assert.True(report.IsPreserved);
        Assert.Empty(report.Issues);
    }

    [Fact]
    public void UpdateMetadata_PreservesSimpleOptionalContentLayersThroughFluentWorkflow() {
        byte[] source = PdfOptionalContentSupport.BuildOptionalContentMetadataPdf();
        PdfDocument layered = PdfDocument.Open(source);
        PdfDocumentPreflight preflight = layered.Preflight();
        Assert.True(preflight.CanRead);
        Assert.True(preflight.CanRewrite);
        Assert.False(preflight.HasRewriteBlocker(PdfRewriteBlockerKind.OptionalContent));
        Assert.True(preflight.DocumentInfo!.HasReadableOptionalContent);

        PdfDocument updated = layered.UpdateMetadata(title: "Layer-preserving update");

        PdfDocumentInfo updatedInfo = updated.Inspect();
        Assert.Equal("Layer-preserving update", updatedInfo.Metadata.Title);
        Assert.True(updatedInfo.HasReadableOptionalContent);
        Assert.Equal(new[] { "Print layer", "Hidden layer" }, updatedInfo.OptionalContentGroupNames);
        PdfOptionalContentGroup hiddenLayer = Assert.Single(updatedInfo.GetOptionalContentGroupsByName("Hidden layer"));
        Assert.False(hiddenLayer.IsInitiallyVisible);
        Assert.True(hiddenLayer.IsLocked);

        PdfRewritePreservationReport report = PdfRewritePreservation.AssertPreserved(
            source,
            updated.ToBytes(),
            new PdfRewritePreservationOptions().AllowMetadataChanges("Title"));
        Assert.True(report.IsPreserved);
        Assert.Empty(report.Issues);
    }

    [Fact]
    public void Save_WritesCurrentBytesToStreamAndPath() {
        PdfDocument document = PdfDocument.Open(BuildThreePagePdf()).Pages.Delete(2);
        using var stream = new MemoryStream();

        document.Save(stream);
        Assert.Equal(document.ToBytes(), stream.ToArray());

        string directory = Path.Combine(Path.GetTempPath(), "officeimo-pdf-workflow-" + Guid.NewGuid().ToString("N"));
        string path = Path.Combine(directory, "saved.pdf");
        try {
            document.Save(path);

            Assert.True(File.Exists(path));
            Assert.Equal(document.ToBytes(), File.ReadAllBytes(path));
        } finally {
            if (Directory.Exists(directory)) {
                Directory.Delete(directory, recursive: true);
            }
        }
    }

    [Fact]
    public async System.Threading.Tasks.Task SaveResult_ReportsOutputWithoutRequiringReadablePdfContent() {
        byte[] invalidPdf = Encoding.ASCII.GetBytes("not a pdf");
        PdfDocument document = PdfDocument.Open(invalidPdf);
        using var stream = new MemoryStream();

        Assert.Empty(document.AnalyzeTextEncoding());

        PdfBytesResult bytesResult = document.TryToBytes();

        Assert.True(bytesResult.Succeeded);
        Assert.Equal(invalidPdf.LongLength, bytesResult.ByteCount);
        Assert.Equal(invalidPdf, bytesResult.Bytes);
        Assert.Equal(invalidPdf, bytesResult.RequireBytes());
        Assert.Empty(bytesResult.Diagnostics);
        Assert.Empty(bytesResult.TextEncodingDiagnostics);
        Assert.Empty(bytesResult.Warnings);

        PdfSaveResult streamResult = document.TrySave(stream);

        Assert.True(streamResult.Succeeded);
        Assert.Null(streamResult.OutputPath);
        Assert.Equal(invalidPdf.LongLength, streamResult.BytesWritten);
        Assert.Empty(streamResult.Diagnostics);
        Assert.Same(streamResult, streamResult.RequireSuccess());
        Assert.Equal(invalidPdf, stream.ToArray());

        using var asyncStream = new MemoryStream();
        PdfSaveResult asyncResult = await document.TrySaveAsync(asyncStream);

        Assert.True(asyncResult.Succeeded);
        Assert.Equal(invalidPdf.LongLength, asyncResult.BytesWritten);
        Assert.Equal(invalidPdf, asyncStream.ToArray());

        string directory = Path.Combine(Path.GetTempPath(), "officeimo-pdf-save-result-" + Guid.NewGuid().ToString("N"));
        string path = Path.Combine(directory, "snapshot.pdf");
        try {
            PdfSaveResult pathResult = document.TrySave(path);

            Assert.True(pathResult.Succeeded);
            Assert.Equal(Path.GetFullPath(path), pathResult.OutputPath);
            Assert.Equal(invalidPdf.LongLength, pathResult.BytesWritten);
            Assert.Equal(invalidPdf, File.ReadAllBytes(path));

            PdfSaveResult directoryResult = document.TrySave(directory);

            Assert.False(directoryResult.Succeeded);
            Assert.Equal(0, directoryResult.BytesWritten);
            Assert.NotEmpty(directoryResult.Diagnostics);
            Assert.Throws<InvalidOperationException>(() => directoryResult.RequireSuccess());
        } finally {
            if (Directory.Exists(directory)) {
                Directory.Delete(directory, recursive: true);
            }
        }

        using var readOnlyStream = new MemoryStream(Array.Empty<byte>(), writable: false);
        PdfSaveResult streamFailure = document.TrySave(readOnlyStream);

        Assert.False(streamFailure.Succeeded);
        Assert.NotEmpty(streamFailure.Diagnostics);
    }

    [Fact]
    public void SaveResult_CarriesTextEncodingDiagnosticsForGeneratedPdfFailures() {
        PdfDocument document = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Snowman \u2603"));
        using var stream = new MemoryStream();

        PdfBytesResult bytesResult = document.TryToBytes();
        Assert.False(bytesResult.Succeeded);
        Assert.Equal(0, bytesResult.ByteCount);
        Assert.Empty(bytesResult.Bytes);
        Assert.NotEmpty(bytesResult.Diagnostics);
        Assert.Throws<InvalidOperationException>(() => bytesResult.RequireBytes());

        PdfSaveResult result = document.TrySave(stream);

        Assert.False(result.Succeeded);
        Assert.Equal(0, result.BytesWritten);
        Assert.Equal(0, stream.Length);
        Assert.NotEmpty(result.Diagnostics);

        PdfTextEncodingDiagnostic diagnostic = Assert.Single(result.TextEncodingDiagnostics);
        PdfTextEncodingDiagnostic bytesDiagnostic = Assert.Single(bytesResult.TextEncodingDiagnostics);

        Assert.Equal("unsupported-text-glyph", diagnostic.Code);
        Assert.Equal("PdfParagraph", diagnostic.Source);
        Assert.Equal("PdfParagraph[0].Run[0]", diagnostic.Location);
        Assert.Equal(0, diagnostic.RunIndex);
        Assert.Equal("U+2603", diagnostic.CodePoint);
        Assert.Equal("\u2603", diagnostic.Text);
        Assert.False(diagnostic.IsControlCharacter);
        Assert.Equal("PDF WinAnsiEncoding", diagnostic.Encoding);
        Assert.Equal("Embedded Unicode fonts are required for this text.", diagnostic.Remediation);
        Assert.Equal(diagnostic.CodePoint, bytesDiagnostic.CodePoint);
        Assert.Equal(diagnostic.Location, bytesDiagnostic.Location);

        PdfConversionWarning warning = Assert.Single(result.Warnings);
        PdfConversionWarning bytesWarning = Assert.Single(bytesResult.Warnings);

        Assert.Equal("OfficeIMO.Pdf", warning.Converter);
        Assert.Equal(diagnostic.Code, warning.Code);
        Assert.Equal(diagnostic.Message, warning.Message);
        Assert.Equal(PdfConversionWarningSeverity.Error, warning.Severity);
        Assert.Equal("U+2603", warning.Details["codePoint"]);
        Assert.Equal("PdfParagraph[0].Run[0]", warning.Details["location"]);
        Assert.Equal("0", warning.Details["runIndex"]);
        Assert.Equal("PDF WinAnsiEncoding", warning.Details["encoding"]);
        Assert.Equal("Embedded Unicode fonts are required for this text.", warning.Details["remediation"]);
        Assert.Equal(warning.Code, bytesWarning.Code);
    }

    [Fact]
    public void AnalyzeTextEncoding_ReturnsAllGeneratedTextDiagnosticsBeforeRender() {
        var options = new PdfOptions {
            ShowHeader = true,
            HeaderFormat = "Header \u2603",
            ShowPageNumbers = true,
            FooterFormat = "Footer \u2602"
        };

        PdfDocument document = PdfDocument.Create(options)
            .Paragraph(paragraph => paragraph.Text("Paragraph \u2603"))
            .H1("Heading \U0001F680")
            .Bullets(new[] { "Bullet \u266b" })
            .Table(new[] { new[] { "Table \u25a0" } }, style: new PdfTableStyle {
                Caption = "Table caption \u2666"
            })
            .Canvas(canvas => canvas
                .Text("Canvas \u260e", 10, 10, 120, 24)
                .FreeTextAnnotation("Callout \u2615", 10, 42, 120, 32)
                .Table(new[] { new[] { "Canvas table" } }, 10, 84, 120, 42, new PdfTableStyle {
                    Caption = "Canvas caption \u273f"
                }))
            .TextField("Person.Name", width: 120, height: 20, value: "Field \u2603");

        IReadOnlyList<PdfTextEncodingDiagnostic> diagnostics = document.AnalyzeTextEncoding();

        Assert.Equal(10, diagnostics.Count);
        Assert.Contains(diagnostics, diagnostic => diagnostic.Source == "PdfHeader" && diagnostic.Location == "PdfHeader[page=1]" && diagnostic.PageNumber == 1 && diagnostic.CodePoint == "U+2603");
        Assert.Contains(diagnostics, diagnostic => diagnostic.Source == "PdfFooter" && diagnostic.Location == "PdfFooter[page=1]" && diagnostic.PageNumber == 1 && diagnostic.CodePoint == "U+2602");
        Assert.Contains(diagnostics, diagnostic => diagnostic.Source == "PdfParagraph" && diagnostic.Location == "PdfParagraph[0].Run[0]" && diagnostic.RunIndex == 0 && diagnostic.CodePoint == "U+2603");
        Assert.Contains(diagnostics, diagnostic => diagnostic.Source == "PdfHeading" && diagnostic.Location == "PdfHeading[1]" && diagnostic.CodePoint == "U+1F680");
        Assert.Contains(diagnostics, diagnostic => diagnostic.Source == "PdfListItem" && diagnostic.Location == "PdfBulletList[2].PdfListItem[0].Run[0]" && diagnostic.RunIndex == 0 && diagnostic.CodePoint == "U+266B");
        Assert.Contains(diagnostics, diagnostic => diagnostic.Source == "PdfTableCaption" && diagnostic.Location == "PdfTable[3].PdfTableCaption" && diagnostic.CodePoint == "U+2666");
        Assert.Contains(diagnostics, diagnostic => diagnostic.Source == "PdfTableCell" && diagnostic.Location == "PdfTable[3].PdfTableCell[0,0].Run[0]" && diagnostic.RunIndex == 0 && diagnostic.TableRowIndex == 0 && diagnostic.TableColumnIndex == 0 && diagnostic.CodePoint == "U+25A0");
        Assert.Contains(diagnostics, diagnostic => diagnostic.Source == "PdfCanvasText" && diagnostic.Location == "PdfCanvas[4].PdfCanvasText[0].Run[0]" && diagnostic.RunIndex == 0 && diagnostic.CodePoint == "U+260E");
        Assert.Contains(diagnostics, diagnostic => diagnostic.Source == "PdfTableCaption" && diagnostic.Location == "PdfCanvas[4].PdfCanvasTable[2].PdfTableCaption" && diagnostic.CodePoint == "U+273F");
        Assert.Contains(diagnostics, diagnostic => diagnostic.Source == "PdfTextField" && diagnostic.Location == "PdfTextField[5]" && diagnostic.FieldName == "Person.Name" && diagnostic.CodePoint == "U+2603");

        PdfBytesResult result = document.TryToBytes();
        using var stream = new MemoryStream();
        PdfSaveResult saveResult = document.TrySave(stream);

        Assert.False(result.Succeeded);
        Assert.Equal(diagnostics.Count, result.TextEncodingDiagnostics.Count);
        Assert.Equal("PdfHeader", result.TextEncodingDiagnostics[0].Source);
        Assert.Equal("PdfHeader[page=1]", result.TextEncodingDiagnostics[0].Location);
        Assert.Equal(1, result.TextEncodingDiagnostics[0].PageNumber);
        Assert.Equal("U+2603", result.TextEncodingDiagnostics[0].CodePoint);
        Assert.Equal("PDF WinAnsiEncoding", result.TextEncodingDiagnostics[0].Encoding);
        Assert.Equal(diagnostics.Count, result.Warnings.Count);
        Assert.Equal("PdfHeader[page=1]", result.Warnings[0].Details["location"]);
        Assert.Equal("1", result.Warnings[0].Details["pageNumber"]);
        Assert.Equal("PDF WinAnsiEncoding", result.Warnings[0].Details["encoding"]);
        Assert.Contains(result.Warnings, warning =>
            warning.Source == "PdfTableCell" &&
            warning.Details["tableRowIndex"] == "0" &&
            warning.Details["tableColumnIndex"] == "0");
        Assert.Contains(result.Warnings, warning =>
            warning.Source == "PdfTextField" &&
            warning.Details["fieldName"] == "Person.Name");
        Assert.Contains("preflight found 10 generated text issues", result.Diagnostics[0], StringComparison.Ordinal);

        ArgumentException preflightException = Assert.ThrowsAny<ArgumentException>(() => document.ToBytes());

        Assert.Equal(1, preflightException.Data["pageNumber"]);

        Assert.False(saveResult.Succeeded);
        Assert.Equal(diagnostics.Count, saveResult.TextEncodingDiagnostics.Count);
        Assert.Equal(0, saveResult.BytesWritten);
        Assert.Equal(0, stream.Length);
    }

    [Fact]
    public void AnalyzeTextEncoding_PreflightsVariantTextWatermarks() {
        var options = new PdfOptions {
            FirstPageTextWatermark = new PdfTextWatermark("First \u2603"),
            EvenPageTextWatermark = new PdfTextWatermark("Even \u2602")
        };

        PdfDocument document = PdfDocument.Create(options);
        IReadOnlyList<PdfTextEncodingDiagnostic> diagnostics = document.AnalyzeTextEncoding();

        Assert.Contains(diagnostics, diagnostic => diagnostic.Source == "PdfFirstPageTextWatermark" && diagnostic.Location == "PdfFirstPageTextWatermark" && diagnostic.CodePoint == "U+2603");
        Assert.Contains(diagnostics, diagnostic => diagnostic.Source == "PdfEvenPageTextWatermark" && diagnostic.Location == "PdfEvenPageTextWatermark" && diagnostic.CodePoint == "U+2602");
    }

    [Fact]
    public void AnalyzeTextEncoding_PreflightsFormWidgetValuesThroughEmbeddedHelveticaPath() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        var options = new PdfOptions();
        options.EmbedStandardFont(PdfStandardFont.Helvetica, fontPath);
        const string value = "\u0105";
        if (PdfTextDiagnostics.AnalyzeGeneratedText(value, options, PdfStandardFont.Helvetica).Count != 0) {
            return;
        }

        PdfDocument document = PdfDocument.Create(options)
            .TextField("Person.City", value: value)
            .ChoiceField("Person.Country", new[] { "PL", value }, value: value);

        IReadOnlyList<PdfTextEncodingDiagnostic> diagnostics = document.AnalyzeTextEncoding();

        Assert.DoesNotContain(diagnostics, diagnostic => diagnostic.FieldName == "Person.City" && diagnostic.CodePoint == "U+0105");
        Assert.DoesNotContain(diagnostics, diagnostic => diagnostic.FieldName == "Person.Country" && diagnostic.CodePoint == "U+0105");

        using var stream = new MemoryStream();
        PdfSaveResult result = document.TrySave(stream);

        Assert.True(result.Succeeded);
        Assert.True(stream.Length > 0);
        Assert.DoesNotContain(result.TextEncodingDiagnostics, diagnostic => diagnostic.FieldName == "Person.City" && diagnostic.CodePoint == "U+0105");
    }

    [Fact]
    public void AnalyzeTextEncoding_PreflightsFreeTextAppearanceThroughWinAnsiPath() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        var options = new PdfOptions();
        options.EmbedStandardFont(PdfStandardFont.Helvetica, fontPath);
        const string value = "\u0105";
        if (PdfTextDiagnostics.AnalyzeGeneratedText(value, options, PdfStandardFont.Helvetica).Count != 0) {
            return;
        }

        PdfDocument document = PdfDocument.Create(options)
            .FreeTextAnnotation(value, width: 120, height: 32);

        IReadOnlyList<PdfTextEncodingDiagnostic> diagnostics = document.AnalyzeTextEncoding();

        Assert.Contains(diagnostics, diagnostic => diagnostic.Source == "PdfFreeTextAnnotation" && diagnostic.CodePoint == "U+0105");

        using var stream = new MemoryStream();
        PdfSaveResult result = document.TrySave(stream);

        Assert.False(result.Succeeded);
        Assert.Equal(0, stream.Length);
        Assert.Contains(result.TextEncodingDiagnostics, diagnostic => diagnostic.Source == "PdfFreeTextAnnotation" && diagnostic.CodePoint == "U+0105");
    }

    [Fact]
    public void ToBytes_ThrowsFullGeneratedTextPreflightBeforeRendering() {
        PdfDocument document = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Paragraph \u2603"))
            .H1("Heading \U0001F680");

        ArgumentException exception = Assert.ThrowsAny<ArgumentException>(() => document.ToBytes());

        Assert.Contains("preflight found 2 generated text issues", exception.Message, StringComparison.Ordinal);
        Assert.Contains("U+2603", exception.Message, StringComparison.Ordinal);
        Assert.Equal("unsupported-text-glyph", exception.Data["code"]);
        Assert.Equal("PdfParagraph", exception.Data["source"]);
        Assert.Equal("PdfParagraph[0].Run[0]", exception.Data["location"]);
        Assert.Equal(0, exception.Data["runIndex"]);
        Assert.Equal(2, exception.Data["diagnosticsCount"]);

        var diagnostics = Assert.IsAssignableFrom<IReadOnlyList<PdfTextEncodingDiagnostic>>(exception.Data["textEncodingDiagnostics"]);
        Assert.Equal(2, diagnostics.Count);
        Assert.Equal("PdfParagraph[0].Run[0]", diagnostics[0].Location);
        Assert.Equal("PdfHeading[1]", diagnostics[1].Location);
    }

    [Fact]
    public void ToBytes_ThrowsTableCellCoordinatesInTextPreflightException() {
        PdfDocument document = PdfDocument.Create()
            .Table(new[] { new[] { "Table \u25a0" } });

        ArgumentException exception = Assert.ThrowsAny<ArgumentException>(() => document.ToBytes());

        Assert.Equal("PdfTableCell", exception.Data["source"]);
        Assert.Equal("PdfTable[0].PdfTableCell[0,0].Run[0]", exception.Data["location"]);
        Assert.Equal(0, exception.Data["tableRowIndex"]);
        Assert.Equal(0, exception.Data["tableColumnIndex"]);

        var diagnostics = Assert.IsAssignableFrom<IReadOnlyList<PdfTextEncodingDiagnostic>>(exception.Data["textEncodingDiagnostics"]);
        PdfTextEncodingDiagnostic diagnostic = Assert.Single(diagnostics);

        Assert.Equal(0, diagnostic.TableRowIndex);
        Assert.Equal(0, diagnostic.TableColumnIndex);
    }

    [Fact]
    public void ToBytes_ThrowsFieldNameForGeneratedFormPreflightException() {
        PdfDocument document = PdfDocument.Create()
            .Table(new[] {
                new[] {
                    PdfTableCell.WithFormFields(
                        "Table form",
                        new[] {
                            PdfTableCellFormField.TextField("Table.DueDate", "Due \u2603", width: 120, height: 18)
                        })
                }
            });

        ArgumentException exception = Assert.ThrowsAny<ArgumentException>(() => document.ToBytes());

        Assert.Equal("PdfTableTextField", exception.Data["source"]);
        Assert.Equal("PdfTable[0].PdfTableCell[0,0].PdfTableTextField", exception.Data["location"]);
        Assert.Equal("Table.DueDate", exception.Data["fieldName"]);
        Assert.Equal(0, exception.Data["tableRowIndex"]);
        Assert.Equal(0, exception.Data["tableColumnIndex"]);

        var diagnostics = Assert.IsAssignableFrom<IReadOnlyList<PdfTextEncodingDiagnostic>>(exception.Data["textEncodingDiagnostics"]);
        PdfTextEncodingDiagnostic diagnostic = Assert.Single(diagnostics);

        Assert.Equal("Table.DueDate", diagnostic.FieldName);
        Assert.Equal(0, diagnostic.TableRowIndex);
        Assert.Equal(0, diagnostic.TableColumnIndex);
    }

    private static byte[] BuildThreePagePdf() {
        return PdfDocument.Create()
            .Meta(title: "Workflow source", author: "OfficeIMO")
            .H1("Page A")
            .Paragraph(p => p.Text("First page body"))
            .PageBreak()
            .H1("Page B")
            .Paragraph(p => p.Text("Second page body"))
            .PageBreak()
            .H1("Page C")
            .Paragraph(p => p.Text("Third page body"))
            .ToBytes();
    }

    private static byte[] BuildThreeBookmarkPdf() {
        return PdfDocument.Create(new PdfOptions {
                CreateOutlineFromHeadings = true
            })
            .H1("Chapter One")
            .Paragraph(p => p.Text("First chapter body"))
            .PageBreak()
            .H1("Chapter Two")
            .Paragraph(p => p.Text("Second chapter body"))
            .PageBreak()
            .H1("Chapter Three")
            .Paragraph(p => p.Text("Third chapter body"))
            .ToBytes();
    }

    private static byte[] BuildOutOfOrderBookmarkPdf() {
        string pageOneContent = BuildStreamObject("BT /F1 12 Tf 72 720 Td (Chapter One) Tj ET");
        string pageTwoContent = BuildStreamObject("BT /F1 12 Tf 72 720 Td (Chapter Two) Tj ET");
        string pageThreeContent = BuildStreamObject("BT /F1 12 Tf 72 720 Td (Chapter Three) Tj ET");
        string pdf = string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /Outlines 9 0 R /PageMode /UseOutlines >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 3 /Kids [3 0 R 4 0 R 5 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Resources << /Font << /F1 6 0 R >> >> /Contents 7 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Resources << /Font << /F1 6 0 R >> >> /Contents 8 0 R >>",
            "endobj",
            "5 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Resources << /Font << /F1 6 0 R >> >> /Contents 10 0 R >>",
            "endobj",
            "6 0 obj",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
            "endobj",
            "7 0 obj",
            pageOneContent,
            "endobj",
            "8 0 obj",
            pageTwoContent,
            "endobj",
            "9 0 obj",
            "<< /Type /Outlines /First 11 0 R /Last 13 0 R /Count 3 >>",
            "endobj",
            "10 0 obj",
            pageThreeContent,
            "endobj",
            "11 0 obj",
            "<< /Title (Chapter Three) /Parent 9 0 R /Dest [5 0 R /Fit] /Next 12 0 R >>",
            "endobj",
            "12 0 obj",
            "<< /Title (Chapter One) /Parent 9 0 R /Dest [3 0 R /Fit] /Prev 11 0 R /Next 13 0 R >>",
            "endobj",
            "13 0 obj",
            "<< /Title (Chapter Two) /Parent 9 0 R /Dest [4 0 R /Fit] /Prev 12 0 R >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 14 >>",
            "startxref",
            "123",
            "%%EOF"
        });

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static string BuildStreamObject(string content) {
        byte[] bytes = Encoding.ASCII.GetBytes(content);
        return "<< /Length " + bytes.Length.ToString(System.Globalization.CultureInfo.InvariantCulture) + " >>\nstream\n" +
            content +
            "\nendstream";
    }

    private static byte[] BuildPdf(string title, string text) {
        return PdfDocument.Create()
            .Meta(title: title, author: "OfficeIMO")
            .H1(title)
            .Paragraph(p => p.Text(text))
            .ToBytes();
    }

    private static byte[] BuildSimpleFormPdf() {
        return PdfDocument.Create()
            .H1("Form")
            .TextField("Person.Name", width: 180, height: 24, value: "Original")
            .ToBytes();
    }
}
