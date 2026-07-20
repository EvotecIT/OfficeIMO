using OfficeIMO.Pdf;
using OfficeIMO.Reader;
using OfficeIMO.Reader.Pdf;
using OfficeIMO.Reader.Rtf;
using OfficeIMO.Rtf;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

[Collection("ReaderRegistryNonParallel")]
public sealed class ReaderPageLocationTests {
    [Fact]
    public void Search_AggregatesEveryPageForARepeatedSourceBlock() {
        var sourceBlock = new OfficeDocumentBlock {
            Id = "paragraph-0042",
            Kind = "paragraph",
            Text = "A paragraph containing the search needle.",
            Location = new ReaderLocation { BlockAnchor = "paragraph-0042" }
        };
        OfficeDocumentReadResult document = new OfficeDocumentReadResult {
            Kind = ReaderInputKind.Word,
            CapabilitiesUsed = new[] { "officeimo.reader.word.pages.computed" },
            Blocks = new[] { sourceBlock },
            Metadata = new[] {
                new OfficeDocumentMetadataEntry {
                    Id = "page-count",
                    Name = "PageCount",
                    Value = "100",
                    ValueType = "count"
                }
            },
            Pages = new[] {
                Page(5, PageBlock(sourceBlock, 5)),
                Page(18, PageBlock(sourceBlock, 18)),
                Page(23, PageBlock(
                    sourceBlock,
                    23,
                    "A needless page fragment with only a substring match."))
            }
        };

        OfficeDocumentSearchResult search = document.Search("needle");
        OfficeDocumentSearchResult wholeWordSearch = document.Search(
            "needle",
            new OfficeDocumentSearchOptions { WholeWord = true });
        OfficeDocumentSearchHit hit = Assert.Single(search.Hits);

        Assert.Equal(new[] { 5, 18, 23 }, search.PageNumbers);
        Assert.Equal(new[] { 5, 18 }, wholeWordSearch.PageNumbers);
        Assert.Equal(100, search.TotalPageCount);
        Assert.Equal(
            new[] { "Page 5 of 100", "Page 18 of 100", "Page 23 of 100" },
            hit.Pages.Select(location => location.Display));
        Assert.Equal(OfficeDocumentPageProvenance.Computed, hit.Pages[0].Provenance);
    }

    [Fact]
    public void Search_WholeWordDoesNotFallBackToSubstringOnlyPageFragments() {
        var sourceBlock = new OfficeDocumentBlock {
            Id = "paragraph-0001",
            Kind = "paragraph",
            Text = "The source contains the whole word needle.",
            Location = new ReaderLocation { BlockAnchor = "paragraph-0001" }
        };
        OfficeDocumentReadResult document = new OfficeDocumentReadResult {
            Kind = ReaderInputKind.Word,
            Blocks = new[] { sourceBlock },
            Pages = new[] {
                Page(1, PageBlock(sourceBlock, 1, "The visible fragment says needless."))
            }
        };

        OfficeDocumentSearchHit hit = Assert.Single(document.Search(
            "needle",
            new OfficeDocumentSearchOptions { WholeWord = true }).Hits);

        Assert.Empty(hit.Pages);
    }

    [Fact]
    public void Search_AssociatesRepeatedOccurrencesWithTheirActualPageFragments() {
        var sourceBlock = new OfficeDocumentBlock {
            Id = "paragraph-0002",
            Kind = "paragraph",
            Text = "First needle before the page boundary. Second needle after it.",
            Location = new ReaderLocation { BlockAnchor = "paragraph-0002" }
        };
        OfficeDocumentReadResult document = new OfficeDocumentReadResult {
            Kind = ReaderInputKind.Word,
            Blocks = new[] { sourceBlock },
            Pages = new[] {
                Page(1, PageBlock(sourceBlock, 1, "First needle before the page boundary. ")),
                Page(2, PageBlock(sourceBlock, 2, "Second needle after it."))
            }
        };

        OfficeDocumentSearchResult result = document.Search("needle");
        OfficeDocumentSearchResult limited = document.Search(
            "needle",
            new OfficeDocumentSearchOptions { MaximumResults = 1 });

        Assert.Equal(2, result.Hits.Count);
        Assert.Equal(new[] { 1 }, result.Hits[0].Pages.Select(page => page.Number!.Value).ToArray());
        Assert.Equal(new[] { 2 }, result.Hits[1].Pages.Select(page => page.Number!.Value).ToArray());
        Assert.Equal(new[] { 1, 2 }, result.PageNumbers);
        OfficeDocumentSearchHit limitedHit = Assert.Single(limited.Hits);
        Assert.Equal(
            new[] { 1 },
            limitedHit.Pages.Select(page => page.Number!.Value).ToArray());
    }

    [Fact]
    public void Search_RetainsEveryContainingPageWhenAMatchCrossesPageFragments() {
        var sourceBlock = new OfficeDocumentBlock {
            Id = "paragraph-cross-page",
            Kind = "paragraph",
            Text = "A needle spans two page fragments.",
            Location = new ReaderLocation { BlockAnchor = "paragraph-cross-page" }
        };
        OfficeDocumentReadResult document = new OfficeDocumentReadResult {
            Kind = ReaderInputKind.Word,
            Blocks = new[] { sourceBlock },
            Pages = new[] {
                Page(1, PageBlock(sourceBlock, 1, "A nee")),
                Page(2, PageBlock(sourceBlock, 2, "dle spans two page fragments."))
            }
        };

        OfficeDocumentSearchHit hit = Assert.Single(document.Search(
            "needle",
            new OfficeDocumentSearchOptions { WholeWord = true }).Hits);

        Assert.Equal(new[] { 1, 2 }, hit.Pages.Select(page => page.Number!.Value).ToArray());
    }

    [Fact]
    public void Search_MapsWholeAndCrossPageOccurrencesIndependently() {
        var sourceBlock = new OfficeDocumentBlock {
            Id = "paragraph-mixed-pages",
            Kind = "paragraph",
            Text = "First needle then second needle.",
            Location = new ReaderLocation { BlockAnchor = "paragraph-mixed-pages" }
        };
        OfficeDocumentReadResult document = new OfficeDocumentReadResult {
            Kind = ReaderInputKind.Word,
            Blocks = new[] { sourceBlock },
            Pages = new[] {
                Page(1, PageBlock(sourceBlock, 1, "First needle then ")),
                Page(2, PageBlock(sourceBlock, 2, "second nee")),
                Page(3, PageBlock(sourceBlock, 3, "dle."))
            }
        };

        OfficeDocumentSearchResult result = document.Search("needle");

        Assert.Equal(2, result.Hits.Count);
        Assert.Equal(new[] { 1 }, result.Hits[0].Pages.Select(page => page.Number!.Value).ToArray());
        Assert.Equal(
            new[] { 2, 3 },
            result.Hits[1].Pages.Select(page => page.Number!.Value).ToArray());
    }

    [Fact]
    public void Search_MapsMismatchedFragmentsByOccurrenceOrderAndOptions() {
        var sourceBlock = new OfficeDocumentBlock {
            Id = "paragraph-mismatched-pages",
            Kind = "paragraph",
            Text = "First Needle then second needle.",
            Location = new ReaderLocation { BlockAnchor = "paragraph-mismatched-pages" }
        };
        OfficeDocumentReadResult document = new OfficeDocumentReadResult {
            Kind = ReaderInputKind.Word,
            Blocks = new[] { sourceBlock },
            Pages = new[] {
                Page(1, PageBlock(sourceBlock, 1, "First NEEDLE then ")),
                Page(2, PageBlock(sourceBlock, 2, "second nee")),
                Page(3, PageBlock(sourceBlock, 3, "dle!"))
            }
        };

        OfficeDocumentSearchResult wholeWord = document.Search(
            "needle",
            new OfficeDocumentSearchOptions { WholeWord = true });
        OfficeDocumentSearchResult matchCase = document.Search(
            "needle",
            new OfficeDocumentSearchOptions {
                MatchCase = true,
                WholeWord = true
            });

        Assert.Equal(2, wholeWord.Hits.Count);
        Assert.Equal(
            new[] { 1 },
            wholeWord.Hits[0].Pages.Select(page => page.Number!.Value).ToArray());
        Assert.Equal(
            new[] { 2, 3 },
            wholeWord.Hits[1].Pages.Select(page => page.Number!.Value).ToArray());
        OfficeDocumentSearchHit caseSensitiveHit = Assert.Single(matchCase.Hits);
        Assert.Equal(
            new[] { 2, 3 },
            caseSensitiveHit.Pages.Select(page => page.Number!.Value).ToArray());
    }

    [Fact]
    public void Search_LimitPreservesMismatchedFragmentOccurrenceCorrelation() {
        var sourceBlock = new OfficeDocumentBlock {
            Id = "paragraph-limited-mismatched-pages",
            Kind = "paragraph",
            Text = "Needle needle.",
            Location = new ReaderLocation { BlockAnchor = "paragraph-limited-mismatched-pages" }
        };
        OfficeDocumentReadResult document = new OfficeDocumentReadResult {
            Kind = ReaderInputKind.Word,
            Blocks = new[] { sourceBlock },
            Pages = new[] {
                Page(1, PageBlock(sourceBlock, 1, "NEEDLE ")),
                Page(2, PageBlock(sourceBlock, 2, "needle."))
            }
        };

        OfficeDocumentSearchResult result = document.Search(
            "needle",
            new OfficeDocumentSearchOptions { MaximumResults = 1 });
        OfficeDocumentSearchHit hit = Assert.Single(result.Hits);

        Assert.Equal(new[] { 1 }, result.PageNumbers);
        Assert.Equal(
            new[] { 1 },
            hit.Pages.Select(page => page.Number!.Value).ToArray());
    }

    [Fact]
    public void Search_DeclinesCitationsWhenFragmentOptionsSwapOccurrenceIdentity() {
        var caseSourceBlock = new OfficeDocumentBlock {
            Id = "paragraph-swapped-case",
            Kind = "paragraph",
            Text = "Needle first, needle second"
        };
        OfficeDocumentReadResult caseDocument = new OfficeDocumentReadResult {
            Blocks = new[] { caseSourceBlock },
            Pages = new[] {
                Page(1, PageBlock(caseSourceBlock, 1, "needle first, ")),
                Page(2, PageBlock(caseSourceBlock, 2, "Needle second"))
            }
        };
        var boundarySourceBlock = new OfficeDocumentBlock {
            Id = "paragraph-swapped-boundary",
            Kind = "paragraph",
            Text = "needleless then needle"
        };
        OfficeDocumentReadResult boundaryDocument = new OfficeDocumentReadResult {
            Blocks = new[] { boundarySourceBlock },
            Pages = new[] {
                Page(1, PageBlock(boundarySourceBlock, 1, "needle then ")),
                Page(2, PageBlock(boundarySourceBlock, 2, "needleless"))
            }
        };

        OfficeDocumentSearchHit caseSensitiveHit = Assert.Single(caseDocument.Search(
            "needle",
            new OfficeDocumentSearchOptions { MatchCase = true }).Hits);
        OfficeDocumentSearchHit wholeWordHit = Assert.Single(boundaryDocument.Search(
            "needle",
            new OfficeDocumentSearchOptions { WholeWord = true }).Hits);

        Assert.Empty(caseSensitiveHit.Pages);
        Assert.Empty(wholeWordHit.Pages);
    }

    [Fact]
    public void Search_DeclinesCitationsWhenFragmentOccurrenceCountsDiffer() {
        var sourceBlock = new OfficeDocumentBlock {
            Id = "paragraph-occurrence-count",
            Kind = "paragraph",
            Text = "needle one; needle two"
        };
        OfficeDocumentReadResult missingDocument = new OfficeDocumentReadResult {
            Blocks = new[] { sourceBlock },
            Pages = new[] {
                Page(1, PageBlock(sourceBlock, 1, "needle one"))
            }
        };
        OfficeDocumentReadResult extraDocument = new OfficeDocumentReadResult {
            Blocks = new[] { sourceBlock },
            Pages = new[] {
                Page(1, PageBlock(
                    sourceBlock,
                    1,
                    "needle one; needle two; needle extra"))
            }
        };

        OfficeDocumentSearchResult missing = missingDocument.Search("needle");
        OfficeDocumentSearchResult extra = extraDocument.Search("needle");

        Assert.Equal(2, missing.Hits.Count);
        Assert.All(missing.Hits, hit => Assert.Empty(hit.Pages));
        Assert.Equal(2, extra.Hits.Count);
        Assert.All(extra.Hits, hit => Assert.Empty(hit.Pages));
    }

    [Fact]
    public void Search_MaximumResultsBoundsOccurrenceCollectionForLargeBlocks() {
        var sourceBlock = new OfficeDocumentBlock {
            Id = "paragraph-large",
            Kind = "paragraph",
            Text = new string('a', 250_000)
        };
        var document = new OfficeDocumentReadResult {
            Blocks = new[] { sourceBlock }
        };

#if NET8_0_OR_GREATER
        long allocatedBefore = GC.GetAllocatedBytesForCurrentThread();
#endif
        OfficeDocumentSearchHit hit = Assert.Single(document.Search(
            "a",
            new OfficeDocumentSearchOptions { MaximumResults = 1 }).Hits);
#if NET8_0_OR_GREATER
        long allocated = GC.GetAllocatedBytesForCurrentThread() - allocatedBefore;
        Assert.InRange(allocated, 0L, 512L * 1024L);
#endif

        Assert.Equal(0, hit.StartIndex);
    }

    [Fact]
    public void PdfReader_ExposesNativePageSearchAndPageMarkedMarkdown() {
        byte[] pdf = PdfDocument.Create(new PdfOptions {
                PageWidth = 420,
                PageHeight = 360,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36
            })
            .Paragraph(paragraph => paragraph.Text("First page body."))
            .PageBreak()
            .Paragraph(paragraph => paragraph.Text("Needle on native page two."))
            .ToBytes();

        OfficeDocumentReadResult document = PdfReaderAdapter.ReadDocument(
            pdf,
            sourceName: "native-pages.pdf");
        OfficeDocumentSearchResult search = document.Search("needle");
        OfficeDocumentPageMarkdown pageTwo = Assert.Single(
            document.GetPageMarkdown(),
            page => page.Page.Number == 2);

        Assert.Contains("officeimo.reader.pdf.pages.native", document.CapabilitiesUsed);
        Assert.Equal(OfficeDocumentPageProvenance.Native, document.GetPageProvenance());
        Assert.Equal(new[] { 2 }, search.PageNumbers);
        Assert.Equal(2, search.TotalPageCount);
        Assert.Contains("page: 2/2; provenance: Native", pageTwo.Markdown, StringComparison.Ordinal);
        Assert.Contains("Needle on native page two.", pageTwo.Markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void RtfReader_UsesExplicitBreaksAndSavedPageCountWithoutClaimingOverflowLayout() {
        RtfDocument rtf = RtfDocument.Create();
        rtf.Info.NumberOfPages = 20;
        rtf.PageSetup.SetPaperSize(12240, 15840);
        rtf.AddParagraph("First page body.");
        RtfParagraph paragraph = rtf.AddParagraph("Needle on explicit page two.");
        paragraph.PageBreakBefore = true;
        RtfField field = paragraph.AddField("FORMTEXT");
        field.AddText("Ada");
        field.SetFormFieldData(data => {
            data.Kind = RtfFormFieldKind.Text;
            data.Name = "Patient";
        });
        paragraph.AddPageBreak();
        paragraph.AddText("Needle on explicit page three.");

        OfficeDocumentReadResult document = RtfReaderAdapter.ReadDocument(
            rtf,
            "explicit-pages.rtf",
            rtfOptions: new ReaderRtfOptions { IncludePageLocations = true });
        OfficeDocumentSearchResult search = document.Search("needle");

        Assert.Contains("officeimo.reader.rtf.pages.explicit", document.CapabilitiesUsed);
        Assert.Equal(OfficeDocumentPageProvenance.ExplicitBreak, document.GetPageProvenance());
        Assert.Equal(3, document.Pages.Count);
        Assert.Equal(612D, document.Pages[0].Width);
        Assert.Equal(792D, document.Pages[0].Height);
        Assert.Equal(20, search.TotalPageCount);
        Assert.Equal(new[] { 2, 3 }, search.PageNumbers);
        Assert.Equal("Patient", Assert.Single(document.Pages[1].Forms).Name);
        Assert.Empty(document.Pages[2].Forms);
        Assert.Contains(document.Diagnostics, diagnostic =>
            diagnostic.Code == "ReaderRtfExplicitPageLocations" &&
            diagnostic.Severity == OfficeDocumentDiagnosticSeverity.Information);
        Assert.Contains("page: 3/20; provenance: ExplicitBreak", document.ToPageMarkedMarkdown(), StringComparison.Ordinal);

        OfficeDocumentReadResult defaultDocument = RtfReaderAdapter.ReadDocument(rtf, "default.rtf");
        Assert.Empty(defaultDocument.Pages);
        Assert.DoesNotContain("officeimo.reader.rtf.pages.explicit", defaultDocument.CapabilitiesUsed);
    }

    [Fact]
    public void RtfReader_PageBreakBeforeEnsuresPageStartWithoutAddingBlankPages() {
        RtfDocument firstBlockBreak = RtfDocument.Create();
        firstBlockBreak.AddParagraph("First body.").PageBreakBefore = true;
        OfficeDocumentReadResult firstResult = RtfReaderAdapter.ReadDocument(
            firstBlockBreak,
            "first-break.rtf",
            rtfOptions: new ReaderRtfOptions { IncludePageLocations = true });

        RtfDocument postBreak = RtfDocument.Create();
        RtfParagraph before = postBreak.AddParagraph("Before explicit break.");
        before.AddPageBreak();
        postBreak.AddParagraph("After explicit break.").PageBreakBefore = true;
        OfficeDocumentReadResult postBreakResult = RtfReaderAdapter.ReadDocument(
            postBreak,
            "post-break.rtf",
            rtfOptions: new ReaderRtfOptions { IncludePageLocations = true });

        Assert.Single(firstResult.Pages);
        Assert.Equal("First body.", Assert.Single(firstResult.Pages[0].Blocks).Text);
        Assert.Equal(2, postBreakResult.Pages.Count);
        Assert.Contains("After explicit break.", postBreakResult.Pages[1].Blocks.Select(block => block.Text));
    }

    [Fact]
    public void RtfReader_PageBreakBeforeUsesSourceOccupancyWhenProjectionOmitsABlock() {
        RtfDocument rtf = RtfDocument.Create();
        rtf.AddImage(RtfImageFormat.Png, new byte[] { 137, 80, 78, 71 });
        rtf.AddParagraph("Paragraph after omitted image.").PageBreakBefore = true;

        OfficeDocumentReadResult document = RtfReaderAdapter.ReadDocument(
            rtf,
            "omitted-image.rtf",
            rtfOptions: new ReaderRtfOptions {
                IncludeImagePlaceholders = false,
                IncludePageLocations = true
            });

        Assert.Equal(2, document.Pages.Count);
        Assert.Empty(document.Pages[0].Blocks);
        Assert.Equal(
            "Paragraph after omitted image.",
            Assert.Single(document.Pages[1].Blocks).Text);
    }

    [Fact]
    public void RtfReader_UsesFirstSectionPageDimensions() {
        RtfDocument rtf = RtfDocument.Create();
        RtfSection section = rtf.AddSection();
        section.PageSetup.SetPaperSize(10000, 12000);
        section.AddParagraph("Section body.");

        OfficeDocumentReadResult document = RtfReaderAdapter.ReadDocument(
            rtf,
            "section-layout.rtf",
            rtfOptions: new ReaderRtfOptions { IncludePageLocations = true });

        OfficeDocumentPage page = Assert.Single(document.Pages);
        Assert.Equal(500D, page.Width);
        Assert.Equal(600D, page.Height);
    }

    [Fact]
    public void RtfReader_FirstSectionBreakAdvancesWhenRootContentPrecedesIt() {
        RtfDocument rtf = RtfDocument.Create();
        rtf.AddParagraph("Root body.");
        RtfSection section = rtf.AddSection();
        section.AddParagraph("Section body.");

        OfficeDocumentReadResult document = RtfReaderAdapter.ReadDocument(
            rtf,
            "root-before-section.rtf",
            rtfOptions: new ReaderRtfOptions { IncludePageLocations = true });

        Assert.Equal(2, document.Pages.Count);
        Assert.Equal("Root body.", Assert.Single(document.Pages[0].Blocks).Text);
        Assert.Equal("Section body.", Assert.Single(document.Pages[1].Blocks).Text);
    }

    [Fact]
    public void WordReader_ComputesPageFragmentsForSearchAndMarkdown() {
        using var stream = new MemoryStream();
        using (WordDocument word = WordDocument.Create(stream)) {
            word.AddParagraph("First page body.");
            word.AddPageBreak();
            word.AddParagraph("Needle on computed page two.");
            word.Save();
        }
        stream.Position = 0;

        OfficeDocumentReadResult document = OfficeIMO.Reader.Tests.ReaderTestReaders
            .Word(includePageLocations: true)
            .ReadDocument(stream, "computed-pages.docx");
        OfficeDocumentSearchResult search = document.Search("needle");

        Assert.Contains("officeimo.reader.word.pages.computed", document.CapabilitiesUsed);
        Assert.Equal(OfficeDocumentPageProvenance.Computed, document.GetPageProvenance());
        Assert.True(document.Pages.Count >= 2);
        Assert.Equal(new[] { 2 }, search.PageNumbers);
        Assert.Contains("page: 2/", document.ToPageMarkedMarkdown(), StringComparison.Ordinal);
        Assert.Contains("Needle on computed page two.", document.Pages[1].Blocks.Select(block => block.Text));
        Assert.NotNull(Assert.Single(search.Hits).Pages[0].Regions.Single());
    }

    [Fact]
    public void WordMapping_DoesNotFallBackToSourceTextForBlankVisualFragments() {
        string text = OfficeIMO.Reader.Word.WordRichMapping.CombineWordFragmentText(
            new[] { string.Empty, "   " });

        Assert.Equal(string.Empty, text);
    }

    private static OfficeDocumentPage Page(int number, OfficeDocumentBlock block) {
        return new OfficeDocumentPage {
            Number = number,
            Name = "Page " + number,
            Location = new ReaderLocation { Page = number },
            Blocks = new[] { block }
        };
    }

    private static OfficeDocumentBlock PageBlock(
        OfficeDocumentBlock source,
        int page,
        string? text = null) {
        return new OfficeDocumentBlock {
            Id = source.Id,
            Kind = source.Kind,
            Text = text ?? source.Text,
            Location = new ReaderLocation {
                Page = page,
                BlockAnchor = source.Id
            }
        };
    }
}
