using OfficeIMO.Reader;
using OfficeIMO.Rtf;
using System.Globalization;

namespace OfficeIMO.Reader.Rtf;

internal static partial class RtfReaderAdapter {
    private static IReadOnlyList<OfficeDocumentPage> BuildRtfPages(
        RtfDocument document,
        RtfRichProjection projection,
        string sourcePath,
        CancellationToken cancellationToken) {
        var pages = new List<RtfPageDraft> {
            CreateRtfPageDraft(document.PageSetup, document.PageSetup)
        };
        int currentPage = 1;
        IReadOnlyDictionary<IRtfBlock, RtfSectionMembership> sectionMembership =
            BuildRtfSectionMembership(document);
        var projectedBySourceIndex = projection.Blocks
            .Where(block => block.Location.SourceBlockIndex.HasValue)
            .GroupBy(block => block.Location.SourceBlockIndex!.Value)
            .ToDictionary(group => group.Key, group => group.First());

        for (int blockIndex = 0; blockIndex < document.Blocks.Count; blockIndex++) {
            cancellationToken.ThrowIfCancellationRequested();
            IRtfBlock sourceBlock = document.Blocks[blockIndex];
            sectionMembership.TryGetValue(sourceBlock, out RtfSectionMembership membership);
            RtfSection? section = membership.Section;
            bool startsSection = membership.StartsSection;
            RtfPageSetup setup = section?.PageSetup ?? document.PageSetup;

            if (startsSection && section != null) {
                if (IsFirstRtfSection(document, section)) {
                    ApplyRtfPageSetup(pages[currentPage - 1], setup, document.PageSetup);
                } else {
                    AdvanceForRtfSectionBreak(
                        pages,
                        ref currentPage,
                        section.BreakKind,
                        setup,
                        document.PageSetup);
                }
            }
            if (sourceBlock is RtfParagraph paragraph &&
                paragraph.PageBreakBefore &&
                pages[currentPage - 1].HasSourceContent) {
                currentPage++;
                EnsureRtfPage(pages, currentPage, setup, document.PageSetup);
            }
            if (!projectedBySourceIndex.TryGetValue(blockIndex, out OfficeDocumentBlock? projectedBlock)) {
                pages[currentPage - 1].HasSourceContent = true;
                continue;
            }

            OfficeDocumentFormField[] sourceForms = projection.Forms
                .Where(form => form.Location.SourceBlockIndex == blockIndex)
                .ToArray();
            IReadOnlyList<int> formFragmentIndexes = sourceBlock is RtfParagraph sourceParagraph
                ? GetRtfFormFragmentIndexes(sourceParagraph)
                : Array.Empty<int>();
            string[] fragments = projectedBlock.Text.Split(new[] { '\f' }, StringSplitOptions.None);
            IReadOnlyList<bool> sourceFragmentOccupancy =
                GetRtfSourceFragmentOccupancy(sourceBlock, fragments.Length);
            for (int fragmentIndex = 0; fragmentIndex < fragments.Length; fragmentIndex++) {
                if (fragmentIndex > 0) {
                    currentPage++;
                    EnsureRtfPage(pages, currentPage, setup, document.PageSetup);
                }
                RtfPageDraft page = pages[currentPage - 1];
                page.HasSourceContent |= sourceFragmentOccupancy[fragmentIndex];
                for (int formIndex = 0; formIndex < sourceForms.Length; formIndex++) {
                    int formFragmentIndex = formIndex < formFragmentIndexes.Count
                        ? formFragmentIndexes[formIndex]
                        : 0;
                    if (formFragmentIndex == fragmentIndex) {
                        page.Forms.Add(sourceForms[formIndex]);
                    }
                }
                string text = fragments[fragmentIndex];
                if (text.Length == 0 && fragments.Length > 1) {
                    continue;
                }
                page.Blocks.Add(CloneRtfPageBlock(projectedBlock, currentPage, text));
            }
        }

        return pages.Select((draft, index) => {
            int pageNumber = index + 1;
            string[] blockIds = draft.Blocks.Select(block => block.Id).ToArray();
            return new OfficeDocumentPage {
                Number = pageNumber,
                Name = "Page " + pageNumber.ToString(CultureInfo.InvariantCulture),
                Width = draft.Width,
                Height = draft.Height,
                Location = new ReaderLocation {
                    Path = sourcePath,
                    Page = pageNumber,
                    SourceBlockKind = "page",
                    BlockAnchor = "rtf-page-" + pageNumber.ToString("D4", CultureInfo.InvariantCulture)
                },
                Blocks = draft.Blocks.AsReadOnly(),
                Tables = projection.Tables.Where(table =>
                    table.Location?.BlockAnchor != null &&
                    blockIds.Contains(table.Location.BlockAnchor, StringComparer.Ordinal)).ToArray(),
                Links = projection.Links.Where(link =>
                    link.Location.BlockAnchor != null &&
                    blockIds.Any(blockId =>
                        link.Location.BlockAnchor.StartsWith(blockId, StringComparison.Ordinal))).ToArray(),
                Forms = draft.Forms.AsReadOnly()
            };
        }).ToArray();
    }

    private static IReadOnlyList<int> GetRtfFormFragmentIndexes(RtfParagraph paragraph) {
        var fragmentIndexes = new List<int>();
        int fragmentIndex = 0;

        void VisitParagraph(RtfParagraph current) {
            foreach (IRtfInline inline in current.Inlines) {
                if (inline is RtfBreak pageBreak &&
                    (pageBreak.Kind == RtfBreakKind.Page ||
                     pageBreak.Kind == RtfBreakKind.SoftPage)) {
                    fragmentIndex++;
                } else if (inline is RtfField field) {
                    if (field.FormFieldData != null) {
                        fragmentIndexes.Add(fragmentIndex);
                    }
                    VisitParagraph(field.Result);
                }
            }
        }

        VisitParagraph(paragraph);
        return fragmentIndexes.AsReadOnly();
    }

    private static IReadOnlyList<bool> GetRtfSourceFragmentOccupancy(
        IRtfBlock sourceBlock,
        int fragmentCount) {
        var occupied = new bool[Math.Max(1, fragmentCount)];
        if (sourceBlock is not RtfParagraph paragraph) {
            occupied[0] = true;
            return occupied;
        }

        int fragmentIndex = 0;
        void VisitParagraph(RtfParagraph current) {
            foreach (IRtfInline inline in current.Inlines) {
                if (inline is RtfBreak pageBreak &&
                    (pageBreak.Kind == RtfBreakKind.Page ||
                     pageBreak.Kind == RtfBreakKind.SoftPage)) {
                    fragmentIndex = Math.Min(fragmentIndex + 1, occupied.Length - 1);
                } else if (inline is RtfRun run) {
                    occupied[fragmentIndex] |= !string.IsNullOrEmpty(run.Text);
                } else if (inline is RtfField field) {
                    occupied[fragmentIndex] |= field.FormFieldData != null;
                    VisitParagraph(field.Result);
                } else if (inline is RtfGeneratedText generatedText) {
                    occupied[fragmentIndex] |= !string.IsNullOrEmpty(generatedText.ToPlainText());
                } else {
                    occupied[fragmentIndex] = true;
                }
            }
        }

        VisitParagraph(paragraph);
        return occupied;
    }

    private static IEnumerable<OfficeDocumentDiagnostic> BuildRtfPageDiagnostics(
        RtfDocument document,
        IReadOnlyList<OfficeDocumentPage> pages,
        string sourcePath) {
        yield return new OfficeDocumentDiagnostic {
            Severity = OfficeDocumentDiagnosticSeverity.Information,
            Category = OfficeDocumentDiagnosticCategory.Adapter,
            Code = "ReaderRtfExplicitPageLocations",
            Message = "RTF page locations were reconstructed from explicit page, soft-page, page-break-before, and page-starting section controls. Automatic text overflow was not calculated.",
            Source = "officeimo.reader.rtf",
            IsRecoverable = true,
            Location = new ReaderLocation {
                Path = sourcePath,
                SourceBlockKind = "page-index",
                BlockAnchor = "rtf-page-index"
            },
            Attributes = new Dictionary<string, string> {
                ["explicitPageCount"] = pages.Count.ToString(CultureInfo.InvariantCulture),
                ["savedPageCount"] = document.Info.NumberOfPages?.ToString(CultureInfo.InvariantCulture) ?? string.Empty
            }
        };
    }

    private static OfficeDocumentBlock CloneRtfPageBlock(
        OfficeDocumentBlock source,
        int pageNumber,
        string text) {
        return new OfficeDocumentBlock {
            Id = source.Id,
            Kind = source.Kind,
            Text = text,
            Level = source.Level,
            Marker = source.Marker,
            Location = new ReaderLocation {
                Path = source.Location.Path,
                SourceBlockIndex = source.Location.SourceBlockIndex,
                SourceBlockKind = source.Location.SourceBlockKind,
                BlockAnchor = source.Location.BlockAnchor,
                TableIndex = source.Location.TableIndex,
                Page = pageNumber
            }
        };
    }

    private static void AdvanceForRtfSectionBreak(
        List<RtfPageDraft> pages,
        ref int currentPage,
        RtfSectionBreakKind breakKind,
        RtfPageSetup setup,
        RtfPageSetup fallbackSetup) {
        switch (breakKind) {
            case RtfSectionBreakKind.Continuous:
            case RtfSectionBreakKind.Column:
                return;
            case RtfSectionBreakKind.EvenPage:
                currentPage++;
                EnsureRtfPage(pages, currentPage, setup, fallbackSetup);
                if (currentPage % 2 != 0) {
                    currentPage++;
                    EnsureRtfPage(pages, currentPage, setup, fallbackSetup);
                }
                return;
            case RtfSectionBreakKind.OddPage:
                currentPage++;
                EnsureRtfPage(pages, currentPage, setup, fallbackSetup);
                if (currentPage % 2 == 0) {
                    currentPage++;
                    EnsureRtfPage(pages, currentPage, setup, fallbackSetup);
                }
                return;
            default:
                currentPage++;
                EnsureRtfPage(pages, currentPage, setup, fallbackSetup);
                return;
        }
    }

    private static IReadOnlyDictionary<IRtfBlock, RtfSectionMembership> BuildRtfSectionMembership(
        RtfDocument document) {
        var membership = new Dictionary<IRtfBlock, RtfSectionMembership>(
            RtfBlockReferenceComparer.Instance);
        foreach (RtfSection section in document.Sections) {
            for (int index = 0; index < section.Blocks.Count; index++) {
                membership[section.Blocks[index]] = new RtfSectionMembership(
                    section,
                    startsSection: index == 0);
            }
        }
        return membership;
    }

    private static bool IsFirstRtfSection(RtfDocument document, RtfSection section) {
        return document.Sections.Count > 0 && ReferenceEquals(document.Sections[0], section);
    }

    private static RtfPageDraft CreateRtfPageDraft(
        RtfPageSetup setup,
        RtfPageSetup fallbackSetup) {
        var draft = new RtfPageDraft();
        ApplyRtfPageSetup(draft, setup, fallbackSetup);
        return draft;
    }

    private static void ApplyRtfPageSetup(
        RtfPageDraft draft,
        RtfPageSetup setup,
        RtfPageSetup fallbackSetup) {
        draft.Width = TwipsToPoints(setup.PaperWidthTwips ?? fallbackSetup.PaperWidthTwips);
        draft.Height = TwipsToPoints(setup.PaperHeightTwips ?? fallbackSetup.PaperHeightTwips);
    }

    private static void EnsureRtfPage(
        List<RtfPageDraft> pages,
        int pageNumber,
        RtfPageSetup setup,
        RtfPageSetup fallbackSetup) {
        while (pages.Count < pageNumber) {
            pages.Add(CreateRtfPageDraft(setup, fallbackSetup));
        }
    }

    private static double? TwipsToPoints(int? value) {
        return value.HasValue && value.Value > 0 ? value.Value / 20D : null;
    }

    private sealed class RtfPageDraft {
        internal double? Width { get; set; }
        internal double? Height { get; set; }
        internal bool HasSourceContent { get; set; }
        internal List<OfficeDocumentBlock> Blocks { get; } = new List<OfficeDocumentBlock>();
        internal List<OfficeDocumentFormField> Forms { get; } = new List<OfficeDocumentFormField>();
    }

    private readonly struct RtfSectionMembership {
        internal RtfSectionMembership(RtfSection section, bool startsSection) {
            Section = section;
            StartsSection = startsSection;
        }

        internal RtfSection? Section { get; }
        internal bool StartsSection { get; }
    }

    private sealed class RtfBlockReferenceComparer : IEqualityComparer<IRtfBlock> {
        internal static RtfBlockReferenceComparer Instance { get; } =
            new RtfBlockReferenceComparer();

        public bool Equals(IRtfBlock? x, IRtfBlock? y) => ReferenceEquals(x, y);

        public int GetHashCode(IRtfBlock obj) =>
            System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(obj);
    }
}
