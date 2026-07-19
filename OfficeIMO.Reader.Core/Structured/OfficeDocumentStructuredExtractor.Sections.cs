using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace OfficeIMO.Reader;

public static partial class OfficeDocumentStructuredExtractor {
    private static IReadOnlyList<OfficeDocumentStructuredSection> BuildSections(
        OfficeDocumentReadResult document,
        ExtractionState state) {
        IReadOnlyList<OfficeDocumentBlock> blocks = GetSectionBlocks(document);
        if (blocks.Count == 0) return Array.Empty<OfficeDocumentStructuredSection>();

        var sections = new List<OfficeDocumentStructuredSection>();
        SectionBuilder? current = null;
        for (int blockIndex = 0; blockIndex < blocks.Count; blockIndex++) {
            state.CancellationToken.ThrowIfCancellationRequested();
            OfficeDocumentBlock block = blocks[blockIndex];
            if (IsHeading(block)) {
                if (current != null) AddSection(sections, current, state);
                if (sections.Count >= state.Options.MaxSections) break;
                current = new SectionBuilder(
                    block.Text?.Trim(),
                    block.Level,
                    block.Location,
                    state.Options.MaxSectionCharacters);
                current.AddBlockId(block.Id);
                continue;
            }

            current ??= new SectionBuilder(
                heading: null,
                level: null,
                block.Location,
                state.Options.MaxSectionCharacters);
            current.AddBlock(block);
        }
        if (current != null && sections.Count < state.Options.MaxSections) AddSection(sections, current, state);
        if (RequiresMoreSections(blocks, sections, current, state.Options.MaxSections)) {
            state.AddLimitDiagnostic("structured-section-limit", state.Options.MaxSections, "sections");
        }
        return sections.Count == 0 ? Array.Empty<OfficeDocumentStructuredSection>() : sections.ToArray();
    }

    private static IReadOnlyList<OfficeDocumentBlock> GetSectionBlocks(OfficeDocumentReadResult document) {
        var blocks = new List<OfficeDocumentBlock>();
        var seenIds = new HashSet<string>(StringComparer.Ordinal);
        foreach (OfficeDocumentBlock block in OfficeDocumentModelTraversal.Blocks(document)) {
            if (string.IsNullOrWhiteSpace(block.Id) || seenIds.Add(block.Id)) blocks.Add(block);
        }
        return blocks.Count == 0 ? Array.Empty<OfficeDocumentBlock>() : blocks.ToArray();
    }

    private static bool IsHeading(OfficeDocumentBlock block) =>
        string.Equals(block.Kind?.Trim(), "heading", StringComparison.OrdinalIgnoreCase);

    private static void AddSection(
        ICollection<OfficeDocumentStructuredSection> sections,
        SectionBuilder builder,
        ExtractionState state) {
        int index = sections.Count;
        sections.Add(new OfficeDocumentStructuredSection {
            Id = "section-" + index.ToString("D4", CultureInfo.InvariantCulture),
            Heading = builder.Heading,
            Level = builder.Level,
            Text = builder.Text,
            BlockIds = builder.BlockIds,
            Location = builder.Location,
            Truncated = builder.Truncated
        });
        if (builder.Truncated) {
            state.AddLimitDiagnostic("structured-section-character-limit", state.Options.MaxSectionCharacters, "section characters");
        }
    }

    private static bool RequiresMoreSections(
        IReadOnlyList<OfficeDocumentBlock> blocks,
        IReadOnlyList<OfficeDocumentStructuredSection> sections,
        SectionBuilder? current,
        int maxSections) {
        if (sections.Count < maxSections) return false;
        int headingCount = 0;
        bool hasPreamble = false;
        for (int index = 0; index < blocks.Count; index++) {
            if (IsHeading(blocks[index])) headingCount++;
            else if (headingCount == 0) hasPreamble = true;
        }
        int required = headingCount + (hasPreamble ? 1 : 0);
        return required > maxSections || (current != null && sections.Count >= maxSections && required > sections.Count);
    }

    private sealed class SectionBuilder {
        private readonly int _maxCharacters;
        private readonly StringBuilder _text = new StringBuilder();
        private readonly List<string> _blockIds = new List<string>();

        internal SectionBuilder(string? heading, int? level, ReaderLocation? location, int maxCharacters) {
            Heading = string.IsNullOrWhiteSpace(heading) ? null : heading;
            Level = level;
            Location = location;
            _maxCharacters = maxCharacters;
        }

        internal string? Heading { get; }
        internal int? Level { get; }
        internal ReaderLocation? Location { get; }
        internal string Text => _text.ToString();
        internal IReadOnlyList<string> BlockIds => _blockIds.Count == 0 ? Array.Empty<string>() : _blockIds.ToArray();
        internal bool Truncated { get; private set; }

        internal void AddBlock(OfficeDocumentBlock block) {
            AddBlockId(block.Id);
            string value = block.Text?.Trim() ?? string.Empty;
            if (value.Length == 0 || Truncated) return;
            int separatorLength = _text.Length == 0 ? 0 : 1;
            int remaining = _maxCharacters - _text.Length - separatorLength;
            if (remaining <= 0) {
                Truncated = true;
                return;
            }
            if (separatorLength > 0) _text.Append('\n');
            if (value.Length <= remaining) {
                _text.Append(value);
                return;
            }
            _text.Append(value, 0, remaining);
            Truncated = true;
        }

        internal void AddBlockId(string? id) {
            if (!string.IsNullOrWhiteSpace(id)) _blockIds.Add(id!);
        }
    }
}
