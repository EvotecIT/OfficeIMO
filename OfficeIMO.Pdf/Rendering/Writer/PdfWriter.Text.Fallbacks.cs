namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private static readonly PdfStandardFont[] FallbackFontSlotFamilies = new[] {
        PdfStandardFont.Helvetica,
        PdfStandardFont.TimesRoman,
        PdfStandardFont.Courier
    };

    private static System.Collections.Generic.IReadOnlyList<TextRun> NormalizeFallbackRuns(System.Collections.Generic.IEnumerable<TextRun> runs, PdfStandardFont baseFont, PdfOptions? options) {
        Guard.NotNull(runs, nameof(runs));
        PdfEmbeddedFontFallbackSet? fallbackSet = options?.EmbeddedFontFallbacksSnapshot;
        if (fallbackSet == null) {
            return runs as System.Collections.Generic.IReadOnlyList<TextRun> ?? runs.ToArray();
        }

        var normalized = new System.Collections.Generic.List<TextRun>();
        foreach (TextRun run in runs) {
            if (run.InlineElement != null) {
                normalized.Add(run);
                continue;
            }

            if (CanWriteRunWithSelectedFont(run, baseFont, options)) {
                normalized.Add(run);
                continue;
            }

            if (TryPlanFallbackTextRuns(fallbackSet, run.Text, run, options, ResolveFontForRun(run, baseFont), out System.Collections.Generic.IReadOnlyList<TextRun> plannedRuns) ||
                TryPlanFallbackRunsPreservingSelectedFont(run, baseFont, options, fallbackSet, out plannedRuns)) {
                normalized.AddRange(plannedRuns);
            } else {
                normalized.Add(run);
            }
        }

        return normalized;
    }

    private static bool TryPlanFallbackRunsPreservingSelectedFont(
        TextRun run,
        PdfStandardFont baseFont,
        PdfOptions? options,
        PdfEmbeddedFontFallbackSet fallbackSet,
        out System.Collections.Generic.IReadOnlyList<TextRun> plannedRuns) {
        plannedRuns = Array.Empty<TextRun>();
        string text = run.Text ?? string.Empty;
        if (text.Length == 0 || IsLayoutControlRun(run)) {
            plannedRuns = new[] { run };
            return true;
        }

        PdfStandardFont fontForRun = ResolveFontForRun(run, baseFont);
        var runs = new System.Collections.Generic.List<TextRun>();

        for (int index = 0; index < text.Length;) {
            char ch = text[index];
            if (IsFallbackLayoutSeparator(ch)) {
                if (ch == '\t') {
                    runs.Add(TextRun.Tab(run.TabLeader, run.TabAlignment));
                } else {
                    runs.Add(TextRun.LineBreak());
                    if (ch == '\r' && index + 1 < text.Length && text[index + 1] == '\n') {
                        index++;
                    }
                }

                index++;
                continue;
            }

            int segmentStart = index;
            bool selectedFontCanWrite = TryGetSelectedFontCoveredFallbackTextLength(text, index, run, fontForRun, options, out int coveredLength);
            index += coveredLength;
            while (index < text.Length &&
                   !IsFallbackLayoutSeparator(text[index]) &&
                   selectedFontCanWrite == TryGetSelectedFontCoveredFallbackTextLength(text, index, run, fontForRun, options, out coveredLength)) {
                index += coveredLength;
            }

            string segment = text.Substring(segmentStart, index - segmentStart);
            if (selectedFontCanWrite) {
                runs.Add(CreateStyledTextRun(segment, run, run.Font));
                continue;
            }

            if (!TryPlanFallbackTextRuns(fallbackSet, segment, run, options, fontForRun, out System.Collections.Generic.IReadOnlyList<TextRun> fallbackRuns)) {
                plannedRuns = Array.Empty<TextRun>();
                return false;
            }

            runs.AddRange(fallbackRuns);
        }

        plannedRuns = runs.AsReadOnly();
        return true;
    }

    private static bool TryPlanFallbackTextRuns(
        PdfEmbeddedFontFallbackSet fallbackSet,
        string? text,
        TextRun styleTemplate,
        PdfOptions? options,
        PdfStandardFont selectedFont,
        out System.Collections.Generic.IReadOnlyList<TextRun> plannedRuns) {
        plannedRuns = Array.Empty<TextRun>();
        string value = text ?? string.Empty;
        PdfTextFallbackPlan plan = fallbackSet.PlanText(
            value,
            shapingMode: options?.TextShapingModeSnapshot ?? PdfTextShapingMode.UnicodeScalar);
        if (!plan.IsFullyCovered ||
            options == null) {
            return false;
        }

        if (fallbackSet.UsesNamedFontFamilies) {
            plannedRuns = plan.ToNamedTextRuns(fallbackSet.FontFamilyNames, styleTemplate);
            return true;
        }

        if (!TryResolveFallbackFontSlots(fallbackSet, plan, selectedFont, options, out System.Collections.Generic.IReadOnlyDictionary<int, PdfStandardFont> fontSlots)) {
            return false;
        }

        fallbackSet.RegisterFonts(options, fontSlots);
        foreach (PdfStandardFont slot in fontSlots.Values) {
            options.MarkEmbeddedFallbackFontFamilySlotUsed(slot);
        }

        plannedRuns = plan.ToTextRuns(fontSlots, styleTemplate);
        return true;
    }

    private static bool TryResolveFallbackFontSlots(
        PdfEmbeddedFontFallbackSet fallbackSet,
        PdfTextFallbackPlan plan,
        PdfStandardFont selectedFont,
        PdfOptions? options,
        out System.Collections.Generic.IReadOnlyDictionary<int, PdfStandardFont> fontSlots) {
        PdfStandardFont selectedFamily = PdfStandardFontMapper.GetFontFamily(selectedFont);
        var resolved = new System.Collections.Generic.Dictionary<int, PdfStandardFont>();
        var used = new System.Collections.Generic.HashSet<PdfStandardFont>();
        var reservedDocumentSlots = CreateReservedDocumentFontSlots(options);
        var plannedCandidateIndexes = new System.Collections.Generic.HashSet<int>(plan.Segments
            .Select(segment => segment.FontIndex)
            .Where(index => index >= 0 && index < fallbackSet.Candidates.Count)
            .Distinct());

        foreach (int index in plannedCandidateIndexes.OrderBy(index => index)) {
            PdfEmbeddedFontFallbackCandidate candidate = fallbackSet.Candidates[index];
            PdfStandardFont requested = PdfStandardFontMapper.GetFontFamily(fallbackSet.FontSlots[index]);
            if (CanUseFallbackFontSlot(requested, candidate, fallbackSet, plannedCandidateIndexes, selectedFamily, used, reservedDocumentSlots, options)) {
                resolved[index] = requested;
                used.Add(requested);
                continue;
            }

            PdfStandardFont? replacement = null;
            foreach (PdfStandardFont family in FallbackFontSlotFamilies) {
                if (CanUseFallbackFontSlot(family, candidate, fallbackSet, plannedCandidateIndexes, selectedFamily, used, reservedDocumentSlots, options)) {
                    replacement = family;
                    break;
                }
            }

            if (!replacement.HasValue) {
                fontSlots = new System.Collections.Generic.Dictionary<int, PdfStandardFont>();
                return false;
            }

            resolved[index] = replacement.Value;
            used.Add(replacement.Value);
        }

        fontSlots = resolved;
        return true;
    }

    private static bool CanUseFallbackFontSlot(
        PdfStandardFont family,
        PdfEmbeddedFontFallbackCandidate candidate,
        PdfEmbeddedFontFallbackSet fallbackSet,
        System.Collections.Generic.HashSet<int> plannedCandidateIndexes,
        PdfStandardFont selectedFamily,
        System.Collections.Generic.HashSet<PdfStandardFont> used,
        System.Collections.Generic.HashSet<PdfStandardFont> reservedDocumentSlots,
        PdfOptions? options) {
        PdfStandardFont normalized = PdfStandardFontMapper.GetFontFamily(family);
        return normalized != selectedFamily &&
               !used.Contains(normalized) &&
               !reservedDocumentSlots.Contains(normalized) &&
               FontFamilySlotIsEmptyOrCandidate(normalized, candidate, fallbackSet, plannedCandidateIndexes, options);
    }

    private static System.Collections.Generic.HashSet<PdfStandardFont> CreateReservedDocumentFontSlots(PdfOptions? options) {
        var reserved = new System.Collections.Generic.HashSet<PdfStandardFont>();
        if (options == null) {
            return reserved;
        }

        PdfOptions.AddRegisteredFontFamilySlot(reserved, options.DefaultFont);
        PdfOptions.AddRegisteredFontFamilySlot(reserved, options.HeaderFont);
        PdfOptions.AddRegisteredFontFamilySlot(reserved, options.FooterFont);
        return reserved;
    }

    private static bool IsFallbackLayoutSeparator(char ch) =>
        ch == '\n' || ch == '\r' || ch == '\t';

    private static bool TryGetSelectedFontCoveredFallbackTextLength(string text, int index, TextRun run, PdfStandardFont fontForRun, PdfOptions? options, out int length) {
        length = GetNextFallbackScalarLength(text, index);
        if (options != null &&
            options.TryResolveNamedFontFace(run.FontFamily, run.Bold, run.Italic, out PdfNamedFontFace namedFace)) {
            if (options.TryGetNamedFontProgram(namedFace, out PdfTrueTypeFontProgram? namedFontProgram) &&
                namedFontProgram != null) {
                return TryGetCoveredTextLength(text, index, namedFontProgram, options.TextShapingModeSnapshot, out length);
            }

            if (options.TryGetNamedOpenTypeCffFontProgram(namedFace, out PdfOpenTypeCffFontProgram? namedCffFontProgram) &&
                namedCffFontProgram != null) {
                return TryGetCoveredTextLength(text, index, namedCffFontProgram, options.TextShapingModeSnapshot, out length);
            }
        }

        if (options != null &&
            options.TryGetEmbeddedStandardFontProgram(fontForRun, out PdfTrueTypeFontProgram? fontProgram) &&
            fontProgram != null) {
            return TryGetCoveredTextLength(text, index, fontProgram, options.TextShapingModeSnapshot, out length);
        }

        if (options != null &&
            options.TryGetEmbeddedStandardOpenTypeCffFontProgram(fontForRun, out PdfOpenTypeCffFontProgram? cffFontProgram) &&
            cffFontProgram != null) {
            return TryGetCoveredTextLength(text, index, cffFontProgram, options.TextShapingModeSnapshot, out length);
        }

        return CanWriteTextWithSelectedFont(text.Substring(index, length), fontForRun, options);
    }

    private static int GetNextFallbackScalarLength(string text, int index) {
        if (index + 1 < text.Length &&
            char.IsHighSurrogate(text[index]) &&
            char.IsLowSurrogate(text[index + 1])) {
            return 2;
        }

        return 1;
    }

    private static bool FontFamilySlotIsEmptyOrCandidate(
        PdfStandardFont family,
        PdfEmbeddedFontFallbackCandidate candidate,
        PdfEmbeddedFontFallbackSet fallbackSet,
        System.Collections.Generic.HashSet<int> plannedCandidateIndexes,
        PdfOptions? options) {
        if (options == null) {
            return true;
        }

        foreach (PdfStandardFont variant in EnumerateFontFamilyVariants(family)) {
            if (options.TryGetEmbeddedStandardFont(variant, out PdfEmbeddedFont? embeddedFont) &&
                embeddedFont != null &&
                !embeddedFont.DataSnapshot.SequenceEqual(candidate.DataSnapshot) &&
                (options.IsEmbeddedFallbackFontFamilySlotUsed(family) ||
                 !EmbeddedFontMatchesUnusedFallbackCandidate(embeddedFont, fallbackSet, plannedCandidateIndexes))) {
                return false;
            }
        }

        return true;
    }

    private static bool EmbeddedFontMatchesUnusedFallbackCandidate(
        PdfEmbeddedFont embeddedFont,
        PdfEmbeddedFontFallbackSet fallbackSet,
        System.Collections.Generic.HashSet<int> plannedCandidateIndexes) {
        byte[] embeddedData = embeddedFont.DataSnapshot;
        for (int index = 0; index < fallbackSet.Candidates.Count; index++) {
            if (!plannedCandidateIndexes.Contains(index) &&
                embeddedData.SequenceEqual(fallbackSet.Candidates[index].DataSnapshot)) {
                return true;
            }
        }

        return false;
    }

    private static System.Collections.Generic.IEnumerable<PdfStandardFont> EnumerateFontFamilyVariants(PdfStandardFont family) {
        PdfStandardFont normalized = PdfStandardFontMapper.GetFontFamily(family);
        yield return normalized;
        yield return PdfStandardFontMapper.GetStyledFont(normalized, bold: true, italic: false);
        yield return PdfStandardFontMapper.GetStyledFont(normalized, bold: false, italic: true);
        yield return PdfStandardFontMapper.GetStyledFont(normalized, bold: true, italic: true);
    }
}
