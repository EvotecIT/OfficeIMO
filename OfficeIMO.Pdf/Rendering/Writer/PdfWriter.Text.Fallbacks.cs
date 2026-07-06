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
            if (ch == '\n' || ch == '\r' || ch == '\t') {
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
            if (ch == ' ') {
                while (index < text.Length && text[index] == ' ') {
                    index++;
                }

                runs.Add(CreateStyledTextRun(text.Substring(segmentStart, index - segmentStart), run, run.Font));
                continue;
            }

            while (index < text.Length &&
                   text[index] != ' ' &&
                   text[index] != '\n' &&
                   text[index] != '\r' &&
                   text[index] != '\t') {
                index++;
            }

            string token = text.Substring(segmentStart, index - segmentStart);
            if (CanWriteTextWithSelectedFont(token, fontForRun, options)) {
                runs.Add(CreateStyledTextRun(token, run, run.Font));
                continue;
            }

            if (!TryPlanFallbackTextRuns(fallbackSet, token, run, options, fontForRun, out System.Collections.Generic.IReadOnlyList<TextRun> fallbackRuns)) {
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
            !TryResolveFallbackFontSlots(fallbackSet, selectedFont, options, out System.Collections.Generic.IReadOnlyList<PdfStandardFont> fontSlots)) {
            return false;
        }

        if (options == null) {
            return false;
        }

        fallbackSet.RegisterFonts(options, fontSlots);
        plannedRuns = plan.ToTextRuns(fontSlots, styleTemplate);
        return true;
    }

    private static bool TryResolveFallbackFontSlots(
        PdfEmbeddedFontFallbackSet fallbackSet,
        PdfStandardFont selectedFont,
        PdfOptions? options,
        out System.Collections.Generic.IReadOnlyList<PdfStandardFont> fontSlots) {
        PdfStandardFont selectedFamily = PdfStandardFontMapper.GetFontFamily(selectedFont);
        var resolved = new System.Collections.Generic.List<PdfStandardFont>(fallbackSet.Candidates.Count);
        var used = new System.Collections.Generic.HashSet<PdfStandardFont>();

        for (int index = 0; index < fallbackSet.Candidates.Count; index++) {
            PdfEmbeddedFontFallbackCandidate candidate = fallbackSet.Candidates[index];
            PdfStandardFont requested = PdfStandardFontMapper.GetFontFamily(fallbackSet.FontSlots[index]);
            if (CanUseFallbackFontSlot(requested, candidate, selectedFamily, used, options)) {
                resolved.Add(requested);
                used.Add(requested);
                continue;
            }

            PdfStandardFont? replacement = null;
            foreach (PdfStandardFont family in FallbackFontSlotFamilies) {
                if (CanUseFallbackFontSlot(family, candidate, selectedFamily, used, options)) {
                    replacement = family;
                    break;
                }
            }

            if (!replacement.HasValue) {
                fontSlots = Array.Empty<PdfStandardFont>();
                return false;
            }

            resolved.Add(replacement.Value);
            used.Add(replacement.Value);
        }

        fontSlots = resolved.AsReadOnly();
        return true;
    }

    private static bool CanUseFallbackFontSlot(
        PdfStandardFont family,
        PdfEmbeddedFontFallbackCandidate candidate,
        PdfStandardFont selectedFamily,
        System.Collections.Generic.HashSet<PdfStandardFont> used,
        PdfOptions? options) {
        PdfStandardFont normalized = PdfStandardFontMapper.GetFontFamily(family);
        return normalized != selectedFamily &&
               !used.Contains(normalized) &&
               FontFamilySlotIsEmptyOrCandidate(normalized, candidate, options);
    }

    private static bool FontFamilySlotIsEmptyOrCandidate(PdfStandardFont family, PdfEmbeddedFontFallbackCandidate candidate, PdfOptions? options) {
        if (options == null) {
            return true;
        }

        foreach (PdfStandardFont variant in EnumerateFontFamilyVariants(family)) {
            if (options.TryGetEmbeddedStandardFont(variant, out PdfEmbeddedFont? embeddedFont) &&
                embeddedFont != null &&
                !embeddedFont.DataSnapshot.SequenceEqual(candidate.DataSnapshot)) {
                return false;
            }
        }

        return true;
    }

    private static System.Collections.Generic.IEnumerable<PdfStandardFont> EnumerateFontFamilyVariants(PdfStandardFont family) {
        PdfStandardFont normalized = PdfStandardFontMapper.GetFontFamily(family);
        yield return normalized;
        yield return PdfStandardFontMapper.GetStyledFont(normalized, bold: true, italic: false);
        yield return PdfStandardFontMapper.GetStyledFont(normalized, bold: false, italic: true);
        yield return PdfStandardFontMapper.GetStyledFont(normalized, bold: true, italic: true);
    }
}
