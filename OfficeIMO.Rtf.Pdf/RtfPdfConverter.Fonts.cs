using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Rtf.Pdf;

internal static partial class RtfPdfConverter {
    private static IReadOnlyDictionary<int, PdfCore.PdfStandardFont> ConfigureDocumentFonts(
        RtfDocument document,
        PdfCore.PdfOptions pdfOptions,
        RtfToPdfOptions options) {
        bool allowSystemFontEmbedding =
            options.ResourcePolicy.AllowSystemFontEmbedding &&
            options.ResourcePolicy.AllowDocumentFontEmbedding;
        bool systemFontBudgetReported = false;
        var fontSlots = new Dictionary<int, PdfCore.PdfStandardFont>();
        var familySlots = new Dictionary<string, PdfCore.PdfStandardFont>(StringComparer.OrdinalIgnoreCase);
        var unresolvedFamilies = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var systemFontFamilies = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        HashSet<int> referencedFontIds = CollectReferencedFontIds(document, options, pdfOptions);
        HashSet<PdfCore.PdfStandardFont> registeredFontSlots = pdfOptions.CreateRegisteredFontFamilySlots(includeDocumentFontSlots: false);
        if (pdfOptions.HasExplicitDefaultFont) PdfCore.PdfOptions.AddRegisteredFontFamilySlot(registeredFontSlots, pdfOptions.DefaultFont);
        if (pdfOptions.HasExplicitHeaderFont) PdfCore.PdfOptions.AddRegisteredFontFamilySlot(registeredFontSlots, pdfOptions.HeaderFont);
        if (pdfOptions.HasExplicitFooterFont) PdfCore.PdfOptions.AddRegisteredFontFamilySlot(registeredFontSlots, pdfOptions.FooterFont);

        RtfFont? defaultFont = document.Settings.DefaultFontId.HasValue
            ? document.Fonts.FirstOrDefault(font => font.Id == document.Settings.DefaultFontId.Value)
            : null;
        if (defaultFont != null) {
            bool preserveConfiguredDefaultFont = pdfOptions.HasExplicitDefaultFont ||
                pdfOptions.HasEmbeddedStandardFontFamily(pdfOptions.DefaultFont);
            if (preserveConfiguredDefaultFont) {
                PdfCore.PdfStandardFont configuredDefaultSlot = PdfCore.PdfStandardFontMapper.GetFontFamily(pdfOptions.DefaultFont);
                registeredFontSlots.Add(configuredDefaultSlot);
                pdfOptions.ApplyInheritedFontFamily(configuredDefaultSlot);
                familySlots[defaultFont.Name] = configuredDefaultSlot;
                fontSlots[defaultFont.Id] = configuredDefaultSlot;
            } else if (TryRegisterDocumentFont(
                defaultFont.Name,
                pdfOptions,
                registeredFontSlots,
                ResolveSystemFontPermission(defaultFont.Name, allowSystemFontEmbedding, options,
                    systemFontFamilies, ref systemFontBudgetReported),
                options,
                out PdfCore.PdfStandardFont defaultSlot)) {
                pdfOptions.ApplyInheritedFontFamily(defaultSlot);
                familySlots[defaultFont.Name] = defaultSlot;
                fontSlots[defaultFont.Id] = defaultSlot;
            } else {
                unresolvedFamilies.Add(defaultFont.Name);
            }
        }

        foreach (RtfFont font in document.Fonts) {
            if (!referencedFontIds.Contains(font.Id)) {
                continue;
            }

            if (familySlots.TryGetValue(font.Name, out PdfCore.PdfStandardFont existingSlot)) {
                fontSlots[font.Id] = existingSlot;
                continue;
            }
            if (unresolvedFamilies.Contains(font.Name)) continue;

            if (TryRegisterDocumentFont(
                font.Name,
                pdfOptions,
                registeredFontSlots,
                ResolveSystemFontPermission(font.Name, allowSystemFontEmbedding, options,
                    systemFontFamilies, ref systemFontBudgetReported),
                options,
                out PdfCore.PdfStandardFont fontSlot)) {
                familySlots[font.Name] = fontSlot;
                fontSlots[font.Id] = fontSlot;
            } else {
                unresolvedFamilies.Add(font.Name);
            }
        }

        string requiredText = string.Join("\n", Writing.RtfDocumentWriter.EnumerateParagraphs(document).Select(paragraph => paragraph.ToPlainText()));
        PdfCore.PdfTextFallbackFeatures fallbacks = PdfCore.PdfTextDiagnostics.ResolveRequiredFallbackFeatures(options.TextFallbacks, new[] { requiredText });
        if (fallbacks != PdfCore.PdfTextFallbackFeatures.None) {
            // Preserve authored and caller font slots; installed fallback faces only
            // supply scalars missing from those faces under the host resource policy.
            foreach (PdfCore.PdfStandardFont slot in fontSlots.Values) registeredFontSlots.Add(slot);
            pdfOptions.UseTextFallbacks(fallbacks, registeredFontSlots,
                options.ResourcePolicy.AllowSystemFontEmbedding, preserveConfiguredFontSlots: true, requiredText: requiredText);
        }
        return fontSlots;
    }

    private static bool ResolveSystemFontPermission(
        string familyName,
        bool allowSystemFontEmbedding,
        RtfToPdfOptions options,
        ISet<string> systemFontFamilies,
        ref bool budgetReported) {
        if (!allowSystemFontEmbedding) return false;
        if (systemFontFamilies.Contains(familyName)) return true;
        if (systemFontFamilies.Count < options.MaximumSystemFontFamilies) {
            systemFontFamilies.Add(familyName);
            return true;
        }
        if (!budgetReported) {
            budgetReported = true;
            AddConversionWarning(
                options,
                "SystemFontResolutionLimitExceeded",
                "Font/" + familyName,
                "The RTF font-family resolution budget was exhausted; remaining families use dependency-free PDF fallbacks.",
                RtfConversionAction.Substituted,
                new Dictionary<string, string> {
                    ["maximumSystemFontFamilies"] = options.MaximumSystemFontFamilies.ToString(System.Globalization.CultureInfo.InvariantCulture)
                });
        }
        return false;
    }

    private static HashSet<int> CollectReferencedFontIds(RtfDocument document, RtfToPdfOptions options, PdfCore.PdfOptions pdfOptions) {
        options.CancellationToken.ThrowIfCancellationRequested();
        var fontIds = new HashSet<int>();
        if (document.Settings.DefaultFontId.HasValue) {
            fontIds.Add(document.Settings.DefaultFontId.Value);
        }

        if (document.Sections.Count > 0) {
            foreach (RtfSection section in document.Sections) {
                CollectReferencedFontIds(section.Blocks, fontIds);
            }
        } else {
            CollectReferencedFontIds(document.Blocks, fontIds);
        }

        foreach (RtfNote note in document.Notes) {
            foreach (RtfParagraph paragraph in note.Paragraphs) {
                CollectReferencedFontIds(paragraph, fontIds);
            }
        }

        if (options.IncludeHeaderFooters) {
            var stories = new HashSet<RtfHeaderFooter>();
            if (document.Sections.Count == 0) {
                CollectStories(document.PageSetup, document.HeaderFooters);
            } else {
                PdfCore.PdfOptions defaults = pdfOptions.Clone();
                ApplyPageSetup(document, document.PageSetup, defaults);
                PdfCore.PdfOptions current = defaults.Clone();
                RtfSection currentSection = document.Sections[0];
                ApplyPageSetup(document, document.GetPageSetup(currentSection), current);
                bool canContinueRunningStories = SupportsContinuousRunningStoryDocument(document);
                for (int index = 0; index < document.Sections.Count; index++) {
                    options.CancellationToken.ThrowIfCancellationRequested();
                    RtfSection section = document.Sections[index];
                    if (index == 0) {
                        CollectStories(document.GetPageSetup(section), document.GetEffectiveHeaderFooters(section));
                    } else if (StartsNewPdfPage(section.BreakKind)) {
                        current = defaults.Clone();
                        ApplyPageSetup(document, section, current, defaults);
                        currentSection = section;
                        CollectStories(document.GetPageSetup(section), document.GetEffectiveHeaderFooters(section));
                    } else if (canContinueRunningStories && TryPrepareContinuousRunningStoryOptions(document, section,
                        currentSection, current, defaults, options, out PdfCore.PdfOptions? requested)) {
                        current = requested!;
                        currentSection = section;
                        CollectStories(document.GetPageSetup(section), document.GetEffectiveHeaderFooters(section));
                    }
                }
            }
            foreach (RtfHeaderFooter story in stories)
                foreach (RtfParagraph paragraph in story.Paragraphs) CollectRunningFontIds(paragraph);

            void CollectStories(RtfPageSetup setup, IReadOnlyList<RtfHeaderFooter> declarations) {
                foreach (bool isFooter in new[] { false, true }) {
                    if (HasRunningStoryPresentationOverrides(options.PdfOptions, isFooter)) continue;
                    foreach (RtfHeaderFooter story in GetSelectedHeaderFooters(document, setup, declarations, isFooter)) stories.Add(story);
                }
            }
        }

        return fontIds;

        void CollectRunningFontIds(RtfParagraph paragraph) {
            foreach (RtfRun run in paragraph.Runs) {
                int? fontId = document.GetRunFormatting(paragraph, run).FontId;
                if (fontId.HasValue) fontIds.Add(fontId.Value);
            }
            // Empty/control-only paragraphs can obtain their line box from a
            // paragraph style without carrying a visible source run.
            int? paragraphFontId = document.GetRunFormatting(paragraph, new RtfRun(string.Empty) {
                UseDefaultCharacterFormatting = paragraph.UseParsedParagraphDefaults
            }).FontId;
            if (paragraphFontId.HasValue) fontIds.Add(paragraphFontId.Value);
            foreach (RtfField field in paragraph.Inlines.OfType<RtfField>()) CollectRunningFontIds(field.Result);
        }
    }

    private static void CollectReferencedFontIds(IEnumerable<IRtfBlock> blocks, HashSet<int> fontIds) {
        foreach (IRtfBlock block in blocks) {
            if (block is RtfParagraph paragraph) {
                CollectReferencedFontIds(paragraph, fontIds);
            } else if (block is RtfTable table) {
                foreach (RtfTableRow row in table.Rows) {
                    foreach (RtfTableCell cell in row.Cells) {
                        CollectReferencedFontIds(cell.Blocks, fontIds);
                    }
                }
            }
        }
    }

    private static void CollectReferencedFontIds(RtfParagraph paragraph, HashSet<int> fontIds) {
        foreach (RtfRun run in paragraph.Runs) {
            if (run.FontId.HasValue) fontIds.Add(run.FontId.Value);
        }

        foreach (IRtfInline inline in paragraph.Inlines) {
            if (inline is RtfField field) {
                CollectReferencedFontIds(field.Result, fontIds);
            }
        }
    }

    private static bool TryRegisterDocumentFont(
        string familyName,
        PdfCore.PdfOptions pdfOptions,
        HashSet<PdfCore.PdfStandardFont> registeredFontSlots,
        bool allowSystemFontEmbedding,
        RtfToPdfOptions options,
        out PdfCore.PdfStandardFont fontSlot) {
        if (PdfCore.PdfStandardFontMapper.TryMapFontFamily(familyName, out PdfCore.PdfStandardFont mappedFont)) {
            PdfCore.PdfStandardFont mappedFamily = PdfCore.PdfStandardFontMapper.GetFontFamily(mappedFont);
            if (!registeredFontSlots.Contains(mappedFamily)) {
                pdfOptions.RegisterOfficeFontFamily(familyName, mappedFamily, allowSystemFontEmbedding);
                registeredFontSlots.Add(mappedFamily);
                fontSlot = mappedFamily;
                return true;
            }

            if (pdfOptions.EmbeddedFontFamilySlotMatches(mappedFamily, familyName)) {
                fontSlot = mappedFamily;
                return true;
            }

            if (allowSystemFontEmbedding &&
                PdfCore.PdfOptions.TrySelectAvailableFontFamilySlot(familyName, registeredFontSlots, out PdfCore.PdfStandardFont distinctSlot) &&
                PdfCore.PdfEmbeddedFontFamily.TryFromSystem(familyName, out PdfCore.PdfEmbeddedFontFamily? distinctFamily) &&
                distinctFamily != null) {
                pdfOptions.RegisterFontFamily(distinctSlot, distinctFamily);
                registeredFontSlots.Add(distinctSlot);
                fontSlot = distinctSlot;
                return true;
            }

            if (pdfOptions.HasEmbeddedStandardFontFamily(mappedFamily)) {
                ReportFontSlotExhaustion(options, familyName, mappedFamily, pdfOptions.GetEmbeddedFontFamilyName(mappedFamily));
                fontSlot = PdfCore.PdfStandardFont.Helvetica;
                return false;
            }

            fontSlot = mappedFamily;
            return true;
        }

        if (allowSystemFontEmbedding) {
            if (!PdfCore.PdfOptions.TrySelectAvailableFontFamilySlot(familyName, registeredFontSlots, out PdfCore.PdfStandardFont availableSlot)) {
                ReportFontSlotExhaustion(options, familyName, null, null);
            } else if (PdfCore.PdfEmbeddedFontFamily.TryFromSystem(familyName, out PdfCore.PdfEmbeddedFontFamily? embeddedFamily) &&
                       embeddedFamily != null) {
                pdfOptions.RegisterFontFamily(availableSlot, embeddedFamily);
                registeredFontSlots.Add(availableSlot);
                fontSlot = availableSlot;
                return true;
            }
        }

        fontSlot = PdfCore.PdfStandardFont.Helvetica;
        return false;
    }

    private static void ReportFontSlotExhaustion(
        RtfToPdfOptions options,
        string familyName,
        PdfCore.PdfStandardFont? fallbackSlot,
        string? occupyingFontFamily) {
        var details = new Dictionary<string, string> {
            ["fontFamily"] = familyName
        };
        if (fallbackSlot.HasValue) details["fallbackSlot"] = fallbackSlot.Value.ToString();
        if (!string.IsNullOrWhiteSpace(occupyingFontFamily)) details["occupyingFontFamily"] = occupyingFontFamily!;
        AddConversionWarning(
            options,
            "FontFamilySlotExhausted",
            "Font/" + familyName,
            "The RTF font family could not receive a distinct embedded PDF family slot; affected runs use the configured document default instead.",
            RtfConversionAction.Substituted,
            details);
    }
}
