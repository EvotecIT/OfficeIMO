using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Rtf.Pdf;

internal static partial class RtfPdfConverter {
    private static IReadOnlyDictionary<int, PdfFontBinding> ConfigureDocumentFonts(
        RtfDocument document,
        PdfCore.PdfOptions pdfOptions,
        RtfToPdfOptions options) {
        bool allowSystemFontEmbedding =
            options.ResourcePolicy.AllowSystemFontEmbedding &&
            options.ResourcePolicy.AllowDocumentFontEmbedding;
        bool systemFontBudgetReported = false;
        var fontSlots = new Dictionary<int, PdfFontBinding>();
        var familySlots = new Dictionary<string, PdfFontBinding>(StringComparer.OrdinalIgnoreCase);
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
                familySlots[defaultFont.Name] = new PdfFontBinding(configuredDefaultSlot);
                fontSlots[defaultFont.Id] = new PdfFontBinding(configuredDefaultSlot);
            } else if (TryRegisterDocumentFont(
                defaultFont.Name,
                pdfOptions,
                registeredFontSlots,
                ResolveSystemFontPermission(defaultFont.Name, allowSystemFontEmbedding &&
                    !pdfOptions.HasNamedFontFamily(OfficeIMO.Drawing.OfficeFontFamilyParser.Quote(defaultFont.Name)), options,
                    systemFontFamilies, ref systemFontBudgetReported),
                options,
                out PdfFontBinding defaultSlot)) {
                pdfOptions.ApplyInheritedFontFamily(defaultSlot.Font);
                familySlots[defaultFont.Name] = defaultSlot;
                fontSlots[defaultFont.Id] = defaultSlot;
            } else {
                unresolvedFamilies.Add(defaultFont.Name);
                ReportUnavailableFont(defaultFont.Name);
            }
        }

        foreach (RtfFont font in document.Fonts) {
            if (!referencedFontIds.Contains(font.Id)) {
                continue;
            }

            if (familySlots.TryGetValue(font.Name, out PdfFontBinding existingSlot)) {
                fontSlots[font.Id] = existingSlot;
                continue;
            }
            if (unresolvedFamilies.Contains(font.Name)) continue;

            if (TryRegisterDocumentFont(
                font.Name,
                pdfOptions,
                registeredFontSlots,
                ResolveSystemFontPermission(font.Name, allowSystemFontEmbedding &&
                    !pdfOptions.HasNamedFontFamily(OfficeIMO.Drawing.OfficeFontFamilyParser.Quote(font.Name)), options,
                    systemFontFamilies, ref systemFontBudgetReported),
                options,
                out PdfFontBinding fontSlot)) {
                familySlots[font.Name] = fontSlot;
                fontSlots[font.Id] = fontSlot;
            } else {
                unresolvedFamilies.Add(font.Name);
                ReportUnavailableFont(font.Name);
            }
        }

        string requiredText = string.Join("\n", Writing.RtfDocumentWriter.EnumerateParagraphs(document).Select(paragraph => paragraph.ToPlainText()));
        PdfCore.PdfTextFallbackFeatures fallbacks = PdfCore.PdfTextDiagnostics.ResolveRequiredFallbackFeatures(options.TextFallbacks, new[] { requiredText });
        if (fallbacks != PdfCore.PdfTextFallbackFeatures.None) {
            // Preserve authored and caller font slots; installed fallback faces only
            // supply scalars missing from those faces under the host resource policy.
            foreach (PdfFontBinding binding in fontSlots.Values) registeredFontSlots.Add(binding.Font);
            pdfOptions.UseTextFallbacks(fallbacks, registeredFontSlots,
                options.ResourcePolicy.AllowSystemFontEmbedding, preserveConfiguredFontSlots: true, requiredText: requiredText);
        }
        return fontSlots;

        void ReportUnavailableFont(string familyName) {
            if (options.Report.Warnings.Any(warning => warning.Code == "FontFamilySlotExhausted" &&
                warning.Details.TryGetValue("fontFamily", out string? reported) &&
                string.Equals(reported, familyName, StringComparison.OrdinalIgnoreCase))) return;
            PdfCore.PdfStandardFont fallback = pdfOptions.DefaultFont;
            string? actualFamily = pdfOptions.GetEmbeddedFontFamilyName(fallback);
            if (defaultFont != null && fontSlots.TryGetValue(defaultFont.Id, out PdfFontBinding defaultBinding)) {
                fallback = defaultBinding.Font;
                actualFamily = defaultBinding.NamedFamily ?? pdfOptions.GetEmbeddedFontFamilyName(fallback);
            }
            options.Report.Add(pdfOptions.CreateFontFamilySubstitutionWarning(
                "OfficeIMO.Rtf.Pdf", "FontFamilySubstituted", "Font/" + familyName,
                familyName, fallback, actualFamily, sourceFontEmbeddingAllowed: allowSystemFontEmbedding));
        }
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
        var referencedNotes = new HashSet<RtfNote>(document.Notes);
        if (document.Settings.DefaultFontId.HasValue) {
            fontIds.Add(document.Settings.DefaultFontId.Value);
        }

        if (document.Sections.Count > 0) {
            foreach (RtfSection section in document.Sections) {
                CollectReferencedFontIds(document, section.Blocks, fontIds, referencedNotes);
            }
        } else {
            CollectReferencedFontIds(document, document.Blocks, fontIds, referencedNotes);
        }

        foreach (RtfNote note in referencedNotes) {
            foreach (RtfParagraph paragraph in note.Paragraphs) {
                CollectReferencedFontIds(document, paragraph, fontIds);
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
                foreach (RtfParagraph paragraph in story.Paragraphs) CollectReferencedFontIds(document, paragraph, fontIds);

            void CollectStories(RtfPageSetup setup, IReadOnlyList<RtfHeaderFooter> declarations) {
                foreach (bool isFooter in new[] { false, true }) {
                    if (HasRunningStoryPresentationOverrides(options.PdfOptions, isFooter)) continue;
                    foreach (RtfHeaderFooter story in GetSelectedHeaderFooters(document, setup, declarations, isFooter)) stories.Add(story);
                }
            }
        }

        return fontIds;
    }

    private static void CollectReferencedFontIds(RtfDocument document, IEnumerable<IRtfBlock> blocks, HashSet<int> fontIds,
        ISet<RtfNote>? referencedNotes = null) {
        foreach (IRtfBlock block in blocks) {
            if (block is RtfParagraph paragraph) {
                CollectReferencedFontIds(document, paragraph, fontIds, referencedNotes);
            } else if (block is RtfTable table) {
                foreach (RtfTableRow row in table.Rows) {
                    foreach (RtfTableCell cell in row.Cells) {
                        CollectReferencedFontIds(document, cell.Blocks, fontIds, referencedNotes);
                    }
                }
            }
        }
    }

    private static void CollectReferencedFontIds(RtfDocument document, RtfParagraph paragraph, HashSet<int> fontIds,
        ISet<RtfNote>? referencedNotes = null) {
        foreach (RtfRun run in paragraph.Runs) {
            int? fontId = document.GetRunFormatting(paragraph, run).FontId;
            if (fontId.HasValue) fontIds.Add(fontId.Value);
            if (run.Note != null) referencedNotes?.Add(run.Note);
        }

        // Font admission follows the same inherited character formatting as
        // paint, including the line-box strut of empty/control-only paragraphs.
        int? paragraphFontId = document.GetRunFormatting(paragraph, new RtfRun(string.Empty) {
            UseDefaultCharacterFormatting = paragraph.UseParsedParagraphDefaults
        }).FontId;
        if (paragraphFontId.HasValue) fontIds.Add(paragraphFontId.Value);
        foreach (RtfField field in paragraph.Inlines.OfType<RtfField>()) {
            CollectReferencedFontIds(document, field.Result, fontIds, referencedNotes);
        }
        // Supported authoring APIs attach notes directly to body references;
        // their bodies need not also appear in the document's note collection.
        foreach (RtfGeneratedText generated in paragraph.Inlines.OfType<RtfGeneratedText>()) {
            if (generated.Note != null) referencedNotes?.Add(generated.Note);
        }
    }

    private static bool TryRegisterDocumentFont(
        string familyName,
        PdfCore.PdfOptions pdfOptions,
        HashSet<PdfCore.PdfStandardFont> registeredFontSlots,
        bool allowSystemFontEmbedding,
        RtfToPdfOptions options,
        out PdfFontBinding binding) {
        string literalFamily = OfficeIMO.Drawing.OfficeFontFamilyParser.Quote(familyName);
        if (pdfOptions.TryResolveNamedOfficeFontFamily(literalFamily, out string? namedFamily) && namedFamily != null) {
            PdfCore.PdfStandardFont fallback = PdfCore.PdfStandardFontMapper.TryMapFontFamily(literalFamily, out PdfCore.PdfStandardFont mapped)
                ? PdfCore.PdfStandardFontMapper.GetFontFamily(mapped) : pdfOptions.DefaultFont;
            binding = new PdfFontBinding(fallback, namedFamily);
            if (pdfOptions.TryResolveFontFamilySubstitution(literalFamily, out _)) {
                string diagnosticFamily = pdfOptions.TryResolveFontFamilySubstitution(familyName, out _) ? familyName : literalFamily;
                PdfCore.PdfConversionWarning warning = pdfOptions.CreateFontFamilySubstitutionWarning(
                    "OfficeIMO.Rtf.Pdf", "FontFamilySubstituted", "Font/" + familyName,
                    diagnosticFamily, fallback, namedFamily);
                if (diagnosticFamily != familyName) {
                    var details = warning.Details.ToDictionary(detail => detail.Key, detail => detail.Value);
                    details["fontFamily"] = familyName;
                    warning = new PdfCore.PdfConversionWarning(warning.Converter, warning.Code, warning.Source,
                        warning.Message, warning.Severity, warning.LossKind, warning.LayoutDiagnostic, details);
                }
                options.Report.Add(warning);
            }
            return true;
        }
        bool registered = TryRegisterDocumentFontSlot(familyName, pdfOptions, registeredFontSlots,
            allowSystemFontEmbedding, options, out PdfCore.PdfStandardFont font);
        binding = new PdfFontBinding(font);
        return registered;
    }

    private static bool TryRegisterDocumentFontSlot(
        string familyName,
        PdfCore.PdfOptions pdfOptions,
        HashSet<PdfCore.PdfStandardFont> registeredFontSlots,
        bool allowSystemFontEmbedding,
        RtfToPdfOptions options,
        out PdfCore.PdfStandardFont fontSlot) {
        if (pdfOptions.TryResolveEmbeddedFontFamilySlot(familyName, out fontSlot)) {
            registeredFontSlots.Add(fontSlot);
            return true;
        }
        if (PdfCore.PdfStandardFontMapper.TryMapFontFamily(
                OfficeIMO.Drawing.OfficeFontFamilyParser.Quote(familyName), out PdfCore.PdfStandardFont mappedFont)) {
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
