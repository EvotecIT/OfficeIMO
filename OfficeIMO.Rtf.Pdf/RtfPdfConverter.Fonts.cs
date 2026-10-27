using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Rtf.Pdf;

internal static partial class RtfPdfConverter {
    private static IReadOnlyDictionary<int, PdfFontBinding> ConfigureDocumentFonts(
        RtfDocument document,
        PdfCore.PdfCreateOptions pdfOptions,
        RtfToPdfOptions options) {
        bool allowSystemFontEmbedding =
            options.ResourcePolicy.AllowSystemFontEmbedding &&
            options.ResourcePolicy.AllowDocumentFontEmbedding;
        bool systemFontBudgetReported = false;
        var fontSlots = new Dictionary<int, PdfFontBinding>();
        var familySlots = new Dictionary<string, PdfFontBinding>(StringComparer.OrdinalIgnoreCase);
        var unresolvedFamilies = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var systemFontFamilies = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        HashSet<int> referencedFontIds = CollectReferencedFontIds(document, options, pdfOptions, out string requiredText);
        HashSet<PdfCore.PdfStandardFont> registeredFontSlots = pdfOptions.CreateRegisteredFontFamilySlots(includeDocumentFontSlots: false);
        if (pdfOptions.HasExplicitDefaultFont) PdfCore.PdfCreateOptions.AddRegisteredFontFamilySlot(registeredFontSlots, pdfOptions.DefaultFont);
        if (pdfOptions.HasExplicitHeaderFont) PdfCore.PdfCreateOptions.AddRegisteredFontFamilySlot(registeredFontSlots, pdfOptions.HeaderFont);
        if (pdfOptions.HasExplicitFooterFont) PdfCore.PdfCreateOptions.AddRegisteredFontFamilySlot(registeredFontSlots, pdfOptions.FooterFont);

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
                familyName, fallback, actualFamily, sourceFontEmbeddingAllowed: allowSystemFontEmbedding,
                sourceFontFamilyRequest: OfficeIMO.Drawing.OfficeFontFamilyParser.Quote(familyName)));
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

    /// <summary>Admits fonts and fallback scalars only from the existing PDF content projection.</summary>
    private static HashSet<int> CollectReferencedFontIds(RtfDocument document, RtfToPdfOptions options,
        PdfCore.PdfCreateOptions pdfOptions, out string requiredText) {
        options.CancellationToken.ThrowIfCancellationRequested();
        var fontIds = new HashSet<int>();
        var text = new List<string>();
        var numbering = new RtfListNumbering(document);
        // Detached notes and references omitted by the body projection do not
        // consume resources needed by visible text. Markers remain body text.
        var referencedNotes = options.IncludeNotes ? new HashSet<RtfNote>() : null;
        if (document.Settings.DefaultFontId.HasValue) {
            fontIds.Add(document.Settings.DefaultFontId.Value);
        }

        if (document.Sections.Count > 0) {
            foreach (RtfSection section in document.Sections) {
                CollectReferencedFontIds(document, section.Blocks, options, fontIds, text, numbering, referencedNotes);
            }
        } else {
            CollectReferencedFontIds(document, document.Blocks, options, fontIds, text, numbering, referencedNotes);
        }

        if (referencedNotes != null) {
            foreach (RtfNote note in referencedNotes) {
                if (note.Kind == RtfNoteKind.Annotation && note.Author != null) text.Add(note.Author);
                foreach (RtfParagraph paragraph in note.Paragraphs) {
                    // Note bodies do not append nested note references.
                    CollectReferencedFontIds(document, paragraph, options, fontIds, text, numbering);
                }
            }
        }

        if (options.IncludeHeaderFooters) {
            var stories = new HashSet<RtfHeaderFooter>();
            if (document.Sections.Count == 0) {
                CollectStories(document.PageSetup, document.HeaderFooters);
            } else {
                PdfCore.PdfCreateOptions defaults = pdfOptions.Clone();
                ApplyPageSetup(document, document.PageSetup, defaults);
                PdfCore.PdfCreateOptions current = defaults.Clone();
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
                        currentSection, current, defaults, options, out PdfCore.PdfCreateOptions? requested)) {
                        current = requested!;
                        currentSection = section;
                        CollectStories(document.GetPageSetup(section), document.GetEffectiveHeaderFooters(section));
                    }
                }
            }
            foreach (RtfHeaderFooter story in stories) {
                var storyNumbering = new RtfListNumbering(document);
                foreach (RtfParagraph paragraph in story.Paragraphs)
                    CollectReferencedFontIds(document, paragraph, options, fontIds, text, storyNumbering,
                        context: PdfInlineProjectionContext.RunningStory);
            }

            void CollectStories(RtfPageSetup setup, IReadOnlyList<RtfHeaderFooter> declarations) {
                foreach (bool isFooter in new[] { false, true }) {
                    bool plain = HasRunningStoryPresentationOverrides(options.PdfOptions, isFooter);
                    foreach (RtfHeaderFooter story in GetSelectedHeaderFooters(document, setup, declarations, isFooter)) {
                        if (plain) text.Add(NormalizeHeaderFooterText(story.ToPlainText()));
                        else stories.Add(story);
                    }
                }
            }
        }

        requiredText = string.Join("\n", text);
        return fontIds;
    }

    private static void CollectReferencedFontIds(RtfDocument document, IEnumerable<IRtfBlock> blocks,
        RtfToPdfOptions options, HashSet<int> fontIds, List<string> text,
        RtfListNumbering numbering, ISet<RtfNote>? referencedNotes = null, bool inTableCell = false) {
        foreach (IRtfBlock block in blocks) {
            options.CancellationToken.ThrowIfCancellationRequested();
            if (block is RtfParagraph paragraph) {
                CollectReferencedFontIds(document, paragraph, options, fontIds, text, numbering, referencedNotes,
                    splitBodyParagraph: !inTableCell);
            } else if (block is RtfTable table && options.IncludeTables) {
                if (inTableCell) {
                    // Nested tables paint plain text in the default family;
                    // source run fonts and note bodies are not projected.
                    CollectFlattenedTableText(table, text);
                    continue;
                }
                foreach (RtfTableRow row in table.Rows) {
                    foreach (RtfTableCell cell in row.Cells) {
                        if (cell.HorizontalMerge == RtfTableCellMerge.Continue ||
                            cell.VerticalMerge == RtfTableCellMerge.Continue) continue;
                        CollectReferencedFontIds(document, cell.Blocks, options, fontIds, text, numbering, referencedNotes,
                            inTableCell: true);
                    }
                }
            } else if (!inTableCell && block is RtfObject rtfObject) {
                text.Add(rtfObject.ToPlainText());
            } else if (!inTableCell && block is RtfShape shape) {
                text.Add(shape.ToPlainText());
            }
        }
    }

    private static void CollectFlattenedTableText(RtfTable table, List<string> text) {
        foreach (RtfTableRow row in table.Rows)
            foreach (RtfTableCell cell in row.Cells)
                foreach (IRtfBlock block in cell.Blocks) {
                    if (block is RtfParagraph paragraph) text.Add(paragraph.ToPlainText());
                    else if (block is RtfTable nested) CollectFlattenedTableText(nested, text);
                }
    }

    private static void CollectReferencedFontIds(RtfDocument document, RtfParagraph paragraph,
        RtfToPdfOptions options, HashSet<int> fontIds, List<string> text, RtfListNumbering numbering,
        ISet<RtfNote>? referencedNotes = null, bool splitBodyParagraph = false,
        bool appendStrut = true, List<PdfCore.PdfTextRun>? parentRuns = null,
        PdfInlineProjectionContext context = PdfInlineProjectionContext.Text) {
        options.CancellationToken.ThrowIfCancellationRequested();
        var runs = parentRuns ?? new List<PdfCore.PdfTextRun>();
        bool emitted = false;
        RtfListMarker? marker = numbering.Next(paragraph);
        if (marker != null) AddPlain(marker.Text + marker.Separator);
        for (int inlineIndex = 0; inlineIndex < paragraph.Inlines.Count; inlineIndex++) {
            IRtfInline inline = paragraph.Inlines[inlineIndex];
            switch (inline) {
                case RtfRun source:
                    RtfRun run = document.Styles.Count > 0 || source.UseDefaultCharacterFormatting
                        ? document.GetRunFormatting(paragraph, source) : source;
                    if (run.Hidden && !options.IncludeHiddenText) continue;
                    if (!string.IsNullOrEmpty(run.Text)) {
                        if (run.FontId.HasValue) fontIds.Add(run.FontId.Value);
                        text.Add(run.Text!);
                        runs.Add(new PdfCore.PdfTextRun(run.Text!, fontSize: run.FontSize));
                    }
                    if (run.Note != null) referencedNotes?.Add(run.Note);
                    break;
                case RtfField field:
                    // Field results append into the parent's runs; they have
                    // no independent paragraph line box, even when empty.
                    CollectReferencedFontIds(document, field.Result, options, fontIds, text, numbering,
                        referencedNotes, appendStrut: false, parentRuns: runs, context: context);
                    break;
                case RtfGeneratedText generated:
                    AddPlain(generated.ToPlainText());
                    if (generated.Note != null) referencedNotes?.Add(generated.Note);
                    break;
                case RtfObject rtfObject:
                    AddPlain(rtfObject.ToPlainText());
                    break;
                case RtfShape shape:
                    AddPlain(shape.ToPlainText());
                    break;
                case RtfImage image when context == PdfInlineProjectionContext.RunningStory:
                    PdfCore.PdfTextRun? prepared = PrepareRunningInlineImage(image, options, paragraph, inlineIndex);
                    if (prepared != null) runs.Add(prepared);
                    break;
                case RtfBreak rtfBreak when splitBodyParagraph &&
                    rtfBreak.Kind is RtfBreakKind.Page or RtfBreakKind.SoftPage or RtfBreakKind.Column:
                case RtfImage when splitBodyParagraph:
                case RtfBookmarkMarker { Kind: RtfBookmarkMarkerKind.Start } when splitBodyParagraph:
                    AdmitInheritedStrut();
                    runs.Clear();
                    emitted = true;
                    break;
                case RtfBreak:
                    runs.Add(PdfCore.PdfTextRun.LineBreak());
                    break;
            }
        }

        if (appendStrut && (runs.Count > 0 || !emitted)) AdmitInheritedStrut();

        void AddPlain(string value) {
            if (value.Length == 0) return;
            text.Add(value);
            runs.Add(new PdfCore.PdfTextRun(value));
        }

        void AdmitInheritedStrut() {
            if (!TryGetParagraphLineBoxControlStyle(runs, 0, out PdfCore.PdfTextRun? controlStyle) ||
                controlStyle != null) return;
            int? paragraphFontId = document.GetRunFormatting(paragraph, new RtfRun(string.Empty) {
                UseDefaultCharacterFormatting = paragraph.UseParsedParagraphDefaults
            }).FontId;
            if (paragraphFontId.HasValue) fontIds.Add(paragraphFontId.Value);
        }
    }

    private static bool TryRegisterDocumentFont(
        string familyName,
        PdfCore.PdfCreateOptions pdfOptions,
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
                options.Report.Add(pdfOptions.CreateFontFamilySubstitutionWarning(
                    "OfficeIMO.Rtf.Pdf", "FontFamilySubstituted", "Font/" + familyName,
                    familyName, fallback, namedFamily, sourceFontFamilyRequest: literalFamily));
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
        PdfCore.PdfCreateOptions pdfOptions,
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
                PdfCore.PdfCreateOptions.TrySelectAvailableFontFamilySlot(familyName, registeredFontSlots, out PdfCore.PdfStandardFont distinctSlot) &&
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
            if (!PdfCore.PdfCreateOptions.TrySelectAvailableFontFamilySlot(familyName, registeredFontSlots, out PdfCore.PdfStandardFont availableSlot)) {
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
