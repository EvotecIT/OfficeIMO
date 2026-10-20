namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private bool _fontLayoutRebuilt;

    /// <summary>
    /// Selects the approved default face from laid-out flow before page admission
    /// and print fitting. Rebuild once with the frozen face, reusing the resource
    /// session, diagnostics and cumulative operation budget. The new engine owns
    /// its own pagination convergence and measurement caches.
    /// </summary>
    private bool TryRebuildWithPreparedFontMetrics(IReadOnlyList<HtmlRenderFlowBlock> blocks,
        out HtmlRenderDocument rendered) {
        rendered = null!;
        if (_options is not Pdf.HtmlToPdfOptions pdfOptions) return false;
        // Preserve the converter's policy gate: explicit and portable font
        // policies do not need page-margin discovery at all.
        var margins = new Lazy<MarginFontPreparation>(() => PrepareMarginFontRequirements(blocks));
        if (!Pdf.HtmlPdfRenderedConverter.TryPrepareLibraryFontMetrics(
                blocks.SelectMany(block => block.Visuals).Concat(EnumerateAdditionalFontVisuals(margins)),
                EnumerateMarginFontRequirements(margins), _fonts, pdfOptions, _cancellationToken)) return false;
        CheckCancellation();
        _fontLayoutRebuilt = true;
        rendered = new HtmlRenderLayoutEngine(_document, _computedStyles, _options,
            _diagnostics, _resources, _pageRules, _fonts, _limits,
            _initialLogicalTextOrder, _initialSemanticNodeId, _operationBudget, _cancellationToken, _fontUsage).Render();
        return true;
    }

    /// <summary>
    /// Includes referenced repeated page content before pagination. First, left and
    /// later right pages cover the supported page selectors; counters need only
    /// representative values for font coverage. Actual margin layout still owns
    /// page counts, running strings, clipping and diagnostics.
    /// </summary>
    private MarginFontPreparation PrepareMarginFontRequirements(
        IReadOnlyList<HtmlRenderFlowBlock> blocks) {
        var prepared = new MarginFontPreparation();
        if (_options.Mode != HtmlRenderMode.Paged) return prepared;
        IReadOnlyList<HtmlRenderFlowBlock> positioned = PreparePositionedFontBlocks(blocks, prepared);
        var stringFonts = new Dictionary<string, HashSet<OfficeFontInfo>>(StringComparer.Ordinal);
        var elementKeys = new HashSet<string>(StringComparer.Ordinal);
        var textRequirements = new HashSet<(string Text, OfficeFontInfo Font)>();

        void AddText(string text, OfficeFontInfo font) {
            if (text.Length > 0 && textRequirements.Add((text, font))) prepared.Text.Add((text, font));
        }

        foreach (string? pageName in blocks.Select(block => block.PageName)
                     .Concat(blocks.SelectMany(block => block.ForcedBreaks.Select(item => item.PageName)))
                     .Concat(new string?[] { null }).Distinct(StringComparer.Ordinal)) {
            for (int pageNumber = 1; pageNumber <= 3; pageNumber++) {
                CheckCancellation();
                ChargeLayoutOperations(2L * _pageRules.RuleCount, "page-margin font rule matching");
                HtmlCssPageGeometry geometry = _pageRules.ResolveGeometry(pageNumber, pageName, _options);
                foreach (HtmlCssPageMarginTemplate box in _pageRules.ResolveMarginBoxes(
                             pageNumber, pageName, geometry, _options).Values) {
                    CheckCancellation();
                    if (box.Content.TryGetRunningElement(out string elementName, out _)) {
                        elementKeys.Add(HtmlCssRunningElementKeys.ForName(elementName));
                        continue;
                    }
                    // Mixed element() content is unsupported by margin layout.
                    if (box.Content.ContainsRunningElement || box.Font.Size <= 0D) continue;
                    ChargeLayoutOperations(box.Content.GetRenderedLength(pageNumber, 3, null), "page-margin font requirements");
                    AddText(box.Content.Render(pageNumber, 3, null), box.Font);
                    foreach (string name in box.Content.ReferencedRunningStringNames) {
                        CheckCancellation();
                        ChargeLayoutOperations(1L, "page-margin font references");
                        if (!stringFonts.TryGetValue(name, out HashSet<OfficeFontInfo>? fonts)) {
                            stringFonts.Add(name, fonts = new HashSet<OfficeFontInfo>());
                        }
                        fonts.Add(box.Font);
                    }
                }
            }
        }

        if (stringFonts.Count == 0 && elementKeys.Count == 0) return prepared;
        var snapshots = new HashSet<int>();
        foreach (HtmlRenderFlowBlock block in blocks.Concat(positioned)) {
            CheckCancellation();
            foreach (HtmlCssRunningStringAssignment assignment in block.RunningStringAssignments) {
                CheckCancellation();
                ChargeLayoutOperations(1L, "page-margin font assignments");
                if (stringFonts.TryGetValue(assignment.Name, out HashSet<OfficeFontInfo>? fonts)) {
                    foreach (OfficeFontInfo font in fonts) {
                        ChargeLayoutOperations(assignment.Value.Length, "page-margin font requirements");
                        AddText(assignment.Value, font);
                    }
                } else if (elementKeys.Contains(assignment.Name)
                    && HtmlCssRunningElementParser.TryParseSnapshotId(assignment.Value, out int snapshotId)
                    && snapshots.Add(snapshotId)
                    && _runningElementSnapshots.TryGetValue(snapshotId, out HtmlCssRunningElementSnapshot? snapshot)) {
                    // Snapshot IDs are bookkeeping, never font-coverage text.
                    // The converter walks the captured visual's actual text and
                    // fonts, including styled descendants and drawing content.
                    ChargeLayoutOperations(snapshot.Block.Visuals.Count, "page-margin font snapshots");
                    prepared.AdditionalVisuals.AddRange(snapshot.Block.Visuals);
                }
            }
        }
        return prepared;
    }

    /// <summary>
    /// Materializes global requests through the same owner and first-page geometry
    /// used by pagination. Its existing cache also retains captures and assignments;
    /// pagination reuses these blocks without changing their paint order.
    /// </summary>
    private IReadOnlyList<HtmlRenderFlowBlock> PreparePositionedFontBlocks(
        IReadOnlyList<HtmlRenderFlowBlock> blocks, MarginFontPreparation prepared) {
        if (_rootPositionedElements.Count == 0 && _fixedPositionedElements.Count == 0) {
            return Array.Empty<HtmlRenderFlowBlock>();
        }
        CheckCancellation();
        ChargeLayoutOperations(_pageRules.RuleCount, "page-margin font positioned rule matching");
        HtmlCssPageGeometry geometry = _pageRules.ResolveGeometry(1, blocks.Count > 0 ? blocks[0].PageName : null, _options);
        SetActivePageGeometry(geometry);
        double contentHeight = Math.Max(1D, ResolvePageBodyContentHeight(1, geometry));
        ValidateSurface(geometry.Width, geometry.Height);
        // Request layout retains assignments on blocks. Running-string page
        // contexts are consumed only during actual pagination and page commit.
        PrepareGlobalPositionedRequests(true, geometry.Width, geometry.Height, geometry.ContentWidth, contentHeight);
        var positioned = new List<HtmlRenderFlowBlock>();
        foreach (PositionedRequestPlacement placement in CreateGlobalPositionedPlacements(
                     true, geometry.Width, geometry.Height, geometry.ContentWidth, contentHeight)) {
            CheckCancellation();
            HtmlRenderFlowBlock block = placement.Request.Resolve(this, placement.Width, placement.Height).Block;
            ChargeLayoutOperations(1L + block.Visuals.Count, "page-margin font positioned blocks");
            positioned.Add(block);
            prepared.AdditionalVisuals.AddRange(block.Visuals);
        }
        return positioned;
    }

    private IEnumerable<HtmlRenderVisual> EnumerateAdditionalFontVisuals(Lazy<MarginFontPreparation> margins) {
        foreach (HtmlRenderVisual visual in margins.Value.AdditionalVisuals) {
            CheckCancellation();
            yield return visual;
        }
    }

    private static IEnumerable<(string Text, OfficeFontInfo Font)> EnumerateMarginFontRequirements(Lazy<MarginFontPreparation> margins) {
        foreach ((string text, OfficeFontInfo font) in margins.Value.Text) yield return (text, font);
    }

    private sealed class MarginFontPreparation {
        internal List<(string Text, OfficeFontInfo Font)> Text { get; } = new List<(string Text, OfficeFontInfo Font)>();
        internal List<HtmlRenderVisual> AdditionalVisuals { get; } = new List<HtmlRenderVisual>();
    }
}
