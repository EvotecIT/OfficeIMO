namespace OfficeIMO.Markdown;

using System;

/// <summary>
/// Explicit transcript-preparation helpers for hosts ingesting or exporting IntelligenceX-style markdown.
/// </summary>
/// <example>
/// <code>
/// var prepared = MarkdownTranscriptPreparation.PrepareIntelligenceXTranscriptBody(markdown);
/// var document = MarkdownTranscriptPreparation.PrepareIntelligenceXTranscriptDocument(markdown);
/// var exportReady = MarkdownTranscriptPreparation.PrepareIntelligenceXTranscriptForExport(markdown);
/// var docxReady = MarkdownTranscriptPreparation.PrepareIntelligenceXTranscriptForDocx(
///     markdown,
///     preservesGroupedDefinitionLikeParagraphs: true);
/// </code>
/// Use these helpers when the host explicitly wants the IntelligenceX transcript contract.
/// Generic markdown ingestion should stay on <see cref="MarkdownReader"/> profiles and document transforms instead.
/// </example>
public static class MarkdownTranscriptPreparation {
    /// <summary>
    /// Applies the explicit IX transcript reader contract to an existing reader configuration in place.
    /// This preserves the caller's security/profile toggles while layering in transcript normalization
    /// and optional AST transforms.
    /// </summary>
    public static void ApplyIntelligenceXTranscriptReaderContract(
        MarkdownReaderOptions options,
        bool preservesGroupedDefinitionLikeParagraphs = true,
        MarkdownVisualFenceLanguageMode? visualFenceLanguageMode = null) {
        if (options == null) {
            throw new ArgumentNullException(nameof(options));
        }

        options.InputNormalization = MarkdownInputNormalizationPresets.CreateIntelligenceXTranscript();
        options.PreferNarrativeSingleLineDefinitions = true;

        if (!preservesGroupedDefinitionLikeParagraphs && options.DefinitionLists) {
            bool hasDefinitionCompatibilityTransform = false;
            for (var i = 0; i < options.DocumentTransforms.Count; i++) {
                if (options.DocumentTransforms[i] is MarkdownSimpleDefinitionListParagraphTransform) {
                    hasDefinitionCompatibilityTransform = true;
                    break;
                }
            }

            if (!hasDefinitionCompatibilityTransform) {
                options.DocumentTransforms.Add(new MarkdownSimpleDefinitionListParagraphTransform());
            }
        }

        if (visualFenceLanguageMode.HasValue) {
            bool hasVisualUpgradeTransform = false;
            for (var i = 0; i < options.DocumentTransforms.Count; i++) {
                if (options.DocumentTransforms[i] is MarkdownJsonVisualCodeBlockTransform existing
                    && existing.LanguageMode == visualFenceLanguageMode.Value) {
                    hasVisualUpgradeTransform = true;
                    break;
                }
            }

            if (!hasVisualUpgradeTransform) {
                options.DocumentTransforms.Add(new MarkdownJsonVisualCodeBlockTransform(visualFenceLanguageMode.Value));
            }
        }
    }

    /// <summary>
    /// Creates reader options for the explicit IX transcript ingestion contract.
    /// </summary>
    /// <param name="readerProfile">Reader profile to compose on top of.</param>
    /// <param name="preservesGroupedDefinitionLikeParagraphs">
    /// Whether grouped simple <c>Label: value</c> lines should remain parsed as definition lists.
    /// When <c>false</c>, a document transform expands simple grouped entries back into paragraphs.
    /// </param>
    /// <param name="visualFenceLanguageMode">
    /// Optional legacy JSON visual-fence upgrade mode to compose into the transcript reader contract.
    /// When supplied, plain JSON code blocks can be upgraded into semantic visual fenced blocks during document transforms.
    /// </param>
    /// <returns>Configured reader options for transcript parsing.</returns>
    public static MarkdownReaderOptions CreateIntelligenceXTranscriptReaderOptions(
        MarkdownReaderOptions.MarkdownDialectProfile readerProfile = MarkdownReaderOptions.MarkdownDialectProfile.OfficeIMO,
        bool preservesGroupedDefinitionLikeParagraphs = true,
        MarkdownVisualFenceLanguageMode? visualFenceLanguageMode = null) {
        var options = MarkdownReaderOptions.CreateProfile(readerProfile);
        ApplyIntelligenceXTranscriptReaderContract(
            options,
            preservesGroupedDefinitionLikeParagraphs,
            visualFenceLanguageMode);
        return options;
    }

    /// <summary>
    /// Applies the shared IX transcript body preparation contract.
    /// </summary>
    /// <param name="markdown">Transcript markdown source.</param>
    /// <returns>Normalized transcript body markdown.</returns>
    public static string PrepareIntelligenceXTranscriptBody(string? markdown) {
        var value = markdown ?? string.Empty;
        if (value.Length == 0) {
            return string.Empty;
        }

        var normalized = MarkdownInputNormalizer.Normalize(
            value,
            MarkdownInputNormalizationPresets.CreateIntelligenceXTranscript());

        var prepared = string.IsNullOrEmpty(normalized) ? value : normalized;
        return MarkdownOrderedLists.SeparateAdjacentOrderedListItems(prepared);
    }

    /// <summary>
    /// Applies the shared IX transcript export-preparation contract after any host-specific transport markers
    /// have already been stripped.
    /// </summary>
    /// <param name="markdown">Transcript markdown source after host-specific marker cleanup.</param>
    /// <returns>Export-prepared transcript markdown.</returns>
    public static string PrepareIntelligenceXTranscriptForExport(string? markdown) {
        return MarkdownBlankLines.CollapseDuplicateBlankLines(
            PrepareIntelligenceXTranscriptBody(markdown));
    }

    /// <summary>
    /// Parses transcript markdown into an AST using the explicit IX transcript contract.
    /// </summary>
    /// <param name="markdown">Transcript markdown source.</param>
    /// <param name="readerProfile">Reader profile to compose on top of.</param>
    /// <param name="visualFenceLanguageMode">
    /// Optional legacy JSON visual-fence upgrade mode to apply while preparing the transcript document.
    /// </param>
    /// <returns>Prepared transcript document.</returns>
    public static MarkdownDoc PrepareIntelligenceXTranscriptDocument(
        string? markdown,
        MarkdownReaderOptions.MarkdownDialectProfile readerProfile = MarkdownReaderOptions.MarkdownDialectProfile.OfficeIMO,
        MarkdownVisualFenceLanguageMode? visualFenceLanguageMode = null) {
        return PrepareIntelligenceXTranscriptDocumentForDocx(
            markdown,
            preservesGroupedDefinitionLikeParagraphs: true,
            readerProfile: readerProfile,
            visualFenceLanguageMode: visualFenceLanguageMode);
    }

    /// <summary>
    /// Applies the shared IX transcript DOCX-preparation contract.
    /// </summary>
    /// <param name="markdown">Transcript markdown source.</param>
    /// <param name="preservesGroupedDefinitionLikeParagraphs">Whether the host renderer preserves grouped definition-like lines without compatibility repair.</param>
    /// <returns>DOCX-prepared transcript markdown.</returns>
    public static string PrepareIntelligenceXTranscriptForDocx(
        string? markdown,
        bool preservesGroupedDefinitionLikeParagraphs) {
        var prepared = PrepareIntelligenceXTranscriptBody(markdown);
        if (preservesGroupedDefinitionLikeParagraphs) {
            return prepared;
        }

        return MarkdownDefinitionLines.SeparateAdjacentDefinitionLikeLinesOutsideFencedCodeBlocks(prepared);
    }

    /// <summary>
    /// Parses transcript markdown for DOCX/export-oriented hosts using the explicit IX transcript contract.
    /// </summary>
    /// <param name="markdown">Transcript markdown source.</param>
    /// <param name="preservesGroupedDefinitionLikeParagraphs">Whether grouped simple definition-like lines should remain definition lists.</param>
    /// <param name="readerProfile">Reader profile to compose on top of.</param>
    /// <param name="visualFenceLanguageMode">
    /// Optional legacy JSON visual-fence upgrade mode to apply while preparing the transcript document.
    /// </param>
    /// <returns>Prepared transcript document.</returns>
    public static MarkdownDoc PrepareIntelligenceXTranscriptDocumentForDocx(
        string? markdown,
        bool preservesGroupedDefinitionLikeParagraphs,
        MarkdownReaderOptions.MarkdownDialectProfile readerProfile = MarkdownReaderOptions.MarkdownDialectProfile.OfficeIMO,
        MarkdownVisualFenceLanguageMode? visualFenceLanguageMode = null) {
        var value = markdown ?? string.Empty;
        if (value.Length == 0) {
            return MarkdownDoc.Create();
        }

        var options = CreateIntelligenceXTranscriptReaderOptions(
            readerProfile,
            preservesGroupedDefinitionLikeParagraphs,
            visualFenceLanguageMode);
        return MarkdownReader.Parse(value, options);
    }
}
