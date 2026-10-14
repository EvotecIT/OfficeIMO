using System;
using System.Collections.Generic;
using OfficeIMO.Word;

namespace OfficeIMO.Word.Legacy;

/// <summary>Legacy word-processing families recognized by the managed importer.</summary>
public enum LegacyWordFormat {
    /// <summary>Corel/Novell WordPerfect documents.</summary>
    WordPerfect,
    /// <summary>MicroPro WordStar documents.</summary>
    WordStar,
    /// <summary>Lotus Ami Pro documents.</summary>
    AmiPro,
    /// <summary>Lotus Word Pro documents.</summary>
    LotusWordPro,
    /// <summary>Microsoft Works word-processing documents.</summary>
    MicrosoftWorks,
    /// <summary>Microsoft Windows Write documents.</summary>
    MicrosoftWrite,
    /// <summary>Selected Microsoft Word for DOS documents.</summary>
    WordForDos
}

/// <summary>Classifies a recovered note-like source object.</summary>
public enum LegacyWordNoteKind {
    /// <summary>A footnote.</summary>
    Footnote,
    /// <summary>An endnote.</summary>
    Endnote,
    /// <summary>An annotation.</summary>
    Annotation,
    /// <summary>A source comment.</summary>
    Comment
}

/// <summary>Describes recovered character formatting without exposing parser internals.</summary>
public sealed class LegacyWordRunContent {
    internal LegacyWordRunContent(LegacyWordRun source) {
        Text = source.Text;
        Bold = source.Bold;
        Italic = source.Italic;
        Strike = source.Strike;
        Underline = source.Underline;
        VerticalPosition = source.VerticalPosition;
        FontSizePoints = source.FontSizePoints;
        FontFamily = source.FontFamily;
        ColorHex = source.ColorHex;
        NoteIndex = source.NoteIndex;
        Image = source.Image == null ? null : new LegacyWordImageContent(source.Image);
    }
    /// <summary>Gets recovered text.</summary>
    public string Text { get; }
    /// <summary>Gets whether bold was recovered.</summary>
    public bool Bold { get; }
    /// <summary>Gets whether italic was recovered.</summary>
    public bool Italic { get; }
    /// <summary>Gets whether strike-through was recovered.</summary>
    public bool Strike { get; }
    /// <summary>Gets the recovered underline style.</summary>
    public WordUnderlineStyle? Underline { get; }
    /// <summary>Gets the recovered vertical text position.</summary>
    public WordVerticalTextPosition? VerticalPosition { get; }
    /// <summary>Gets the recovered font size in points.</summary>
    public double? FontSizePoints { get; }
    /// <summary>Gets the recovered font family.</summary>
    public string? FontFamily { get; }
    /// <summary>Gets the recovered RGB color.</summary>
    public string? ColorHex { get; }
    /// <summary>Gets an anchored note's zero-based index in <see cref="LegacyWordContent.Notes"/>, when this run represents a note reference.</summary>
    public int? NoteIndex { get; }
    /// <summary>Gets a recovered inline image where the profile qualifies graphic recovery.</summary>
    public LegacyWordImageContent? Image { get; }
}

/// <summary>A recovered graphic with its original WPG1 bytes and a bounded PNG projection.</summary>
public sealed class LegacyWordImageContent {
    private readonly byte[] _source, _png;
    internal LegacyWordImageContent(LegacyWordImage source) {
        _source = source.SourceBytes; _png = source.PngBytes;
        WidthPoints = source.WidthPoints; HeightPoints = source.HeightPoints;
    }
    /// <summary>Gets the recovered graphic canvas width in points, before unsupported box overrides.</summary>
    public double WidthPoints { get; }
    /// <summary>Gets the recovered graphic canvas height in points, before unsupported box overrides.</summary>
    public double HeightPoints { get; }
    /// <summary>Returns an independent copy of the original WPG1 source data.</summary>
    public byte[] GetSourceBytes() => (byte[])_source.Clone();
    /// <summary>Returns an independent copy of the recovered PNG image.</summary>
    public byte[] GetPngBytes() => (byte[])_png.Clone();
}

/// <summary>Describes one recovered source paragraph-style definition.</summary>
public sealed class LegacyWordStyleContent {
    internal LegacyWordStyleContent(LegacyWordStyle source) {
        Name = source.Name;
        Bold = source.Bold;
        Italic = source.Italic;
        Underline = source.Underline;
        FontSizePoints = source.FontSizePoints;
        FontFamily = source.FontFamily;
        ColorHex = source.ColorHex;
        Alignment = source.Alignment;
        LineSpacingPoints = source.LineSpacingPoints;
        SpacingBeforePoints = source.SpacingBeforePoints;
        SpacingAfterPoints = source.SpacingAfterPoints;
        PageBreakBefore = source.PageBreakBefore;
        KeepWithNext = source.KeepWithNext;
        KeepLinesTogether = source.KeepLinesTogether;
    }
    /// <summary>Gets the recovered source style name.</summary>
    public string Name { get; }
    /// <summary>Gets whether bold is part of the style.</summary>
    public bool Bold { get; }
    /// <summary>Gets whether italic is part of the style.</summary>
    public bool Italic { get; }
    /// <summary>Gets the recovered underline style.</summary>
    public WordUnderlineStyle? Underline { get; }
    /// <summary>Gets the recovered font size in points.</summary>
    public double? FontSizePoints { get; }
    /// <summary>Gets the recovered font family.</summary>
    public string? FontFamily { get; }
    /// <summary>Gets the recovered RGB color.</summary>
    public string? ColorHex { get; }
    /// <summary>Gets recovered paragraph alignment.</summary>
    public WordParagraphAlignment? Alignment { get; }
    /// <summary>Gets recovered line spacing in points.</summary>
    public double? LineSpacingPoints { get; }
    /// <summary>Gets recovered spacing before the paragraph in points.</summary>
    public double? SpacingBeforePoints { get; }
    /// <summary>Gets recovered spacing after the paragraph in points.</summary>
    public double? SpacingAfterPoints { get; }
    /// <summary>Gets whether the style requests a page break before the paragraph.</summary>
    public bool PageBreakBefore { get; }
    /// <summary>Gets whether the style keeps a paragraph with the next paragraph.</summary>
    public bool KeepWithNext { get; }
    /// <summary>Gets whether the style keeps paragraph lines together.</summary>
    public bool KeepLinesTogether { get; }
}

/// <summary>Describes one recovered source paragraph.</summary>
public sealed class LegacyWordParagraphContent : LegacyWordBlockContent {
    internal LegacyWordParagraphContent(LegacyWordParagraph source) {
        Text = source.Text;
        Runs = source.Runs.ConvertAll(static run => new LegacyWordRunContent(run)).AsReadOnly();
        IsList = source.IsList;
        ListLevel = source.ListLevel;
        Alignment = source.Alignment;
        PageBreakBefore = source.PageBreakBefore;
        KeepWithNext = source.KeepWithNext;
        KeepLinesTogether = source.KeepLinesTogether;
        LineSpacingPoints = source.LineSpacingPoints;
        SpacingBeforePoints = source.SpacingBeforePoints;
        SpacingAfterPoints = source.SpacingAfterPoints;
        StyleName = source.StyleName;
    }
    /// <summary>Gets the combined paragraph text.</summary>
    public string Text { get; }
    /// <summary>Gets recovered formatted runs.</summary>
    public IReadOnlyList<LegacyWordRunContent> Runs { get; }
    /// <summary>Gets whether the paragraph is a list item.</summary>
    public bool IsList { get; }
    /// <summary>Gets the recovered list nesting level.</summary>
    public int ListLevel { get; }
    /// <summary>Gets recovered alignment.</summary>
    public WordParagraphAlignment? Alignment { get; }
    /// <summary>Gets whether the source requested a page break before this paragraph.</summary>
    public bool PageBreakBefore { get; }
    /// <summary>Gets whether the source requested keeping this paragraph with the next.</summary>
    public bool KeepWithNext { get; }
    /// <summary>Gets whether the source requested keeping paragraph lines together.</summary>
    public bool KeepLinesTogether { get; }
    /// <summary>Gets recovered line spacing in points.</summary>
    public double? LineSpacingPoints { get; }
    /// <summary>Gets recovered spacing before the paragraph in points.</summary>
    public double? SpacingBeforePoints { get; }
    /// <summary>Gets recovered spacing after the paragraph in points.</summary>
    public double? SpacingAfterPoints { get; }
    /// <summary>Gets the recovered source style name, when available.</summary>
    public string? StyleName { get; }
}

/// <summary>Describes a recovered note.</summary>
public sealed class LegacyWordNoteContent {
    internal LegacyWordNoteContent(LegacyWordNote source) {
        Kind = source.Kind; Text = source.Text; IsAnchored = source.IsAnchored;
        Paragraphs = source.Paragraphs.ConvertAll(paragraph => new LegacyWordParagraphContent(paragraph)).AsReadOnly();
    }
    /// <summary>Gets the note kind.</summary>
    public LegacyWordNoteKind Kind { get; }
    /// <summary>Gets bounded note text.</summary>
    public string Text { get; }
    /// <summary>Gets whether a recovered run retains the note's source anchor.</summary>
    public bool IsAnchored { get; }
    /// <summary>Gets formatted note paragraphs where the profile decodes them.</summary>
    public IReadOnlyList<LegacyWordParagraphContent> Paragraphs { get; }
}

/// <summary>Describes an inert source resource reference. Import never resolves it.</summary>
public sealed class LegacyWordResourceReference {
    internal LegacyWordResourceReference(LegacyWordResource source) { Kind = source.Kind; Reference = source.Reference; }
    /// <summary>Gets the source resource kind.</summary>
    public string Kind { get; }
    /// <summary>Gets the bounded source reference.</summary>
    public string Reference { get; }
}

/// <summary>Provides a source-oriented snapshot alongside the projected DOCX model.</summary>
public sealed class LegacyWordContent {
    internal LegacyWordContent(LegacyWordModel source) {
        Paragraphs = source.Paragraphs.ConvertAll(static paragraph => new LegacyWordParagraphContent(paragraph)).AsReadOnly();
        Styles = source.Styles.ConvertAll(static style => new LegacyWordStyleContent(style)).AsReadOnly();
        Notes = source.Notes.ConvertAll(static note => new LegacyWordNoteContent(note)).AsReadOnly();
        Resources = source.Resources.ConvertAll(static resource => new LegacyWordResourceReference(resource)).AsReadOnly();
        if (source.Sections.Count == 0) {
            var section = new LegacyWordSection();
            section.Blocks.AddRange(source.Paragraphs);
            Sections = Array.AsReadOnly(new[] { new LegacyWordSectionContent(section) });
        } else {
            Sections = source.Sections.ConvertAll(section => new LegacyWordSectionContent(section)).AsReadOnly();
        }
    }
    /// <summary>Gets body paragraphs outside tables. Table and running-story paragraphs are available through <see cref="Sections"/>.</summary>
    public IReadOnlyList<LegacyWordParagraphContent> Paragraphs { get; }
    /// <summary>Gets recovered paragraph-style definitions.</summary>
    public IReadOnlyList<LegacyWordStyleContent> Styles { get; }
    /// <summary>Gets recovered notes.</summary>
    public IReadOnlyList<LegacyWordNoteContent> Notes { get; }
    /// <summary>Gets inert resource references.</summary>
    public IReadOnlyList<LegacyWordResourceReference> Resources { get; }
    /// <summary>Gets source-ordered body sections, including tables, page geometry, and running stories when recovered.</summary>
    public IReadOnlyList<LegacyWordSectionContent> Sections { get; }
}

/// <summary>Describes one bounded legacy-word source profile match.</summary>
public sealed class LegacyWordDetection {
    internal LegacyWordDetection(LegacyWordFormat format, string profileId, int confidence, string reason) {
        Format = format;
        ProfileId = profileId;
        Confidence = confidence;
        Reason = reason;
    }

    /// <summary>Gets the detected product family.</summary>
    public LegacyWordFormat Format { get; }
    /// <summary>Gets a stable adapter/profile identifier.</summary>
    public string ProfileId { get; }
    /// <summary>Gets confidence from 0 through 100.</summary>
    public int Confidence { get; }
    /// <summary>Gets the bounded evidence used for detection.</summary>
    public string Reason { get; }
}

/// <summary>Options for safe read-only legacy-word import.</summary>
public sealed class LegacyWordImportOptions {
    /// <summary>Gets or sets hard resource limits.</summary>
    public OfficeLegacyImportLimits Limits { get; set; } = new();
    /// <summary>Gets or sets an explicit family when the source signature is weak or damaged.</summary>
    public LegacyWordFormat? FormatHint { get; set; }
    /// <summary>Gets or sets the source name used for extension-assisted detection.</summary>
    public string? SourceName { get; set; }
    /// <summary>Gets or sets whether salvage-quality output must be rejected.</summary>
    public bool RequireStructured { get; set; }
}

/// <summary>Owns an imported editable Word model and its source-loss report.</summary>
public sealed class LegacyWordImportResult : IDisposable {
    internal LegacyWordImportResult(WordDocument document, LegacyWordDetection detection, OfficeLegacyImportReport report, string plainText, IReadOnlyDictionary<string, string> metadata, LegacyWordContent content) {
        Value = document;
        Detection = detection;
        Report = report;
        PlainText = plainText;
        Metadata = metadata;
        Content = content;
    }

    /// <summary>Gets the normal OfficeIMO Word model used by DOCX and converter packages.</summary>
    public WordDocument Value { get; }
    /// <summary>Gets detected family and profile information.</summary>
    public LegacyWordDetection Detection { get; }
    /// <summary>Gets structured/salvage quality, inert-content flags, and explicit losses.</summary>
    public OfficeLegacyImportReport Report { get; }
    /// <summary>Gets the bounded recovered plain text.</summary>
    public string PlainText { get; }
    /// <summary>Gets recovered source metadata.</summary>
    public IReadOnlyDictionary<string, string> Metadata { get; }
    /// <summary>Gets the source-oriented semantic recovery snapshot.</summary>
    public LegacyWordContent Content { get; }
    /// <summary>Gets whether the import used salvage recovery or omitted, blocked, or kept source content inert.</summary>
    public bool HasLoss => Report.HasLoss;
    /// <summary>Returns the imported Word document.</summary>
    public WordDocument RequireValue() => Value;
    /// <summary>Returns the imported Word document or throws when the import was lossy.</summary>
    public WordDocument RequireNoLoss() {
        try {
            Report.RequireNoLoss();
            return Value;
        } catch {
            Value.Dispose();
            throw;
        }
    }
    /// <inheritdoc />
    public void Dispose() => Value.Dispose();
}
