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
    public int? FontSizePoints { get; }
    /// <summary>Gets the recovered font family.</summary>
    public string? FontFamily { get; }
    /// <summary>Gets the recovered RGB color.</summary>
    public string? ColorHex { get; }
}

/// <summary>Describes one recovered source paragraph.</summary>
public sealed class LegacyWordParagraphContent {
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
    internal LegacyWordNoteContent(LegacyWordNote source) { Kind = source.Kind; Text = source.Text; }
    /// <summary>Gets the note kind.</summary>
    public LegacyWordNoteKind Kind { get; }
    /// <summary>Gets bounded note text.</summary>
    public string Text { get; }
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
        Notes = source.Notes.ConvertAll(static note => new LegacyWordNoteContent(note)).AsReadOnly();
        Resources = source.Resources.ConvertAll(static resource => new LegacyWordResourceReference(resource)).AsReadOnly();
    }
    /// <summary>Gets recovered paragraphs and formatted runs.</summary>
    public IReadOnlyList<LegacyWordParagraphContent> Paragraphs { get; }
    /// <summary>Gets recovered notes.</summary>
    public IReadOnlyList<LegacyWordNoteContent> Notes { get; }
    /// <summary>Gets inert resource references.</summary>
    public IReadOnlyList<LegacyWordResourceReference> Resources { get; }
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
        Document = document;
        Detection = detection;
        Report = report;
        PlainText = plainText;
        Metadata = metadata;
        Content = content;
    }

    /// <summary>Gets the normal OfficeIMO Word model used by DOCX and converter packages.</summary>
    public WordDocument Document { get; }
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
    /// <inheritdoc />
    public void Dispose() => Document.Dispose();
}
