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
    internal LegacyWordImportResult(WordDocument document, LegacyWordDetection detection, OfficeLegacyImportReport report, string plainText, IReadOnlyDictionary<string, string> metadata) {
        Document = document;
        Detection = detection;
        Report = report;
        PlainText = plainText;
        Metadata = metadata;
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
    /// <inheritdoc />
    public void Dispose() => Document.Dispose();
}
