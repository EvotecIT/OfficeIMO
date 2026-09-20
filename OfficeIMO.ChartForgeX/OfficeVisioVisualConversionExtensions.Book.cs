using System;
using System.Collections.Generic;
using System.Linq;
using global::ChartForgeX.VisualArtifacts;
using OfficeIMO.Visio;

namespace OfficeIMO.ChartForgeX;

public static partial class OfficeVisioVisualConversionExtensions {
    /// <summary>
    /// Projects a sequence of semantic envelopes into one editable multi-page Visio document.
    /// Input order becomes page order; every page retains its own fidelity report.
    /// </summary>
    public static OfficeVisioVisualBookResult ToOfficeVisioBook(
        this IEnumerable<VisualArtifactInterchangeEnvelope> envelopes, OfficeVisioVisualOptions? options = null) {
        if (envelopes == null) throw new ArgumentNullException(nameof(envelopes));
        options ??= new OfficeVisioVisualOptions();
        var source = envelopes.ToList();
        if (source.Count == 0) throw new ArgumentException("At least one page is required.", nameof(envelopes));
        foreach (var envelope in source) {
            if (envelope == null) throw new ArgumentException("Pages cannot contain a null envelope.", nameof(envelopes));
            envelope.Validate();
        }
        var document = VisioDocument.Create();
        var pages = new List<OfficeVisioVisualConversionResult>();
        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var envelope in source) {
            string baseName = string.IsNullOrWhiteSpace(envelope.Title) ? options.PageName : envelope.Title;
            string name = baseName;
            int suffix = 2;
            while (!names.Add(name)) name = baseName + " (" + suffix++ + ")";
            pages.Add(ProjectPage(envelope, document, options.ForPage(name)));
        }
        return new OfficeVisioVisualBookResult(document, pages);
    }
}

/// <summary>A multi-page native Visio document with a fidelity report for every input envelope.</summary>
public sealed class OfficeVisioVisualBookResult {
    internal OfficeVisioVisualBookResult(VisioDocument document, List<OfficeVisioVisualConversionResult> pages) {
        Document = document; Pages = pages.AsReadOnly();
    }
    /// <summary>Gets the document containing every projected page.</summary>
    public VisioDocument Document { get; }
    /// <summary>Gets page results in input order, including each page's semantic fidelity diagnostics.</summary>
    public IReadOnlyList<OfficeVisioVisualConversionResult> Pages { get; }
}
