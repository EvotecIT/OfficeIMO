using System;
using System.Collections.Generic;
using System.Linq;
using OfficeIMO;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Services.PowerPoint;

/// <summary>Resolves live slide targets and manual placement on the native twelve-column grid.</summary>
internal static class PowerPointManualPlacement {
    /// <summary>Explicit coordinates select fixed placement instead of the automatic content zones.</summary>
    internal static bool HasExplicitBounds(IEnumerable<string> parameters) =>
        parameters.Any(name => name is "X" or "Y" or "Width" or "Height");

    internal static (PowerPointSlide Slide, PowerPointContentScope? Scope) Resolve(PowerPointSlideSelector? selection, PowerPointDocument? document = null) {
        var scope = PowerPointContentScope.Current;
        document ??= OfficeDocumentContext.Find<PowerPointDocument>();
        PowerPointSlide slide = selection?.Resolve(document) ??
            (scope != null && (document == null || ReferenceEquals(scope.Document, document)) ? scope.NativeSlide : null)
            ?? throw new ArgumentException("Supply a native Slide or invoke this command inside PowerPointSlide.");
        return (slide, scope != null && ReferenceEquals(scope.NativeSlide, slide) ? scope : null);
    }

    internal static (OfficeLength? X, OfficeLength? Width) Horizontal(PowerPointSlide slide,
        OfficeLength? x, OfficeLength? width, int? column, int? span) {
        if (!column.HasValue && !span.HasValue) return (x, width);
        int start = column ?? 1, count = span ?? 12;
        if (start < 1 || count < 1 || start > 12 || count > 13 - start)
            throw new ArgumentException("Column and Span must fit within the twelve-column slide grid.");
        var document = slide.OwnerDocument ?? throw new ArgumentException("The slide has no owning document.");
        var context = new PowerPointLayoutContext(document, PowerPointLayoutTheme.From(slide.AuthoringTheme));
        return (x ?? OfficeLength.Points(context.Grid.X(start)), width ?? OfficeLength.Points(context.Grid.W(count)));
    }
}
