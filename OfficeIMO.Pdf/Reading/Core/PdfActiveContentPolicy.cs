using System;
using System.Collections.Generic;

namespace OfficeIMO.Pdf;

internal static class PdfActiveContentPolicy {
    private static readonly HashSet<string> UnsafeActionTypes = new(StringComparer.Ordinal) {
        "JavaScript", "Launch", "GoToR", "GoToE", "SubmitForm", "ImportData", "Movie", "Rendition", "RichMedia"
    };

    internal static readonly string[] MarkerNames = {
        "JavaScript", "JS", "AA", "Launch", "GoToR", "GoToE", "SubmitForm", "ImportData", "Movie", "Rendition", "RichMedia"
    };

    internal static bool IsUnsafeActionType(string actionType) => UnsafeActionTypes.Contains(actionType);
}
