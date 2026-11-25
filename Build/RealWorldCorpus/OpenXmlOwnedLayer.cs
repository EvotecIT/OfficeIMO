using System.Globalization;
#if OFFICEIMO_OWNED_OPENXML
using OfficeIMO.Core.Internal.OpenXml;
#endif

namespace OfficeIMO.RealWorldCorpus;

/// <summary>
/// OfficeIMO's own Open XML reads that the differential probe compares with the SDK. The project defines
/// <c>OFFICEIMO_OWNED_OPENXML</c> when the package layer is present in the checkout; without it these lanes report
/// <c>unavailable</c> instead of failing, so the runner builds on every branch.
/// </summary>
internal static class OpenXmlOwnedLayer {
#if OFFICEIMO_OWNED_OPENXML
    internal static bool IsAvailable => true;

    /// <summary>The parts the package reports a content type for, and every relationship of the package and those parts.</summary>
    internal static OpenXmlPackageView ReadPackage(byte[] snapshot, IEnumerable<string> candidatePartNames) {
        using var stream = new MemoryStream(snapshot, writable: false);
        using OfficeOpcPackage package = OfficeOpcPackage.Open(stream, leaveOpen: true);
        var view = new OpenXmlPackageView();
        AddRelationships(view, package, null);
        foreach (string name in candidatePartNames.Distinct(StringComparer.Ordinal)) {
            if (OpenXmlPackageView.IsRelationshipsPart(name) || !package.PartExists(name)) continue;
            string? contentType = package.GetContentType(name);
            if (contentType == null) continue;
            if (!view.AddPart(name, contentType)) continue;
            AddRelationships(view, package, name);
        }
        return view;
    }

    private static void AddRelationships(OpenXmlPackageView view, OfficeOpcPackage package, string? source) {
        string relationshipsPart = OfficeOpcPartNames.GetRelationshipsPartName(source);
        if (!package.PartExists(relationshipsPart)) return;
        foreach (OfficeOpcRelationship relationship in package.GetRelationships(source)) {
            view.AddRelationship(source, relationship.Id, relationship.Type, relationship.IsExternal,
                relationship.IsExternal ? relationship.Target : relationship.TargetPartName);
        }
    }

    internal static List<string> ReadWordBodyParagraphs(byte[] snapshot) {
        using var stream = new MemoryStream(snapshot, writable: false);
        return OfficeOpenXmlTextReader.ReadWordBodyParagraphs(stream).ToList();
    }

    internal static List<string> ReadSlideParagraphs(byte[] snapshot) {
        using var stream = new MemoryStream(snapshot, writable: false);
        return OfficeOpenXmlTextReader.ReadSlideParagraphs(stream)
            .Select(paragraph => paragraph.SlideIndex.ToString(CultureInfo.InvariantCulture) + ":" + paragraph.Text)
            .ToList();
    }
#else
    internal static bool IsAvailable => false;

    internal static OpenXmlPackageView ReadPackage(byte[] snapshot, IEnumerable<string> candidatePartNames) => throw new NotSupportedException();

    internal static List<string> ReadWordBodyParagraphs(byte[] snapshot) => throw new NotSupportedException();

    internal static List<string> ReadSlideParagraphs(byte[] snapshot) => throw new NotSupportedException();
#endif

#if OFFICEIMO_WORD_STREAMED_BODY
    internal static bool WordStreamedTextAvailable => true;

    /// <summary>Whether the streamed Word text pass reads the document itself, with the Reader's default limits.</summary>
    /// <returns>Null when it streams; otherwise why it declined, or the exception type it threw.</returns>
    internal static string? GetWordStreamedDeclineReason(byte[] snapshot, OfficeIMO.Reader.ReaderOptions options) {
        using var stream = new MemoryStream(snapshot, writable: false);
        try {
            List<string>? lines = global::OfficeIMO.Word.WordStreamedBodyText.TryReadLines(stream, options.MaxChars, options.MaxTableRows,
                options.OpenXmlMaxCharactersInPart, new OfficeIMO.Reader.Word.ReaderWordOptions().IncludeFootnotes, CancellationToken.None,
                out string? declineReason);
            return lines == null ? declineReason ?? "declined" : null;
        } catch (Exception exception) when (exception is not OutOfMemoryException) {
            return "threw " + exception.GetType().FullName;
        }
    }
#else
    internal static bool WordStreamedTextAvailable => false;

    internal static string? GetWordStreamedDeclineReason(byte[] snapshot, OfficeIMO.Reader.ReaderOptions options) => throw new NotSupportedException();
#endif
}
