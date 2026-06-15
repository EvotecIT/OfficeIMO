using System.Collections.Generic;

namespace OfficeIMO.Word.Rtf;

public static partial class WordRtfConverterExtensions {
    private static int? GetOrAddRevisionAuthorIndex(RtfDocument document, string? author, Dictionary<string, int> revisionAuthorIndexes) {
        if (string.IsNullOrWhiteSpace(author)) {
            return null;
        }

        if (revisionAuthorIndexes.TryGetValue(author!, out int existing)) {
            return existing;
        }

        int index = document.AddRevisionAuthor(author!);
        revisionAuthorIndexes[author!] = index;
        return index;
    }

    private static string GetRevisionAuthorName(RtfDocument? document, int? authorIndex) {
        if (document != null && authorIndex.HasValue && authorIndex.Value >= 0 && authorIndex.Value < document.RevisionAuthors.Count) {
            string name = document.RevisionAuthors[authorIndex.Value].Name;
            if (!string.IsNullOrWhiteSpace(name)) {
                return name;
            }
        }

        return "OfficeIMO";
    }
}
