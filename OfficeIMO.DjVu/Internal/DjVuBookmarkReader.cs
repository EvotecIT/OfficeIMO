using System.Globalization;

namespace OfficeIMO.DjVu;

internal static class DjVuBookmarkReader {
    internal static IReadOnlyList<DjVuBookmark> Read(DjVuChunk root, DjVuDocument document, DjVuReadBudget budget, out string? diagnostic) {
        diagnostic = null;
        var result = new List<DjVuBookmark>();
        var chunks = root.Children.Where(c => c.Id == "NAVM").ToArray();
        if (chunks.Length == 0) return result.AsReadOnly();
        try {
            if (root.FormType != "DJVM" || chunks.Length != 1 || root.Children.Count < 2 || root.Children[1].Id != "NAVM")
                throw new InvalidDataException("DjVu outline must immediately follow the document directory.");
            var chunk = chunks[0];
            byte[] data = BzzDecoder.Decode(chunk.Source, chunk.Offset, chunk.Length, budget);
            if (data.Length < 2) throw new InvalidDataException("Truncated DjVu outline.");
            int count = DjVuBinary.U16(data, 0), position = 2, remaining = count, characters = 0;
            if (count > budget.Options.MaxBookmarks) throw new DjVuResourceLimitException(nameof(DjVuReadOptions.MaxBookmarks));
            while (remaining != 0) result.Add(Entry(1));
            if (position != data.Length) throw new InvalidDataException("Trailing DjVu outline records.");
            return result.AsReadOnly();

            DjVuBookmark Entry(int depth) {
                budget.Cancellation.ThrowIfCancellationRequested();
                if (depth > budget.Options.MaxBookmarkDepth) throw new DjVuResourceLimitException(nameof(DjVuReadOptions.MaxBookmarkDepth));
                if (remaining-- <= 0 || position >= data.Length) throw new InvalidDataException("Invalid DjVu outline child count.");
                int childCount = data[position++];
                string title = String(), target = String();
                if (childCount > remaining) throw new InvalidDataException("DjVu outline children exceed the declared record count.");
                var children = new List<DjVuBookmark>();
                for (int i = 0; i < childCount; i++) children.Add(Entry(depth + 1));
                return new DjVuBookmark(title, target, Resolve(target, document), children);
            }
            string String() {
                if (data.Length - position < 3) throw new InvalidDataException("Truncated DjVu outline string.");
                int size = DjVuBinary.U24(data, position); position += 3;
                if (size > data.Length - position) throw new InvalidDataException("DjVu outline string exceeds the chunk.");
                int characterCount = DjVuBinary.Utf8.GetCharCount(data, position, size);
                if (characterCount > budget.Options.MaxBookmarkCharacters - characters)
                    throw new DjVuResourceLimitException(nameof(DjVuReadOptions.MaxBookmarkCharacters));
                characters += characterCount;
                string value = DjVuBinary.Utf8.GetString(data, position, size); position += size;
                return value;
            }
        } catch (InvalidDataException error) { diagnostic = error.Message; return Array.Empty<DjVuBookmark>(); }
        catch (DecoderFallbackException error) { diagnostic = error.Message; return Array.Empty<DjVuBookmark>(); }
    }

    private static int? Resolve(string target, DjVuDocument document) {
        string id = target.StartsWith("#", StringComparison.Ordinal) ? target.Substring(1) : target;
        var page = document.Pages.FirstOrDefault(p => p.Id == id || p.Component.Name == id);
        if (page != null) return page.Number;
        if (target.StartsWith("#", StringComparison.Ordinal) && int.TryParse(id, NumberStyles.None, CultureInfo.InvariantCulture, out int number) && number > 0 && number <= document.Pages.Count)
            return number;
        return null;
    }
}
