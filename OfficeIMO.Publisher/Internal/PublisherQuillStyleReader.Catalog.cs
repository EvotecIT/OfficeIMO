namespace OfficeIMO.Publisher.Internal;

internal sealed partial class PublisherQuillStyleReader {
    private PublisherQuillStyleCatalog ReadCatalog() {
        var libraries = new List<PublisherQuillStyleLibrary>();
        foreach (PublisherQuillChunk chunk in _chunks.Where(item => item.Tag == "STSH" && item.Id is 0 or 1 or 2)) {
            if (libraries.Any(item => item.Id == chunk.Id)) throw new InvalidDataException("Duplicate Publisher STSH library.");
            _data.Range(chunk.Offset, 20, chunk.End);
            int count = _data.Offset(_data.U32(chunk.Offset + 4));
            if (count > (long)_context.Options.Limits.MaxItems * 2)
                throw new InvalidDataException("Publisher indexed style library exceeds the item limit.");
            int basis = checked(chunk.Offset + 20), end = checked(basis + _data.Offset(_data.U32(chunk.Offset)));
            _data.Range(basis, end - basis, chunk.End);
            _data.Range(basis, checked(count * 4), end);
            var offsets = new int[count];
            for (int index = 0; index < count; index++) {
                _context.Record();
                int offset = checked(basis + _data.Offset(_data.U32(basis + index * 4)));
                if (offset < basis + count * 4 || offset >= end) throw new InvalidDataException("Publisher style entry points outside its indexed payload.");
                offsets[index] = offset;
            }
            // Index order is identity; physical order determines each raw extent.
            // Shared offsets retain distinct ordinals without guessing a new entry.
            int[] physical = offsets.Distinct().OrderBy(offset => offset).ToArray();
            var lengths = physical.Select((offset, index) => (Offset: offset,
                Length: (index + 1 < physical.Length ? physical[index + 1] : end) - offset)).ToDictionary(item => item.Offset, item => item.Length);
            libraries.Add(new PublisherQuillStyleLibrary(chunk.Id, new PublisherQuillSourceSpan(chunk.Offset, chunk.Length),
                offsets.Select(offset => new PublisherQuillSourceSpan(offset, lengths[offset]))));
        }
        var definitions = new List<PublisherQuillStyleDefinition>();
        PublisherQuillStyleLibrary? source = libraries.SingleOrDefault(item => item.Id == 1);
        if (source != null) {
            if ((source.Entries.Count & 1) != 0) throw new InvalidDataException("Publisher default style pairs are incomplete.");
            for (int index = 0; index < source.Entries.Count; index += 2) {
                PublisherQuillSourceSpan character = source.Entries[index], paragraph = source.Entries[index + 1];
                PublisherQuillSourceSpan characterDefinition = Definition(character), paragraphDefinition = Definition(paragraph);
                definitions.Add(new PublisherQuillStyleDefinition(index / 2,
                    ReadCharacter(characterDefinition.Offset, characterDefinition.Offset + characterDefinition.Length),
                    ReadParagraph(paragraphDefinition.Offset, paragraphDefinition.Offset + paragraphDefinition.Length),
                    character, paragraph, characterDefinition, paragraphDefinition));
            }
            if (definitions.Count > 1) _context.Add("PUB_NAMED_STYLE_METADATA_UNASSESSED",
                "Referenced native style values are projected. Names, builtin selectors and auxiliary inheritance metadata remain opaque and are not an editable style library.",
                OfficeConversionLossKind.Unassessed, "Quill/STSH");
        }
        return new PublisherQuillStyleCatalog(definitions, libraries, _fonts, _colors);
    }

    private PublisherQuillSourceSpan Definition(PublisherQuillSourceSpan entry) {
        _data.Range(entry.Offset, 6, entry.Offset + entry.Length);
        int offset = entry.Offset + 2, length = _data.Offset(_data.U32(offset));
        if (length < 4) throw new InvalidDataException("Publisher style definition is smaller than its OPL header.");
        _data.Range(offset, length, entry.Offset + entry.Length);
        return new PublisherQuillSourceSpan(offset, length);
    }
}
