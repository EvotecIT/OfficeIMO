namespace OfficeIMO.Bibliography;

/// <summary>Copies semantic values and their codec bindings without serialization or shared mutable children.</summary>
internal sealed class BibliographyModelCopy {
    private readonly BibliographyReferenceOptions _options;
    private readonly CancellationToken _cancellationToken;
    private long _characters;
    private int _values;

    internal BibliographyModelCopy(BibliographyReferenceOptions options, CancellationToken cancellationToken) {
        _options = options;
        _cancellationToken = cancellationToken;
    }

    internal string? Value(string? value) {
        _cancellationToken.ThrowIfCancellationRequested();
        if (value == null) return null;
        if (_values >= _options.MaximumValues || value.Length > _options.MaximumExpandedCharacters - _characters)
            throw new InvalidOperationException("Bibliography reference resolution exceeds its copied-value or expanded-character limit.");
        _values++;
        _characters += value.Length;
        return value;
    }

    internal BibliographyNativeField Field(BibliographyNativeField source) {
        Value(source.Name); Value(source.Value); Value(source.RawValue);
        return source.Copy();
    }

    internal BibliographyContributor Contributor(BibliographyContributor source) {
        Value(string.Empty);
        var name = new BibliographyName {
            Given = Value(source.Name.Given), Family = Value(source.Name.Family), Literal = Value(source.Name.Literal),
            Suffix = Value(source.Name.Suffix), DroppingParticle = Value(source.Name.DroppingParticle),
            NonDroppingParticle = Value(source.Name.NonDroppingParticle)
        };
        foreach (BibliographyNativeField field in source.Name.NativeFields) name.NativeFields.Add(Field(field));
        return new BibliographyContributor(source.Role, name);
    }

    internal BibliographyDate Date(BibliographyDate source) {
        // Charge date objects even when all of their public values are empty.
        Value(string.Empty);
        var date = new BibliographyDate { Role = source.Role, Year = source.Year, Month = source.Month, Day = source.Day,
            EndYear = source.EndYear, EndMonth = source.EndMonth, EndDay = source.EndDay, Literal = Value(source.Literal) };
        foreach (BibliographyNativeField field in source.NativeFields) date.NativeFields.Add(Field(field));
        return date;
    }

    internal BibliographyIdentifier Identifier(BibliographyIdentifier source) =>
        new BibliographyIdentifier(Value(source.Scheme)!, Value(source.Value)!);

    internal BibliographyItem Item(BibliographyItem source) {
        var item = new BibliographyItem {
            Key = Value(source.Key)!, Type = source.Type, NativeType = Value(source.NativeType), Title = Value(source.Title),
            ContainerTitle = Value(source.ContainerTitle), CollectionTitle = Value(source.CollectionTitle), Publisher = Value(source.Publisher),
            PublisherPlace = Value(source.PublisherPlace), Edition = Value(source.Edition), Volume = Value(source.Volume), Issue = Value(source.Issue),
            Pages = Value(source.Pages), Abstract = Value(source.Abstract), Language = Value(source.Language), Url = Value(source.Url),
            RisPageStart = Value(source.RisPageStart), RisPageEnd = Value(source.RisPageEnd), BibMonthWasNumeric = source.BibMonthWasNumeric,
            CslNumericKey = Value(source.CslNumericKey), CslNumericKeyRaw = Value(source.CslNumericKeyRaw)
        };
        CopyBindings(source.BibFieldNames, item.BibFieldNames);
        CopyBindings(source.EndNoteFieldNames, item.EndNoteFieldNames);
        CopyBindings(source.TaggedFieldNames, item.TaggedFieldNames);
        foreach (string binding in source.TaggedScalarBindings) item.TaggedScalarBindings.Add(Value(binding)!);
        foreach (BibliographyContributor contributor in source.Contributors) {
            BibliographyContributor copy = Contributor(contributor);
            item.Contributors.Add(copy);
            if (source.TaggedContributorTags.TryGetValue(contributor, out string? tag)) item.TaggedContributorTags[copy] = Value(tag)!;
        }
        foreach (BibliographyDate date in source.Dates) {
            BibliographyDate copy = Date(date);
            item.Dates.Add(copy);
            if (source.TaggedDateTags.TryGetValue(date, out string? tag)) item.TaggedDateTags[copy] = Value(tag)!;
        }
        foreach (BibliographyIdentifier identifier in source.Identifiers) {
            BibliographyIdentifier copy = Identifier(identifier);
            item.Identifiers.Add(copy);
            if (source.TaggedIdentifierTags.TryGetValue(identifier, out string? tag)) item.TaggedIdentifierTags[copy] = Value(tag)!;
        }
        foreach (string keyword in source.Keywords) item.Keywords.Add(Value(keyword)!);
        foreach (string note in source.Notes) item.Notes.Add(Value(note)!);
        foreach (BibliographyNativeField field in source.NativeFields) {
            BibliographyNativeField copy = Field(field);
            item.NativeFields.Add(copy);
            if (ReferenceEquals(source.NbibTypeBinding, field)) item.NbibTypeBinding = copy;
        }
        return item;
    }

    private void CopyBindings(IDictionary<string, string> source, IDictionary<string, string> destination) {
        foreach (KeyValuePair<string, string> pair in source) destination[Value(pair.Key)!] = Value(pair.Value)!;
    }
}
