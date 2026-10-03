namespace OfficeIMO.Bibliography;

internal sealed partial class BibliographyReferenceResolver {
    private static readonly Field[] ScalarFields = {
        new Field("title", item => item.Title, (item, value) => item.Title = value),
        new Field("container-title", item => item.ContainerTitle, (item, value) => item.ContainerTitle = value),
        new Field("collection-title", item => item.CollectionTitle, (item, value) => item.CollectionTitle = value),
        new Field("publisher", item => item.Publisher, (item, value) => item.Publisher = value),
        new Field("publisher-place", item => item.PublisherPlace, (item, value) => item.PublisherPlace = value),
        new Field("edition", item => item.Edition, (item, value) => item.Edition = value),
        new Field("volume", item => item.Volume, (item, value) => item.Volume = value),
        new Field("issue", item => item.Issue, (item, value) => item.Issue = value),
        new Field("pages", item => item.Pages, (item, value) => item.Pages = value),
        new Field("abstract", item => item.Abstract, (item, value) => item.Abstract = value),
        new Field("language", item => item.Language, (item, value) => item.Language = value),
        new Field("url", item => item.Url, (item, value) => item.Url = value)
    };
    private static readonly ISet<string> ControlFields = new HashSet<string>(new[] {
        "crossref", "xdata", "xref", "ids", "entryset", "entrysubtype", "execute", "label", "options", "presort",
        "related", "relatedoptions", "relatedstring", "relatedtype", "shorthand", "shorthandintro", "sortkey"
    }, StringComparer.OrdinalIgnoreCase);

    private void Merge(Reference reference) {
        BibliographyItem child = _items[reference.Child], parent = _items[reference.Parent];
        bool overwrite = reference.Relation == "xdata" && _options.XDataOverridesExistingFields;
        bool containerTitle = reference.Relation == "crossref" && HasContainingTitle(parent, child);
        foreach (Field source in ScalarFields) {
            _cancellationToken.ThrowIfCancellationRequested();
            Field target = containerTitle && source.Name == "title" ? ScalarFields[1] : source;
            string? value = source.Get(parent);
            if (value == null || !overwrite && target.Get(child) != null) continue;
            target.Set(child, _copy.Value(value));
            child.BibFieldNames.Remove(target.Name);
            if (parent.BibFieldNames.TryGetValue(source.Name, out string? binding)) child.BibFieldNames[target.Name] = _copy.Value(binding)!;
            if (containerTitle && source.Name == "title") child.BibFieldNames["container-title"] = child.Type == BibliographyItemType.ArticleJournal ? "journaltitle" : "booktitle";
            if (source.Name == "pages") { child.RisPageStart = _copy.Value(parent.RisPageStart); child.RisPageEnd = _copy.Value(parent.RisPageEnd); }
            Record(reference, target.Name, source.Name);
        }
        MergeContributors(reference, overwrite);
        MergeDates(reference, overwrite);
        MergeIdentifiers(reference, overwrite);
        MergeStrings(reference, "keywords", child.Keywords, parent.Keywords, overwrite);
        MergeStrings(reference, "notes", child.Notes, parent.Notes, overwrite);
        MergeNativeFields(reference, overwrite, containerTitle);
    }

    private static bool HasContainingTitle(BibliographyItem parent, BibliographyItem child) =>
        (child.Type == BibliographyItemType.Chapter && (parent.Type == BibliographyItemType.Book ||
            string.Equals(parent.NativeType, "collection", StringComparison.OrdinalIgnoreCase) || string.Equals(parent.NativeType, "reference", StringComparison.OrdinalIgnoreCase))) ||
        (child.Type == BibliographyItemType.PaperConference && parent.Type == BibliographyItemType.Proceedings) ||
        (child.Type == BibliographyItemType.ArticleJournal && string.Equals(parent.NativeType, "periodical", StringComparison.OrdinalIgnoreCase));

    private void MergeContributors(Reference reference, bool overwrite) {
        BibliographyItem child = _items[reference.Child], parent = _items[reference.Parent];
        foreach (IGrouping<BibliographyContributorRole, BibliographyContributor> group in parent.Contributors.GroupBy(contributor => contributor.Role)) {
            _cancellationToken.ThrowIfCancellationRequested();
            if (!overwrite && child.Contributors.Any(contributor => contributor.Role == group.Key)) continue;
            for (int index = child.Contributors.Count - 1; index >= 0; index--) {
                _cancellationToken.ThrowIfCancellationRequested();
                BibliographyContributor contributor = child.Contributors[index];
                if (contributor.Role != group.Key) continue;
                child.TaggedContributorTags.Remove(contributor);
                child.Contributors.RemoveAt(index);
            }
            foreach (BibliographyContributor contributor in group) {
                BibliographyContributor copy = _copy.Contributor(contributor);
                child.Contributors.Add(copy);
                if (parent.TaggedContributorTags.TryGetValue(contributor, out string? tag)) child.TaggedContributorTags[copy] = _copy.Value(tag)!;
            }
            string field = "contributors." + group.Key.ToString().ToLowerInvariant();
            Record(reference, field, field);
        }
    }

    private void MergeDates(Reference reference, bool overwrite) {
        BibliographyItem child = _items[reference.Child], parent = _items[reference.Parent];
        foreach (IGrouping<BibliographyDateRole, BibliographyDate> group in parent.Dates.GroupBy(date => date.Role)) {
            _cancellationToken.ThrowIfCancellationRequested();
            if (!overwrite && child.Dates.Any(date => date.Role == group.Key)) continue;
            for (int index = child.Dates.Count - 1; index >= 0; index--) {
                _cancellationToken.ThrowIfCancellationRequested();
                BibliographyDate date = child.Dates[index];
                if (date.Role != group.Key) continue;
                child.TaggedDateTags.Remove(date);
                child.Dates.RemoveAt(index);
            }
            foreach (BibliographyDate date in group) {
                BibliographyDate copy = _copy.Date(date);
                child.Dates.Add(copy);
                if (parent.TaggedDateTags.TryGetValue(date, out string? tag)) child.TaggedDateTags[copy] = _copy.Value(tag)!;
            }
            if (group.Key == BibliographyDateRole.Issued) child.BibMonthWasNumeric = parent.BibMonthWasNumeric;
            string field = "dates." + group.Key.ToString().ToLowerInvariant();
            Record(reference, field, field);
        }
    }

    private void MergeIdentifiers(Reference reference, bool overwrite) {
        BibliographyItem child = _items[reference.Child], parent = _items[reference.Parent];
        foreach (IGrouping<string, BibliographyIdentifier> group in parent.Identifiers.GroupBy(identifier => identifier.Scheme, StringComparer.OrdinalIgnoreCase)) {
            _cancellationToken.ThrowIfCancellationRequested();
            if (!overwrite && child.Identifiers.Any(identifier => string.Equals(identifier.Scheme, group.Key, StringComparison.OrdinalIgnoreCase))) continue;
            for (int index = child.Identifiers.Count - 1; index >= 0; index--) {
                _cancellationToken.ThrowIfCancellationRequested();
                BibliographyIdentifier identifier = child.Identifiers[index];
                if (!string.Equals(identifier.Scheme, group.Key, StringComparison.OrdinalIgnoreCase)) continue;
                child.TaggedIdentifierTags.Remove(identifier);
                child.Identifiers.RemoveAt(index);
            }
            foreach (BibliographyIdentifier identifier in group) {
                BibliographyIdentifier copy = _copy.Identifier(identifier);
                child.Identifiers.Add(copy);
                if (parent.TaggedIdentifierTags.TryGetValue(identifier, out string? tag)) child.TaggedIdentifierTags[copy] = _copy.Value(tag)!;
            }
            string field = "identifiers." + group.Key.ToLowerInvariant();
            Record(reference, field, field);
        }
    }

    private void MergeStrings(Reference reference, string field, IList<string> child, IList<string> parent, bool overwrite) {
        _cancellationToken.ThrowIfCancellationRequested();
        if (parent.Count == 0 || !overwrite && child.Count != 0) return;
        child.Clear();
        foreach (string value in parent) child.Add(_copy.Value(value)!);
        Record(reference, field, field);
    }

    private void MergeNativeFields(Reference reference, bool overwrite, bool containerTitle) {
        BibliographyItem child = _items[reference.Child], parent = _items[reference.Parent];
        var groups = new List<IGrouping<string, BibliographyNativeField>>();
        foreach (IGrouping<string, BibliographyNativeField> group in parent.NativeFields.GroupBy(NativeKey, StringComparer.Ordinal)) {
            _cancellationToken.ThrowIfCancellationRequested();
            BibliographyNativeField first = group.First();
            if (IsBib(first.Format) && (ControlFields.Contains(first.Name) || containerTitle &&
                (first.Name.Equals("shorttitle", StringComparison.OrdinalIgnoreCase) || first.Name.Equals("sorttitle", StringComparison.OrdinalIgnoreCase) ||
                 first.Name.Equals("indextitle", StringComparison.OrdinalIgnoreCase) || first.Name.Equals("indexsorttitle", StringComparison.OrdinalIgnoreCase)))) continue;
            groups.Add(group);
        }

        var childKeys = new HashSet<string>(StringComparer.Ordinal);
        foreach (BibliographyNativeField field in child.NativeFields) {
            _cancellationToken.ThrowIfCancellationRequested();
            childKeys.Add(NativeKey(field));
        }

        if (overwrite) {
            var replacedKeys = new HashSet<string>(groups.Select(group => group.Key), StringComparer.Ordinal);
            int retained = 0;
            for (int index = 0; index < child.NativeFields.Count; index++) {
                _cancellationToken.ThrowIfCancellationRequested();
                BibliographyNativeField field = child.NativeFields[index];
                if (replacedKeys.Contains(NativeKey(field))) {
                    if (ReferenceEquals(child.NbibTypeBinding, field)) child.NbibTypeBinding = null;
                    continue;
                }
                child.NativeFields[retained++] = field;
            }
            while (child.NativeFields.Count > retained) child.NativeFields.RemoveAt(child.NativeFields.Count - 1);
        }

        foreach (IGrouping<string, BibliographyNativeField> group in groups) {
            _cancellationToken.ThrowIfCancellationRequested();
            if (!overwrite && childKeys.Contains(group.Key)) continue;
            foreach (BibliographyNativeField field in group) child.NativeFields.Add(_copy.Field(field));
            string name = "native." + group.Key;
            Record(reference, name, name);
        }
    }

    private static string NativeKey(BibliographyNativeField field) => ((int)field.Format).ToString(CultureInfo.InvariantCulture) + "." + field.Name.ToLowerInvariant();

    private sealed class Field {
        internal Field(string name, Func<BibliographyItem, string?> get, Action<BibliographyItem, string?> set) { Name = name; Get = get; Set = set; }
        internal string Name { get; }
        internal Func<BibliographyItem, string?> Get { get; }
        internal Action<BibliographyItem, string?> Set { get; }
    }
}
