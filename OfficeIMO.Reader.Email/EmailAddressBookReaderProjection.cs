using OfficeIMO.Email;
using OfficeIMO.Email.AddressBook;

namespace OfficeIMO.Reader.Email;

internal static class EmailAddressBookReaderProjection {
    internal static ReaderChunk CreateChunk(OfflineAddressBookEntry entry,
        string logicalPath, ReaderOptions readerOptions, ReaderEmailAddressBookOptions options) {
        var text = new StringBuilder();
        var markdown = new StringBuilder();
        Add(text, markdown, "Name", entry.DisplayName);
        Add(text, markdown, "Object type", entry.ObjectType.ToString());
        Add(text, markdown, "Primary email", entry.SmtpAddress);
        Add(text, markdown, "X500 address", entry.X500Address);
        Add(text, markdown, "Target address", entry.TargetAddress);
        Add(text, markdown, "Account", entry.Account);
        Add(text, markdown, "Given name", entry.GivenName);
        Add(text, markdown, "Surname", entry.Surname);
        Add(text, markdown, "Company", entry.CompanyName);
        Add(text, markdown, "Department", entry.Department);
        Add(text, markdown, "Title", entry.JobTitle);
        Add(text, markdown, "Office", entry.OfficeLocation);
        Add(text, markdown, "Business phone", entry.BusinessTelephone);
        Add(text, markdown, "Home phone", entry.HomeTelephone);
        Add(text, markdown, "Mobile phone", entry.MobileTelephone);
        Add(text, markdown, "Fax", entry.PrimaryFax);
        Add(text, markdown, "Street", entry.StreetAddress);
        Add(text, markdown, "City", entry.Locality);
        Add(text, markdown, "State or province", entry.StateOrProvince);
        Add(text, markdown, "Postal code", entry.PostalCode);
        Add(text, markdown, "Country", entry.Country);
        Add(text, markdown, "Comment", entry.Comment);
        AddValues(text, markdown, "Proxy address", entry.ProxyAddresses, options.MaxMultiValueItems);
        AddValues(text, markdown, "Additional business phone", entry.BusinessTelephone2, options.MaxMultiValueItems);
        AddValues(text, markdown, "Additional home phone", entry.HomeTelephone2, options.MaxMultiValueItems);
        if (entry.IsDistributionList) {
            long count = entry.DistributionListMemberCount.HasValue
                ? entry.DistributionListMemberCount.Value
                : entry.MemberDistinguishedNames.Count;
            Add(text, markdown, "Distribution-list member count", count.ToString(CultureInfo.InvariantCulture));
        }
        if (options.IncludeMembershipValues) {
            AddValues(text, markdown, "Member", entry.MemberDistinguishedNames, options.MaxMultiValueItems);
            AddValues(text, markdown, "Member of", entry.MemberOfDistinguishedNames, options.MaxMultiValueItems);
        }
        if (entry.TruncatedPropertyTags.Count > 0) {
            Add(text, markdown, "Truncated property count",
                entry.TruncatedPropertyTags.Count.ToString(CultureInfo.InvariantCulture));
        }

        string textValue = text.ToString().TrimEnd();
        string markdownValue = markdown.ToString().TrimEnd();
        var warnings = new List<string>();
        if (textValue.Length > readerOptions.MaxChars) {
            textValue = textValue.Substring(0, readerOptions.MaxChars);
            markdownValue = markdownValue.Length <= readerOptions.MaxChars
                ? markdownValue
                : markdownValue.Substring(0, readerOptions.MaxChars);
            warnings.Add("Address-book entry projection was truncated due to ReaderOptions.MaxChars.");
        }
        string sourceId = BuildSourceId(entry.AddressList.SourcePath);
        var chunk = new ReaderChunk {
            Id = string.Concat("oab:", entry.Reference.AddressListIndex.ToString("D4", CultureInfo.InvariantCulture),
                ":", entry.Reference.EntryIndex.ToString("D10", CultureInfo.InvariantCulture)),
            Kind = ReaderInputKind.Email,
            Location = new ReaderLocation {
                Path = logicalPath,
                BlockIndex = entry.Reference.EntryIndex <= int.MaxValue
                    ? (int?)entry.Reference.EntryIndex
                    : null,
                SourceBlockIndex = entry.Reference.EntryIndex <= int.MaxValue
                    ? (int?)entry.Reference.EntryIndex
                    : null,
                HeadingPath = entry.AddressList.Name,
                SourceBlockKind = entry.IsDistributionList ? "address-book-distribution-list" : "address-book-entry",
                BlockAnchor = entry.Reference.Id
            },
            SourceId = sourceId,
            SourceLengthBytes = entry.AddressList.SourceLength,
            Text = textValue,
            Markdown = markdownValue,
            TokenEstimate = textValue.Length == 0 ? 0 : Math.Max(1, (textValue.Length + 3) / 4),
            Warnings = warnings.Count == 0 ? null : warnings
        };
        if (readerOptions.ComputeHashes) chunk.ChunkHash = ComputeSha256(textValue + "\n" + markdownValue);
        return chunk;
    }

    internal static EmailDiagnostic ProjectionError(Exception exception, string location) =>
        new EmailDiagnostic(
            "OAB_READER_ENTRY_SKIPPED",
            exception.Message,
            exception is OfflineAddressBookLimitExceededException
                ? EmailDiagnosticSeverity.Warning
                : EmailDiagnosticSeverity.Error,
            location);

    internal static string LogicalPath(string sourceName, OfflineAddressBookEntryReference reference) =>
        string.Concat(Path.GetFileName(sourceName), "/address-list-",
            reference.AddressListIndex.ToString("D4", CultureInfo.InvariantCulture), "/entry-",
            reference.EntryIndex.ToString("D10", CultureInfo.InvariantCulture));

    internal static OfficeDocumentDiagnostic MapDiagnostic(EmailDiagnostic diagnostic) =>
        new OfficeDocumentDiagnostic {
            Severity = diagnostic.Severity == EmailDiagnosticSeverity.Error
                ? OfficeDocumentDiagnosticSeverity.Error
                : diagnostic.Severity == EmailDiagnosticSeverity.Warning
                    ? OfficeDocumentDiagnosticSeverity.Warning
                    : OfficeDocumentDiagnosticSeverity.Information,
            Category = diagnostic.Code.IndexOf("LIMIT", StringComparison.OrdinalIgnoreCase) >= 0
                ? OfficeDocumentDiagnosticCategory.Limit
                : OfficeDocumentDiagnosticCategory.Adapter,
            Code = diagnostic.Code,
            Message = diagnostic.Message,
            Source = "OfficeIMO.Reader.Email",
            IsRecoverable = diagnostic.Severity != EmailDiagnosticSeverity.Error,
            Location = diagnostic.Location == null ? null : new ReaderLocation { Path = diagnostic.Location }
        };

    private static void Add(StringBuilder text, StringBuilder markdown, string label, string? value) {
        string normalized = Normalize(value);
        if (normalized.Length == 0) return;
        text.Append(label).Append(": ").AppendLine(normalized);
        markdown.Append("- **").Append(label).Append(":** ").AppendLine(EscapeMarkdown(normalized));
    }

    private static void AddValues(StringBuilder text, StringBuilder markdown, string label,
        IReadOnlyList<string> values, int maximum) {
        int count = Math.Min(values.Count, maximum);
        for (int index = 0; index < count; index++) Add(text, markdown, label, values[index]);
        if (values.Count > maximum) Add(text, markdown, label + " values omitted",
            (values.Count - maximum).ToString(CultureInfo.InvariantCulture));
    }

    private static string Normalize(string? value) {
        if (string.IsNullOrWhiteSpace(value)) return string.Empty;
        var builder = new StringBuilder(value!.Length);
        bool pendingSpace = false;
        foreach (char character in value) {
            if (char.IsWhiteSpace(character) || char.IsControl(character)) {
                pendingSpace = builder.Length > 0;
                continue;
            }
            if (pendingSpace) builder.Append(' ');
            pendingSpace = false;
            builder.Append(character);
        }
        return builder.ToString();
    }

    private static string EscapeMarkdown(string value) => value
        .Replace("\\", "\\\\")
        .Replace("*", "\\*")
        .Replace("_", "\\_")
        .Replace("[", "\\[")
        .Replace("]", "\\]")
        .Replace("<", "\\<")
        .Replace(">", "\\>");

    private static string BuildSourceId(string sourceKey) => "src:" + ComputeSha256(
        Path.DirectorySeparatorChar == '\\' ? sourceKey.ToLowerInvariant() : sourceKey);

    private static string ComputeSha256(string value) {
        using (SHA256 sha = SHA256.Create()) {
            byte[] hash = sha.ComputeHash(Encoding.UTF8.GetBytes(value));
            var result = new StringBuilder(hash.Length * 2);
            foreach (byte item in hash) result.Append(item.ToString("x2", CultureInfo.InvariantCulture));
            return result.ToString();
        }
    }
}
