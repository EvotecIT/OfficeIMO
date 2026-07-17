namespace OfficeIMO.Email;

internal static partial class VCardCodec {
    internal static bool TryProject(string text, EmailDocument document, IList<EmailDiagnostic> diagnostics,
        string location) {
        int projectionDiagnosticStart = diagnostics.Count;
        List<VCardProperty> properties;
        try {
            properties = Parse(text.TrimStart('\uFEFF'), diagnostics, location);
        } catch (InvalidDataException exception) {
            diagnostics.Add(new EmailDiagnostic("EMAIL_VCARD_PARSE_INVALID", exception.Message,
                EmailDiagnosticSeverity.Warning, location));
            document.MimeSemanticProjectionIsIncomplete = true;
            return false;
        }
        if (!properties.Any(property => property.Name == "BEGIN" &&
            property.Value.Equals("VCARD", StringComparison.OrdinalIgnoreCase))) return false;
        document.MimeSemanticProjectionIsIncomplete |= diagnostics.Skip(projectionDiagnosticStart).Any(diagnostic =>
            diagnostic.Code == "EMAIL_MIME_QUOTED_PRINTABLE_INVALID" ||
            diagnostic.Code == "EMAIL_MIME_CHARSET_UNSUPPORTED");

        int cardCount = properties.Count(property => property.Name == "BEGIN" &&
            property.Value.Equals("VCARD", StringComparison.OrdinalIgnoreCase));
        VCardProperty[] versions = properties.Where(property => property.Name == "VERSION").ToArray();
        document.MimeSemanticProjectionIsIncomplete |= cardCount > 1 ||
            versions.Length != 1 || !versions[0].Value.Trim().Equals("3.0", StringComparison.OrdinalIgnoreCase) ||
            properties.Any(property => property.Group != null) ||
            HasDuplicateVCardSingletons(properties) ||
            HasUnprojectedExtension(properties) ||
            HasUnpreservedPropertyParameters(properties) ||
            properties.Count(property => property.Name == "EMAIL") > 3 || properties.Any(property =>
            property.Name == "PHOTO" || property.Name == "KEY" || property.Name == "LOGO" ||
            property.Name == "GENDER" || property.Name == "GEO" || property.Name == "TZ" ||
            property.Name == "RELATED" || property.Name == "MEMBER" || property.Name == "AGENT" ||
            property.Name == "SORT-STRING" || property.Name == "MAILER" || property.Name == "UID" ||
            property.Name == "SOURCE" || property.Name == "FBURL" || property.Name == "CALURI" ||
            property.Name == "CALADRURI" || property.Name == "SOUND" || property.Name == "PRODID" ||
            property.Name == "REV" || property.Name == "KIND" &&
            !property.Value.Trim().Equals("individual", StringComparison.OrdinalIgnoreCase) ||
            property.Name == "CLASS" &&
            !ParseVCardSensitivity(property.Value).HasValue);

        var contact = document.Contact ?? new OutlookContact();
        document.Contact = contact;
        document.OutlookItemKind = OutlookItemKind.Contact;
        document.MessageClass = "IPM.Contact";

        string[] name = SplitEscaped(GetValue(properties, "N"), ';');
        contact.Surname = ValueAt(name, 0);
        contact.GivenName = ValueAt(name, 1);
        contact.MiddleName = ValueAt(name, 2);
        contact.Prefix = ValueAt(name, 3);
        contact.Generation = ValueAt(name, 4);
        contact.DisplayName = Unescape(GetValue(properties, "FN"));
        string[] nickNames = SplitEscaped(GetValue(properties, "NICKNAME"), ',');
        contact.NickName = ValueAt(nickNames, 0);
        document.MimeSemanticProjectionIsIncomplete |= nickNames.Skip(1)
            .Any(value => !string.IsNullOrWhiteSpace(value));
        if (string.IsNullOrWhiteSpace(document.Subject)) document.Subject = contact.DisplayName;

        string[] organization = SplitEscaped(GetValue(properties, "ORG"), ';');
        contact.CompanyName = ValueAt(organization, 0);
        contact.Department = ValueAt(organization, 1);
        document.MimeSemanticProjectionIsIncomplete |= organization.Skip(2)
            .Any(value => !string.IsNullOrWhiteSpace(value));
        contact.JobTitle = Unescape(GetValue(properties, "TITLE"));
        contact.Profession = Unescape(GetValue(properties, "ROLE"));
        contact.Language = Unescape(GetValue(properties, "LANG"));
        VCardProperty[] instantMessagingAddresses = properties.Where(property => property.Name == "IMPP").ToArray();
        contact.InstantMessagingAddress = instantMessagingAddresses.Length == 0
            ? null
            : Unescape(instantMessagingAddresses[0].Value);
        document.MimeSemanticProjectionIsIncomplete |= instantMessagingAddresses.Length > 1;
        string? birthday = GetValue(properties, "BDAY");
        string? anniversary = GetValue(properties, "ANNIVERSARY");
        contact.Birthday = ParseDate(birthday);
        contact.WeddingAnniversary = ParseDate(anniversary);
        document.MimeSemanticProjectionIsIncomplete |=
            !string.IsNullOrWhiteSpace(birthday) && !contact.Birthday.HasValue ||
            !string.IsNullOrWhiteSpace(anniversary) && !contact.WeddingAnniversary.HasValue;
        int? sensitivity = ParseVCardSensitivity(GetValue(properties, "CLASS"));
        contact.IsPrivate = sensitivity.HasValue ? sensitivity.Value != 0 : (bool?)null;
        if (sensitivity.HasValue) document.MessageMetadata.Sensitivity = sensitivity;

        document.MimeSemanticProjectionIsIncomplete |= ApplyEmails(properties, contact) ||
            ApplyPhones(properties, contact.Phones) ||
            HasAddressSlotOverflow(properties) || HasUnprojectedAddressComponents(properties) ||
            HasUnsupportedAddressTypes(properties) || HasUrlSlotOverflow(properties) ||
            HasUnsupportedUrlTypes(properties);
        ApplyAddresses(properties, contact);
        ApplyUrls(properties, contact);
        ApplyExtensions(properties, contact);
        foreach (VCardProperty property in properties.Where(property => property.Name == "CATEGORIES")) {
            foreach (string category in SplitEscaped(property.Value, ',')) {
                if (!string.IsNullOrWhiteSpace(category) && !document.MessageMetadata.Categories.Any(existing =>
                    string.Equals(existing, category, StringComparison.OrdinalIgnoreCase))) {
                    document.MessageMetadata.Categories.Add(category);
                }
            }
        }
        string? note = Unescape(GetValue(properties, "NOTE"));
        document.MimeSemanticSourceHasTextBody |= !string.IsNullOrWhiteSpace(note);
        if (!string.IsNullOrWhiteSpace(note)) {
            if (document.Body.Text == null) document.Body.Text = note;
            else if (!string.Equals(document.Body.Text, note, StringComparison.Ordinal)) {
                document.MimeSemanticProjectionIsIncomplete = true;
            }
        }
        return true;
    }

    internal static EmailAttachment? FindSemanticAttachment(EmailDocument document) {
        if (document.OutlookItemKind != OutlookItemKind.Contact) return null;
        return document.Attachments.FirstOrDefault(attachment => attachment.IsProjectedSemanticContent &&
            IsVCardContentType(attachment.ContentType,
                attachment.ContentTypeParameters.TryGetValue("profile", out string? profile) ? profile : null));
    }

    internal static EmailAttachment CreateAttachment(EmailDocument document, EmailAttachment? source = null) {
        if (source != null) return source;
        byte[] content = Create(document);
        var attachment = new EmailAttachment {
            ContentType = "text/vcard",
            Content = content,
            Length = content.LongLength,
            IsProjectedSemanticContent = true,
            IsMimeBodyPart = true
        };
        attachment.ContentTypeParameters["charset"] = "utf-8";
        return attachment;
    }

    internal static bool HasOpaqueContactState(OutlookContact contact) =>
        HasOpaqueAddress(contact.Email1) || HasOpaqueAddress(contact.Email2) || HasOpaqueAddress(contact.Email3);

    private static bool HasOpaqueAddress(OutlookContactEmailAddress address) =>
        address.OriginalEntryId != null && string.IsNullOrWhiteSpace(address.Address) ||
        !string.IsNullOrWhiteSpace(address.Address) && !string.IsNullOrWhiteSpace(address.AddressType) &&
        !string.Equals(address.AddressType, "SMTP", StringComparison.OrdinalIgnoreCase);

    internal static bool IsVCardContentType(string? contentType, string? profile = null) =>
        string.Equals(contentType, "text/vcard", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(contentType, "text/x-vcard", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(contentType, "application/vcard", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(contentType, "text/directory", StringComparison.OrdinalIgnoreCase) &&
        string.Equals(profile, "vcard", StringComparison.OrdinalIgnoreCase);

    private static byte[] Create(EmailDocument document) {
        OutlookContact contact = document.Contact ?? new OutlookContact();
        var output = new StringBuilder();
        AppendLine(output, "BEGIN:VCARD");
        AppendLine(output, "VERSION:3.0");
        AppendLine(output, string.Concat("N:", Escape(contact.Surname), ";", Escape(contact.GivenName), ";",
            Escape(contact.MiddleName), ";", Escape(contact.Prefix), ";", Escape(contact.Generation)));
        string displayName = contact.DisplayName ?? document.Subject ?? string.Join(" ",
            new[] { contact.GivenName, contact.Surname }.Where(value => !string.IsNullOrWhiteSpace(value)));
        AppendText(output, "FN", displayName);
        AppendText(output, "NICKNAME", contact.NickName);
        if (!string.IsNullOrWhiteSpace(contact.CompanyName) || !string.IsNullOrWhiteSpace(contact.Department)) {
            AppendLine(output, string.Concat("ORG:", Escape(contact.CompanyName), ";", Escape(contact.Department)));
        }
        AppendText(output, "TITLE", contact.JobTitle);
        AppendText(output, "ROLE", contact.Profession);
        AppendText(output, "LANG", contact.Language);
        if (contact.Birthday.HasValue) AppendLine(output, string.Concat("BDAY:", FormatDate(contact.Birthday.Value)));
        if (contact.WeddingAnniversary.HasValue) AppendLine(output,
            string.Concat("ANNIVERSARY:", FormatDate(contact.WeddingAnniversary.Value)));
        if (document.MessageMetadata.Sensitivity == 3) AppendLine(output, "CLASS:CONFIDENTIAL");
        else if (document.MessageMetadata.Sensitivity == 1 || document.MessageMetadata.Sensitivity == 2 ||
            contact.IsPrivate == true) AppendLine(output, "CLASS:PRIVATE");
        else if (document.MessageMetadata.Sensitivity == 0 || contact.IsPrivate == false) {
            AppendLine(output, "CLASS:PUBLIC");
        }

        WriteEmail(output, contact.Email1, "WORK", true);
        WriteEmail(output, contact.Email2, "HOME", false);
        WriteEmail(output, contact.Email3, "OTHER", false);
        WritePhones(output, contact.Phones);
        WriteAddress(output, contact.BusinessAddress, "WORK");
        WriteAddress(output, contact.HomeAddress, "HOME");
        WriteAddress(output, contact.OtherAddress, "OTHER");
        WriteAddress(output, contact.WorkAddress, "WORK");
        if (!string.IsNullOrWhiteSpace(contact.BusinessHomePage)) AppendLine(output,
            string.Concat("URL;TYPE=WORK:", EscapeUri(contact.BusinessHomePage!)));
        if (!string.IsNullOrWhiteSpace(contact.PersonalHomePage)) AppendLine(output,
            string.Concat("URL;TYPE=HOME:", EscapeUri(contact.PersonalHomePage!)));
        if (!string.IsNullOrWhiteSpace(contact.InstantMessagingAddress)) AppendLine(output,
            string.Concat("IMPP:", EscapeUri(contact.InstantMessagingAddress!)));

        AppendText(output, "X-MS-MANAGER", contact.ManagerName);
        AppendText(output, "X-MS-ASSISTANT", contact.AssistantName);
        AppendText(output, "X-MS-SPOUSE", contact.SpouseName);
        foreach (string child in contact.Children) AppendText(output, "X-MS-CHILD", child);
        AppendText(output, "X-MS-LOCATION", contact.Location);
        AppendText(output, "X-MS-OFFICE", contact.OfficeLocation);
        AppendText(output, "X-MS-FILE-AS", contact.FileAs);
        AppendText(output, "X-MS-INITIALS", contact.Initials);
        AppendText(output, "X-MS-HTML", contact.Html);
        if (contact.HasPicture.HasValue) AppendLine(output,
            string.Concat("X-OFFICEIMO-HAS-PICTURE:", contact.HasPicture.Value ? "TRUE" : "FALSE"));
        WriteEmailMetadata(output, contact.Email1, 1);
        WriteEmailMetadata(output, contact.Email2, 2);
        WriteEmailMetadata(output, contact.Email3, 3);
        WriteAddressMetadata(output, contact.BusinessAddress, "BUSINESS");
        WriteAddressMetadata(output, contact.HomeAddress, "HOME");
        WriteAddressMetadata(output, contact.OtherAddress, "OTHER");
        WriteAddressMetadata(output, contact.WorkAddress, "WORK");
        if (document.MessageMetadata.Categories.Count > 0) AppendLine(output, string.Concat("CATEGORIES:",
            string.Join(",", document.MessageMetadata.Categories.Where(category => !string.IsNullOrWhiteSpace(category))
                .Select(Escape))));
        AppendText(output, "NOTE", document.Body.Text);
        AppendLine(output, "END:VCARD");
        return Encoding.UTF8.GetBytes(output.ToString());
    }

    private static bool ApplyEmails(IEnumerable<VCardProperty> properties, OutlookContact contact) {
        VCardProperty[] emails = properties.Where(property => property.Name == "EMAIL").ToArray();
        OutlookContactEmailAddress[] targets = { contact.Email1, contact.Email2, contact.Email3 };
        var used = new bool[targets.Length];
        bool incomplete = false;
        foreach (VCardProperty email in emails) {
            string types = email.Parameters.TryGetValue("TYPE", out string? type) ? type : string.Empty;
            int preferredIndex = ContainsType(types, "HOME") ? 1 : ContainsType(types, "OTHER") ? 2 :
                ContainsType(types, "WORK") ? 0 : -1;
            int index = preferredIndex >= 0 && !used[preferredIndex]
                ? preferredIndex
                : Array.FindIndex(used, value => !value);
            if (index < 0) break;
            incomplete |= preferredIndex >= 0 && index != preferredIndex ||
                HasUnsupportedEmailSlot(email, index, types);
            used[index] = true;
            targets[index].Address = Unescape(email.Value);
            targets[index].AddressType = "SMTP";
            string prefix = string.Concat("X-OFFICEIMO-EMAIL",
                (index + 1).ToString(CultureInfo.InvariantCulture));
            targets[index].DisplayName = UnescapeOrNull(GetValue(properties,
                string.Concat(prefix, "-DISPLAY-NAME"))) ?? contact.DisplayName;
            targets[index].OriginalDisplayName = UnescapeOrNull(GetValue(properties,
                string.Concat(prefix, "-ORIGINAL-DISPLAY-NAME")));
        }
        return incomplete;
    }

    private static bool ApplyPhones(IEnumerable<VCardProperty> properties, OutlookContactPhones phones) {
        var counts = new Dictionary<VCardPhoneSlot, int>();
        bool incomplete = false;
        foreach (VCardProperty property in properties.Where(property => property.Name == "TEL")) {
            string types = property.Parameters.TryGetValue("TYPE", out string? type) ? type : string.Empty;
            string value = Unescape(property.Value);
            VCardPhoneSlot slot = GetPhoneSlot(types);
            int count = counts.TryGetValue(slot, out int current) ? current + 1 : 1;
            counts[slot] = count;
            int capacity = slot == VCardPhoneSlot.Home || slot == VCardPhoneSlot.Work ||
                slot == VCardPhoneSlot.General ? 2 : 1;
            incomplete |= count > capacity || HasUnsupportedPhoneSlot(property, slot, count, types);
            switch (slot) {
                case VCardPhoneSlot.Mobile: phones.Mobile = value; break;
                case VCardPhoneSlot.Assistant: phones.Assistant = value; break;
                case VCardPhoneSlot.Company: phones.CompanyMain = value; break;
                case VCardPhoneSlot.Car: phones.Car = value; break;
                case VCardPhoneSlot.Radio: phones.Radio = value; break;
                case VCardPhoneSlot.Callback: phones.Callback = value; break;
                case VCardPhoneSlot.Telex: phones.Telex = value; break;
                case VCardPhoneSlot.Text: phones.TextTelephone = value; break;
                case VCardPhoneSlot.Isdn: phones.Isdn = value; break;
                case VCardPhoneSlot.PrimaryFax: phones.PrimaryFax = value; break;
                case VCardPhoneSlot.HomeFax: phones.HomeFax = value; break;
                case VCardPhoneSlot.BusinessFax: phones.BusinessFax = value; break;
                case VCardPhoneSlot.Home:
                    if (count == 1) phones.Home = value;
                    else phones.Home2 = value;
                    break;
                case VCardPhoneSlot.Pager: phones.Pager = value; break;
                case VCardPhoneSlot.Other: phones.Other = value; break;
                case VCardPhoneSlot.Work:
                    if (count == 1) phones.Business = value;
                    else phones.Business2 = value;
                    break;
                default:
                    if (count == 1) phones.Primary = value;
                    else phones.Other = value;
                    break;
            }
        }
        return incomplete;
    }

    private static VCardPhoneSlot GetPhoneSlot(string types) {
        if (ContainsType(types, "CELL")) return VCardPhoneSlot.Mobile;
        if (ContainsType(types, "X-ASSISTANT")) return VCardPhoneSlot.Assistant;
        if (ContainsType(types, "X-COMPANY")) return VCardPhoneSlot.Company;
        if (ContainsType(types, "CAR")) return VCardPhoneSlot.Car;
        if (ContainsType(types, "X-RADIO")) return VCardPhoneSlot.Radio;
        if (ContainsType(types, "X-CALLBACK")) return VCardPhoneSlot.Callback;
        if (ContainsType(types, "X-TELEX")) return VCardPhoneSlot.Telex;
        if (ContainsType(types, "TEXT")) return VCardPhoneSlot.Text;
        if (ContainsType(types, "ISDN")) return VCardPhoneSlot.Isdn;
        if (ContainsType(types, "FAX") && ContainsType(types, "PREF")) return VCardPhoneSlot.PrimaryFax;
        if (ContainsType(types, "FAX") && ContainsType(types, "HOME")) return VCardPhoneSlot.HomeFax;
        if (ContainsType(types, "FAX")) return VCardPhoneSlot.BusinessFax;
        if (ContainsType(types, "HOME")) return VCardPhoneSlot.Home;
        if (ContainsType(types, "PAGER")) return VCardPhoneSlot.Pager;
        if (ContainsType(types, "OTHER")) return VCardPhoneSlot.Other;
        if (ContainsType(types, "WORK")) return VCardPhoneSlot.Work;
        return VCardPhoneSlot.General;
    }

    private static void ApplyAddresses(IEnumerable<VCardProperty> properties, OutlookContact contact) {
        int workIndex = 0;
        foreach (VCardProperty property in properties.Where(property => property.Name == "ADR")) {
            string types = property.Parameters.TryGetValue("TYPE", out string? type) ? type : string.Empty;
            OutlookPostalAddress target = ContainsType(types, "HOME") ? contact.HomeAddress :
                ContainsType(types, "WORK") && workIndex++ > 0 ? contact.WorkAddress :
                ContainsType(types, "WORK") ? contact.BusinessAddress : contact.OtherAddress;
            string[] values = SplitEscaped(property.Value, ';');
            target.PostOfficeBox = ValueAt(values, 0);
            target.Street = ValueAt(values, 2);
            target.City = ValueAt(values, 3);
            target.StateOrProvince = ValueAt(values, 4);
            target.PostalCode = ValueAt(values, 5);
            target.Country = ValueAt(values, 6);
            if (property.Parameters.TryGetValue("LABEL", out string? label)) target.Formatted = label;
        }
        int labelWorkIndex = 0;
        foreach (VCardProperty property in properties.Where(property => property.Name == "LABEL")) {
            string types = property.Parameters.TryGetValue("TYPE", out string? type) ? type : string.Empty;
            OutlookPostalAddress target = ContainsType(types, "HOME") ? contact.HomeAddress :
                ContainsType(types, "WORK") && labelWorkIndex++ > 0 ? contact.WorkAddress :
                ContainsType(types, "WORK") ? contact.BusinessAddress : contact.OtherAddress;
            target.Formatted = Unescape(property.Value);
        }
        contact.BusinessAddress.CountryCode = Unescape(GetValue(properties, "X-OFFICEIMO-BUSINESS-COUNTRY-CODE"));
        contact.HomeAddress.CountryCode = Unescape(GetValue(properties, "X-OFFICEIMO-HOME-COUNTRY-CODE"));
        contact.OtherAddress.CountryCode = Unescape(GetValue(properties, "X-OFFICEIMO-OTHER-COUNTRY-CODE"));
        contact.WorkAddress.CountryCode = Unescape(GetValue(properties, "X-OFFICEIMO-WORK-COUNTRY-CODE"));
    }

    private static void ApplyUrls(IEnumerable<VCardProperty> properties, OutlookContact contact) {
        foreach (VCardProperty property in properties.Where(property => property.Name == "URL")) {
            string types = property.Parameters.TryGetValue("TYPE", out string? type) ? type : string.Empty;
            if (ContainsType(types, "HOME")) contact.PersonalHomePage = Unescape(property.Value);
            else contact.BusinessHomePage = Unescape(property.Value);
        }
    }

    private static void ApplyExtensions(IEnumerable<VCardProperty> properties, OutlookContact contact) {
        contact.ManagerName = Unescape(GetValue(properties, "X-MS-MANAGER"));
        contact.AssistantName = Unescape(GetValue(properties, "X-MS-ASSISTANT"));
        contact.SpouseName = Unescape(GetValue(properties, "X-MS-SPOUSE"));
        contact.Location = Unescape(GetValue(properties, "X-MS-LOCATION"));
        contact.OfficeLocation = Unescape(GetValue(properties, "X-MS-OFFICE"));
        contact.FileAs = Unescape(GetValue(properties, "X-MS-FILE-AS"));
        contact.Initials = Unescape(GetValue(properties, "X-MS-INITIALS"));
        contact.Html = Unescape(GetValue(properties, "X-MS-HTML"));
        contact.HasPicture = ParseBoolean(GetValue(properties, "X-OFFICEIMO-HAS-PICTURE"));
        foreach (VCardProperty child in properties.Where(property => property.Name == "X-MS-CHILD")) {
            contact.Children.Add(Unescape(child.Value));
        }
    }

    private static void WriteEmail(StringBuilder output, OutlookContactEmailAddress email, string type, bool preferred) {
        if (string.IsNullOrWhiteSpace(email.Address)) return;
        AppendLine(output, string.Concat("EMAIL;TYPE=", type, preferred ? ",PREF:" : ":", Escape(email.Address)));
    }

    private static void WriteEmailMetadata(StringBuilder output, OutlookContactEmailAddress email, int index) {
        string prefix = string.Concat("X-OFFICEIMO-EMAIL", index.ToString(CultureInfo.InvariantCulture));
        AppendText(output, string.Concat(prefix, "-DISPLAY-NAME"), email.DisplayName);
        AppendText(output, string.Concat(prefix, "-ORIGINAL-DISPLAY-NAME"), email.OriginalDisplayName);
    }

    private static void WritePhones(StringBuilder output, OutlookContactPhones phones) {
        WritePhone(output, phones.Business, "WORK,VOICE");
        WritePhone(output, phones.Business2, "WORK,VOICE");
        WritePhone(output, phones.Home, "HOME,VOICE");
        WritePhone(output, phones.Home2, "HOME,VOICE");
        WritePhone(output, phones.Mobile, "CELL,VOICE");
        WritePhone(output, phones.Other, "OTHER,VOICE");
        WritePhone(output, phones.Primary, "PREF,VOICE");
        WritePhone(output, phones.BusinessFax, "WORK,FAX");
        WritePhone(output, phones.HomeFax, "HOME,FAX");
        WritePhone(output, phones.PrimaryFax, "PREF,FAX");
        WritePhone(output, phones.Assistant, "X-ASSISTANT");
        WritePhone(output, phones.CompanyMain, "WORK,X-COMPANY");
        WritePhone(output, phones.Car, "CAR");
        WritePhone(output, phones.Radio, "X-RADIO");
        WritePhone(output, phones.Pager, "PAGER");
        WritePhone(output, phones.Callback, "X-CALLBACK");
        WritePhone(output, phones.Telex, "X-TELEX");
        WritePhone(output, phones.TextTelephone, "TEXT");
        WritePhone(output, phones.Isdn, "ISDN");
    }

    private static void WritePhone(StringBuilder output, string? value, string type) {
        if (!string.IsNullOrWhiteSpace(value)) AppendLine(output,
            string.Concat("TEL;TYPE=", type, ":", Escape(value)));
    }

    private static void WriteAddress(StringBuilder output, OutlookPostalAddress address, string type) {
        if (IsEmpty(address)) return;
        AppendLine(output, string.Concat("ADR;TYPE=", type, ":", Escape(address.PostOfficeBox), ";;",
            Escape(address.Street), ";", Escape(address.City), ";", Escape(address.StateOrProvince), ";",
            Escape(address.PostalCode), ";", Escape(address.Country)));
        if (!string.IsNullOrWhiteSpace(address.Formatted)) AppendText(output, string.Concat("LABEL;TYPE=", type), address.Formatted);
    }

    private static void WriteAddressMetadata(StringBuilder output, OutlookPostalAddress address, string type) =>
        AppendText(output, string.Concat("X-OFFICEIMO-", type, "-COUNTRY-CODE"), address.CountryCode);

    private static bool IsEmpty(OutlookPostalAddress address) => new[] { address.Formatted, address.Street,
        address.City, address.StateOrProvince, address.PostalCode, address.Country, address.PostOfficeBox,
        address.CountryCode }
        .All(string.IsNullOrWhiteSpace);

    private static List<VCardProperty> Parse(string text, IList<EmailDiagnostic> diagnostics, string location) {
        var result = new List<VCardProperty>();
        foreach (ContentLineComponent root in VCardDocument.Parse(text).Cards) {
            result.Add(new VCardProperty("BEGIN", root.Name, null));
            foreach (ContentLineProperty source in root.Properties) {
                var property = new VCardProperty(source.Name.ToUpperInvariant(), source.Value, source.Group);
                foreach (ContentLineParameter parameter in source.Parameters) {
                    string parameterName = parameter.Values.Count == 0 ? "TYPE" : parameter.Name;
                    string parameterValue = parameter.Values.Count == 0
                        ? parameter.Name
                        : string.Join(",", parameter.Values);
                    if (property.Parameters.TryGetValue(parameterName, out string? prior))
                        property.Parameters[parameterName] = string.Concat(prior, ",", parameterValue);
                    else property.Parameters[parameterName] = parameterValue;
                }
                DecodePropertyValue(property, diagnostics, location);
                result.Add(property);
            }
            result.Add(new VCardProperty("END", root.Name, null));
        }
        return result;
    }

    private static void DecodePropertyValue(VCardProperty property, IList<EmailDiagnostic> diagnostics,
        string location) {
        if (!property.Parameters.TryGetValue("ENCODING", out string? encoding) ||
            !(encoding.Equals("QUOTED-PRINTABLE", StringComparison.OrdinalIgnoreCase) ||
              encoding.Equals("QP", StringComparison.OrdinalIgnoreCase))) return;
        byte[] decoded = MimeTextCodec.DecodeQuotedPrintable(Encoding.ASCII.GetBytes(property.Value), false,
            diagnostics, string.Concat(location, "/", property.Name));
        property.Parameters.TryGetValue("CHARSET", out string? charset);
        property.Value = MimeTextCodec.DecodeText(decoded, charset, diagnostics,
            string.Concat(location, "/", property.Name));
    }

    private static string? GetValue(IEnumerable<VCardProperty> properties, string name) =>
        properties.FirstOrDefault(property => property.Name == name)?.Value;

    private static DateTimeOffset? ParseDate(string? value) {
        if (DateTime.TryParseExact(value, new[] { "yyyyMMdd", "yyyy-MM-dd" }, CultureInfo.InvariantCulture,
            DateTimeStyles.None, out DateTime date)) return new DateTimeOffset(date, TimeSpan.Zero);
        return null;
    }

    private static bool? ParseBoolean(string? value) =>
        string.Equals(value, "TRUE", StringComparison.OrdinalIgnoreCase) ? true :
        string.Equals(value, "FALSE", StringComparison.OrdinalIgnoreCase) ? false : (bool?)null;

    private static int? ParseVCardSensitivity(string? value) =>
        string.Equals(value, "PUBLIC", StringComparison.OrdinalIgnoreCase) ? 0 :
        string.Equals(value, "PRIVATE", StringComparison.OrdinalIgnoreCase) ? 2 :
        string.Equals(value, "CONFIDENTIAL", StringComparison.OrdinalIgnoreCase) ? 3 : (int?)null;

    private static string FormatDate(DateTimeOffset value) => value.Date.ToString("yyyyMMdd", CultureInfo.InvariantCulture);

    private static bool ContainsType(string types, string type) => types.Split(',')
        .Any(value => value.Trim().Equals(type, StringComparison.OrdinalIgnoreCase));

    private static string[] SplitEscaped(string? value, char separator) {
        if (value == null) return Array.Empty<string>();
        var result = new List<string>();
        var current = new StringBuilder();
        bool escaped = false;
        foreach (char character in value) {
            if (escaped) {
                current.Append('\\').Append(character);
                escaped = false;
            } else if (character == '\\') escaped = true;
            else if (character == separator) {
                result.Add(Unescape(current.ToString()));
                current.Clear();
            } else current.Append(character);
        }
        if (escaped) current.Append('\\');
        result.Add(Unescape(current.ToString()));
        return result.ToArray();
    }

    private static string? ValueAt(string[] values, int index) => index < values.Length &&
        !string.IsNullOrWhiteSpace(values[index]) ? values[index] : null;

    private static void AppendText(StringBuilder output, string name, string? value) {
        if (!string.IsNullOrWhiteSpace(value)) AppendLine(output, string.Concat(name, ":", Escape(value)));
    }

    private static void AppendLine(StringBuilder output, string line) {
        const int maximumOctets = 75;
        var current = new StringBuilder();
        int octets = 0;
        for (int index = 0; index < line.Length;) {
            int length = char.IsHighSurrogate(line[index]) && index + 1 < line.Length &&
                char.IsLowSurrogate(line[index + 1]) ? 2 : 1;
            string character = line.Substring(index, length);
            int bytes = Encoding.UTF8.GetByteCount(character);
            if (current.Length > 0 && octets + bytes > maximumOctets) {
                output.Append(current).Append("\r\n "); current.Clear(); octets = 1;
            }
            current.Append(character); octets += bytes; index += length;
        }
        output.Append(current).Append("\r\n");
    }

    private static string Escape(string? value) => (value ?? string.Empty).Replace("\\", "\\\\")
        .Replace(";", "\\;").Replace(",", "\\,").Replace("\r\n", "\\n").Replace("\r", "\\n").Replace("\n", "\\n");

    private static string Unescape(string? value) {
        if (string.IsNullOrEmpty(value)) return string.Empty;
        var result = new StringBuilder(value!.Length);
        for (int index = 0; index < value.Length; index++) {
            char character = value[index];
            if (character != '\\' || index + 1 >= value.Length) {
                result.Append(character);
                continue;
            }
            char escaped = value[++index];
            if (escaped == 'n' || escaped == 'N') result.Append('\n');
            else if (escaped == ',' || escaped == ';' || escaped == '\\') result.Append(escaped);
            else result.Append('\\').Append(escaped);
        }
        return result.ToString();
    }

    private static string? UnescapeOrNull(string? value) => string.IsNullOrWhiteSpace(value)
        ? null
        : Unescape(value);

    private static string EscapeUri(string value) => value.Replace("\r", string.Empty).Replace("\n", string.Empty);

    private sealed class VCardProperty {
        internal VCardProperty(string name, string value, string? group) {
            Name = name;
            Value = value;
            Group = group;
        }
        internal string Name { get; }
        internal string Value { get; set; }
        internal string? Group { get; }
        internal IDictionary<string, string> Parameters { get; } =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
    }

    private enum VCardPhoneSlot {
        General,
        Work,
        Home,
        Mobile,
        PrimaryFax,
        HomeFax,
        BusinessFax,
        Assistant,
        Company,
        Car,
        Radio,
        Pager,
        Other,
        Callback,
        Telex,
        Text,
        Isdn
    }
}
