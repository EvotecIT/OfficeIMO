namespace OfficeIMO.Bibliography;

internal static class CodecMappings {
    internal static BibliographyItemType ParseType(string? type) {
        switch ((type ?? string.Empty).Trim().ToLowerInvariant()) {
            case "article": case "article-journal": case "jour": case "journal article": return BibliographyItemType.ArticleJournal;
            case "article-magazine": case "mgzn": case "magazine article": return BibliographyItemType.ArticleMagazine;
            case "article-newspaper": case "news": case "newspaper article": return BibliographyItemType.ArticleNewspaper;
            case "book": return BibliographyItemType.Book;
            case "inbook": case "incollection": case "chapter": case "chap": case "book section": return BibliographyItemType.Chapter;
            case "inproceedings": case "conference": case "paper-conference": case "conf": case "conference paper": return BibliographyItemType.PaperConference;
            case "proceedings": return BibliographyItemType.Proceedings;
            case "report": case "techreport": case "rprt": return BibliographyItemType.Report;
            case "phdthesis": case "mastersthesis": case "thesis": case "thes": return BibliographyItemType.Thesis;
            case "webpage": case "web": case "web page": return BibliographyItemType.WebPage;
            case "dataset": case "data": return BibliographyItemType.Dataset;
            case "software": case "computer program": return BibliographyItemType.Software;
            case "patent": case "pat": return BibliographyItemType.Patent;
            case "legal_case": case "case": return BibliographyItemType.LegalCase;
            case "unpublished": case "manuscript": return BibliographyItemType.Manuscript;
            case "personal_communication": case "pcomm": return BibliographyItemType.PersonalCommunication;
            case "document": case "generic": case "gen": return BibliographyItemType.Document;
            default: return BibliographyItemType.Unknown;
        }
    }

    internal static string ToCslType(BibliographyItemType type) {
        switch (type) {
            case BibliographyItemType.ArticleJournal: return "article-journal";
            case BibliographyItemType.ArticleMagazine: return "article-magazine";
            case BibliographyItemType.ArticleNewspaper: return "article-newspaper";
            case BibliographyItemType.Book: return "book";
            case BibliographyItemType.Chapter: return "chapter";
            case BibliographyItemType.PaperConference: return "paper-conference";
            case BibliographyItemType.Proceedings: return "book";
            case BibliographyItemType.Report: return "report";
            case BibliographyItemType.Thesis: return "thesis";
            case BibliographyItemType.WebPage: return "webpage";
            case BibliographyItemType.Dataset: return "dataset";
            case BibliographyItemType.Software: return "software";
            case BibliographyItemType.Patent: return "patent";
            case BibliographyItemType.LegalCase: return "legal_case";
            case BibliographyItemType.PersonalCommunication: return "personal_communication";
            case BibliographyItemType.Manuscript: return "manuscript";
            default: return "document";
        }
    }

    internal static string ToBibType(BibliographyItemType type) {
        switch (type) {
            case BibliographyItemType.ArticleJournal: case BibliographyItemType.ArticleMagazine: case BibliographyItemType.ArticleNewspaper: return "article";
            case BibliographyItemType.Book: return "book";
            case BibliographyItemType.Chapter: return "incollection";
            case BibliographyItemType.PaperConference: return "inproceedings";
            case BibliographyItemType.Proceedings: return "proceedings";
            case BibliographyItemType.Report: return "techreport";
            case BibliographyItemType.Thesis: return "phdthesis";
            case BibliographyItemType.Manuscript: return "unpublished";
            default: return "misc";
        }
    }

    internal static string ToRisType(BibliographyItemType type) {
        switch (type) {
            case BibliographyItemType.ArticleJournal: return "JOUR";
            case BibliographyItemType.ArticleMagazine: return "MGZN";
            case BibliographyItemType.ArticleNewspaper: return "NEWS";
            case BibliographyItemType.Book: return "BOOK";
            case BibliographyItemType.Chapter: return "CHAP";
            case BibliographyItemType.PaperConference: return "CONF";
            case BibliographyItemType.Report: return "RPRT";
            case BibliographyItemType.Thesis: return "THES";
            case BibliographyItemType.WebPage: return "WEB";
            case BibliographyItemType.Dataset: return "DATA";
            case BibliographyItemType.Software: return "COMP";
            case BibliographyItemType.Patent: return "PAT";
            case BibliographyItemType.PersonalCommunication: return "PCOMM";
            default: return "GEN";
        }
    }

    internal static BibliographyName ParseCommaName(string value) {
        string trimmed = value.Trim();
        if (trimmed.Length >= 2 && trimmed[0] == '{' && trimmed[trimmed.Length - 1] == '}') return new BibliographyName { Literal = trimmed.Substring(1, trimmed.Length - 2) };
        string[] parts = value.Split(new[] { ',' }, 3);
        if (parts.Length == 1) {
            string[] words = value.Trim().Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            if (words.Length <= 1) return new BibliographyName { Literal = value.Trim() };
            return new BibliographyName { Given = string.Join(" ", words.Take(words.Length - 1)), Family = words[words.Length - 1] };
        }
        return new BibliographyName { Family = parts[0].Trim(), Given = parts[1].Trim(), Suffix = parts.Length > 2 ? parts[2].Trim() : null };
    }

    internal static string FormatName(BibliographyName name) {
        if (!string.IsNullOrWhiteSpace(name.Literal)) return name.Literal!;
        string family = string.Join(" ", new[] { name.NonDroppingParticle, name.Family }.Where(static part => !string.IsNullOrWhiteSpace(part)));
        string given = string.Join(" ", new[] { name.Given, name.DroppingParticle }.Where(static part => !string.IsNullOrWhiteSpace(part)));
        return string.Join(", ", new[] { family, given, name.Suffix }.Where(static part => !string.IsNullOrWhiteSpace(part)));
    }

    internal static void AddIdentifier(BibliographyItem item, string scheme, string? value) {
        if (!string.IsNullOrWhiteSpace(value)) item.Identifiers.Add(new BibliographyIdentifier(scheme, value!));
    }

    internal static string InferSerialScheme(string value) {
        string compact = new string(value.Where(char.IsLetterOrDigit).ToArray());
        if (compact.Length == 8) return "ISSN";
        if (compact.Length == 10 || compact.Length == 13) return "ISBN";
        return "SN";
    }

    internal static BibliographyDate ParseDate(BibliographyDateRole role, string value) {
        string literal = value.Trim();
        var date = new BibliographyDate { Role = role };
        int rangeSeparator = literal.IndexOf('/');
        if (rangeSeparator > 0 && rangeSeparator + 1 < literal.Length && LooksLikeRangeEnd(literal.Substring(rangeSeparator + 1))) {
            bool validStart = TryParseDatePart(literal.Substring(0, rangeSeparator), out int? year, out int? month, out int? day);
            bool validEnd = TryParseDatePart(literal.Substring(rangeSeparator + 1), out int? endYear, out int? endMonth, out int? endDay);
            if (validStart && validEnd) {
                date.Year = year; date.Month = month; date.Day = day;
                date.EndYear = endYear; date.EndMonth = endMonth; date.EndDay = endDay;
            } else date.Literal = literal;
            return date;
        }
        if (TryParseDatePart(literal, out int? singleYear, out int? singleMonth, out int? singleDay)) {
            date.Year = singleYear; date.Month = singleMonth; date.Day = singleDay;
        } else date.Literal = literal;
        return date;
    }

    internal static string FormatDate(BibliographyDate date) {
        string start = FormatDatePart(date.Year, date.Month, date.Day);
        if (start.Length == 0) return date.Literal ?? string.Empty;
        string end = FormatDatePart(date.EndYear, date.EndMonth, date.EndDay);
        return end.Length == 0 ? start : start + "/" + end;
    }

    internal static string OutputKey(BibliographyItem item, int zeroBasedIndex) =>
        string.IsNullOrWhiteSpace(item.Key) ? "item-" + (zeroBasedIndex + 1).ToString(CultureInfo.InvariantCulture) : item.Key;

    internal static int? ParseMonth(string value) {
        if (int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int numeric) && numeric >= 1 && numeric <= 12) return numeric;
        string candidate = value.Trim().TrimEnd('.');
        for (int month = 1; month <= 12; month++) {
            DateTimeFormatInfo format = CultureInfo.InvariantCulture.DateTimeFormat;
            if (string.Equals(candidate, format.GetMonthName(month), StringComparison.OrdinalIgnoreCase) || string.Equals(candidate, format.GetAbbreviatedMonthName(month).TrimEnd('.'), StringComparison.OrdinalIgnoreCase)) return month;
        }
        return null;
    }

    private static bool LooksLikeRangeEnd(string value) {
        string trimmed = value.TrimStart();
        int digitCount = 0;
        while (digitCount < trimmed.Length && char.IsDigit(trimmed[digitCount])) digitCount++;
        return digitCount >= 4;
    }

    private static bool TryParseDatePart(string value, out int? year, out int? month, out int? day) {
        year = null; month = null; day = null;
        string trimmed = value.Trim();
        string[] pieces = trimmed.Split(new[] { '-', '/', ' ', ',' }, StringSplitOptions.RemoveEmptyEntries).Select(static piece => piece.Trim()).ToArray();
        if (pieces.Length == 0 || pieces.Length > 3) return false;
        if (int.TryParse(pieces[0], NumberStyles.Integer, CultureInfo.InvariantCulture, out int leadingYear) && leadingYear >= 1 && pieces[0].Length >= 4) {
            year = leadingYear;
            if (pieces.Length > 1) month = ParseMonth(pieces[1]);
            if (pieces.Length > 1 && !month.HasValue) return false;
            if (pieces.Length > 2 && (!int.TryParse(pieces[2], NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsedDay) || parsedDay < 1 || parsedDay > 31)) return false;
            if (pieces.Length > 2) day = int.Parse(pieces[2], CultureInfo.InvariantCulture);
            return true;
        }
        if (pieces.Length >= 2 && int.TryParse(pieces[pieces.Length - 1], NumberStyles.Integer, CultureInfo.InvariantCulture, out int trailingYear) && trailingYear >= 1 && pieces[pieces.Length - 1].Length >= 4) {
            year = trailingYear;
            month = ParseMonth(pieces[0]);
            if (!month.HasValue) return false;
            if (pieces.Length == 3 && (!int.TryParse(pieces[1], NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsedDay) || parsedDay < 1 || parsedDay > 31)) return false;
            if (pieces.Length == 3) day = int.Parse(pieces[1], CultureInfo.InvariantCulture);
            return true;
        }
        return false;
    }

    private static string FormatDatePart(int? year, int? month, int? day) {
        if (!year.HasValue) return string.Empty;
        var builder = new StringBuilder(year.Value.ToString("0000", CultureInfo.InvariantCulture));
        if (month.HasValue) builder.Append('-').Append(month.Value.ToString("00", CultureInfo.InvariantCulture));
        if (day.HasValue) builder.Append('-').Append(day.Value.ToString("00", CultureInfo.InvariantCulture));
        return builder.ToString();
    }
}
