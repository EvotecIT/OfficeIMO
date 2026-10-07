using OfficeIMO.Workflows;

internal static class SubjectFixtures {
    // Classification assignments are synthetic; schema validity does not establish code membership or suitability.
    internal static BookOnixSubject[] Create() => [
        new() { Scheme = BookOnixSubjectScheme.Dewey, Code = "823.92", IsMain = true },
        new() { Scheme = BookOnixSubjectScheme.LibraryOfCongressClassification, Code = "PR", IsMain = true },
        new() { Scheme = BookOnixSubjectScheme.LibraryOfCongressHeading, Headings = [new("Children's stories", "eng")] },
        new() { Scheme = BookOnixSubjectScheme.Bisac, Code = "JUV000000", IsMain = true, SchemeVersion = "2025" },
        new() { Scheme = BookOnixSubjectScheme.Keywords, Headings = [new("stories; adventure", "eng"), new("opowieści; przygoda", "pol")] },
        new() { Scheme = BookOnixSubjectScheme.Proprietary, SchemeName = "Example & partners", Code = "A1", IsMain = true },
        new() { Scheme = BookOnixSubjectScheme.Thema, Code = "YFB", SchemeVersion = "1.5", IsMain = true,
            Headings = [new("Children's fiction", "eng"), new("Literatura dziecięca", "pol")] },
        new() { Scheme = BookOnixSubjectScheme.ThemaGeographical, Code = "1DBK" },
        new() { Scheme = BookOnixSubjectScheme.ThemaLanguage, Code = "2ACB" },
        new() { Scheme = BookOnixSubjectScheme.ThemaTimePeriod, Code = "3MR" },
        new() { Scheme = BookOnixSubjectScheme.ThemaEducationalPurpose, Code = "4CA" },
        new() { Scheme = BookOnixSubjectScheme.ThemaInterest, Code = "5AJ" },
        new() { Scheme = BookOnixSubjectScheme.ThemaStyle, Code = "6BA" }
    ];
}
