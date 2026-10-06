using OfficeIMO.Workflows;

internal static class AudienceFixtures {
    internal static BookOnixAudienceMetadata? Create(string profile) => profile switch {
        "audience-grades" => new() {
            Categories = [new(BookOnixAudienceType.PrimaryAndSecondaryEducation, true)],
            GradeRanges = [new(BookOnixGradeSystem.UnitedStates, BookOnixGrade.Preschool, BookOnixGrade.Kindergarten),
                new(BookOnixGradeSystem.CanadaExcludingQuebec, BookOnixGrade.Grade9, BookOnixGrade.Grade12),
                new(BookOnixGradeSystem.China, BookOnixGrade.Grade13)],
            Descriptions = [new("Synthetic school and college grade assertions", "eng")] },
        "audience" => new() {
            Categories = Enum.GetValues<BookOnixAudienceType>().Select(type => new BookOnixAudience(type, type == BookOnixAudienceType.Children)).ToArray(),
            AgeRanges = [new(BookOnixAgeRangeType.InterestYears, 8, 12), new(BookOnixAgeRangeType.ReadingYears, 7, 7)],
            Descriptions = [new("Synthetic category and age assertions", "eng"), new("Przykładowi czytelnicy", "pol")] },
        "audience-months" => new() { Categories = [new(BookOnixAudienceType.Children, true)],
            AgeRanges = [new(BookOnixAgeRangeType.InterestMonths, 36, 42)] },
        "audience-open" => new() { AgeRanges = [new(BookOnixAgeRangeType.InterestYears, 12), new(BookOnixAgeRangeType.ReadingYears, null, 10)] },
        _ => null
    };
}
