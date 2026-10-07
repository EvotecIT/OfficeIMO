namespace OfficeIMO.Workflows;

/// <summary>ONIX list 145 usage types.</summary>
public enum BookOnixUsageType {
    /// <summary>No particular constraints; use unlimited status alone (00).</summary>
    NoConstraints = 0,
    /// <summary>Preview before purchase (01).</summary>
    Preview = 1,
    /// <summary>Physical copies of extracts (02).</summary>
    Print = 2,
    /// <summary>Digital copies of extracts (03).</summary>
    CopyPaste = 3,
    /// <summary>Use across concurrent devices (04).</summary>
    Share = 4,
    /// <summary>Text-to-speech reading (05).</summary>
    TextToSpeech = 5,
    /// <summary>Lending to another account (06).</summary>
    Lend = 6,
    /// <summary>Time-limited license (07).</summary>
    TimeLimitedLicense = 7,
    /// <summary>Consecutive loans or extensions (08).</summary>
    LibraryLoanRenewal = 8,
    /// <summary>Concurrent-user license (09).</summary>
    MultiUserLicense = 9,
    /// <summary>Preview at the retail or library premises (10).</summary>
    PreviewOnPremises = 10,
    /// <summary>Text and data mining (11).</summary>
    TextAndDataMining = 11,
    /// <summary>AI-enhanced private purchase experience; list issue 73 or later (13).</summary>
    PrivatePurchaseAi = 13,
    /// <summary>AI-enhanced private reading experience; list issue 73 or later (14).</summary>
    PrivateReadingAi = 14,
    /// <summary>Library lending (16).</summary>
    LibraryLoan = 16,
}

/// <summary>ONIX list 146 permission statuses.</summary>
public enum BookOnixUsageStatus {
    /// <summary>Permitted without a quantity limit (01).</summary>
    Unlimited = 1,
    /// <summary>Permitted subject to an explicit limit (02).</summary>
    Limited = 2,
    /// <summary>Prohibited (03).</summary>
    Prohibited = 3
}

/// <summary>ONIX list 147 quantity units. Date and media-time units require their typed factories.</summary>
public enum BookOnixUsageUnit {
    /// <summary>Maximum extract copies (01).</summary>
    Copies = 1,
    /// <summary>Maximum extract characters (02).</summary>
    Characters = 2,
    /// <summary>Maximum extract words (03).</summary>
    Words = 3,
    /// <summary>Maximum extract pages (04).</summary>
    Pages = 4,
    /// <summary>Maximum percentage of content (05).</summary>
    Percentage = 5,
    /// <summary>Maximum devices in a sharing group (06).</summary>
    Devices = 6,
    /// <summary>Concurrent users; zero explicitly means unlimited (07).</summary>
    ConcurrentUsers = 7,
    /// <summary>Content percentage per separately supplied period (08).</summary>
    PercentagePerPeriod = 8,
    /// <summary>Days from purchase or activation (09).</summary>
    Days = 9,
    /// <summary>Maximum lifetime usage events (10).</summary>
    Times = 10,
    /// <summary>First allowed absolute page, counting the cover as page one; requires fixed pagination (11).</summary>
    StartPage = 11,
    /// <summary>Last allowed absolute page; requires an explicit start (12).</summary>
    EndPage = 12,
    /// <summary>Time period in weeks (13).</summary>
    Weeks = 13,
    /// <summary>Time period in months (14).</summary>
    Months = 14,
    /// <summary>Licensed individuals, independent of concurrency (15).</summary>
    Users = 15,
    /// <summary>Maximum extract duration in whole seconds (16).</summary>
    MediaDuration = 16,
    /// <summary>First allowed media position, with centisecond precision (17).</summary>
    StartTime = 17,
    /// <summary>Last allowed media position; requires an explicit start (18).</summary>
    EndTime = 18,
    /// <summary>Maximum concurrent learner groups (19).</summary>
    ConcurrentClasses = 19,
    /// <summary>Maximum learner groups, independent of concurrency (20).</summary>
    Classes = 20,
    /// <summary>Maximum extract resolution in dots per inch (21).</summary>
    DotsPerInch = 21,
    /// <summary>Maximum usage events per day (22).</summary>
    TimesPerDay = 22,
    /// <summary>Maximum usage events per month (23).</summary>
    TimesPerMonth = 23,
    /// <summary>Maximum usage events per year (24).</summary>
    TimesPerYear = 24,
    /// <summary>Maximum extract resolution in dots per centimeter (26).</summary>
    DotsPerCentimeter = 26,
    /// <summary>Days beginning at publication, independent of purchase (27).</summary>
    DaysFromPublication = 27,
    /// <summary>Weeks beginning at publication (28).</summary>
    WeeksFromPublication = 28,
    /// <summary>Months beginning at publication (29).</summary>
    MonthsFromPublication = 29,
    /// <summary>Maximum licensed institutions (31).</summary>
    Institutions = 31,
    /// <summary>Calendar date when the constraint begins (98).</summary>
    ValidFrom = 98,
    /// <summary>Calendar date when the constraint ends (99).</summary>
    ValidUntil = 99,
}

/// <summary>Explicit usage assertion; serialization does not enforce permissions or interpret a license.</summary>
public sealed record BookOnixUsageConstraint(BookOnixUsageType Type, BookOnixUsageStatus Status) {
    /// <summary>Up to 32 limits with distinct units. Date-only limits may accompany any status.</summary>
    public IReadOnlyList<BookOnixUsageLimit> Limits { get; init; } = [];
}
