using System.Globalization;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Workflows.Tests;

public sealed class BookOnixUsageTests {
    private static readonly XNamespace Ns = BookProject.OnixNamespace;
    private static BookOnixUsageLimit N(BookOnixUsageUnit unit, decimal value) => BookOnixUsageLimit.Number(unit, value);
    private static BookOnixUsageConstraint C(params BookOnixUsageLimit[] limits) => new(BookOnixUsageType.Preview, BookOnixUsageStatus.Limited) { Limits = limits };
    private static BookOnixExportOptions Options(params BookOnixUsageConstraint[] constraints) => BookOnixTests.Options() with {
        CollateralTexts = [new() { Type = BookOnixTextType.Excerpt, Audiences = [BookOnixContentAudience.EndCustomers],
            Texts = [new("Synthetic excerpt")], SourceLinks = ["https://example.org/excerpt"],
            UsageConstraints = constraints, Licenses = [new() { Names = [new("Terms")] }] }]
    };

    [Fact]
    public void ConstraintsPreserveCodesOrderCultureAndPublication() {
        var project = BookOnixTests.Project(); byte[] before = project.ToProjectBytes();
        var baseline = project.ExportOnix(BookOnixTests.Options(), BookOnixTests.TestSchema());
        var culture = CultureInfo.CurrentCulture;
        try {
            CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("pl-PL");
            var result = project.ExportOnix(Options(C(N(BookOnixUsageUnit.PercentagePerPeriod, 12.50m), N(BookOnixUsageUnit.Days, 7)),
                new(BookOnixUsageType.TextAndDataMining, BookOnixUsageStatus.Prohibited)), BookOnixTests.TestSchema());
            var content = XDocument.Load(new MemoryStream(result.Bytes)).Descendants(Ns + "TextContent").Single();
            var constraints = content.Elements(Ns + "EpubUsageConstraint").ToArray();
            Assert.Equal(new[] { "01", "11" }, constraints.Select(e => e.Element(Ns + "EpubUsageType")!.Value));
            Assert.Equal(new[] { "02", "03" }, constraints.Select(e => e.Element(Ns + "EpubUsageStatus")!.Value));
            Assert.Equal(new[] { "12.50", "7" }, constraints[0].Descendants(Ns + "Quantity").Select(e => e.Value));
            Assert.Equal(new[] { "08", "09" }, constraints[0].Descendants(Ns + "EpubUsageUnit").Select(e => e.Value));
            Assert.Equal("TextSourceLink", constraints[0].ElementsBeforeSelf().Last().Name.LocalName);
            Assert.Equal("EpubLicense", constraints[1].ElementsAfterSelf().First().Name.LocalName);
            Assert.Equal(before, project.ToProjectBytes()); Assert.Equal(baseline.Publication.Bytes, result.Publication.Bytes);
            Assert.Equal(result.Bytes, BookOnixMessage.Create([result], BookOnixTests.TestSchema()).Bytes);
        } finally { CultureInfo.CurrentCulture = culture; }
    }

    [Fact]
    public void DateAndTimeFactoriesPreservePrecisionAndLeadingZeros() {
        var result = BookOnixTests.Project().ExportOnix(Options(C(
            BookOnixUsageLimit.Time(BookOnixUsageUnit.StartTime, TimeSpan.FromMilliseconds(1230)),
            BookOnixUsageLimit.Time(BookOnixUsageUnit.EndTime, TimeSpan.FromHours(27) + TimeSpan.FromSeconds(5)),
            BookOnixUsageLimit.Time(BookOnixUsageUnit.MediaDuration, TimeSpan.FromSeconds(30)),
            BookOnixUsageLimit.Date(BookOnixUsageUnit.ValidFrom, new(999, 1, 1)),
            BookOnixUsageLimit.Date(BookOnixUsageUnit.ValidUntil, new(2026, 12, 31)))), BookOnixTests.TestSchema());
        Assert.Equal(new[] { "000000123", "0270005", "0000030", "09990101", "20261231" },
            XDocument.Load(new MemoryStream(result.Bytes)).Descendants(Ns + "Quantity").Select(e => e.Value));
        Assert.Equal("9995959", BookOnixUsageLimit.Time(BookOnixUsageUnit.MediaDuration, TimeSpan.FromHours(1000) - TimeSpan.FromSeconds(1)).Quantity);
    }

    [Fact]
    public void ExplicitZeroAndDatedPermissionStatusesRemainDistinct() {
        var result = BookOnixTests.Project().ExportOnix(Options(
            new(BookOnixUsageType.MultiUserLicense, BookOnixUsageStatus.Limited) { Limits = [N(BookOnixUsageUnit.ConcurrentUsers, 0)] },
            new(BookOnixUsageType.TimeLimitedLicense, BookOnixUsageStatus.Limited) { Limits = [N(BookOnixUsageUnit.Days, 0)] },
            new(BookOnixUsageType.Print, BookOnixUsageStatus.Prohibited) { Limits = [BookOnixUsageLimit.Date(BookOnixUsageUnit.ValidUntil, new(2026, 12, 31))] },
            new(BookOnixUsageType.Print, BookOnixUsageStatus.Unlimited) { Limits = [BookOnixUsageLimit.Date(BookOnixUsageUnit.ValidFrom, new(2027, 1, 1))] }), BookOnixTests.TestSchema());
        Assert.Equal(new[] { "02", "02", "03", "01" }, XDocument.Load(new MemoryStream(result.Bytes)).Descendants(Ns + "EpubUsageStatus").Select(e => e.Value));
    }

    [Theory]
    [InlineData(BookOnixUsageType.NoConstraints, "00")]
    [InlineData(BookOnixUsageType.Preview, "01")]
    [InlineData(BookOnixUsageType.Print, "02")]
    [InlineData(BookOnixUsageType.CopyPaste, "03")]
    [InlineData(BookOnixUsageType.Share, "04")]
    [InlineData(BookOnixUsageType.TextToSpeech, "05")]
    [InlineData(BookOnixUsageType.Lend, "06")]
    [InlineData(BookOnixUsageType.TimeLimitedLicense, "07")]
    [InlineData(BookOnixUsageType.LibraryLoanRenewal, "08")]
    [InlineData(BookOnixUsageType.MultiUserLicense, "09")]
    [InlineData(BookOnixUsageType.PreviewOnPremises, "10")]
    [InlineData(BookOnixUsageType.TextAndDataMining, "11")]
    [InlineData(BookOnixUsageType.PrivatePurchaseAi, "13")]
    [InlineData(BookOnixUsageType.PrivateReadingAi, "14")]
    [InlineData(BookOnixUsageType.LibraryLoan, "16")]
    public void UsageCodesAreExplicit(BookOnixUsageType type, string code) {
        var result = BookOnixTests.Project().ExportOnix(Options(new BookOnixUsageConstraint(type, BookOnixUsageStatus.Unlimited)), BookOnixTests.TestSchema());
        Assert.Equal(code, XDocument.Load(new MemoryStream(result.Bytes)).Descendants(Ns + "EpubUsageType").Single().Value);
    }

    [Theory]
    [InlineData("negative")][InlineData("percentage")][InlineData("fractional-count")][InlineData("page-zero")]
    [InlineData("numeric-time")][InlineData("numeric-date")][InlineData("unknown-unit")]
    [InlineData("time-unit")][InlineData("negative-time")][InlineData("huge-time")][InlineData("subsecond-duration")][InlineData("subcentisecond")][InlineData("date-unit")]
    public void FactoriesRejectAmbiguousOrUnrepresentableValues(string kind) {
        Assert.ThrowsAny<ArgumentException>(() => {
            _ = kind switch {
                "negative" => N(BookOnixUsageUnit.Words, -1), "percentage" => N(BookOnixUsageUnit.Percentage, 101),
                "fractional-count" => N(BookOnixUsageUnit.Users, 1.5m), "page-zero" => N(BookOnixUsageUnit.StartPage, 0),
                "numeric-time" => N(BookOnixUsageUnit.StartTime, 1), "numeric-date" => N(BookOnixUsageUnit.ValidFrom, 20260101),
                "unknown-unit" => N((BookOnixUsageUnit)999, 1),
                "time-unit" => BookOnixUsageLimit.Time(BookOnixUsageUnit.Words, TimeSpan.Zero),
                "negative-time" => BookOnixUsageLimit.Time(BookOnixUsageUnit.StartTime, TimeSpan.FromTicks(-1)),
                "huge-time" => BookOnixUsageLimit.Time(BookOnixUsageUnit.StartTime, TimeSpan.FromHours(1000)),
                "subsecond-duration" => BookOnixUsageLimit.Time(BookOnixUsageUnit.MediaDuration, TimeSpan.FromMilliseconds(10)),
                "subcentisecond" => BookOnixUsageLimit.Time(BookOnixUsageUnit.StartTime, TimeSpan.FromMilliseconds(1)),
                _ => BookOnixUsageLimit.Date(BookOnixUsageUnit.Words, new(2026, 1, 1))
            };
        });
    }

    [Theory]
    [InlineData("missing-limit")][InlineData("unlimited-quantity")][InlineData("prohibited-quantity")][InlineData("duplicate-unit")]
    [InlineData("reversed-pages")][InlineData("reversed-times")][InlineData("reversed-dates")]
    [InlineData("end-page-alone")][InlineData("start-page-alone")][InlineData("start-time-alone")][InlineData("end-time-alone")][InlineData("missing-period")]
    [InlineData("tdm-limited")][InlineData("wrong-license-period")][InlineData("missing-users")]
    [InlineData("no-constraints-mixed")][InlineData("no-constraints-prohibited")]
    [InlineData("unknown-type")][InlineData("unknown-status")][InlineData("null-limits")][InlineData("null-limit")][InlineData("too-many-limits")]
    [InlineData("null-list")][InlineData("null-constraint")][InlineData("too-many-constraints")]
    public void InvalidConstraintsFailWithoutMutation(string kind) {
        var constraint = kind switch {
            "missing-limit" => C(), "unlimited-quantity" => C(N(BookOnixUsageUnit.Words, 1)) with { Status = BookOnixUsageStatus.Unlimited },
            "prohibited-quantity" => C(N(BookOnixUsageUnit.Words, 1)) with { Status = BookOnixUsageStatus.Prohibited },
            "duplicate-unit" => C(N(BookOnixUsageUnit.Words, 1), N(BookOnixUsageUnit.Words, 2)),
            "reversed-pages" => C(N(BookOnixUsageUnit.StartPage, 3), N(BookOnixUsageUnit.EndPage, 2)),
            "reversed-times" => C(BookOnixUsageLimit.Time(BookOnixUsageUnit.StartTime, TimeSpan.FromSeconds(2)), BookOnixUsageLimit.Time(BookOnixUsageUnit.EndTime, TimeSpan.FromSeconds(1))),
            "reversed-dates" => C(BookOnixUsageLimit.Date(BookOnixUsageUnit.ValidFrom, new(2027, 1, 1)), BookOnixUsageLimit.Date(BookOnixUsageUnit.ValidUntil, new(2026, 1, 1))),
            "end-page-alone" => C(N(BookOnixUsageUnit.EndPage, 2)), "start-page-alone" => C(N(BookOnixUsageUnit.StartPage, 1)),
            "start-time-alone" => C(BookOnixUsageLimit.Time(BookOnixUsageUnit.StartTime, TimeSpan.Zero)),
            "end-time-alone" => C(BookOnixUsageLimit.Time(BookOnixUsageUnit.EndTime, TimeSpan.FromSeconds(1))),
            "missing-period" => C(N(BookOnixUsageUnit.PercentagePerPeriod, 10)),
            "tdm-limited" => C(N(BookOnixUsageUnit.Words, 1)) with { Type = BookOnixUsageType.TextAndDataMining },
            "wrong-license-period" => C(N(BookOnixUsageUnit.Words, 1)) with { Type = BookOnixUsageType.TimeLimitedLicense },
            "missing-users" => C(N(BookOnixUsageUnit.Words, 1)) with { Type = BookOnixUsageType.MultiUserLicense },
            "no-constraints-mixed" => new(BookOnixUsageType.NoConstraints, BookOnixUsageStatus.Unlimited),
            "no-constraints-prohibited" => new(BookOnixUsageType.NoConstraints, BookOnixUsageStatus.Prohibited),
            "unknown-type" => C() with { Type = (BookOnixUsageType)999 }, "unknown-status" => C() with { Status = (BookOnixUsageStatus)999 },
            "null-limits" => C() with { Limits = null! }, "null-limit" => C([null!]),
            "too-many-limits" => C(Enumerable.Repeat(N(BookOnixUsageUnit.Words, 1), 33).ToArray()), _ => C(N(BookOnixUsageUnit.Words, 1))
        };
        BookOnixUsageConstraint[] constraints = kind switch {
            "null-list" => null!, "null-constraint" => [null!], "too-many-constraints" => Enumerable.Repeat(constraint, 33).ToArray(),
            "no-constraints-mixed" => [constraint, C(N(BookOnixUsageUnit.Words, 1))], _ => [constraint]
        };
        var project = BookOnixTests.Project(); byte[] before = project.ToProjectBytes();
        Assert.ThrowsAny<ArgumentException>(() => project.ExportOnix(Options(constraints), BookOnixTests.TestSchema()));
        Assert.Equal(before, project.ToProjectBytes());
    }
}
