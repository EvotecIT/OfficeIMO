using OfficeIMO.Email;

namespace OfficeIMO.Email.AddressBook.Tests;

public sealed class OfflineAddressBookIdentityIndexTests {
    [Fact]
    public void Resolves_primary_alias_legacy_and_account_without_display_name_guessing() {
        using var stream = new MemoryStream(new OabV4Fixture().Build());
        using OfflineAddressBookSession session = OfflineAddressBookSession.Open(stream, "identity.oab");
        OfflineAddressBookIdentityIndex index = session.BuildIdentityIndex();

        OfflineAddressBookIdentityResolution primary = index.Resolve("ada@example.test");
        OfflineAddressBookIdentityResolution alias = index.Resolve("SMTP:alias-ada@example.test");
        OfflineAddressBookIdentityResolution legacy = index.Resolve(new EmailAddress(
            "/o=Example/ou=Recipients/cn=ada", "Ada") { AddressType = "EX" });
        OfflineAddressBookIdentityResolution account = index.Resolve("ada");
        OfflineAddressBookIdentityResolution display = index.Resolve("Ada Lovelace",
            options: new OfflineAddressBookIdentityResolutionOptions(allowDisplayNameMatch: true));

        Assert.True(index.IsComplete);
        Assert.Equal(3, index.EntriesScanned);
        Assert.Equal(OfflineAddressBookIdentityResolutionStatus.Resolved, primary.Status);
        Assert.Equal(OfflineAddressBookIdentityMatchSource.PrimarySmtpAddress,
            primary.Candidate!.MatchSource);
        Assert.Equal("ada@example.test", primary.Candidate.ToEmailAddress().Address);
        Assert.Equal(OfflineAddressBookIdentityMatchSource.ProxyAddress, alias.Candidate!.MatchSource);
        Assert.True(alias.Candidate.IsAuthoritativeAddress);
        Assert.Equal("ada@example.test", legacy.Candidate!.PrimarySmtpAddress);
        Assert.Equal(OfflineAddressBookIdentityMatchSource.AccountName, account.Candidate!.MatchSource);
        Assert.False(account.Candidate.IsAuthoritativeAddress);
        Assert.Equal(OfflineAddressBookIdentityResolutionStatus.NotFound, display.Status);
    }

    [Fact]
    public void Display_names_are_opt_in_at_build_and_query_time() {
        using var stream = new MemoryStream(new OabV4Fixture().Build());
        using OfflineAddressBookSession session = OfflineAddressBookSession.Open(stream, "display.oab");
        OfflineAddressBookIdentityIndex index = session.BuildIdentityIndex(
            new OfflineAddressBookIdentityIndexOptions(includeDisplayNames: true));

        Assert.Equal(OfflineAddressBookIdentityResolutionStatus.NotFound,
            index.Resolve("Ada Lovelace").Status);
        OfflineAddressBookIdentityResolution resolved = index.Resolve("Ada Lovelace",
            options: new OfflineAddressBookIdentityResolutionOptions(allowDisplayNameMatch: true));

        Assert.Equal(OfflineAddressBookIdentityResolutionStatus.Resolved, resolved.Status);
        Assert.Equal(OfflineAddressBookIdentityMatchSource.DisplayName, resolved.Candidate!.MatchSource);
        Assert.False(resolved.Candidate.IsAuthoritativeAddress);
    }

    [Fact]
    public void Duplicate_and_bounded_indexes_report_ambiguity_and_incomplete_not_false_not_found() {
        var fixture = new OabV4Fixture()
            .AddPerson("Ada Duplicate", "ada@example.test", "ada-2", "Ada", "Duplicate", "Research");
        using var stream = new MemoryStream(fixture.Build());
        using OfflineAddressBookSession session = OfflineAddressBookSession.Open(stream, "duplicates.oab");

        OfflineAddressBookIdentityResolution ambiguous = session.BuildIdentityIndex()
            .Resolve("ada@example.test");
        OfflineAddressBookIdentityIndex bounded = session.BuildIdentityIndex(
            new OfflineAddressBookIdentityIndexOptions(maxEntries: 1));
        OfflineAddressBookIdentityResolution missing = bounded.Resolve("grace@example.test");

        Assert.Equal(OfflineAddressBookIdentityResolutionStatus.Ambiguous, ambiguous.Status);
        Assert.Equal(2, ambiguous.Candidates.Count);
        Assert.False(bounded.IsComplete);
        Assert.Equal(OfflineAddressBookIdentityResolutionStatus.Incomplete, missing.Status);
        Assert.Contains(bounded.Diagnostics, diagnostic =>
            diagnostic.Code == "OAB_IDENTITY_INDEX_ENTRY_LIMIT");
    }

    [Fact]
    public void Identity_count_counts_distinct_keys_not_duplicate_candidates() {
        using var baselineStream = new MemoryStream(new OabV4Fixture().Build());
        using OfflineAddressBookSession baselineSession =
            OfflineAddressBookSession.Open(baselineStream, "baseline.oab");
        int baselineIdentityCount = baselineSession.BuildIdentityIndex().IdentityCount;

        var duplicateFixture = new OabV4Fixture()
            .AddPerson("Ada Duplicate", "ada@example.test", "ada", "Ada", "Duplicate", "Research");
        using var duplicateStream = new MemoryStream(duplicateFixture.Build());
        using OfflineAddressBookSession duplicateSession =
            OfflineAddressBookSession.Open(duplicateStream, "duplicate-identities.oab");
        OfflineAddressBookIdentityIndex duplicateIndex = duplicateSession.BuildIdentityIndex();

        Assert.Equal(baselineIdentityCount, duplicateIndex.IdentityCount);
        Assert.Equal(OfflineAddressBookIdentityResolutionStatus.Ambiguous,
            duplicateIndex.Resolve("ada@example.test").Status);
    }
}
