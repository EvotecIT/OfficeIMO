namespace OfficeIMO.Entities.Tests;

public sealed class InternetEntityTests {
    [Theory]
    [InlineData("jan.kowalski@example.com", "jan.kowalski@example.com", "jan.kowalski@example.com")]
    [InlineData("Jan@Bücher.Example", "Jan@bücher.example", "Jan@xn--bcher-kva.example")]
    [InlineData("\"a b\"@example.com", "\"a b\"@example.com", "\"a b\"@example.com")]
    public void EmailKeepsLocalPartAndIdnaMapsOnlyTheComparisonDomain(string written, string normalized, string key) {
        EntityMatch match = Assert.Single(new EntityFinder().Find("Email: " + written + ".").Matches);
        Assert.Same(InternetSchemes.Email, match.Scheme);
        Assert.Equal(normalized, match.Normalized);
        Assert.Equal(key, match.Scheme.Decode(normalized)!.Properties["key"]);
    }

    [Theory]
    [InlineData("one..two@example.com")]
    [InlineData("user(comment)@example.com")]
    [InlineData("user@-example.com")]
    [InlineData("user@example..com")]
    public void MailboxValidationRejectsInvalidRfcShapes(string written) =>
        Assert.False(InternetSchemes.Email.IsValid(written));

    [Fact]
    public void MailboxValidationAcceptsAddressLiteralsAndAppliesOctetLimits() {
        var email = InternetSchemes.Email;
        Assert.True(email.IsValid("user@[192.0.2.1]"));
        Assert.True(email.IsValid("user@[IPv6:2001:db8::1]"));
        Assert.Empty(new EntityFinder().Find("user@[192.0.2.1]").Matches);
        Assert.True(email.IsValid(new string('é', 32) + "@example.com"));
        Assert.Equal(EntityValidationStatus.InvalidLength, email.Validate(new string('é', 33) + "@example.com").Status);
        Assert.False(email.IsValid("user@" + new string('a', 64) + ".example"));
    }

    [Theory]
    [InlineData("Visit https://example.com/a?b=c.", "https://example.com/a?b=c")]
    [InlineData("(HTTP://Example.COM/A(B)).", "http://example.com/A(B)")]
    [InlineData("See WWW.Example.COM/X!", "www.example.com/X")]
    [InlineData("ftp://LOCALHOST/pub", "ftp://localhost/pub")]
    [InlineData("mailto:Jan@Example.COM?subject=Hi", "mailto:Jan@example.com?subject=Hi")]
    public void UrlKeepsPathAndQueryAndTrimsSentencePunctuation(string text, string expected) {
        EntityMatch match = Assert.Single(new EntityFinder().Find(text).Matches);
        Assert.Same(InternetSchemes.Url, match.Scheme);
        Assert.Equal(expected, match.Normalized);
        Assert.Equal(expected.Length, match.Length);
    }

    [Theory]
    [InlineData("file:///tmp/document")]
    [InlineData("https://intranet/a")]
    [InlineData("https:///example.com")]
    public void UrlValidationRequiresAnAllowedSchemeAndAnExplicitHost(string written) =>
        Assert.False(InternetSchemes.Url.IsValid(written));
}
