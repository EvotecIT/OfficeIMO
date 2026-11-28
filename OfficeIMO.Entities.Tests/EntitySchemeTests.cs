using System;

namespace OfficeIMO.Entities.Tests;

public sealed class EntitySchemeTests {
    [Fact]
    public void PeselValidatesDateAndDecodesBirthDateAndSex() {
        EntityScheme scheme = PolandSchemes.Pesel;
        // Weighted sum of the ten leading digits is 101, so the check digit is 9.
        EntityValidationResult valid = scheme.Validate("44051401359");
        Assert.Equal(EntityCheckOutcome.Passed, valid.CheckDigit);
        Assert.Equal(EntityCheckProvenance.IssuerPublished, valid.CheckProvenance);
        EntityDetails details = Assert.IsType<EntityDetails>(scheme.Decode(valid.Normalized!));
        Assert.Equal(new DateTime(1944, 5, 14), details.BirthDate);
        Assert.Equal(EntitySex.Male, details.Sex);
        EntityValidationResult invalid = scheme.Validate("44051401358");
        Assert.Equal(EntityValidationStatus.InvalidCheckDigit, invalid.Status);
        Assert.Equal("44051401358", invalid.Normalized);
        Assert.DoesNotContain(invalid.Normalized!, invalid.ToString());
        Assert.Equal(EntityValidationStatus.InvalidDate, scheme.Validate("44023101359").Status);
        Assert.Equal(EntityValidationStatus.UnsupportedVariant, scheme.Validate("44451401359").Status);
        Assert.Equal(EntityValidationStatus.UnsupportedVariant, scheme.Validate("44651401359").Status);
    }

    [Theory]
    [InlineData("PL.REGON", "123456785", "123456784")] // Sum 192 mod 11 = 5.
    [InlineData("PL.REGON", "12345678500010", "12345678500011")] // Weighted sum 219 mod 11 = 10, represented as 0.
    [InlineData("PL.IDCARD", "ABA300000", "ABA400000")] // 7×10 + 3×11 + 10 = 113.
    [InlineData("PL.PASSPORT", "ZS4001107", "ZS5001107")] // Ministry's 2006 specimen uses the third-position check.
    public void PolishCommunityChecksReportProvenanceOnPassAndFail(string id, string value, string badValue) {
        EntityScheme scheme = EntityCatalog.Default.GetScheme(id);
        EntityValidationResult valid = scheme.Validate(value);
        Assert.Equal(EntityValidationStatus.Valid, valid.Status);
        Assert.Equal(EntityCheckOutcome.Passed, valid.CheckDigit);
        Assert.Equal(EntityCheckProvenance.CommunityDocumented, valid.CheckProvenance);
        EntityValidationResult invalid = scheme.Validate(badValue);
        Assert.Equal(EntityValidationStatus.InvalidCheckDigit, invalid.Status);
        Assert.Equal(EntityCheckProvenance.CommunityDocumented, invalid.CheckProvenance);
        Assert.NotNull(scheme.Specification);
    }

    [Fact]
    public void OlderPassportSeriesRetainsFormatOnlyValidation() {
        EntityValidationResult result = PolandSchemes.Passport.Validate("BM1234567");
        Assert.True(result.IsValid);
        Assert.Equal(EntityCheckOutcome.NotDefined, result.CheckDigit);
        Assert.Equal(EntityCheckProvenance.None, result.CheckProvenance);
    }

    [Theory]
    [InlineData("pl 526-000-12-46", "5260001246", "526-000-12-46")] // BMF example.
    [InlineData("123-45-63-218", "1234563218", "123-456-32-18")] // Weighted sum 118 mod 11 = 8.
    public void NipAcceptsBothPrintGroupingsAndCanonicalizesOptionalPrefix(string written, string expected, string display) {
        EntityScheme nip = PolandSchemes.Nip;
        Assert.Equal(expected, nip.Normalize(written));
        Assert.Equal(display, nip.Format(expected));
        Assert.Equal(expected, nip.Format(expected, EntityFormatStyle.Compact));
    }

    [Theory]
    [InlineData("GB.NINO", "AB 12 34 56 A", "AB123456A")]
    [InlineData("GB.NINO", "AB123456", "AB123456")]
    [InlineData("GB.COMPANY", "02050399", "02050399")] // Companies House URI guide.
    [InlineData("GB.COMPANY", "sc002180", "SC002180")] // Companies House URI guide.
    [InlineData("US.SSN", "123-45-6789", "123456789")]
    [InlineData("US.EIN", "12-3456789", "123456789")]
    [InlineData("BIC", "chasus33xxx", "CHASUS33XXX")]
    public void FormatOnlySchemesExposeCanonicalValuesWithoutCheck(string id, string written, string expected) {
        EntityValidationResult result = EntityCatalog.Default.Validate(id, written);
        Assert.Equal(EntityValidationStatus.Valid, result.Status);
        Assert.Equal(EntityCheckOutcome.NotDefined, result.CheckDigit);
        Assert.Equal(expected, result.Normalized);
    }

    [Theory]
    [InlineData("GB.NINO", "QQ 12 34 56 A")] // HMRC's deliberate specimen.
    [InlineData("GB.NINO", "BG123456A")]
    [InlineData("US.SSN", "666-45-6789")]
    [InlineData("US.SSN", "999-45-6789")]
    [InlineData("US.SSN", "123-00-6789")]
    [InlineData("US.SSN", "123-45-0000")]
    [InlineData("US.EIN", "00-3456789")]
    public void LocalAllocationRulesRejectReservedShapes(string id, string written) {
        Assert.Equal(EntityValidationStatus.NotAllocated, EntityCatalog.Default.Validate(id, written).Status);
    }

    [Theory]
    [InlineData("PL61 1090 1014 0000 0712 1981 2874", "PL61109010140000071219812874")]
    [InlineData("GB82 WEST 1234 5698 7654 32", "GB82WEST12345698765432")]
    [InlineData("DE89 3704 0044 0532 0130 00", "DE89370400440532013000")]
    public void RegisteredIbanExamplesRoundTripInPrintAndReturnCountryAndBban(string written, string expected) {
        EntityScheme iban = FinancialSchemes.Iban;
        EntityValidationResult result = iban.Validate(written);
        Assert.Equal(expected, result.Normalized);
        Assert.Equal(EntityCheckOutcome.Passed, result.CheckDigit);
        Assert.Equal(EntityCheckProvenance.IssuerPublished, result.CheckProvenance);
        Assert.Equal(written, iban.Format(expected));
        EntityDetails details = Assert.IsType<EntityDetails>(iban.Decode(expected));
        Assert.Equal(expected.Substring(0, 2), details.Country);
        Assert.Equal(expected.Substring(4), details.Properties["bban"]);
    }

    [Theory]
    [InlineData("ISIN", "US0378331005", "US0378331006")]
    [InlineData("LEI", "5493001KJTIIGC8Y1R12", "5493001KJTIIGC8Y1R13")]
    [InlineData("CARD", "4111 1111 1111 1111", "4111 1111 1111 1112")]
    [InlineData("IBAN", "DE89 3704 0044 0532 0130 00", "DE88 3704 0044 0532 0130 00")]
    public void FinancialChecksRejectAnOtherwiseIdenticalWrongCheck(string id, string value, string change) {
        EntityScheme scheme = EntityCatalog.Default.GetScheme(id);
        Assert.Equal(EntityCheckOutcome.Passed, scheme.Validate(value).CheckDigit);
        Assert.Equal(EntityValidationStatus.InvalidCheckDigit, scheme.Validate(change).Status);
    }
}
