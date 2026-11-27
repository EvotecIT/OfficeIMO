namespace OfficeIMO.Entities.Tests;

public sealed class EuVatTests {
    // Authority examples from the BMF November 2020 VAT construction rules.
    [Theory]
    [InlineData("ATU10223006")]
    [InlineData("BE0776091951")]
    [InlineData("DE111111125")]
    [InlineData("DK88146328")]
    [InlineData("EE100207415")]
    [InlineData("FI09853608")]
    [InlineData("HU21376414")]
    [InlineData("IE8Z49289F")]
    [InlineData("IE3628739L")]
    [InlineData("IE3628739UA")]
    [InlineData("IT00000010215")]
    [InlineData("LT213179412")]
    [InlineData("LT290061371314")]
    [InlineData("LU10000356")]
    [InlineData("NL010000446B01")]
    [InlineData("NL123456789B13")]
    [InlineData("PL5260001246")]
    [InlineData("PT502757191")]
    [InlineData("SE556188840401")]
    [InlineData("SI15012557")]
    [InlineData("SK4030000007")]
    [InlineData("XI434031494")]
    [InlineData("HR12345678903")] // Croatian OIB worked example in the register check guide.
    [InlineData("ESK1234567L")] // The AEAT note applies mod 23 to the seven digits, with remainder 19.
    [InlineData("ES12345678Z")] // Spanish Ministry of the Interior, number mod 23 = 14 (Z).
    public void PublishedVatChecksAcceptTheirWorkedExample(string value) {
        EntityScheme scheme = EuVatSchemes.Get(value.Substring(0, 2));
        EntityValidationResult result = scheme.Validate(value);
        Assert.Equal(EntityValidationStatus.Valid, result.Status);
        Assert.Equal(value, result.Normalized);
        Assert.Equal(EntityCheckOutcome.Passed, result.CheckDigit);
        Assert.Equal(EntityCheckProvenance.IssuerPublished, result.CheckProvenance);
    }

    [Theory]
    [InlineData("BG831094393", "BG831094394")] // Weighted sum 179 mod 11 = 3.
    [InlineData("CY00532445O", "CY00532445P")] // The transformed sum is 40, giving O.
    [InlineData("CZ46505334", "CZ46505335")] // Weighted sum 139, next multiple of 11 is 143.
    [InlineData("EL123456783", "EL123456784")] // Powers-of-two weighted sum is 1004, remainder 3.
    [InlineData("ESB12345674", "ESB12345675")] // CIF digit is the complement of weighted digit sum 26.
    [InlineData("FR40303265045", "FR41303265045")] // 12 + 3×(303265045 mod 97) gives 40 mod 97.
    [InlineData("LV40003245752", "LV40003245753")] // Weighted residue is 1, subtracted from 3.
    [InlineData("MT12345634", "MT12345635")] // Weighted sum 151 mod 37 = 3, so the final pair is 34.
    [InlineData("RO4221306", "RO4221307")] // Romanian ANAF SAF-T example; multiply the weighted sum by ten and take mod 11.
    public void CommunityVatChecksDistinguishPassAndFailure(string value, string badValue) {
        EntityScheme scheme = EuVatSchemes.Get(value.Substring(0, 2));
        EntityValidationResult valid = scheme.Validate(value);
        Assert.Equal(EntityValidationStatus.Valid, valid.Status);
        Assert.Equal(EntityCheckProvenance.CommunityDocumented, valid.CheckProvenance);
        EntityValidationResult invalid = scheme.Validate(badValue);
        Assert.Equal(EntityValidationStatus.InvalidCheckDigit, invalid.Status);
        Assert.Equal(EntityCheckOutcome.Failed, invalid.CheckDigit);
        Assert.Equal(EntityCheckProvenance.CommunityDocumented, invalid.CheckProvenance);
    }

    [Theory]
    [InlineData("BG3071234567")]
    [InlineData("CZ395601439")]
    [InlineData("LV32123456789")]
    [InlineData("XIGD499")]
    public void UncheckedVatVariantsHaveNoCheckProvenance(string value) {
        EntityValidationResult result = EuVatSchemes.Get(value.Substring(0, 2)).Validate(value);
        Assert.True(result.IsValid);
        Assert.Equal(EntityCheckOutcome.NotDefined, result.CheckDigit);
        Assert.Equal(EntityCheckProvenance.None, result.CheckProvenance);
    }

    [Fact]
    public void CyprusReservedPrefixFailsBeforeItsOtherwisePassingCheck() {
        var scheme = EuVatSchemes.Get("CY");
        Assert.True(scheme.IsValid("CY10259033P")); // Transformed sum 41, remainder 15 maps to P.
        EntityValidationResult invalid = scheme.Validate("CY12000000F"); // Transformed sum 5 maps to F.
        Assert.Equal(EntityValidationStatus.NotAllocated, invalid.Status);
        Assert.Null(invalid.Check);
    }

    [Fact]
    public void SpanishSpecialPersonalVatReportsTheIssuerCheckOnFailure() {
        EntityValidationResult result = EuVatSchemes.Get("ES").Validate("ESK1234567A");
        Assert.Equal(EntityValidationStatus.InvalidCheckDigit, result.Status);
        Assert.Equal(EntityCheckProvenance.IssuerPublished, result.CheckProvenance);
        Assert.Equal("www.caib.es", result.Check!.Source!.Host);
    }

    [Fact]
    public void FrenchNationalKeyCanMatchItsVatCountryPrefix() {
        // The alphanumeric key FR gives the same modulus 11 residue as the serial portion.
        var scheme = EuVatSchemes.Get("FR");
        Assert.Equal("FRFR100000009", scheme.Normalize("FR100000009"));
    }

    [Fact]
    public void DutchVatChecksPunctuationButRejectsAnOutOfRangeSuffixBeforeChecking() {
        // NL expands to 2321 and each asterisk to 37. The resulting number has remainder 1.
        EntityMatch match = Assert.Single(new EntityFinder().Find("NL**0000000007").Matches);
        Assert.Equal("NL**0000000007", match.Normalized);
        Assert.Equal(EntityCheckKind.Mod97_10, match.Check!.Kind);
        EntityValidationResult invalid = EuVatSchemes.Get("NL").Validate("NLAAAAAAAAAA01");
        Assert.Equal(EntityValidationStatus.InvalidFormat, invalid.Status);
        Assert.Null(invalid.Check);
    }

    [Fact]
    public void SpanishVatReportsTheSourceOfTheApplicableCheck() {
        EntityScheme scheme = EuVatSchemes.Get("ES");
        EntityCheckInfo personal = Assert.IsType<EntityCheckInfo>(scheme.Validate("ES12345678Z").Check);
        EntityMatch company = Assert.Single(new EntityFinder().Find("ESB12345674").Matches);
        EntityCheckInfo business = Assert.IsType<EntityCheckInfo>(company.Check);

        Assert.Contains(personal, scheme.Checks);
        Assert.Contains(business, scheme.Checks);
        Assert.Equal(EntityCheckKind.Mod23, personal.Kind);
        Assert.Equal(EntityCheckKind.Mod10, business.Kind);
        Assert.Equal("www.interior.gob.es", personal.Source!.Host);
        Assert.NotEqual(personal.Source, business.Source);
        Assert.True(company.Evidence.HasFlag(EntityEvidence.CommunityDocumentedCheck));
    }

    [Fact]
    public void GreeceAndNorthernIrelandUseIsoCountriesInMetadata() {
        Assert.Equal("GR", EuVatSchemes.Get("el").Country);
        Assert.Equal("GB", EuVatSchemes.Get("XI").Country);
    }

    [Theory]
    [InlineData("434031494", "GB434031494", EntityCheckOutcome.Passed)] // BMF numeric example.
    [InlineData("GB434031494001", "GB434031494001", EntityCheckOutcome.Passed)]
    [InlineData("GBGD000", "GBGD000", EntityCheckOutcome.NotDefined)]
    [InlineData("GBHA500", "GBHA500", EntityCheckOutcome.NotDefined)]
    public void UnitedKingdomVatSupportsTheBranchAndPublicBodyVariants(string value, string expected, EntityCheckOutcome check) {
        EntityValidationResult result = UnitedKingdomSchemes.Vat.Validate(value);
        Assert.Equal(EntityValidationStatus.Valid, result.Status);
        Assert.Equal(expected, result.Normalized);
        Assert.Equal(check, result.CheckDigit);
    }

    [Theory]
    [InlineData("GBGD500")]
    [InlineData("GBHA499")]
    [InlineData("GB434031494000")]
    public void UnitedKingdomVatRejectsReservedBranchAndPublicBodyRanges(string value) =>
        Assert.Equal(EntityValidationStatus.NotAllocated, UnitedKingdomSchemes.Vat.Validate(value).Status);
}
