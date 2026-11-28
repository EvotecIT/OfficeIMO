using System;
using OfficeIMO.Entities.Internal;

namespace OfficeIMO.Entities;

/// <summary>Polish passport number. The check follows <see href="https://romek.info/ut/paszport.html">the documented positional check</see>.</summary>
internal sealed class PassportScheme : CompactEntityScheme {
    public override string Id => "PL.PASSPORT";
    public override EntityKind Kind => EntityKind.IdentityDocument;
    public override string Country => "PL";
    public override string DisplayName => "Polish passport";
    public override string Issuer => "Polish Ministry of the Interior and Administration";
    public override Uri Specification { get; } = new("https://romek.info/ut/paszport.html");
    public override int MinLength => 9;
    public override int MaxLength => 9;
    public override EntityCheckKind Check => EntityCheckKind.Mod10;

    protected override EntityCheckProvenance DefaultCheckProvenance => EntityCheckProvenance.CommunityDocumented;

    protected override EntityValidationResult ValidateCompact(string value) {
        if (!EntityCharacters.IsLetters(value, 0, 2) || !EntityCharacters.IsDigits(value, 2, 7)) return Result(EntityValidationStatus.InvalidFormat);
        // The issuer used the BM series without an in-number check digit.
        if (value[0] == 'B' && value[1] == 'M') return Result(EntityValidationStatus.Valid, value);
        return Checked(value, CheckDigits.WeightedSum(value, "730173173") % 10 == value[2] - '0');
    }
}
