using System;
using OfficeIMO.Entities.Internal;

namespace OfficeIMO.Entities;

/// <summary>Polish identity card number. The check follows <see href="https://itmock.pl/blog/numer-dowodu-osobistego">the documented positional check</see>.</summary>
internal sealed class IdentityCardScheme : CompactEntityScheme {
    public override string Id => "PL.IDCARD";
    public override EntityKind Kind => EntityKind.IdentityDocument;
    public override string Country => "PL";
    public override string DisplayName => "Polish identity card";
    public override string Issuer => "Polish Ministry of the Interior and Administration";
    public override Uri Specification { get; } = new("https://itmock.pl/blog/numer-dowodu-osobistego");
    public override int MinLength => 9;
    public override int MaxLength => 9;
    public override EntityCheckKind Check => EntityCheckKind.Mod10;

    protected override EntityCheckProvenance DefaultCheckProvenance => EntityCheckProvenance.CommunityDocumented;

    protected override EntityValidationResult ValidateCompact(string value) {
        if (!EntityCharacters.IsLetters(value, 0, 3) || !EntityCharacters.IsDigits(value, 3, 6)) return Result(EntityValidationStatus.InvalidFormat);
        return Checked(value, CheckDigits.WeightedSum(value, "731073173") % 10 == value[3] - '0');
    }
}
