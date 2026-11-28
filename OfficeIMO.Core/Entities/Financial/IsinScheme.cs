using System;
using OfficeIMO.Entities.Internal;

namespace OfficeIMO.Entities;

/// <summary>International Securities Identification Number per ISO 6166 and <see href="https://anna-web.org/identifiers/">ANNA</see>.</summary>
internal sealed class IsinScheme : CompactEntityScheme {
    public override string Id => "ISIN";
    public override EntityKind Kind => EntityKind.SecurityId;
    public override string DisplayName => "International Securities Identification Number";
    public override string Issuer => "Association of National Numbering Agencies";
    public override Uri Specification { get; } = new("https://anna-web.org/identifiers/");
    public override int MinLength => 12;
    public override int MaxLength => 12;
    public override EntityCheckKind Check => EntityCheckKind.Luhn;
    public override EntitySensitivity Sensitivity => EntitySensitivity.Public;

    protected override EntityValidationResult ValidateCompact(string value) {
        if (!EntityCharacters.IsLetters(value, 0, 2) || !EntityCharacters.IsDigit(value[11])) return Result(EntityValidationStatus.InvalidFormat);
        return Checked(value, CheckDigits.Luhn(value));
    }
}
