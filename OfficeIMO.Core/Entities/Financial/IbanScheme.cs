using System;
using System.Collections.Generic;
using System.Text;
using OfficeIMO.Entities.Internal;

namespace OfficeIMO.Entities;

/// <summary>IBAN rules per ISO 13616 and the <see href="https://www.swift.com/standards/data-standards/iban-international-bank-account-number">SWIFT IBAN Registry</see>.</summary>
internal sealed class IbanScheme : EntityScheme {
    public override string Id => "IBAN";
    public override EntityKind Kind => EntityKind.BankAccount;
    public override string DisplayName => "IBAN";
    public override string Issuer => "SWIFT (ISO 13616 registration authority)";
    public override Uri Specification { get; } = new("https://www.swift.com/standards/data-standards/iban-international-bank-account-number");
    public override int MinLength => 15;
    public override int MaxLength => 34;
    public override EntityCheckKind Check => EntityCheckKind.Mod97_10;
    public override EntitySensitivity Sensitivity => EntitySensitivity.Financial;

    public override EntityValidationResult Validate(string? value) {
        if (string.IsNullOrWhiteSpace(value)) return new(this, EntityValidationStatus.Empty);
        var compact = new StringBuilder(MaxLength);
        foreach (char character in value!) {
            if (char.IsWhiteSpace(character)) continue;
            if (!EntityCharacters.IsAlphaNumeric(character)) return new(this, EntityValidationStatus.InvalidCharacters);
            if (compact.Length == MaxLength) return new(this, EntityValidationStatus.InvalidLength);
            compact.Append(EntityCharacters.Upper(character));
        }
        if (compact.Length < MinLength) return new(this, EntityValidationStatus.InvalidLength);

        string normalized = compact.ToString();
        if (!EntityCharacters.IsLetters(normalized, 0, 2) || !EntityCharacters.IsDigits(normalized, 2, 2) ||
            !IbanStructure.HasRegisteredStructure(normalized)) {
            return new(this, EntityValidationStatus.InvalidFormat);
        }

        int digits = EntityCharacters.Number(normalized, 2, 2);
        bool valid = digits >= 2 && digits <= 98 && CheckDigits.Mod97(normalized, rotate: 4) == 1;
        return new(this, valid ? EntityValidationStatus.Valid : EntityValidationStatus.InvalidCheckDigit, normalized,
            valid ? EntityCheckOutcome.Passed : EntityCheckOutcome.Failed, Checks[0]);
    }

    public override string Format(string normalized, EntityFormatStyle style = EntityFormatStyle.Display) {
        string value = base.Format(normalized, style);
        return style == EntityFormatStyle.Display ? EntityCharacters.Group(value, 4) : value;
    }

    public override string MaskPartial(string normalized) {
        string? value = Normalize(normalized);
        if (value == null) return base.MaskPartial(normalized);
        return Format(value.Substring(0, 4) + new string('•', value.Length - 8) + value.Substring(value.Length - 4));
    }

    public override EntityDetails? Decode(string normalized) {
        string? value = Normalize(normalized);
        return value == null ? null : new EntityDetails(country: value.Substring(0, 2),
            properties: new Dictionary<string, string> { ["bban"] = value.Substring(4) });
    }
}
