using System;
using System.Collections.Generic;
using OfficeIMO.Entities.Internal;

namespace OfficeIMO.Entities;

/// <summary>Payment card number using the Luhn check in <see href="https://www.iso.org/standard/70484.html">ISO/IEC 7812-1</see>.</summary>
internal sealed class CardScheme : CompactEntityScheme {
    public override string Id => "CARD";
    public override EntityKind Kind => EntityKind.PaymentCard;
    public override string DisplayName => "Payment card number";
    public override string Issuer => "ISO/IEC 7812 registration authority";
    public override Uri Specification { get; } = new("https://www.iso.org/standard/70484.html");
    public override int MinLength => 12;
    public override int MaxLength => 19;
    public override EntityCheckKind Check => EntityCheckKind.Luhn;
    public override EntitySensitivity Sensitivity => EntitySensitivity.Financial;

    protected override EntityValidationResult ValidateCompact(string value) =>
        EntityCharacters.IsDigits(value, 0, value.Length) ? Checked(value, CheckDigits.Luhn(value)) : Result(EntityValidationStatus.InvalidCharacters);

    public override string MaskPartial(string normalized) {
        string? value = Normalize(normalized);
        return value == null ? base.MaskPartial(normalized) :
            EntityCharacters.Group(new string('•', value.Length - 4) + value.Substring(value.Length - 4), 4);
    }

    public override EntityDetails? Decode(string normalized) {
        string? value = Normalize(normalized);
        return value == null ? null : new EntityDetails(properties: new Dictionary<string, string> { ["mii"] = value.Substring(0, 1) });
    }
}
