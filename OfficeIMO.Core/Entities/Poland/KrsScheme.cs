using System;
using OfficeIMO.Entities.Internal;

namespace OfficeIMO.Entities;

/// <summary>Polish KRS number as stored in <see href="https://api.sejm.gov.pl/eli/acts/DU/2026/649/text.pdf">the court register</see>.</summary>
internal sealed class KrsScheme : CompactEntityScheme {
    public override string Id => "PL.KRS";
    public override EntityKind Kind => EntityKind.CompanyId;
    public override string Country => "PL";
    public override string DisplayName => "KRS";
    public override string Issuer => "Polish Ministry of Justice";
    public override Uri Specification { get; } = new("https://api.sejm.gov.pl/eli/acts/DU/2026/649/text.pdf");
    public override int MinLength => 10;
    public override int MaxLength => 10;
    public override EntitySensitivity Sensitivity => EntitySensitivity.Organizational;

    protected override EntityValidationResult ValidateCompact(string value) =>
        EntityCharacters.IsDigits(value, 0, value.Length) ? Result(EntityValidationStatus.Valid, value) : Result(EntityValidationStatus.InvalidCharacters);
}
