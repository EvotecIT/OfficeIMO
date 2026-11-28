using System;
using OfficeIMO.Entities.Internal;

namespace OfficeIMO.Entities;

/// <summary>Company numbers from the <see href="https://assets.publishing.service.gov.uk/media/5d08c0f340f0b6094a379078/uniformResourceIdentifiersCustomerGuide.pdf">Companies House URI guide</see>.</summary>
internal sealed class CompaniesHouseScheme : CompactEntityScheme {
    private const string Prefixes = " AC ZC FC GE LP OC SE SA SZ SF GS SL SO SC ES NA NZ NF GN NL NC R0 NI EN IP SP IC SI NP NV RC SR NR NO BR OE ";

    public override string Id => "GB.COMPANY";
    public override EntityKind Kind => EntityKind.CompanyId;
    public override string Country => "GB";
    public override string DisplayName => "Companies House number";
    public override string Issuer => "Companies House";
    public override Uri Specification { get; } = new("https://assets.publishing.service.gov.uk/media/5d08c0f340f0b6094a379078/uniformResourceIdentifiersCustomerGuide.pdf");
    public override int MinLength => 8;
    public override int MaxLength => 8;
    public override EntitySensitivity Sensitivity => EntitySensitivity.Organizational;

    protected override EntityValidationResult ValidateCompact(string value) {
        if (!EntityCharacters.IsDigits(value, 2, 6)) return Result(EntityValidationStatus.InvalidFormat);
        bool valid = EntityCharacters.IsDigits(value, 0, 2) || Prefixes.IndexOf(" " + value.Substring(0, 2) + " ", StringComparison.Ordinal) >= 0;
        return Result(valid ? EntityValidationStatus.Valid : EntityValidationStatus.NotAllocated, value);
    }
}
