using System;
using OfficeIMO.Entities.Internal;

namespace OfficeIMO.Entities;

internal abstract class PrefixedVatScheme : CompactEntityScheme {
    internal const string ViesSpecification = "https://ec.europa.eu/taxation_customs/vies/#/faq";
    internal const string BmfSpecification = NipScheme.Source;
    internal const string CdqSpecification = "https://meta.cdq.com/Public:Rule_Category/IDENTIFIER_CHECKDIGIT";

    protected PrefixedVatScheme(string prefix, int nationalMinLength, int nationalMaxLength,
        EntityCheckKind check, string specification, string? country = null, string? id = null) {
        Prefix = prefix;
        Id = id ?? "EU.VAT." + prefix;
        Country = country ?? prefix;
        MinLength = nationalMinLength + 2;
        MaxLength = nationalMaxLength + 2;
        Check = check;
        Specification = new Uri(specification);
    }

    internal string Prefix { get; }
    public sealed override string Id { get; }
    public sealed override string Country { get; }
    public sealed override EntityKind Kind => EntityKind.VatId;
    public sealed override string DisplayName => Prefix + " VAT";
    public override string Issuer => Country + " tax authority";
    public sealed override Uri Specification { get; }
    public sealed override int MinLength { get; }
    public sealed override int MaxLength { get; }
    public sealed override EntityCheckKind Check { get; }
    public sealed override EntitySensitivity Sensitivity => EntitySensitivity.Organizational;

    protected override string Canonicalize(string compact) => compact.StartsWith(Prefix, StringComparison.Ordinal) ? compact : Prefix + compact;

    protected sealed override EntityValidationResult ValidateCompact(string value) => ValidateNational(value.Substring(2));

    protected abstract EntityValidationResult ValidateNational(string number);

    protected EntityValidationResult NationalResult(EntityValidationStatus status, string number) => Result(status, Prefix + number);

    protected EntityValidationResult NationalCheck(string number, bool passes,
        EntityCheckInfo? check = null) => Checked(Prefix + number, passes, check);

    protected EntityValidationResult NationalDigits(string number, Func<string, bool> passes,
        EntityCheckInfo? check = null) =>
        EntityCharacters.IsDigits(number, 0, number.Length) ? NationalCheck(number, passes(number), check) : Result(EntityValidationStatus.InvalidFormat);
}
