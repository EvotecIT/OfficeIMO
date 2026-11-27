using System;
using OfficeIMO.Entities.Internal;

namespace OfficeIMO.Entities;

/// <summary>Polish NIP checksum from the <see href="https://www.bmf.gv.at/dam/jcr:d6794f8f-d321-43df-9840-1a841f9bf5dc/BMF_UID_Konstruktionsregeln_Stand_November%202020.pdf">BMF construction rules</see>.</summary>
internal sealed class NipScheme : CompactEntityScheme {
    internal const string Source = "https://www.bmf.gv.at/dam/jcr:d6794f8f-d321-43df-9840-1a841f9bf5dc/BMF_UID_Konstruktionsregeln_Stand_November%202020.pdf";

    public override string Id => "PL.NIP";
    public override EntityKind Kind => EntityKind.TaxId;
    public override string Country => "PL";
    public override string DisplayName => "NIP";
    public override string Issuer => "Polish tax administration";
    public override Uri Specification { get; } = new(Source);
    public override int MinLength => 10;
    public override int MaxLength => 10;
    public override EntityCheckKind Check => EntityCheckKind.Mod11;
    public override EntitySensitivity Sensitivity => EntitySensitivity.Organizational;

    protected override string Canonicalize(string compact) => compact.StartsWith("PL", StringComparison.Ordinal) ? compact.Substring(2) : compact;

    protected override EntityValidationResult ValidateCompact(string value) =>
        EntityCharacters.IsDigits(value, 0, value.Length) ? Checked(value, PassesCheck(value)) : Result(EntityValidationStatus.InvalidCharacters);

    internal static bool PassesCheck(string value, int start = 0) =>
        CheckDigits.WeightedSum(value, "657234567", start) % 11 == value[start + 9] - '0';

    public override string Format(string normalized, EntityFormatStyle style = EntityFormatStyle.Display) {
        string value = base.Format(normalized, style);
        return style == EntityFormatStyle.Display && value.Length == 10
            ? value.Insert(8, "-").Insert(6, "-").Insert(3, "-")
            : value;
    }
}
