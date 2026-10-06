using System.Globalization;
using System.Xml.Linq;

namespace OfficeIMO.Workflows;

public sealed partial class BookProject {
    private static readonly HashSet<string> OnixResourceFormats = new(StringComparer.Ordinal) {
        "A103", "A104", "A105", "A106", "A107", "A108", "A111",
        "D101", "D102", "D103", "D104", "D105", "D106", "D107", "D108", "D109", "D401",
        "D501", "D502", "D503", "D504", "D505", "D506", "D507", "D508", "D509", "D510", "D511",
        "E101", "E105", "E107", "E112", "E113", "E115", "E116", "E139", "E140"
    };

    private static void AddOnixSupportingResources(XElement collateral, IReadOnlyList<BookOnixSupportingResource> resources,
        ref int textBudget, CancellationToken token) {
        XNamespace ns = OnixNamespace;
        int sequence = 0;
        foreach (var resource in resources) {
            token.ThrowIfCancellationRequested();
            ArgumentNullException.ThrowIfNull(resource);
            if (!Enum.IsDefined(resource.Type) || !Enum.IsDefined(resource.Mode)) throw new ArgumentOutOfRangeException(nameof(resources));
            ArgumentNullException.ThrowIfNull(resource.Versions);
            if (resource.Versions.Count is < 1 or > 16) throw new ArgumentException("Supply one to 16 resource versions.", nameof(resources));
            if (resource.LengthMinutes < 0 || (resource.LengthMinutes != null && resource.Mode is not (BookOnixResourceMode.Audio or BookOnixResourceMode.Video)))
                throw new ArgumentException("Length in minutes describes audio/video resources and cannot be negative.", nameof(resources));
            var element = new XElement(ns + "SupportingResource", new XElement(ns + "SequenceNumber", ++sequence),
                new XElement(ns + "ResourceContentType", ((int)resource.Type).ToString("00", CultureInfo.InvariantCulture)));
            element.Add(BuildOnixContentAudiences(resource.Audiences));
            if (resource.Territory != null) element.Add(ReadOnixTerritory(resource.Territory).ToXml());
            element.Add(new XElement(ns + "ResourceMode", ((int)resource.Mode).ToString("00", CultureInfo.InvariantCulture)));
            foreach (var (code, notes) in new[] { ("01", resource.Credits), ("02", resource.Captions), ("03", resource.CopyrightHolders), ("07", resource.AlternativeTexts) }) {
                var values = BuildOnixCollateralValues(notes, "FeatureNote", 4096, false, ref textBudget, token);
                if (values.Count > 0) element.Add(new XElement(ns + "ResourceFeature", new XElement(ns + "ResourceFeatureType", code), values));
            }
            if (resource.LengthMinutes != null) element.Add(new XElement(ns + "ResourceFeature", new XElement(ns + "ResourceFeatureType", "04"),
                new XElement(ns + "FeatureValue", resource.LengthMinutes.Value.ToString(CultureInfo.InvariantCulture))));
            foreach (var version in resource.Versions)
                element.Add(BuildOnixResourceVersion(version, resource.Mode, ref textBudget, token));
            collateral.Add(element);
        }
    }

    private static XElement BuildOnixResourceVersion(BookOnixResourceVersion version, BookOnixResourceMode mode, ref int textBudget, CancellationToken token) {
        token.ThrowIfCancellationRequested();
        ArgumentNullException.ThrowIfNull(version);
        if (!Enum.IsDefined(version.Form)) throw new ArgumentOutOfRangeException(nameof(version.Form));
        ArgumentNullException.ThrowIfNull(version.Links);
        if (version.Links.Count is < 1 or > 16) throw new ArgumentException("Supply one to 16 resource links.", nameof(version));
        if (version.ImageWidth <= 0 || version.ImageHeight <= 0 || version.ByteLength < 0)
            throw new ArgumentException("Image dimensions must be positive and byte length nonnegative.", nameof(version));
        if ((version.ImageWidth != null || version.ImageHeight != null) && mode != BookOnixResourceMode.Image)
            throw new ArgumentException("Image dimensions require image mode.", nameof(version));
        if (version.UsableFrom > version.UsableUntil) throw new ArgumentException("Resource usage dates are reversed.", nameof(version));
        XNamespace ns = OnixNamespace;
        var result = new XElement(ns + "ResourceVersion", new XElement(ns + "ResourceForm", ((int)version.Form).ToString("00", CultureInfo.InvariantCulture)));
        void Feature(string code, string? value) {
            if (value != null) result.Add(new XElement(ns + "ResourceVersionFeature", new XElement(ns + "ResourceVersionFeatureType", code), new XElement(ns + "FeatureValue", value)));
        }
        if (version.FileFormatCode != null && !OnixResourceFormats.Contains(version.FileFormatCode))
            throw new ArgumentException("Unknown ONIX list 178 file format.", nameof(version.FileFormatCode));
        if (version.FileName != null) {
            RequireOnixText(version.FileName, nameof(version.FileName));
            if (version.FileName.Length > 255 || version.FileName is "." or ".." || version.FileName.Any(c => c is '/' or '\\' || char.IsControl(c)))
                throw new ArgumentException("Supply a filename, not a path.", nameof(version.FileName));
        }
        if (version.Sha256 != null && (version.Sha256.Length != 64 || version.Sha256.Any(c => !Uri.IsHexDigit(c))))
            throw new ArgumentException("SHA-256 requires exactly 64 hexadecimal digits.", nameof(version.Sha256));
        Feature("01", version.FileFormatCode); Feature("02", version.ImageHeight?.ToString(CultureInfo.InvariantCulture));
        Feature("03", version.ImageWidth?.ToString(CultureInfo.InvariantCulture)); Feature("04", version.FileName);
        Feature("07", version.ByteLength?.ToString(CultureInfo.InvariantCulture)); Feature("08", version.Sha256);
        if (version.FileName != null) ConsumeOnixResourceText(version.FileName, ref textBudget);
        var identities = new HashSet<(string, string?)>();
        bool multilingual = version.Links.Where(link => link != null).Select(link => link.LanguageCode).Distinct(StringComparer.Ordinal).Count() > 1;
        foreach (var link in version.Links) {
            token.ThrowIfCancellationRequested();
            ArgumentNullException.ThrowIfNull(link);
            RequireOnixHttpUrl(link.Url, nameof(link.Url));
            RequireOnixTranslationLanguage(link.LanguageCode, multilingual ? 2 : 1, nameof(link.LanguageCode));
            if (!identities.Add((link.Url, link.LanguageCode))) throw new ArgumentException("Resource URL/language pairs must be distinct.", nameof(version.Links));
            ConsumeOnixResourceText(link.Url, ref textBudget);
            result.Add(new XElement(ns + "ResourceLink", link.LanguageCode != null ? new XAttribute("language", link.LanguageCode) : null, link.Url));
        }
        result.Add(BuildOnixUsageConstraints(version.UsageConstraints, token));
        result.Add(BuildOnixLicenses(version.Licenses, ref textBudget, token));
        foreach (var (role, date) in new[] { ("14", version.UsableFrom), ("15", version.UsableUntil), ("17", version.UpdatedOn) })
            if (date != null) result.Add(new XElement(ns + "ContentDate", new XElement(ns + "ContentDateRole", role),
                new XElement(ns + "Date", new XAttribute("dateformat", "00"), date.Value.ToString("yyyyMMdd", CultureInfo.InvariantCulture))));
        return result;
    }

    private static void ConsumeOnixResourceText(string value, ref int budget) {
        if (value.Length > budget) throw new ArgumentException("Supporting resources exceed the aggregate collateral text budget.", nameof(value));
        budget -= value.Length;
    }

    private static IReadOnlyList<XElement> BuildOnixContentAudiences(IReadOnlyList<BookOnixContentAudience> audiences) {
        ArgumentNullException.ThrowIfNull(audiences);
        if (audiences.Count is < 1 or > 13 || audiences.Distinct().Count() != audiences.Count ||
            (audiences.Count != 1 && audiences.Contains(BookOnixContentAudience.Unrestricted)))
            throw new ArgumentException("Supply distinct collateral recipients; Unrestricted cannot accompany other codes.", nameof(audiences));
        XNamespace ns = OnixNamespace;
        return audiences.Select(audience => {
            if (!Enum.IsDefined(audience)) throw new ArgumentOutOfRangeException(nameof(audiences));
            return new XElement(ns + "ContentAudience", ((int)audience).ToString("00", CultureInfo.InvariantCulture));
        }).ToArray();
    }
}
