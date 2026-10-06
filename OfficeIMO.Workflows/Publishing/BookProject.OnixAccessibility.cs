using System.Globalization;
using System.Net.Mail;
using System.Xml.Linq;

namespace OfficeIMO.Workflows;

public sealed partial class BookProject {
    private static IReadOnlyList<XElement> BuildOnixAccessibility(BookOnixAccessibilityMetadata? metadata, CancellationToken token) {
        if (metadata == null) return [];
        ArgumentNullException.ThrowIfNull(metadata.Features);
        if (metadata.Features.Count > 32) throw new ArgumentException("Supply at most 32 accessibility features.", nameof(metadata));
        var features = metadata.Features.ToArray();
        if (features.Distinct().Count() != features.Length)
            throw new ArgumentException("Accessibility features must be distinct.", nameof(metadata));
        if (metadata.Status == BookOnixAccessibilityStatus.Unknown && metadata.Conformance != null)
            throw new ArgumentException("Unknown accessibility cannot also assert conformance.", nameof(metadata));
        XNamespace onix = OnixNamespace;
        var result = new List<XElement>();
        void Add(string code, string? description = null) {
            token.ThrowIfCancellationRequested();
            if (description != null) RequireOnixText(description, nameof(metadata));
            result.Add(new XElement(onix + "ProductFormFeature", new XElement(onix + "ProductFormFeatureType", "09"),
                new XElement(onix + "ProductFormFeatureValue", code),
                description == null ? null : new XElement(onix + "ProductFormFeatureDescription", description)));
        }
        if (metadata.Summary != null) Add("00", metadata.Summary);
        if (metadata.Status is { } status) Add(status switch {
            BookOnixAccessibilityStatus.Unknown => "08", BookOnixAccessibilityStatus.Limited => "09",
            _ => throw new ArgumentOutOfRangeException(nameof(metadata.Status))
        });
        foreach (var feature in features) {
            if (!Enum.IsDefined(feature)) throw new ArgumentOutOfRangeException(nameof(metadata.Features));
            Add(((int)feature).ToString("00", CultureInfo.InvariantCulture));
        }
        if (metadata.Conformance is { } conformance) {
            Add("04");
            Add(conformance.Version switch {
                BookOnixWcagVersion.V2_0 => "80", BookOnixWcagVersion.V2_1 => "81", BookOnixWcagVersion.V2_2 => "82",
                _ => throw new ArgumentOutOfRangeException(nameof(conformance.Version))
            });
            Add(conformance.Level switch {
                BookOnixWcagLevel.A => "84", BookOnixWcagLevel.AA => "85", BookOnixWcagLevel.AAA => "86",
                _ => throw new ArgumentOutOfRangeException(nameof(conformance.Level))
            });
        }
        if (metadata.AssessmentDate is { } date) Add("91", date.ToString("yyyyMMdd", CultureInfo.InvariantCulture));
        void AddUrl(string code, string? url) {
            if (url == null) return;
            RequireOnixText(url, nameof(metadata));
            if (url.Any(char.IsWhiteSpace) || !Uri.TryCreate(url, UriKind.Absolute, out var parsed) ||
                (parsed.Scheme != Uri.UriSchemeHttps && parsed.Scheme != Uri.UriSchemeHttp) || parsed.UserInfo.Length != 0)
                throw new ArgumentException("Accessibility information requires an absolute HTTP(S) URL without credentials.", nameof(metadata));
            Add(code, url);
        }
        void AddEmail(string code, string? email) {
            if (email == null) return;
            RequireOnixText(email, nameof(metadata));
            if (!MailAddress.TryCreate(email, out var parsed) || parsed.Address != email || parsed.DisplayName.Length != 0)
                throw new ArgumentException("Supply a plain accessibility contact email address.", nameof(metadata));
            Add(code, email);
        }
        if (metadata.Certification is { } certification) {
            RequireOnixText(certification.Name, nameof(certification.Name));
            RequireOnixText(certification.Url, nameof(certification.Url));
            if (certification.CredentiallingOrganizationName != null) Add("88", certification.CredentiallingOrganizationName);
            AddUrl("89", certification.CredentiallingOrganizationUrl);
            Add("90", certification.Name);
            AddUrl("93", certification.Url);
        }
        AddUrl("94", metadata.IndependentReportUrl);
        AddUrl("95", metadata.IntermediaryInformationUrl);
        AddUrl("96", metadata.PublisherInformationUrl);
        AddUrl("97", metadata.CompatibilityReportUrl);
        AddEmail("98", metadata.IntermediaryContactEmail);
        AddEmail("99", metadata.PublisherContactEmail);
        if (result.Count == 0) throw new ArgumentException("Supply at least one accessibility assertion, or omit Accessibility.", nameof(metadata));
        return result;
    }
}
