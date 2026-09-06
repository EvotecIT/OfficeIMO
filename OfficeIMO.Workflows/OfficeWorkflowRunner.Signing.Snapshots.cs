using OfficeIMO.Pdf;

namespace OfficeIMO.Workflows;

public sealed partial class OfficeWorkflowRunner {
    internal static PdfExternalSignatureOptions? SnapshotSignatureOptions(PdfExternalSignatureOptions? options) {
        if (options is null) return null;
        PdfVisibleSignatureAppearanceOptions? appearance = options.VisibleAppearance is null ? null : new PdfVisibleSignatureAppearanceOptions {
            PageNumber = options.VisibleAppearance.PageNumber,
            X = options.VisibleAppearance.X,
            Y = options.VisibleAppearance.Y,
            Width = options.VisibleAppearance.Width,
            Height = options.VisibleAppearance.Height,
            Text = options.VisibleAppearance.Text,
            ShowText = options.VisibleAppearance.ShowText,
            FontSize = options.VisibleAppearance.FontSize,
            BackgroundColor = options.VisibleAppearance.BackgroundColor,
            BorderColor = options.VisibleAppearance.BorderColor,
            TextColor = options.VisibleAppearance.TextColor,
            ImageBytes = options.VisibleAppearance.ImageBytes?.ToArray(),
            ImageFit = options.VisibleAppearance.ImageFit,
            ImagePadding = options.VisibleAppearance.ImagePadding
        };
        return new PdfExternalSignatureOptions {
            MaxInputBytes = options.MaxInputBytes,
            CancellationToken = options.CancellationToken,
            Profile = options.Profile,
            CertificationPermission = options.CertificationPermission,
            VisibleAppearance = appearance,
            FieldName = options.FieldName,
            Filter = options.Filter,
            SubFilter = options.SubFilter,
            Name = options.Name,
            Reason = options.Reason,
            Location = options.Location,
            ContactInfo = options.ContactInfo,
            SigningTime = options.SigningTime,
            ReservedSignatureContentsBytes = options.ReservedSignatureContentsBytes
        };
    }

}
