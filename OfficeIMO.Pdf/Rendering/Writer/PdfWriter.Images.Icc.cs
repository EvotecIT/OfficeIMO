using OfficeIMO.Drawing;
using System.Globalization;
using System.Threading;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    /// <summary>Retains the source device profile without requiring a managed color transform.</summary>
    private static bool TryAttachImageIccProfile(byte[] data, OfficeImageFormat format, PdfImageStream image,
        int components, CancellationToken cancellationToken, out string? reason) {
        byte[]? profile = OfficeImageMetadataInspector.ReadIccProfile(data, format, 4 * 1024 * 1024,
            cancellationToken, out bool hasProfile);
        reason = null;
        if (!hasProfile) return true;
        if (profile == null || !OfficeIccProfileValidator.TryValidate(profile, 0, profile.Length, cancellationToken)) {
            reason = "The embedded image ICC profile is invalid or exceeds the supported profile size limit.";
            return false;
        }
        string space = Encoding.ASCII.GetString(profile, 16, 4);
        string profileClass = Encoding.ASCII.GetString(profile, 12, 4);
        string pcs = Encoding.ASCII.GetString(profile, 20, 4);
        int profileComponents = space == "GRAY" ? 1 : space == "RGB " ? 3 : space == "CMYK" ? 4 : 0;
        if (profileComponents != components || profileComponents == 0 ||
            (profileClass != "scnr" && profileClass != "mntr" && profileClass != "prtr" && profileClass != "spac") ||
            (pcs != "XYZ " && pcs != "Lab ")) {
            reason = "The embedded ICC profile is not a compatible Gray, RGB, or CMYK image profile.";
            return false;
        }
        image.IccProfile = profile;
        image.IccComponents = components;
        return true;
    }

    internal static string BuildImageIccDictionary(PdfImageStream image) =>
        "<< /N " + image.IccComponents.ToString(CultureInfo.InvariantCulture) + " /Alternate " +
        GetImageIccAlternate(image) + " /Length " + image.IccProfile!.Length.ToString(CultureInfo.InvariantCulture) + " >>";

    internal static string GetImageIccAlternate(PdfImageStream image) =>
        image.IccComponents == 1 ? "/DeviceGray" : image.IccComponents == 4 ? "/DeviceCMYK" : "/DeviceRGB";

    /// <summary>Adds dependent ICC and alpha streams before an image in typed rewrite/append pipelines.</summary>
    internal static int AddImageXObjectResource(PdfImageStream source, Func<PdfStream, int> addStream,
        PdfFileVersion version = PdfFileVersion.Pdf14) {
        PdfImageStream image = PrepareImageStreamForVersion(source, version);
        int? mask = image.SoftMask == null ? null : AddImageXObjectResource(image.SoftMask, addStream, version);
        int? profile = null;
        if (image.IccProfile != null) {
            var dictionary = new PdfDictionary();
            dictionary.Items["N"] = new PdfNumber(image.IccComponents);
            dictionary.Items["Alternate"] = new PdfName(GetImageIccAlternate(image).Substring(1));
            dictionary.Items["Filter"] = new PdfName("FlateDecode");
            profile = addStream(new PdfStream(dictionary, PdfFlateEncoder.Compress(image.IccProfile)));
        }
        return addStream(PdfImageXObjectDictionaryBuilder.BuildStreamObject(image, mask, profile));
    }
}
