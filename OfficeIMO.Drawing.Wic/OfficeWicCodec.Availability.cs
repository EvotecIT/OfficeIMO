using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using OfficeIMO.Drawing.Wic.Interop;

namespace OfficeIMO.Drawing.Wic;

public sealed partial class OfficeWicCodec {
    /// <summary>Enumerates installed WIC decoders, then decodes a real embedded HEVC image to RGBA.</summary>
    /// <remarks>Results are not cached, so an extension installed during the process lifetime can be detected.
    /// A registered HEIF decoder that fails the probe reports WindowsHevcExtensionMissing; its message
    /// distinguishes failed decoding from proof that a Store package is absent.</remarks>
    public OfficeRasterCodecAvailability GetAvailability() {
        if (!IsSupported) return new(OfficeRasterCodecUnavailableReason.NoPlatformCodec,
            "Windows Imaging Component is available only on Windows.");
        try {
            if (!HasHeifDecoder()) return new(OfficeRasterCodecUnavailableReason.WindowsHeifExtensionMissing,
                "Microsoft's HEIF WIC decoder is not registered. Install HEIF Image Extensions through Windows or Microsoft Store.");
        } catch (COMException ex) {
            return new(OfficeRasterCodecUnavailableReason.NoPlatformCodec,
                $"WIC decoder enumeration failed (HRESULT 0x{ex.HResult:X8}). HEIF/HEVC availability could not be established.");
        }

        try {
            OfficeRasterImage? image = DecodeWindows(ReadProbe());
            if (image is { Width: 16, Height: 16 }) return new(OfficeRasterCodecUnavailableReason.None);
            return ProbeUnavailable("The HEVC probe produced no valid 16 by 16 raster.");
        } catch (COMException ex) {
            return ProbeUnavailable($"The HEVC pixel decode failed (HRESULT 0x{ex.HResult:X8}).");
        }
    }

    private static OfficeRasterCodecAvailability ProbeUnavailable(string detail) => new(
        OfficeRasterCodecUnavailableReason.WindowsHevcExtensionMissing,
        "The HEIF WIC decoder is registered, but HEVC decoding is unavailable. Install or repair HEVC Video Extensions through Windows or Microsoft Store. " +
        detail + " A probe failure may also indicate a broken system codec; it does not prove the package is absent.");

    [SupportedOSPlatform("windows")]
    private static bool HasHeifDecoder() {
        using var apartment = new WicApartment();
        using var factory = WicNative.CreateFactory();
        // WICDecoder = 1. Refresh the enumeration after a user installs an extension.
        WicNative.Check(factory.Value.CreateComponentEnumerator(1, 1, out nint enumeratorPointer));
        using var enumerator = new WicObject<IEnumUnknown>(enumeratorPointer);
        while (true) {
            int hr = enumerator.Value.Next(1, out nint item, out uint fetched);
            // Adopt any returned reference even if the enumerator reports an error.
            if (item != 0) {
                using var info = new WicObject<IWicComponentInfo>(item);
                WicNative.Check(hr);
                WicNative.Check(info.Value.GetCLSID(out Guid clsid));
                if (fetched == 1 && clsid == WicNative.HeifDecoderClass) return true;
            } else {
                WicNative.Check(hr);
            }
            if (hr == 1 || fetched == 0) return false; // S_FALSE: end of sequence
        }
    }

    private static byte[] ReadProbe() {
        using Stream stream = typeof(OfficeWicCodec).Assembly.GetManifestResourceStream("OfficeIMO.Drawing.Wic.Probe.heic")
            ?? throw new InvalidOperationException("The embedded HEVC availability probe is missing.");
        byte[] bytes = new byte[checked((int)stream.Length)];
        stream.ReadExactly(bytes);
        return bytes;
    }
}
