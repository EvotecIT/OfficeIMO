using System;
using System.Threading;

namespace OfficeIMO.Drawing;

internal static partial class OfficeImageMetadataInspector {
    private static void InspectJpegXr(byte[] data, OfficeImageMetadataSnapshot snapshot, CancellationToken token) {
        try {
            var container = OfficeJpegXrDecoder.ReadContainer(data, token);
            snapshot.Kinds |= container.MetadataKinds;
            snapshot.HasColorRenderingMetadata = container.HasColorRenderingMetadata;
            snapshot.HasPhysicalResolution = (container.MetadataKinds & OfficeImageMetadataKinds.Resolution) != 0;
            bool swap = container.Transform >= 4;
            snapshot.PhysicalDpiX = swap ? container.DpiY : container.DpiX;
            snapshot.PhysicalDpiY = swap ? container.DpiX : container.DpiY;
        } catch (FormatException) { snapshot.HasColorRenderingMetadata = true; }
        catch (OverflowException) { snapshot.HasColorRenderingMetadata = true; }
    }

    private static byte[]? ReadJpegXrIcc(byte[] data, int maximumBytes, CancellationToken token, out bool hasProfile) {
        hasProfile = false;
        try {
            var container = OfficeJpegXrDecoder.ReadContainer(data, token);
            hasProfile = container.HasIcc;
            if (!hasProfile || container.IccLength == 0 || container.IccLength > maximumBytes) return null;
            return Slice(data, container.IccOffset, container.IccLength, token);
        } catch (FormatException) { return null; }
        catch (OverflowException) { return null; }
    }
}
