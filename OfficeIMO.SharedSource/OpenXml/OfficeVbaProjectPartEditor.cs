using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.OpenXml.Internal;

/// <summary>Shared bounded VBA part operations for Word, Excel, and PowerPoint.</summary>
internal static class OfficeVbaProjectPartEditor {
    internal static byte[] Read(VbaProjectPart part, long maximumBytes) {
        using Stream input = part.GetStream(FileMode.Open, FileAccess.Read);
        return OfficeStreamReader.ReadAllBytes(input, maximumBytes);
    }

    internal static bool Apply(VbaProjectPart part, byte[] bytes, OfficeVbaWriteOptions options) {
        if (Read(part, options.MaximumProjectBytes).SequenceEqual(bytes)) return false;
        OpenXmlPart[] signatures = part.Parts.Where(pair =>
            pair.OpenXmlPart.RelationshipType.IndexOf("vbaProjectSignature", StringComparison.OrdinalIgnoreCase) >= 0
            || pair.OpenXmlPart.ContentType.IndexOf("vbaProjectSignature", StringComparison.OrdinalIgnoreCase) >= 0)
            .Select(pair => pair.OpenXmlPart).ToArray();
        if (signatures.Length > 0 && !options.AllowSignatureRemoval) {
            throw new InvalidOperationException("Changing the VBA project invalidates its signatures. Explicit signature removal is required.");
        }
        using var input = new MemoryStream(bytes, writable: false);
        part.FeedData(input);
        foreach (OpenXmlPart signature in signatures) part.DeletePart(signature);
        return true;
    }
}
