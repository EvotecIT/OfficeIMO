namespace OfficeIMO.Pdf;

internal static partial class ResourceResolver {
    private static bool TryGetJpxPayload(PdfStream stream, Dictionary<int, PdfIndirectObject> objects,
        string colorSpace, int maximumBytes, out byte[] payload) {
        payload = Array.Empty<byte>();
        // An RGBA codec can honor the codestream's own Gray/RGB samples. PDF-specific masks,
        // alternate color spaces, and output-intent conversion require sample-level normalization.
        if (colorSpace is not ("" or "DeviceGray" or "G" or "DeviceRGB" or "RGB") ||
            GetTransparencyMaskKind(stream.Dictionary, objects) != null ||
            PdfImageMaskNormalizer.IsImageMask(stream, objects)) return false;
        if (stream.Dictionary.Items.TryGetValue("SMaskInData", out PdfObject? embeddedMask) &&
            PdfObjectLookup.ResolveChain(objects, embeddedMask) is not (PdfNull or PdfNumber { Value: 0 })) return false;
        PdfObject? filter = stream.Dictionary.Items.TryGetValue("Filter", out PdfObject? value) ? value : null;
        PdfObject? resolved = PdfObjectLookup.ResolveChain(objects, filter);
        var filters = new List<PdfObject>();
        if (resolved is PdfName name) filters.Add(name);
        else if (resolved is PdfArray array) filters.AddRange(array.Items);
        else return false;
        if (filters.Count == 0 || PdfObjectLookup.ResolveChain(objects, filters[filters.Count - 1]) is not PdfName { Name: "JPXDecode" }) return false;
        var prefix = new PdfDictionary();
        var prefixFilters = new PdfArray();
        for (int i = 0; i < filters.Count - 1; i++) prefixFilters.Items.Add(filters[i]);
        prefix.Items["Filter"] = prefixFilters;
        if (stream.Dictionary.Items.TryGetValue("DecodeParms", out PdfObject? parameters)) {
            PdfObject? resolvedParameters = PdfObjectLookup.ResolveChain(objects, parameters);
            if (resolvedParameters is PdfArray parameterArray) {
                if (parameterArray.Items.Count != filters.Count ||
                    PdfObjectLookup.ResolveChain(objects, parameterArray.Items[filters.Count - 1]) is not PdfNull) return false;
                var prefixParameters = new PdfArray();
                for (int i = 0; i < filters.Count - 1; i++) prefixParameters.Items.Add(parameterArray.Items[i]);
                prefix.Items["DecodeParms"] = prefixParameters;
            } else if (resolvedParameters is not PdfNull) return false;
        }
        try {
            payload = Filters.StreamDecoder.DecodeRequired(prefix, stream.Data, objects, maximumBytes);
            // SMaskInData=0 (including absence) requires ignoring encoded alpha. Until sample-level
            // normalization is available, only prove opaque Gray/RGB headers safe for pass-through.
            if (!OfficeIMO.Drawing.OfficeJpeg2000Header.TryGetOpaqueComponents(payload, out int components)) return false;
            return colorSpace == "" || (components == 1 ? colorSpace is "DeviceGray" or "G" : colorSpace is "DeviceRGB" or "RGB");
        } catch (InvalidDataException) {
            return false;
        }
    }
}
