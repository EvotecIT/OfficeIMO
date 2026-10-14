using System;
using System.IO;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeManagedTextShapingProvider {
    private static LatinFont? ResolveLatinFont(OfficeTextShapingRequest request) {
        if (!request.ApplyDefaultLatinLigatures) return null;
        // Built-in instance keys include the selected face and axes. Raw bytes alone
        // identify only the source program, so they are safe for default-face requests.
        object? key = request.FontProgramCacheKeyForShaping is OfficeTrueTypeFont or OfficeOpenTypeCffFont
            ? request.FontProgramCacheKeyForShaping : null;
        if (key == null && request.VariationCoordinatesForShaping.Count == 0 &&
            !request.FontCollectionIndex.HasValue) key = request.FontDataForShaping;
        return key == null ? new LatinFont(request) : LatinFonts.GetValue(key, _ => new LatinFont(request));
    }

    private static IOfficeFontProgram? LoadShapingFont(OfficeTextShapingRequest request) {
        if (request.IsOpenTypeCff)
            return OfficeOpenTypeCffFont.TryLoad(request.FontDataForShaping,
                request.VariationCoordinatesForShaping, out _);
        if (request.VariationCoordinatesForShaping.Count == 0)
            return OfficeTrueTypeFont.TryLoad(request.FontDataForShaping, request.FontCollectionIndex);

        // Variable collections use the same boundary as font registration: callers
        // must extract the selected face before supplying variation coordinates.
        byte[] data = request.FontDataForShaping;
        if (request.FontCollectionIndex > 0 ||
            data.Length >= 4 && data[0] == 't' && data[1] == 't' && data[2] == 'c' && data[3] == 'f') return null;
        try {
            OfficeOpenTypeReader? reader = OfficeOpenTypeReader.TryCreate(data);
            if (reader == null) return null;
            OfficeFontVariationModel model = OfficeFontVariationModel.Create(reader,
                request.VariationCoordinatesForShaping);
            return OfficeTrueTypeFont.TryLoad(data, model, out _);
        } catch (Exception exception) when (exception is InvalidDataException
                                            || exception is ArgumentException
                                            || exception is NotSupportedException
                                            || exception is OverflowException
                                            || exception is IndexOutOfRangeException) {
            return null;
        }
    }
}
