using OfficeIMO.Excel.Xlsb.Biff12;
using OfficeIMO.Excel.Xlsb.Model;
using OfficeIMO.Excel.Xlsb.Projection;

namespace OfficeIMO.Excel.Xlsb.Write;

internal static partial class XlsbNativePackageWriter {
    private static byte[] RewriteCalculationProperties(ExcelDocument document, XlsbWorkbook source, byte[] original, long maximumBytes) {
        var properties = document.WorkbookRoot.CalculationProperties;
        if (XlsbCalculationPropertiesProjector.Matches(properties, source.CalculationProperties)) return original;
        using var input = new MemoryStream(original); var records = XlsbRecordReader.ReadAll(input);
        using var output = new OfficeIMO.Core.Internal.OfficeBoundedMemoryStream(maximumBytes); bool written = false;
        foreach (var record in records) {
            if (record.Type == 157 || record.Type == 132 && !written) {
                if (!written) XlsbCalculationPropertiesWriter.Write(output, properties);
                written = true; if (record.Type == 157) continue;
            }
            XlsbRecordWriter.Write(output, record.Type, record.Data);
        }
        return output.ToArray();
    }

    /// <summary>Changes caches and recalc flags without decoding or replacing formula tokens, including opaque native formulas.</summary>
    private static byte[] ApplyFormulaCachePolicy(byte[] original, ExcelNativeCalculationPolicy policy, long maximumBytes) {
        if (!policy.ClearCachedResults && !policy.RecalculateOnOpen) return original;
        using var input = new MemoryStream(original); var records = XlsbRecordReader.ReadAll(input);
        using var output = new OfficeIMO.Core.Internal.OfficeBoundedMemoryStream(maximumBytes);
        bool inData = false; int futureDepth = 0;
        foreach (var record in records) {
            byte[] payload = record.Data; int type = record.Type;
            if (type == 145) inData = true;
            if (type == 146) inData = false;
            if (type == 35) futureDepth++;
            if (type == 36) futureDepth--;
            if (inData && futureDepth == 0 && type is >= 8 and <= 11) {
                int formula = type == 8 ? checked(12 + (int)XlsbFormulaTokenScanner.ReadUInt32(payload, 8) * 2) : type == 9 ? 16 : 9;
                if (formula > payload.Length - 6) throw new InvalidDataException("The native XLSB formula cache is truncated.");
                if (policy.ClearCachedResults) {
                    // BIFF12 has no empty-cache formula record. A zero-length string cache plus full recalculation is the native blank representation.
                    byte[] cleared = new byte[checked(12 + payload.Length - formula)];
                    Buffer.BlockCopy(payload, 0, cleared, 0, 8); Buffer.BlockCopy(payload, formula, cleared, 12, payload.Length - formula);
                    payload = cleared; formula = 12; type = 8;
                } else payload = (byte[])payload.Clone();
                if (policy.RecalculateOnOpen) payload[formula] |= 2; // GrbitFmla.fAlwaysCalc (bit 1; bit 0 is reserved)
            }
            XlsbRecordWriter.Write(output, type, payload);
        }
        return output.ToArray();
    }
}
