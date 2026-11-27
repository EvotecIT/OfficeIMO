using OfficeIMO.Excel.Xlsb.Biff12;
using OfficeIMO.Excel.Xlsb.Model;
using OfficeIMO.Excel.Xlsb.Projection;

namespace OfficeIMO.Excel.Xlsb.Write {
    /// <summary>Validates and encodes workbook calculation properties as BrtCalcProp.</summary>
    internal static class XlsbCalculationPropertiesWriter {
        private const int BrtCalcProp = 157;

        internal static void Write(Stream output, OfficeXmlNode properties) {
            ArgumentNullException.ThrowIfNull(output);
            if (properties.IsNone) return;
            XlsbCalculationProperties values = ReadValidated(properties);
            using var payload = new MemoryStream(26);
            XlsbPayloadWriter.WriteUInt32(payload, values.CalculationId);
            XlsbPayloadWriter.WriteUInt32(payload, values.CalculationMode);
            XlsbPayloadWriter.WriteUInt32(payload, values.IterationCount);
            XlsbPayloadWriter.WriteDouble(payload, values.IterationDelta);
            XlsbPayloadWriter.WriteUInt32(payload, checked((uint)values.ConcurrentThreadCount));
            XlsbPayloadWriter.WriteUInt16(payload, values.Flags);
            XlsbRecordWriter.Write(output, BrtCalcProp, payload.ToArray());
        }

        internal static void Validate(OfficeXmlNode properties) {
            if (!properties.IsNone) ReadValidated(properties);
        }

        private static XlsbCalculationProperties ReadValidated(OfficeXmlNode properties) {
            if (properties.FirstChild.Exists) {
                throw new NotSupportedException("Native XLSB generation does not support child content in calculation properties.");
            }
            string? unsupported = XlsbCalculationPropertiesProjector.UnsupportedAttribute(properties);
            if (unsupported != null) {
                throw new NotSupportedException($"Native XLSB generation does not yet support calculation property '{unsupported}'.");
            }

            XlsbCalculationProperties values = XlsbCalculationPropertiesProjector.Read(properties);
            double delta = values.IterationDelta;
            if (double.IsNaN(delta) || double.IsInfinity(delta) || delta < 0D) {
                throw new NotSupportedException("Native XLSB generation requires a finite, non-negative iteration delta.");
            }
            if (values.HasManualConcurrentCount &&
                (!values.ConcurrentCalculation || values.ConcurrentThreadCount < 1 || values.ConcurrentThreadCount > 1024)) {
                throw new NotSupportedException("Native XLSB generation requires concurrentManualCount from 1 through 1024 with concurrent calculation enabled.");
            }
            return values;
        }
    }
}
