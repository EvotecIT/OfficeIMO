using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel.Xlsb.Write {
    /// <summary>Validates and encodes workbook calculation properties as BrtCalcProp.</summary>
    internal static class XlsbCalculationPropertiesWriter {
        private const int BrtCalcProp = 157;

        internal static void Write(Stream output, CalculationProperties? properties) {
            if (output == null) throw new ArgumentNullException(nameof(output));
            if (properties == null) return;
            Xlsb.Biff12.XlsbRecordWriter.Write(output, BrtCalcProp, CreatePayload(properties));
        }

        internal static void Validate(CalculationProperties? properties) {
            if (properties == null) return;
            CreatePayload(properties);
        }

        private static byte[] CreatePayload(CalculationProperties properties) {
            if (properties.HasChildren) {
                throw new NotSupportedException("Native XLSB generation does not support child content in calculation properties.");
            }
            EnsureOnlyAttributes(properties,
                "calcId", "calcMode", "fullCalcOnLoad", "refMode", "iterate", "iterateCount",
                "iterateDelta", "fullPrecision", "calcCompleted", "calcOnSave", "concurrentCalc",
                "concurrentManualCount", "forceFullCalc");

            CalculateModeValues? calculationMode = properties.CalculationMode?.Value;
            if (calculationMode.HasValue
                && calculationMode.Value != CalculateModeValues.Manual
                && calculationMode.Value != CalculateModeValues.Auto
                && calculationMode.Value != CalculateModeValues.AutoNoTable) {
                throw new NotSupportedException($"Native XLSB generation cannot encode calculation mode '{calculationMode.Value}'.");
            }
            ReferenceModeValues? referenceMode = properties.ReferenceMode?.Value;
            if (referenceMode.HasValue
                && referenceMode.Value != ReferenceModeValues.A1
                && referenceMode.Value != ReferenceModeValues.R1C1) {
                throw new NotSupportedException($"Native XLSB generation cannot encode reference mode '{referenceMode.Value}'.");
            }

            uint mode = calculationMode == CalculateModeValues.Manual
                ? 0U
                : calculationMode == CalculateModeValues.AutoNoTable ? 2U : 1U;
            uint iterations = properties.IterateCount?.Value ?? 100U;
            double delta = properties.IterateDelta?.Value ?? 0.001D;
            if (double.IsNaN(delta) || double.IsInfinity(delta) || delta < 0D) {
                throw new NotSupportedException("Native XLSB generation requires a finite, non-negative iteration delta.");
            }

            uint? manualCount = properties.ConcurrentManualCount?.Value;
            bool concurrent = properties.ConcurrentCalculation?.Value ?? true;
            if (manualCount.HasValue && (!concurrent || manualCount.Value < 1U || manualCount.Value > 1024U)) {
                throw new NotSupportedException("Native XLSB generation requires concurrentManualCount from 1 through 1024 with concurrent calculation enabled.");
            }

            ushort flags = (ushort)((properties.FullCalculationOnLoad?.Value == true ? 0x0001 : 0)
                | (referenceMode == ReferenceModeValues.R1C1 ? 0 : 0x0002)
                | (properties.Iterate?.Value == true ? 0x0004 : 0)
                | (properties.FullPrecision?.Value == false ? 0 : 0x0008)
                | (properties.CalculationCompleted?.Value == false ? 0x0010 : 0)
                | (properties.CalculationOnSave?.Value == false ? 0 : 0x0020)
                | (concurrent ? 0x0040 : 0)
                | (manualCount.HasValue ? 0x0080 : 0)
                | (properties.ForceFullCalculation?.Value == true ? 0x0100 : 0));

            using var payload = new MemoryStream(26);
            WriteUInt32(payload, properties.CalculationId?.Value ?? 0U);
            WriteUInt32(payload, mode);
            WriteUInt32(payload, iterations);
            WriteDouble(payload, delta);
            WriteInt32(payload, manualCount.HasValue ? checked((int)manualCount.Value) : 1);
            WriteUInt16(payload, flags);
            return payload.ToArray();
        }

        private static void EnsureOnlyAttributes(OpenXmlElement element, params string[] allowedNames) {
            var allowed = new HashSet<string>(allowedNames, StringComparer.Ordinal);
            OpenXmlAttribute? unsupported = element.GetAttributes()
                .Cast<OpenXmlAttribute?>()
                .FirstOrDefault(attribute => attribute.HasValue
                    && !string.Equals(attribute.Value.NamespaceUri, "http://www.w3.org/2000/xmlns/", StringComparison.Ordinal)
                    && !allowed.Contains(attribute.Value.LocalName));
            if (unsupported.HasValue) {
                throw new NotSupportedException($"Native XLSB generation does not yet support calculation property '{unsupported.Value.LocalName}'.");
            }
        }

        private static void WriteDouble(Stream output, double value) {
            byte[] bytes = BitConverter.GetBytes(value);
            output.Write(bytes, 0, bytes.Length);
        }

        private static void WriteInt32(Stream output, int value) => WriteUInt32(output, unchecked((uint)value));

        private static void WriteUInt16(Stream output, ushort value) {
            output.WriteByte((byte)value);
            output.WriteByte((byte)(value >> 8));
        }

        private static void WriteUInt32(Stream output, uint value) {
            output.WriteByte((byte)value);
            output.WriteByte((byte)(value >> 8));
            output.WriteByte((byte)(value >> 16));
            output.WriteByte((byte)(value >> 24));
        }
    }
}
