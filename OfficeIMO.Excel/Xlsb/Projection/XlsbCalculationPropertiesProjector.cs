using OfficeIMO.Excel.Xlsb.Model;

namespace OfficeIMO.Excel.Xlsb.Projection {
    /// <summary>Projects and compares the SpreadsheetML representation of BrtCalcProp.</summary>
    internal static class XlsbCalculationPropertiesProjector {
        internal const string ElementName = "calcPr";
        private static readonly string[] AttributeNames = {
            "calcId", "calcMode", "fullCalcOnLoad", "refMode", "iterate", "iterateCount",
            "iterateDelta", "fullPrecision", "calcCompleted", "calcOnSave", "concurrentCalc",
            "concurrentManualCount", "forceFullCalc"
        };

        internal static void Apply(OfficeXmlNode workbook, XlsbCalculationProperties source) {
            OfficeXmlNode properties = XlsbWorkbookMarkup.Add(workbook, ElementName);
            properties.SetUInt32Attribute("calcId", source.CalculationId);
            properties.SetAttribute("calcMode", source.CalculationMode switch { 0 => "manual", 2 => "autoNoTable", _ => "auto" });
            properties.SetBooleanAttribute("fullCalcOnLoad", source.FullCalculationOnLoad);
            properties.SetAttribute("refMode", source.UsesA1References ? "A1" : "R1C1");
            properties.SetBooleanAttribute("iterate", source.IterationEnabled);
            properties.SetUInt32Attribute("iterateCount", source.IterationCount);
            properties.SetDoubleAttribute("iterateDelta", source.IterationDelta);
            properties.SetBooleanAttribute("fullPrecision", source.FullPrecision);
            properties.SetBooleanAttribute("calcCompleted", source.CalculationCompleted);
            properties.SetBooleanAttribute("calcOnSave", source.CalculationOnSave);
            properties.SetBooleanAttribute("concurrentCalc", source.ConcurrentCalculation);
            properties.SetBooleanAttribute("forceFullCalc", source.ForceFullCalculation);
            if (source.HasManualConcurrentCount) properties.SetUInt32Attribute("concurrentManualCount", checked((uint)source.ConcurrentThreadCount));
        }

        internal static bool Matches(OfficeXmlNode actual, XlsbCalculationProperties? expected) {
            if (actual.IsNone || expected == null) return actual.IsNone && expected == null;
            if (actual.FirstChild.Exists || UnsupportedAttribute(actual) != null) return false;

            XlsbCalculationProperties current = Read(actual);
            return current.CalculationId == expected.CalculationId
                && current.CalculationMode == expected.CalculationMode
                && current.IterationCount == expected.IterationCount
                && current.IterationDelta.Equals(expected.IterationDelta)
                && current.Flags == (expected.Flags & 0x01FF)
                && (!expected.HasManualConcurrentCount || current.ConcurrentThreadCount == expected.ConcurrentThreadCount);
        }

        internal static string? UnsupportedAttribute(OfficeXmlNode properties) =>
            XlsbWorkbookMarkup.UnsupportedAttribute(properties, AttributeNames);

        /// <summary>Reads the BIFF12 fields with the SpreadsheetML defaults for omitted attributes.</summary>
        internal static XlsbCalculationProperties Read(OfficeXmlNode properties) {
            uint mode = properties.Attribute("calcMode") switch {
                null or "auto" => 1U,
                "manual" => 0U,
                "autoNoTable" => 2U,
                string value => throw new NotSupportedException($"Native XLSB generation cannot encode calculation mode '{value}'.")
            };
            bool a1 = properties.Attribute("refMode") switch {
                null or "A1" => true,
                "R1C1" => false,
                string value => throw new NotSupportedException($"Native XLSB generation cannot encode reference mode '{value}'.")
            };

            uint? manualCount = XlsbWorkbookMarkup.ReadUInt32(properties, "concurrentManualCount");
            ushort flags = (ushort)((XlsbWorkbookMarkup.ReadBoolean(properties, "fullCalcOnLoad") == true ? 0x0001 : 0)
                | (a1 ? 0x0002 : 0)
                | (XlsbWorkbookMarkup.ReadBoolean(properties, "iterate") == true ? 0x0004 : 0)
                | (XlsbWorkbookMarkup.ReadBoolean(properties, "fullPrecision") == false ? 0 : 0x0008)
                | (XlsbWorkbookMarkup.ReadBoolean(properties, "calcCompleted") == false ? 0x0010 : 0)
                | (XlsbWorkbookMarkup.ReadBoolean(properties, "calcOnSave") == false ? 0 : 0x0020)
                | (XlsbWorkbookMarkup.ReadBoolean(properties, "concurrentCalc") == false ? 0 : 0x0040)
                | (manualCount.HasValue ? 0x0080 : 0)
                | (XlsbWorkbookMarkup.ReadBoolean(properties, "forceFullCalc") == true ? 0x0100 : 0));
            if (manualCount > int.MaxValue) {
                throw new NotSupportedException("Native XLSB generation requires concurrentManualCount from 1 through 1024 with concurrent calculation enabled.");
            }

            double delta = properties.Attribute("iterateDelta") is string valueText ? System.Xml.XmlConvert.ToDouble(valueText) : 0.001D;
            return new XlsbCalculationProperties(XlsbWorkbookMarkup.ReadUInt32(properties, "calcId") ?? 0U, mode,
                XlsbWorkbookMarkup.ReadUInt32(properties, "iterateCount") ?? 100U, delta, (int)(manualCount ?? 1U), flags);
        }
    }
}
