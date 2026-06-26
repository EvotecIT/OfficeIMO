namespace OfficeIMO.Excel.LegacyXls.Model {
    internal static class LegacyXlsChartLayoutModeName {
        internal static string GetName(ushort mode) {
            return mode switch {
                0x0000 => "Automatic",
                0x0001 => "Factor",
                0x0002 => "Edge",
                _ => $"Unknown:0x{mode:X4}"
            };
        }
    }
}
