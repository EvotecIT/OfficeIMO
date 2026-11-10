using System.Text;
using OfficeIMO.Data.Conversions;

namespace OfficeIMO.Data.Conversions.Tests;

internal static class Convert {
    public static (string Output, OfficeConversionReport Report) Run(string input, DataFormat source, DataFormat target, DataConversionOptions? options = null) {
        using var output = new MemoryStream();
        OfficeConversionReport report = DataConversionEngine.Convert(new MemoryStream(Encoding.UTF8.GetBytes(input)), source, output, target, options);
        return (Encoding.UTF8.GetString(output.ToArray()), report);
    }

    public static string[] Codes(OfficeConversionReport report) => report.Diagnostics.Select(d => d.Code).ToArray();
}
