#if NET10_0_OR_GREATER
namespace OfficeIMO.Excel.Benchmarks;

/// <summary>Runs the validated native writer workload without benchmark reporting during CPU sampling.</summary>
internal static class ExcelNativeGeneratedWriteProfileRunner {
    internal static void Run(string[] args) {
        ExcelFileFormat format = args.Length > 1 ? Enum.Parse<ExcelFileFormat>(args[1], ignoreCase: true) : ExcelFileFormat.Xlsb;
        int iterations = args.Length > 2 ? int.Parse(args[2], System.Globalization.CultureInfo.InvariantCulture) : 200;
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(iterations);
        if (format != ExcelFileFormat.Xls && format != ExcelFileFormat.Xlsb) throw new ArgumentException("Choose XLS or XLSB.");
        var workload = new ExcelReaderGeneratedWriteBenchmarks { Format = format, RowCount = 50_000 };
        workload.SetupOfficeIMO();
        for (int index = 0; index < 8; index++) workload.OfficeIMO();
        byte[] output = [];
        for (int index = 0; index < iterations; index++) output = workload.OfficeIMO();
        workload.Validate(output);
        Console.WriteLine($"Validated {iterations} {format} writes of {workload.RowCount} rows; final package {output.Length} bytes.");
    }
}
#endif
