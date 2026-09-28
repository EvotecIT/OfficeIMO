using System.Diagnostics;
using System.Globalization;
using OfficeIMO.Excel;

if (args.Length != 3 || args[0] is not ("typed-scan" or "sheet-names-100"))
    throw new ArgumentException("Pass profile, XLSX fixture path, and hexadecimal processor affinity mask.");

if (OperatingSystem.IsWindows())
    Process.GetCurrentProcess().ProcessorAffinity = new IntPtr(long.Parse(args[2], NumberStyles.HexNumber, CultureInfo.InvariantCulture));
Process.GetCurrentProcess().PriorityClass = ProcessPriorityClass.Normal;
Console.WriteLine("profile,iteration,elapsedMs,allocatedBytes,checksum");

if (args[0] == "sheet-names-100") {
    for (int batch = 0; batch < 21; batch++) {
        long before = GC.GetAllocatedBytesForCurrentThread();
        long started = Stopwatch.GetTimestamp();
        for (int repetition = 0; repetition < 100; repetition++) {
            IReadOnlyList<string> sheets = ExcelDocument.GetSheetNames(args[1]);
            if (sheets.Count != 1 || string.IsNullOrEmpty(sheets[0]))
                throw new InvalidDataException("Unexpected worksheet names.");
        }
        Console.WriteLine(FormattableString.Invariant(
            $"sheet-names-100,{batch},{Stopwatch.GetElapsedTime(started).TotalMilliseconds:F3},{GC.GetAllocatedBytesForCurrentThread() - before},"));
    }
} else {
    Span<int> bits = stackalloc int[4];
    for (int iteration = 0; iteration < 24; iteration++) {
        long before = GC.GetAllocatedBytesForCurrentThread();
        long started = Stopwatch.GetTimestamp();
        using var reader = ExcelDocument.OpenDataReader(args[1], new ExcelReadOptions { NumericAsDecimal = true });
        int rows = 0;
        ulong checksum = 14695981039346656037UL;
        while (reader.Read()) {
            for (int column = 0; column < 5; column++)
                checksum = unchecked(checksum * 1099511628211UL + (uint)reader.GetString(column).Length);
            checksum = unchecked(checksum * 1099511628211UL + (ulong)reader.GetDateTime(5).Ticks);
            checksum = unchecked(checksum * 1099511628211UL + (uint)reader.GetInt32(6));
            checksum = unchecked(checksum * 1099511628211UL + (ulong)reader.GetDateTime(7).Ticks);
            checksum = unchecked(checksum * 1099511628211UL + (uint)reader.GetInt32(8));
            for (int column = 9; column < 14; column++) {
                decimal.GetBits(reader.GetDecimal(column), bits);
                checksum = unchecked(checksum * 1099511628211UL + (uint)bits[0]);
            }
            rows++;
        }
        if (rows != 65535) throw new InvalidDataException($"Unexpected row count: {rows}.");
        reader.Dispose();
        Console.WriteLine(FormattableString.Invariant(
            $"typed-scan,{iteration},{Stopwatch.GetElapsedTime(started).TotalMilliseconds:F3},{GC.GetAllocatedBytesForCurrentThread() - before},{checksum}"));
    }
}
