using System.IO.Compression;
using OfficeIMO.Excel;

/// <summary>Exercises buffered XLSB compression under NativeAOT.</summary>
internal static class BufferedXlsbSaveSmoke {
    internal static void Run() {
        foreach (byte[] bytes in SaveWithConstrainedThreadPool()) Validate(bytes);
        Console.WriteLine("PASS | Buffered XLSB NativeAOT save: 4096 ordered rows.");
    }

    private static byte[][] SaveWithConstrainedThreadPool() {
        ThreadPool.GetMinThreads(out int minimum, out int minimumIo);
        ThreadPool.GetMaxThreads(out int maximum, out int maximumIo);
        using var ready = new CountdownEvent(2);
        using var start = new ManualResetEventSlim();
        Task<byte[]>[] saves = [];
        try {
            // The Windows native pool does not support changing its limits.
            if (!ThreadPool.SetMinThreads(1, minimumIo) || !ThreadPool.SetMaxThreads(2, maximumIo)) {
                Console.WriteLine("SKIP | Bounded-pool save check: this pool does not support configurable limits.");
                return [Save()];
            }
            ThreadPool.GetMaxThreads(out int appliedMaximum, out _);
            if (appliedMaximum != 2) throw new InvalidOperationException("The isolated pool maximum was not applied.");
            saves = Enumerable.Range(0, 2).Select(_ => Task.Run(() => {
                ready.Signal();
                start.Wait();
                return Save();
            })).ToArray();
            if (!ready.Wait(TimeSpan.FromSeconds(10))) throw new TimeoutException("Concurrent save callers did not start.");
            start.Set();
            if (!Task.WaitAll(saves, TimeSpan.FromSeconds(10)))
                throw new TimeoutException("Buffered XLSB saves stalled in a bounded thread pool.");
            Console.WriteLine("PASS | Concurrent buffered XLSB saves with a two-worker pool maximum.");
            return saves.Select(save => save.Result).ToArray();
        } finally {
            start.Set();
            ThreadPool.SetMaxThreads(maximum, maximumIo);
            ThreadPool.SetMinThreads(minimum, minimumIo);
            if (!Task.WaitAll(saves, TimeSpan.FromSeconds(20)))
                throw new TimeoutException("Concurrent saves did not finish after restoring the pool.");
        }
    }

    private static byte[] Save() {
        using ExcelDocument document = ExcelDocument.Create();
        document.AddSheet("Buffered").WriteData(Enumerable.Range(0, 4096),
            ("Id", static row => row), ("Label", static row => "_x0041_ & " + row));
        return document.ToBytes(new ExcelSaveOptions { Format = ExcelFileFormat.Xlsb });
    }

    private static void Validate(byte[] bytes) {
        using var source = new MemoryStream(bytes, writable: false);
        using (var package = new ZipArchive(source, ZipArchiveMode.Read, leaveOpen: true)) {
            if (package.GetEntry("xl/worksheets/sheet1.bin") == null)
                throw new InvalidOperationException("The buffered XLSB smoke did not produce a binary worksheet.");
        }
        source.Position = 0;
        using ExcelWorkbookDataReader reader = ExcelDocument.OpenReader(source);
        if (reader.FieldCount != 2 || reader.GetName(0) != "Id" || reader.GetName(1) != "Label")
            throw new InvalidOperationException("The buffered XLSB smoke lost its schema.");
        for (int row = 0; row < 4096; row++) {
            if (!reader.Read() || reader.GetInt32(0) != row || reader.GetString(1) != "_x0041_ & " + row)
                throw new InvalidOperationException("The buffered XLSB smoke lost values or their order.");
        }
        if (reader.Read() || reader.NextResult())
            throw new InvalidOperationException("The buffered XLSB smoke added rows or worksheets.");
    }
}
