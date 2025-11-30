using System;
using System.Diagnostics.Tracing;
using System.IO;
using System.IO.Compression;
using System.Linq;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests
{
    public partial class Excel
    {
        [Fact]
        public void Load_LargeWorkbook_StreamsWithoutLargeArrayAllocation()
        {
            string filePath = Path.Combine(_directoryWithFiles, $"LargeLoad_{Guid.NewGuid():N}.xlsx");

            try
            {
                CreateLargeWorkbook(filePath, payloadBytes: 6 * 1024 * 1024);

                long fileSize = new FileInfo(filePath).Length;
                Assert.True(fileSize > 5 * 1024 * 1024, $"Expected a file larger than 5MB, but found {fileSize} bytes.");

                using var allocationListener = new AllocationListener();
                using (var document = ExcelDocument.Load(filePath, readOnly: true))
                {
                    Assert.NotEmpty(document.Sheets);
                    Assert.Contains(document.Sheets, sheet => string.Equals(sheet.Name, "Large", StringComparison.Ordinal));
                }

                // Ensure we did not allocate a buffer roughly equal to the full package size.
                Assert.True(
                    allocationListener.MaxAllocation < fileSize * 0.9,
                    $"Expected maximum single allocation ({allocationListener.MaxAllocation} bytes) to stay below 90% of the file size ({fileSize} bytes)."
                );
            }
            finally
            {
                if (File.Exists(filePath))
                {
                    File.Delete(filePath);
                }
            }
        }

        private static void CreateLargeWorkbook(string filePath, int payloadBytes)
        {
            using var document = ExcelDocument.Create(filePath);
            var sheet = document.AddWorkSheet("Large");
            sheet.CellValue(1, 1, "Payload");
            document.Save();

            var buffer = new byte[payloadBytes];
            new Random(42).NextBytes(buffer);

            using var archive = ZipFile.Open(filePath, ZipArchiveMode.Update);
            var entry = archive.CreateEntry("xl/media/payload.bin", CompressionLevel.NoCompression);
            using var entryStream = entry.Open();
            entryStream.Write(buffer, 0, buffer.Length);
        }

        private sealed class AllocationListener : EventListener
        {
            private const string RuntimeProviderName = "System.Runtime";

            public long MaxAllocation { get; private set; }

            protected override void OnEventSourceCreated(EventSource eventSource)
            {
                if (eventSource?.Name == RuntimeProviderName)
                {
                    EnableEvents(eventSource, EventLevel.Informational, (EventKeywords)0x1);
                }
            }

            protected override void OnEventWritten(EventWrittenEventArgs eventData)
            {
                if (eventData == null || eventData.PayloadNames == null || eventData.Payload == null)
                {
                    return;
                }

                if (!string.Equals(eventData.EventName, "GCAllocationTick", StringComparison.Ordinal))
                {
                    return;
                }

                for (int i = 0; i < eventData.PayloadNames.Count; i++)
                {
                    string? payloadName = eventData.PayloadNames[i];
                    if (!string.Equals(payloadName, "AllocationAmount64", StringComparison.Ordinal) &&
                        !string.Equals(payloadName, "AllocationAmount", StringComparison.Ordinal))
                    {
                        continue;
                    }

                    var value = eventData.Payload[i];
                    long allocation = value switch
                    {
                        long longValue => longValue,
                        int intValue => intValue,
                        _ => 0
                    };

                    if (allocation > MaxAllocation)
                    {
                        MaxAllocation = allocation;
                    }
                }
            }
        }
    }
}
