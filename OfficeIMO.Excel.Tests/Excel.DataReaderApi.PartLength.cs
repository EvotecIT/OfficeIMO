using System.IO.Compression;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Excel.Tests;

public partial class Excel {
    [Theory]
    [InlineData(true, false)]
    [InlineData(true, true)]
    [InlineData(false, false)]
    public void DataReader_RejectsWorksheetWithOverstatedPackageLength(bool declaredDimension, bool prefetch) {
        string path = CreateCompactFastPathWorkbook();
        try {
            string dimension = declaredDimension ? "<dimension ref=\"A1:B500002\"/>" : "";
            string xml = "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" + dimension
                + "<sheetData><row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>Id</t></is></c>"
                + "<c r=\"B1\" t=\"inlineStr\"><is><t>Value</t></is></c></row>"
                + "<row r=\"2\"><c r=\"A2\"><v>42</v></c></row>" + new string(' ', 70000)
                + "<row r=\"500002\"><c r=\"A500002\"><v>43</v></c></row></sheetData></worksheet>";
            ReplaceZipEntry(path, "xl/worksheets/sheet1.xml", Encoding.UTF8.GetBytes(xml));
            byte[] bytes = File.ReadAllBytes(path);
            OpenXmlPartLengthTests.SetDeclaredLength(bytes, "xl/worksheets/sheet1.xml", 40 * 1024 * 1024);
            File.WriteAllBytes(path, bytes);

            Exception? error = Record.Exception(() => {
                using var reader = ExcelDocument.OpenDataReader(path,
                    new ExcelReadOptions { EnableWorksheetPrefetch = prefetch });
            });
            Assert.True(error is InvalidDataException or EndOfStreamException,
                error?.ToString() ?? "The corrupt worksheet exposed a data reader.");
        } finally {
            File.Delete(path);
        }
    }
}

public class OpenXmlPartLengthTests {
    [Theory]
    [InlineData(128, 140)]
    [InlineData(8193, 8194)]
    [InlineData(16384, 16385)]
    public async Task BufferedAndStreamingPartsRejectPrematureEnd(int actualLength, int declaredLength) {
        foreach (string route in new[] { "array", "byte", "async" }) {
            byte[] package = Package(actualLength);
            SetDeclaredLength(package, "xl/test.xml", declaredLength);
            using var parts = OpenXmlPackagePartBufferReader.TryOpen(package)!;
            Exception? error = await Record.ExceptionAsync(async () => {
                using Stream part = parts.OpenPart("xl/test.xml", 32768);
                byte[] buffer = new byte[257];
                if (route == "byte") {
                    while (part.ReadByte() >= 0) { }
                } else if (route == "async") {
                    while (await part.ReadAsync(buffer, 0, buffer.Length) != 0) { }
                } else {
                    while (part.Read(buffer, 0, buffer.Length) != 0) { }
                }
            });
            Assert.True(error is InvalidDataException or EndOfStreamException,
                $"{route}: {error?.ToString() ?? "The corrupt part was accepted."}");
        }
    }

    [Fact]
    public async Task EmptyReadsAndEarlyDisposalDoNotRequireEndOfPart() {
        using var parts = OpenXmlPackagePartBufferReader.TryOpen(Package(8193))!;
        using (Stream part = parts.OpenPart("xl/test.xml", 8193)) {
            Assert.Equal(0, part.Read(Array.Empty<byte>(), 0, 0));
            Assert.Equal(0, await part.ReadAsync(Array.Empty<byte>(), 0, 0));
#if NET8_0_OR_GREATER
            Assert.Equal(0, part.Read(Span<byte>.Empty));
            Assert.Equal(0, await part.ReadAsync(Memory<byte>.Empty));
#endif
            Assert.Equal(0, part.ReadByte());
        }
        using Stream reopened = parts.OpenPart("xl/test.xml", 8193);
        using var output = new MemoryStream();
        await reopened.CopyToAsync(output);
        Assert.Equal(Content(8193), output.ToArray());
        Assert.Equal(-1, reopened.ReadByte());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task PartLengthStreamRejectsUnclippedOverrun(bool asynchronous) {
        using Stream part = new OpenXmlPartLengthStream(new MemoryStream(Content(16384)), 16383, "xl/test.xml");
        byte[] buffer = new byte[257];
        await Assert.ThrowsAsync<InvalidDataException>(async () => {
            while ((asynchronous ? await part.ReadAsync(buffer, 0, buffer.Length) : part.Read(buffer, 0, buffer.Length)) != 0) { }
        });
    }

#if NET8_0_OR_GREATER
    [Theory]
    [InlineData(16383)]
    [InlineData(16384)]
    [InlineData(16385)]
    public async Task SpanAndMemoryReadsEnforceTheSamePartLength(int declaredLength) {
        foreach (bool asynchronous in new[] { false, true }) {
            // Deflate streams on modern .NET can cap output at the ZIP length.
            // Exercise the wrapper's two-sided invariant with an unclipped stream.
            using Stream part = new OpenXmlPartLengthStream(new MemoryStream(Content(16384)), declaredLength, "xl/test.xml");
            using var output = new MemoryStream();
            byte[] buffer = new byte[257];
            Exception? error = await Record.ExceptionAsync(async () => {
                int read;
                while ((read = asynchronous ? await part.ReadAsync(buffer.AsMemory()) : part.Read(buffer.AsSpan())) != 0) {
                    output.Write(buffer, 0, read);
                }
            });
            if (declaredLength == 16384) {
                Assert.Null(error);
                Assert.Equal(Content(16384), output.ToArray());
            } else {
                Assert.IsType<InvalidDataException>(error);
            }
        }
    }
#endif

    private static byte[] Content(int length) => Enumerable.Range(0, length).Select(index => (byte)(index % 251)).ToArray();

    private static byte[] Package(int length) {
        using var output = new MemoryStream();
        using (var archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true)) {
            using Stream entry = archive.CreateEntry("xl/test.xml").Open();
            byte[] content = Content(length);
            entry.Write(content, 0, content.Length);
        }
        return output.ToArray();
    }

    internal static void SetDeclaredLength(byte[] package, string partName, int length) {
        for (int offset = 0; offset <= package.Length - 46; offset++) {
            if (BitConverter.ToUInt32(package, offset) != 0x02014B50U) continue;
            int nameLength = BitConverter.ToUInt16(package, offset + 28);
            if (offset + 46 + nameLength > package.Length) continue;
            if (Encoding.UTF8.GetString(package, offset + 46, nameLength) != partName) continue;
            Array.Copy(BitConverter.GetBytes(length), 0, package, offset + 24, 4);
            return;
        }
        throw new InvalidOperationException($"Package part '{partName}' was not found.");
    }
}
