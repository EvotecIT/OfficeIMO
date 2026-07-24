using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.CSV;
using Xunit;

namespace OfficeIMO.CSV.Tests;

public class CsvSaveApiTests
{
    [Fact]
    public void Save_Stream_And_ToBytes_Produce_Equivalent_Output()
    {
        var document = CreateDocument();
        var options = new CsvSaveOptions { NewLine = "\n" };
        byte[] expected = document.ToBytes(options);

        using var stream = new MemoryStream();
        document.Save(stream, options);
        using MemoryStream encoded = document.ToStream(options);

        Assert.Equal(expected, stream.ToArray());
        Assert.Equal(expected, encoded.ToArray());
        Assert.Equal(0, encoded.Position);
        Assert.Equal("Name,Value\nAlpha,1\n", Encoding.UTF8.GetString(expected));
    }

    [Fact]
    public async Task SaveAsync_Path_Infers_GZip_Compression()
    {
        string path = Path.Combine(Path.GetTempPath(), "OfficeIMO.CSV.Async." + Guid.NewGuid().ToString("N") + ".csv.gz");
        try
        {
            await CreateDocument().SaveAsync(path, new CsvSaveOptions { NewLine = "\n" });

            using var file = File.OpenRead(path);
            using var gzip = new GZipStream(file, CompressionMode.Decompress);
            using var reader = new StreamReader(gzip, Encoding.UTF8);
            Assert.Equal("Name,Value\nAlpha,1\n", await reader.ReadToEndAsync());
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public async Task SaveAsync_Append_Honors_Path_Policy_Without_Leaking_It_Into_Serialization()
    {
        string path = Path.Combine(Path.GetTempPath(), "OfficeIMO.CSV.AsyncAppend." + Guid.NewGuid().ToString("N") + ".csv");
        try
        {
            File.WriteAllText(path, "Name,Value\nAlpha,1\n");
            await new CsvDocument()
                .WithHeader("Name", "Value")
                .AddRow("Beta", 2)
                .SaveAsync(path, new CsvSaveOptions { Append = true, IncludeHeader = false, NewLine = "\n" });

            Assert.Equal("Name,Value\nAlpha,1\nBeta,2\n", File.ReadAllText(path));
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task SaveAsync_Append_Does_Not_Write_Encoding_Preamble_Into_Existing_Content(bool useUtf16)
    {
        string path = Path.Combine(Path.GetTempPath(), "OfficeIMO.CSV.AsyncAppendBom." + Guid.NewGuid().ToString("N") + ".csv");
        Encoding encoding = useUtf16 ? Encoding.Unicode : new UTF8Encoding(encoderShouldEmitUTF8Identifier: true);
        try
        {
            File.WriteAllText(path, "Name,Value\nAlpha,1\n", encoding);
            await new CsvDocument()
                .WithHeader("Name", "Value")
                .AddRow("Beta", 2)
                .SaveAsync(path, new CsvSaveOptions {
                    Append = true,
                    IncludeHeader = false,
                    NewLine = "\n",
                    Encoding = encoding
                });

            byte[] bytes = File.ReadAllBytes(path);
            byte[] preamble = encoding.GetPreamble();
            Assert.True(bytes.Take(preamble.Length).SequenceEqual(preamble));
            Assert.Equal(-1, FindSequence(bytes, preamble, preamble.Length));
            Assert.Equal("Name,Value\nAlpha,1\nBeta,2\n", File.ReadAllText(path, encoding));
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void Save_Path_Creates_Missing_Parent_Directory()
    {
        string directory = CreateMissingDirectoryPath();
        string path = Path.Combine(directory, "document.csv");
        try
        {
            CreateDocument().Save(path, new CsvSaveOptions { NewLine = "\n" });

            Assert.Equal("Name,Value\nAlpha,1\n", File.ReadAllText(path));
        }
        finally
        {
            DeleteDirectoryIfExists(directory);
        }
    }

    [Fact]
    public void Save_Append_Creates_Missing_Parent_Directory()
    {
        string directory = CreateMissingDirectoryPath();
        string path = Path.Combine(directory, "document.csv");
        try
        {
            CreateDocument().Save(path, new CsvSaveOptions { Append = true, NewLine = "\n" });

            Assert.Equal("Name,Value\nAlpha,1\n", File.ReadAllText(path));
        }
        finally
        {
            DeleteDirectoryIfExists(directory);
        }
    }

    [Fact]
    public void CreateTextWriter_NoClobberClaimsDestinationAtomically()
    {
        string path = Path.Combine(Path.GetTempPath(), "OfficeIMO.CSV.NoClobber." + Guid.NewGuid().ToString("N") + ".csv");
        try
        {
            File.WriteAllText(path, "existing");

            Assert.Throws<IOException>(() =>
                CsvFile.CreateTextWriter(path, new CsvSaveOptions { NoClobber = true }));

            Assert.Equal("existing", File.ReadAllText(path));
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public async Task SaveAsync_Append_Creates_Missing_Parent_Directory()
    {
        string directory = CreateMissingDirectoryPath();
        string path = Path.Combine(directory, "document.csv");
        try
        {
            await CreateDocument().SaveAsync(path, new CsvSaveOptions { Append = true, NewLine = "\n" });

            Assert.Equal("Name,Value\nAlpha,1\n", File.ReadAllText(path));
        }
        finally
        {
            DeleteDirectoryIfExists(directory);
        }
    }

    [Fact]
    public async Task SaveAsync_Rejects_Compressed_Append_Inferred_From_Path()
    {
        string path = Path.Combine(Path.GetTempPath(), "OfficeIMO.CSV.AsyncAppend." + Guid.NewGuid().ToString("N") + ".csv.gz");
        await Assert.ThrowsAsync<NotSupportedException>(() =>
            CreateDocument().SaveAsync(path, new CsvSaveOptions { Append = true }));
        Assert.False(File.Exists(path));
    }

    [Fact]
    public async Task SaveAsync_PreCanceled_Path_Does_Not_Create_File()
    {
        string path = Path.Combine(Path.GetTempPath(), "OfficeIMO.CSV.Canceled." + Guid.NewGuid().ToString("N") + ".csv");
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            CreateDocument().SaveAsync(path, cancellationToken: new CancellationToken(canceled: true)));
        Assert.False(File.Exists(path));
    }

    private static CsvDocument CreateDocument() => new CsvDocument()
        .WithHeader("Name", "Value")
        .AddRow("Alpha", 1);

    private static int FindSequence(byte[] bytes, byte[] sequence, int startIndex)
    {
        for (int index = startIndex; index <= bytes.Length - sequence.Length; index++)
        {
            bool matches = true;
            for (int offset = 0; offset < sequence.Length; offset++)
            {
                if (bytes[index + offset] == sequence[offset]) continue;
                matches = false;
                break;
            }

            if (matches) return index;
        }

        return -1;
    }

    private static string CreateMissingDirectoryPath() =>
        Path.Combine(Path.GetTempPath(), "OfficeIMO.CSV.Save." + Guid.NewGuid().ToString("N"));

    private static void DeleteDirectoryIfExists(string path)
    {
        if (Directory.Exists(path)) Directory.Delete(path, recursive: true);
    }
}
