using System.Text;
using PSWriteOffice.Services;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Tests;

public sealed class AtomicFileWriterTests
{
    [Fact]
    public void WriteUnique_writes_payload_once_when_names_collide()
    {
        var directory = Path.Combine(Path.GetTempPath(), $"PSWriteOffice-{Guid.NewGuid():N}");
        Directory.CreateDirectory(directory);

        try
        {
            File.WriteAllText(Path.Combine(directory, "report.txt"), "existing-1");
            File.WriteAllText(Path.Combine(directory, "report-2.txt"), "existing-2");
            var writeCount = 0;

            var outputPath = AtomicFileWriter.WriteUnique(directory, "report.txt", temporaryPath =>
            {
                writeCount++;
                File.WriteAllText(temporaryPath, "new-payload", new UTF8Encoding(false));
            });

            Assert.Equal(1, writeCount);
            Assert.Equal("report-3.txt", Path.GetFileName(outputPath));
            Assert.Equal("new-payload", File.ReadAllText(outputPath));
            Assert.Equal("existing-1", File.ReadAllText(Path.Combine(directory, "report.txt")));
            Assert.Equal("existing-2", File.ReadAllText(Path.Combine(directory, "report-2.txt")));
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void WriteUnique_cleans_temporary_file_when_writer_fails()
    {
        var directory = Path.Combine(Path.GetTempPath(), $"PSWriteOffice-{Guid.NewGuid():N}");
        Directory.CreateDirectory(directory);

        try
        {
            Assert.Throws<InvalidOperationException>(() =>
                AtomicFileWriter.WriteUnique(directory, "report.txt", temporaryPath =>
                {
                    File.WriteAllText(temporaryPath, "partial");
                    throw new InvalidOperationException("writer failed");
                }));

            Assert.Empty(Directory.GetFiles(directory));
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void WriteUnique_keeps_temporary_name_within_component_limit()
    {
        var directory = Path.Combine(Path.GetTempPath(), $"PSWriteOffice-{Guid.NewGuid():N}");
        Directory.CreateDirectory(directory);

        try
        {
            var fileName = new string('a', 220) + ".bin";

            var outputPath = AtomicFileWriter.WriteUnique(directory, fileName, Encoding.UTF8.GetBytes("payload"));

            Assert.Equal(fileName, Path.GetFileName(outputPath));
            Assert.Equal("payload", File.ReadAllText(outputPath));
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void WriteUnique_truncates_initial_name_to_portable_component_limit()
    {
        var directory = Path.Combine(Path.GetTempPath(), $"PSWriteOffice-{Guid.NewGuid():N}");
        Directory.CreateDirectory(directory);

        try
        {
            var fileName = new string('a', 260) + ".bin";

            var outputPath = AtomicFileWriter.WriteUnique(directory, fileName, Encoding.UTF8.GetBytes("payload"));
            var outputName = Path.GetFileName(outputPath);

            Assert.Equal(255, outputName.Length);
            Assert.Equal(new string('a', 251) + ".bin", outputName);
            Assert.Equal("payload", File.ReadAllText(outputPath));
            Assert.Empty(Directory.GetFiles(directory, ".*.tmp*"));
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void WriteUnique_truncates_initial_unicode_name_to_portable_byte_limit()
    {
        var directory = Path.Combine(Path.GetTempPath(), $"PSWriteOffice-{Guid.NewGuid():N}");
        Directory.CreateDirectory(directory);

        try
        {
            var fileName = new string('\u00e9', 130) + ".bin";

            var outputPath = AtomicFileWriter.WriteUnique(directory, fileName, Encoding.UTF8.GetBytes("payload"));
            var outputName = Path.GetFileName(outputPath);

            Assert.True(Encoding.UTF8.GetByteCount(outputName) <= 255);
            Assert.EndsWith(".bin", outputName, StringComparison.Ordinal);
            Assert.Equal("payload", File.ReadAllText(outputPath));
            Assert.Empty(Directory.GetFiles(directory, ".*.tmp*"));
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void GetUniquePath_bounds_initial_and_collision_names()
    {
        var directory = Path.Combine(Path.GetTempPath(), $"PSWriteOffice-{Guid.NewGuid():N}");
        Directory.CreateDirectory(directory);

        try
        {
            var fileName = new string('a', 260) + ".pdf";
            var initialPath = PdfCommandUtilities.GetUniquePath(directory, fileName);
            var initialName = Path.GetFileName(initialPath);
            Assert.Equal(255, initialName.Length);
            Assert.EndsWith(".pdf", initialName, StringComparison.Ordinal);

            File.WriteAllText(initialPath, "existing");
            var collisionPath = PdfCommandUtilities.GetUniquePath(directory, fileName);
            var collisionName = Path.GetFileName(collisionPath);
            Assert.Equal(255, collisionName.Length);
            Assert.EndsWith("-2.pdf", collisionName, StringComparison.Ordinal);
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void WriteUnique_keeps_collision_name_within_component_limit()
    {
        var directory = Path.Combine(Path.GetTempPath(), $"PSWriteOffice-{Guid.NewGuid():N}");
        Directory.CreateDirectory(directory);

        try
        {
            var fileName = new string('a', 251) + ".bin";
            File.WriteAllText(Path.Combine(directory, fileName), "existing");

            var outputPath = AtomicFileWriter.WriteUnique(directory, fileName, Encoding.UTF8.GetBytes("payload"));
            var outputName = Path.GetFileName(outputPath);

            Assert.Equal(255, outputName.Length);
            Assert.Equal(new string('a', 249) + "-2.bin", outputName);
            Assert.Equal("payload", File.ReadAllText(outputPath));
            Assert.Equal("existing", File.ReadAllText(Path.Combine(directory, fileName)));
            Assert.Empty(Directory.GetFiles(directory, ".*.tmp*"));
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void WriteUnique_keeps_unicode_collision_name_within_portable_byte_limit()
    {
        var directory = Path.Combine(Path.GetTempPath(), $"PSWriteOffice-{Guid.NewGuid():N}");
        Directory.CreateDirectory(directory);

        try
        {
            var fileName = new string('\u00e9', 125) + ".bin";
            File.WriteAllText(Path.Combine(directory, fileName), "existing");

            var outputPath = AtomicFileWriter.WriteUnique(directory, fileName, Encoding.UTF8.GetBytes("payload"));
            var outputName = Path.GetFileName(outputPath);

            Assert.True(Encoding.UTF8.GetByteCount(outputName) <= 255);
            Assert.EndsWith("-2.bin", outputName, StringComparison.Ordinal);
            Assert.Equal("payload", File.ReadAllText(outputPath));
            Assert.Equal("existing", File.ReadAllText(Path.Combine(directory, fileName)));
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void Write_keeps_temporary_name_within_component_limit()
    {
        var directory = Path.Combine(Path.GetTempPath(), $"PSWriteOffice-{Guid.NewGuid():N}");
        Directory.CreateDirectory(directory);

        try
        {
            var fileName = new string('a', 220) + ".pdf";
            var outputPath = Path.Combine(directory, fileName);

            AtomicFileWriter.Write(outputPath, Encoding.UTF8.GetBytes("payload"), overwrite: false);

            Assert.Equal("payload", File.ReadAllText(outputPath));
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }
    }
}
