#if NET8_0_OR_GREATER
using System.Diagnostics;

namespace OfficeIMO.Archives.Tests;

/// <summary>Checks that independent tools accept the writers' output. Opt in with OFFICEIMO_RUN_ARCHIVE_TOOL_INTEROP=1;
/// it needs unzip, tar (GNU tar or bsdtar), gzip and python3 on PATH.</summary>
public sealed class ArchiveToolInteropTests {
    [ToolInteropFact]
    public void UnzipTarGzipAndPythonAcceptWrittenArchives() {
        using var fixture = new ArchiveFixtures();
        string zip = Path.Combine(fixture.Root, "out.zip");
        string tar = Path.Combine(fixture.Root, "out.tar");
        string tgz = Path.Combine(fixture.Root, "out.tar.gz");
        string gz = Path.Combine(fixture.Root, "readme.txt.gz");
        ArchiveWriteEntry[] entries = ZipArchiveWriterTests.SampleEntries().Append(
            ArchiveWriteEntry.File("docs/" + new string('n', 160) + ".txt", Encoding.UTF8.GetBytes("pax path"))).ToArray();
        ZipArchiveWriter.Create(zip, entries);
        TarArchiveWriter.Create(tar, entries);
        TarArchiveWriter.Create(tgz, entries, new TarArchiveWriterOptions { GZip = true });
        using (var output = File.Create(gz)) GZipArchive.Compress(new MemoryStream(ZipArchiveWriterTests.Text), output);

        Run("unzip", "-t", zip);
        Assert.Contains("docs/zażółć-gęślą-jaźń.txt", Run("python3", "-I", "-c",
            "import sys,zipfile; z=zipfile.ZipFile(sys.argv[1]); assert z.testzip() is None; print('\\n'.join(z.namelist()))", zip));
        foreach (string archive in new[] { tar, tgz }) {
            string listing = Run("tar", "-tvf", archive);
            Assert.Contains(new string('n', 160) + ".txt", listing);
            Assert.Equal(Encoding.UTF8.GetString(ZipArchiveWriterTests.Text), Run("tar", "-xOf", archive, "docs/readme.txt"));
        }
        Run("gzip", "-t", gz);
        Assert.Equal(Encoding.UTF8.GetString(ZipArchiveWriterTests.Text), Run("gzip", "-dc", gz));
    }

    private static string Run(string tool, params string[] arguments) {
        var start = new ProcessStartInfo(tool) { RedirectStandardOutput = true, RedirectStandardError = true, UseShellExecute = false,
            StandardOutputEncoding = Encoding.UTF8 };
        foreach (string argument in arguments) start.ArgumentList.Add(argument);
        using Process process = Process.Start(start)!;
        string output = process.StandardOutput.ReadToEnd();
        string error = process.StandardError.ReadToEnd();
        process.WaitForExit();
        Assert.True(process.ExitCode == 0, tool + " failed (" + process.ExitCode + "): " + error + output);
        return output;
    }

    private sealed class ToolInteropFactAttribute : FactAttribute {
        public ToolInteropFactAttribute() {
            if (Environment.GetEnvironmentVariable("OFFICEIMO_RUN_ARCHIVE_TOOL_INTEROP") != "1")
                Skip = "Set OFFICEIMO_RUN_ARCHIVE_TOOL_INTEROP=1 to check the output with unzip, tar, gzip and python3.";
        }
    }
}
#endif
