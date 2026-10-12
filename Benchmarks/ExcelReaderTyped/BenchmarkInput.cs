using OfficeIMO.Benchmarks;
using System.Diagnostics;
using System.IO.Compression;
using System.Reflection;
using System.Security.Cryptography;

namespace OfficeIMO.Excel.ReaderComparison.Benchmarks {
    internal static class BenchmarkInput {
        private static bool _described;

        // Called by setup as well as the host so BDN child-process identity is visible.
        internal static void WriteDescription() {
            if (_described) return;
            BenchmarkProcessPriority.ApplyConfiguredPriority();
            using Process process = Process.GetCurrentProcess();
            string affinity = OperatingSystem.IsWindows()
                ? BenchmarkProcessorAffinity.Format(process.ProcessorAffinity) : "unavailable";
            Console.WriteLine($"Benchmark process: priority={process.PriorityClass}, affinity={affinity}.");
            AssemblyMetadataAttribute[] metadata = Assembly.GetExecutingAssembly().GetCustomAttributes<AssemblyMetadataAttribute>().ToArray();
            string? packageVersion = metadata.Single(attribute => attribute.Key == "OfficeIMOBenchmarkPackageVersion").Value;
            string? directory = metadata.Single(attribute => attribute.Key == "OfficeIMOBenchmarkAssemblyDirectory").Value;
            string? newApis = metadata.Single(attribute => attribute.Key == "OfficeIMOBenchmarkNewApis").Value;
            string? csv = metadata.Single(attribute => attribute.Key == "OfficeIMOBenchmarkCsv").Value;
            string? arrow = metadata.Single(attribute => attribute.Key == "OfficeIMOBenchmarkArrow").Value;
        string? generatedMapping = metadata.Single(attribute => attribute.Key == "OfficeIMOBenchmarkGeneratedMapping").Value;
            string input = !string.IsNullOrEmpty(packageVersion) ? $"OfficeIMO NuGet {packageVersion}"
                : !string.IsNullOrEmpty(directory) ? $"saved OfficeIMO assemblies from {directory}" : "OfficeIMO source";
            Console.WriteLine($"Typed XLSX input: {input}; ExcelReader.NET 6.0.0; Sylvan.Data.Excel 0.5.8; Sylvan.Data 0.2.17.");
            Console.WriteLine($"OfficeIMO new API cases enabled: {newApis}.");
            Console.WriteLine($"CSV comparison cases enabled: {csv}.");
            Console.WriteLine($"Arrow comparison cases enabled: {arrow}.");
            Console.WriteLine($"Generated mapping cases enabled: {generatedMapping}.");
            List<Assembly> assemblies = new List<Assembly> { typeof(ExcelDocument).Assembly, Assembly.Load("OfficeIMO.Core"),
                typeof(ExcelReader.Core.Reader.Excel).Assembly, typeof(Sylvan.Data.Excel.ExcelDataReader).Assembly };
#if OFFICEIMO_BENCHMARK_CSV
            assemblies.Add(typeof(OfficeIMO.CSV.CsvDocument).Assembly);
            assemblies.Add(typeof(global::Sylvan.Data.Csv.CsvDataReader).Assembly);
#endif
#if OFFICEIMO_BENCHMARK_ARROW
            assemblies.Add(Assembly.Load("OfficeIMO.Data.Arrow"));
            assemblies.Add(Assembly.Load("ExcelReader.Arrow"));
            assemblies.Add(Assembly.Load("Apache.Arrow"));
#endif
            foreach (Assembly assembly in assemblies) {
                string name = assembly.GetName().Name!;
                string version = assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion ?? "unknown";
                string hash = Hash(assembly.Location);
                if (!string.IsNullOrEmpty(directory) && name is "OfficeIMO.Excel" or "OfficeIMO.Core" or "OfficeIMO.CSV" or "OfficeIMO.Data.Arrow") {
                    if (hash != Hash(Path.Combine(directory, name + ".dll")))
                        throw new InvalidDataException($"The loaded {name} assembly differs from the saved input.");
                }
                Console.WriteLine($"{name}: {version}; SHA256={hash}.");
            }
            _described = true;
        }

        private static string Hash(string path) {
            using FileStream stream = File.OpenRead(path);
            return Convert.ToHexString(SHA256.HashData(stream));
        }

        // Setup-only: identifies the bytes actually consumed, without rewriting or retaining them.
        internal static void WriteFixtureIdentity(string id, ReadOnlySpan<byte> bytes) =>
            Console.WriteLine($"Fixture bytes: id={id}; bytes={bytes.Length}; SHA256={Convert.ToHexString(SHA256.HashData(bytes))}.");

        // Entry payloads distinguish ZIP timestamp/compression changes from changed document content.
        internal static void WriteWorkbookFixtureIdentity(string id, byte[] bytes, bool zipPackage = true) {
            WriteFixtureIdentity(id, bytes);
            if (!zipPackage) return;
            using MemoryStream stream = new MemoryStream(bytes, writable: false);
            using ZipArchive package = new ZipArchive(stream, ZipArchiveMode.Read);
            foreach (ZipArchiveEntry entry in package.Entries.OrderBy(entry => entry.FullName, StringComparer.Ordinal)) {
                using Stream input = entry.Open();
                Console.WriteLine($"Fixture ZIP part: id={id}; name={entry.FullName}; bytes={entry.Length}; "
                    + $"compressedBytes={entry.CompressedLength}; modified={entry.LastWriteTime:O}; "
                    + $"SHA256={Convert.ToHexString(SHA256.HashData(input))}.");
            }
        }
    }
}
