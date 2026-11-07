using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Text;
using OfficeIMO.CSV;

internal static class CsvPropertyWriteSmoke
{
    [DynamicDependency(DynamicallyAccessedMemberTypes.PublicProperties, typeof(PreservedRow))]
    [UnconditionalSuppressMessage("Trimming", "IL2026", Justification = "This fixture explicitly preserves the row's public properties for the annotated reflection compatibility API.")]
    internal static void Run()
    {
        using var output = new MemoryStream();
        using (var writer = CsvRowWriter.CreateStream(output,
            new CsvSaveOptions { NewLine = "\n", Culture = CultureInfo.InvariantCulture }, leaveOpen: true))
        {
            writer.WriteObject(new PreservedRow { Amount = 1234.50m, Name = "zażółć,🙂", Optional = 0.2500m });
            writer.WriteObject(new PreservedRow { Amount = -0.25m, Name = "short", Optional = null });
        }
        string actual = Encoding.UTF8.GetString(output.ToArray());
        const string expected = "Amount,Name,Optional\n1234.50,\"zażółć,🙂\",0.2500\n-0.25,short,\n";
        if (actual != expected) throw new InvalidOperationException("The metadata-preserved object writer changed values under NativeAOT.");
        Console.WriteLine("PASS | Metadata-preserved WriteObject compatibility path | dynamic code " +
            System.Runtime.CompilerServices.RuntimeFeature.IsDynamicCodeSupported);
    }

    private sealed class PreservedRow
    {
        public decimal Amount { get; set; }
        public string Name { get; set; } = "";
        public decimal? Optional { get; set; }
    }
}
