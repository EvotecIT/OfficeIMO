using System.Globalization;
using System.Text;

// Writes the round-trip text of a fixed, seeded set of doubles and floats, one value per line, to
// <output-directory>/<runtime>-double.txt and <runtime>-float.txt. On .NET 10 the text is the runtime's "R"; on
// .NET Framework it comes from OfficeRoundTripNumberText's own digits, so the files must be identical.
// Usage: OfficeIMO.RoundTripNumberTextProbe <count> <output-directory>
internal static class Program {
    private static int Main(string[] args) {
        if (args.Length != 2 || !int.TryParse(args[0], out int count) || count < 1) {
            Console.Error.WriteLine("Usage: OfficeIMO.RoundTripNumberTextProbe <count> <output-directory>");
            return 2;
        }
        string runtime = Environment.Version.Major >= 5 ? "core" : "framework";
        Directory.CreateDirectory(args[1]);
        Write(Path.Combine(args[1], runtime + "-double.txt"), count, single: false);
        Write(Path.Combine(args[1], runtime + "-float.txt"), count, single: true);
        return 0;
    }

    // Raw bit patterns, uniform fractions, scaled values, short decimals with a binary tail, ratios of small
    // integers and doubles near the subnormal and overflow ends.
    private static IEnumerable<double> Values(int count) {
        var random = new Random(20261121);
        var bytes = new byte[8];
        for (int index = 0; index < count; index++) {
            switch (index % 6) {
                case 0:
                    random.NextBytes(bytes);
                    yield return BitConverter.ToDouble(bytes, 0);
                    break;
                case 1:
                    yield return random.NextDouble();
                    break;
                case 2:
                    yield return (random.NextDouble() - 0.5) * Math.Pow(10, random.Next(-8, 9));
                    break;
                case 3:
                    yield return Math.Round(random.NextDouble() * 1000, random.Next(1, 8)) + random.Next(0, 3) * Math.Pow(2, -random.Next(10, 40));
                    break;
                case 4:
                    yield return random.Next(1, 1000000) / Math.Pow(10, random.Next(0, 12));
                    break;
                default:
                    yield return BitConverter.Int64BitsToDouble(((long)random.Next(0, 4096) << 52) | ((long)random.Next() << 20) | (long)random.Next(0, 3));
                    break;
            }
        }
    }

    private static void Write(string path, int count, bool single) {
        using (var writer = new StreamWriter(path, false, new UTF8Encoding(false))) {
            int index = 0;
            foreach (double value in Values(count)) {
                writer.Write(index++.ToString(CultureInfo.InvariantCulture));
                writer.Write(' ');
                writer.WriteLine(single ? OfficeIMO.OfficeRoundTripNumberText.Format((float)value, CultureInfo.InvariantCulture)
                    : OfficeIMO.OfficeRoundTripNumberText.Format(value, CultureInfo.InvariantCulture));
            }
            foreach (string name in new[] { "de-DE", "fr-FR" }) {
                CultureInfo culture = CultureInfo.GetCultureInfo(name);
                foreach (double value in new[] { 1234.5, -1.5e-20, 0.000123, 1e300, -0.0 }) {
                    writer.WriteLine(name + " " + (single ? OfficeIMO.OfficeRoundTripNumberText.Format((float)value, culture)
                        : OfficeIMO.OfficeRoundTripNumberText.Format(value, culture)));
                }
            }
        }
    }
}
