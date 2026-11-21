using System.Globalization;
using System.Text;

// Opt-in check of OfficeRoundTripNumberText, run by Build/Test-RoundTripNumberText.ps1.
//
// On .NET Framework it writes the text of a fixed, seeded set of doubles and floats to framework-double.txt and
// framework-float.txt in the output directory, one "culture bits text" line per value, and fails when that
// runtime's own parser reads any text back as another value. On .NET 10 it reads those files and fails when a
// correct parser reads any text back as another value; it also counts the texts that equal .NET 10's "R" and the
// longer ones written because .NET Framework's parser misreads the shortest.
//
// Usage: OfficeIMO.RoundTripNumberTextProbe <count> <output-directory>
internal static class Program {
    private static readonly string[] Cultures = { "de-DE", "fr-FR" };

    private static int Main(string[] args) {
        if (args.Length != 2 || !int.TryParse(args[0], out int count) || count < 1) {
            Console.Error.WriteLine("Usage: OfficeIMO.RoundTripNumberTextProbe <count> <output-directory>");
            return 2;
        }
        Directory.CreateDirectory(args[1]);
        string doubles = Path.Combine(args[1], "framework-double.txt");
        string floats = Path.Combine(args[1], "framework-float.txt");
        if (Environment.Version.Major < 5) {
            int misread = WriteDoubles(doubles, count) + WriteFloats(floats, count);
            Console.WriteLine(misread == 0 ? ".NET Framework's parser reads every text back exactly."
                : $".NET Framework's parser reads {misread} texts back as another value.");
            return misread == 0 ? 0 : 1;
        }
        bool doublesPass = Check(doubles, single: false);
        bool floatsPass = Check(floats, single: true);
        return doublesPass && floatsPass ? 0 : 1;
    }

    // Raw bit patterns, uniform fractions, scaled values, short decimals with a binary tail, ratios of small
    // integers and doubles near the subnormal and overflow ends; then a few values in two other cultures.
    private static IEnumerable<(CultureInfo Culture, double Value)> Values(int count) {
        var random = new Random(20261121);
        var bytes = new byte[8];
        for (int index = 0; index < count; index++) {
            double value;
            switch (index % 6) {
                case 0:
                    random.NextBytes(bytes);
                    value = BitConverter.ToDouble(bytes, 0);
                    break;
                case 1:
                    value = random.NextDouble();
                    break;
                case 2:
                    value = (random.NextDouble() - 0.5) * Math.Pow(10, random.Next(-8, 9));
                    break;
                case 3:
                    value = Math.Round(random.NextDouble() * 1000, random.Next(1, 8)) + random.Next(0, 3) * Math.Pow(2, -random.Next(10, 40));
                    break;
                case 4:
                    value = random.Next(1, 1000000) / Math.Pow(10, random.Next(0, 12));
                    break;
                default:
                    value = BitConverter.Int64BitsToDouble(((long)random.Next(0, 4096) << 52) | ((long)random.Next() << 20) | (long)random.Next(0, 3));
                    break;
            }
            if (!double.IsNaN(value) && !double.IsInfinity(value)) yield return (CultureInfo.InvariantCulture, value);
        }
        foreach (string name in Cultures) {
            foreach (double value in new[] { 1234.5, -1.5e-20, 0.000123, 1e300, -0.0 }) yield return (CultureInfo.GetCultureInfo(name), value);
        }
    }

    private static int WriteDoubles(string path, int count) {
        int misread = 0;
        using (var writer = new StreamWriter(path, false, new UTF8Encoding(false))) {
            foreach ((CultureInfo culture, double value) in Values(count)) {
                string text = OfficeIMO.OfficeRoundTripNumberText.Format(value, culture);
                long bits = BitConverter.DoubleToInt64Bits(value);
                // .NET Framework reads -0 as 0, whatever the text, so a zero only has to read back as zero there.
                if (!double.TryParse(text, NumberStyles.Float, culture, out double parsed) ||
                    (BitConverter.DoubleToInt64Bits(parsed) != bits && !(value == 0D && parsed == 0D))) {
                    if (misread++ < 10) Console.WriteLine($"double {bits}: '{text}' reads back as {parsed.ToString("R", CultureInfo.InvariantCulture)}");
                }
                writer.WriteLine(Name(culture) + " " + bits.ToString(CultureInfo.InvariantCulture) + " " + text);
            }
        }
        return misread;
    }

    private static int WriteFloats(string path, int count) {
        int misread = 0;
        using (var writer = new StreamWriter(path, false, new UTF8Encoding(false))) {
            foreach ((CultureInfo culture, double source) in Values(count)) {
                float value = (float)source;
                if (float.IsInfinity(value)) continue;
                string text = OfficeIMO.OfficeRoundTripNumberText.Format(value, culture);
                int bits = BitConverter.ToInt32(BitConverter.GetBytes(value), 0);
                if (!float.TryParse(text, NumberStyles.Float, culture, out float parsed) ||
                    (BitConverter.ToInt32(BitConverter.GetBytes(parsed), 0) != bits && !(value == 0F && parsed == 0F))) {
                    if (misread++ < 10) Console.WriteLine($"float {bits}: '{text}' reads back as {parsed.ToString("R", CultureInfo.InvariantCulture)}");
                }
                writer.WriteLine(Name(culture) + " " + bits.ToString(CultureInfo.InvariantCulture) + " " + text);
            }
        }
        return misread;
    }

    private static bool Check(string path, bool single) {
        int same = 0, longer = 0, wrong = 0;
        foreach (string line in File.ReadLines(path)) {
            string[] parts = line.Split(new[] { ' ' }, 3);
            CultureInfo culture = parts[0] == "inv" ? CultureInfo.InvariantCulture : CultureInfo.GetCultureInfo(parts[0]);
            string text = parts[2];
            string runtime;
            bool readsBack;
            if (single) {
                int bits = int.Parse(parts[1], CultureInfo.InvariantCulture);
                float value = BitConverter.ToSingle(BitConverter.GetBytes(bits), 0);
                runtime = value.ToString("R", culture);
                readsBack = float.TryParse(text, NumberStyles.Float, culture, out float parsed) && BitConverter.ToInt32(BitConverter.GetBytes(parsed), 0) == bits;
            } else {
                long bits = long.Parse(parts[1], CultureInfo.InvariantCulture);
                double value = BitConverter.Int64BitsToDouble(bits);
                runtime = value.ToString("R", culture);
                readsBack = double.TryParse(text, NumberStyles.Float, culture, out double parsed) && BitConverter.DoubleToInt64Bits(parsed) == bits;
            }
            if (!readsBack) {
                if (wrong++ < 10) Console.WriteLine($"{(single ? "float" : "double")} {parts[1]}: '{text}' does not read back; .NET 10 writes '{runtime}'");
            } else if (text == runtime) {
                same++;
            } else {
                longer++;
            }
        }
        string kind = single ? "float" : "double";
        Console.WriteLine($"{kind}: {same} texts equal .NET 10's, {longer} are longer for .NET Framework's parser, {wrong} do not read back.");
        return wrong == 0;
    }

    private static string Name(CultureInfo culture) => culture.Name.Length == 0 ? "inv" : culture.Name;
}
