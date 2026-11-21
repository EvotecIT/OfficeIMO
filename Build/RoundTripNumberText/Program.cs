using System.Diagnostics;
using System.Globalization;
using System.Numerics;
using System.Text;

// Opt-in check of OfficeRoundTripNumberText, run by Build/Test-RoundTripNumberText.ps1 in three steps.
//
//   corpus <count> <dir>  .NET 10: writes parse-corpus.txt, strings with .NET 10's double and float parse results.
//   write <count> <dir>   .NET Framework: writes the text of seeded doubles and floats, fails when OfficeIMO's
//                         parser does not read every text back exactly, and fails when OfficeIMO's parser disagrees
//                         with .NET 10 on any corpus string. It also times both parsers on the corpus.
//   check <count> <dir>   .NET 10: fails unless every text written on .NET Framework is byte-identical to .NET 10's
//                         "R" and reads back exactly.
internal static class Program {
    private static readonly string[] Cultures = { "de-DE", "fr-FR" };

    private static int Main(string[] args) {
        if (args.Length != 3 || !int.TryParse(args[1], out int count) || count < 1) {
            Console.Error.WriteLine("Usage: OfficeIMO.RoundTripNumberTextProbe corpus|write|check <count> <output-directory>");
            return 2;
        }
        Directory.CreateDirectory(args[2]);
        string doubles = Path.Combine(args[2], "framework-double.txt");
        string floats = Path.Combine(args[2], "framework-float.txt");
        string corpus = Path.Combine(args[2], "parse-corpus.txt");
        switch (args[0]) {
            case "corpus":
                WriteCorpus(corpus, count);
                return 0;
            case "write":
                int misread = WriteDoubles(doubles, count) + WriteFloats(floats, count);
                Console.WriteLine(misread == 0 ? "OfficeIMO's parser reads every written text back exactly."
                    : $"OfficeIMO's parser reads {misread} written texts back as another value.");
                int disagreements = CheckCorpus(corpus);
                return misread == 0 && disagreements == 0 ? 0 : 1;
            case "check":
                bool doublesPass = CheckTexts(doubles, single: false);
                bool floatsPass = CheckTexts(floats, single: true);
                return doublesPass && floatsPass ? 0 : 1;
            default:
                Console.Error.WriteLine("Unknown step " + args[0]);
                return 2;
        }
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
                if (!OfficeIMO.OfficeRoundTripNumberText.TryParse(text, NumberStyles.Float, culture, out double parsed) ||
                    BitConverter.DoubleToInt64Bits(parsed) != bits) {
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
                int bits = SingleBits(value);
                if (!OfficeIMO.OfficeRoundTripNumberText.TryParse(text, NumberStyles.Float, culture, out float parsed) || SingleBits(parsed) != bits) {
                    if (misread++ < 10) Console.WriteLine($"float {bits}: '{text}' reads back as {parsed.ToString("R", CultureInfo.InvariantCulture)}");
                }
                writer.WriteLine(Name(culture) + " " + bits.ToString(CultureInfo.InvariantCulture) + " " + text);
            }
        }
        return misread;
    }

    private static bool CheckTexts(string path, bool single) {
        int same = 0, different = 0, wrong = 0;
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
                readsBack = float.TryParse(text, NumberStyles.Float, culture, out float parsed) && SingleBits(parsed) == bits;
            } else {
                long bits = long.Parse(parts[1], CultureInfo.InvariantCulture);
                double value = BitConverter.Int64BitsToDouble(bits);
                runtime = value.ToString("R", culture);
                readsBack = double.TryParse(text, NumberStyles.Float, culture, out double parsed) && BitConverter.DoubleToInt64Bits(parsed) == bits;
            }
            if (!readsBack) {
                if (wrong++ < 10) Console.WriteLine($"{(single ? "float" : "double")} {parts[1]}: '{text}' does not read back");
            } else if (text == runtime) {
                same++;
            } else if (different++ < 10) {
                Console.WriteLine($"{(single ? "float" : "double")} {parts[1]}: .NET Framework wrote '{text}', .NET 10 writes '{runtime}'");
            }
        }
        string kind = single ? "float" : "double";
        Console.WriteLine($"{kind}: {same} texts equal .NET 10's, {different} differ, {wrong} do not read back.");
        return wrong == 0 && different == 0;
    }

    // Parse corpus: long and short digit strings over the whole exponent range, decimals at and next to the exact
    // midpoint between two doubles or two floats, formatted values and spellings, each with a number style and a
    // culture, followed by .NET 10's double and float results.
    private static readonly NumberStyles[] Styles = {
        NumberStyles.Float, NumberStyles.Float | NumberStyles.AllowThousands, NumberStyles.Any, NumberStyles.Number, NumberStyles.Integer
    };

    private static readonly string[] Spellings = {
        "Infinity", "infinity", "INFINITY", "+Infinity", "-Infinity", "-infinity", " Infinity ", "Inf", "NaN", "nan", "-NaN", "+NaN",
        " NaN", "NaN ", "1e400", "-1e400", "1e-400", "-1e-400", "-0", "-0.0", "+0", "0e5", "1.7976931348623157E+308",
        "1.7976931348623158E+308", "1.7976931348623159E+308", "2.4703282292062328E-324", "2.4703282292062327E-324", "5e-324",
        "3.4028235E+38", "3.4028236E+38", "3.4028235677973366E+38", "3.4028235677973367E+38", "1.401298464324817E-45",
        "7.006492321624085E-46", "7.006492321624086E-46", "1.", ".5", ".", "-.5", "+.5", "1e", "1e+", "1E5", "1e05", "1e+0005",
        "0x10", "1_0", "1,000", "1,000.5", "1,0,0", ",5", "1 000", " 1", "1 ", "\t1\n", "1\0", "--1", "+-1", "(1)", "$1", "1%",
        "1.5e3.2", "00001.25", "1e999999999999", "1e-999999999999", "0.6971128700134806015", "0.25642034204212829653",
        "9007199254740993", "123456789012345678901234567890", "0.000000000000000000000000000001",
    };

    private static void WriteCorpus(string path, int count) {
        var random = new Random(20261122);
        using var writer = new StreamWriter(path, false, new UTF8Encoding(false));
        void Add(string text, NumberStyles style, CultureInfo culture) {
            bool doubleOk = double.TryParse(text, style, culture, out double d);
            bool floatOk = float.TryParse(text, style, culture, out float f);
            writer.WriteLine(((int)style).ToString(CultureInfo.InvariantCulture) + "\t" + Name(culture) + "\t" + Escape(text) + "\t" +
                (doubleOk ? BitConverter.DoubleToInt64Bits(d).ToString(CultureInfo.InvariantCulture) : "-") + "\t" +
                (floatOk ? SingleBits(f).ToString(CultureInfo.InvariantCulture) : "-"));
        }
        foreach (string spelling in Spellings) foreach (NumberStyles style in Styles) Add(spelling, style, CultureInfo.InvariantCulture);
        foreach (string name in Cultures) {
            CultureInfo culture = CultureInfo.GetCultureInfo(name);
            foreach (string text in new[] { "1,5", "1.234,5", "-1,5E-20", "1e3", "Infinity", "NaN", "0,30000000000000004", "123456789,123456789123" }) {
                Add(text, NumberStyles.Float | NumberStyles.AllowThousands, culture);
            }
        }
        var bytes = new byte[8];
        for (int index = 0; index < count; index++) {
            string text;
            switch (index % 8) {
                case 0:
                case 1:
                    text = RandomDigits(random, random.Next(1, 26), random.Next(-30, 31));
                    break;
                case 2:
                    text = RandomDigits(random, random.Next(15, 26), random.Next(-345, 310));
                    break;
                case 3:
                    random.NextBytes(bytes);
                    double raw = BitConverter.ToDouble(bytes, 0);
                    if (double.IsNaN(raw) || double.IsInfinity(raw)) raw = random.NextDouble();
                    text = Midpoint(Math.Abs(raw), random, single: false);
                    break;
                case 4:
                    text = Midpoint(Math.Abs((double)(float)(random.NextDouble() * Math.Pow(10, random.Next(-40, 38)))), random, single: true);
                    break;
                case 5:
                    double value = (random.NextDouble() - 0.5) * Math.Pow(10, random.Next(-20, 21));
                    string[] formats = { "R", "E16", "G17", "F2", "G15", "E20" };
                    text = value.ToString(formats[random.Next(formats.Length)], CultureInfo.InvariantCulture);
                    break;
                case 6:
                    text = RandomDigits(random, random.Next(1, 20), random.Next(-330, -300));
                    break;
                default:
                    text = RandomDigits(random, random.Next(1, 10), random.Next(-50, 50));
                    break;
            }
            Add(text, index % 16 == 7 ? NumberStyles.Float | NumberStyles.AllowThousands : NumberStyles.Float, CultureInfo.InvariantCulture);
        }
    }

    private static string RandomDigits(Random random, int digits, int exponent) {
        var builder = new StringBuilder();
        if (random.Next(4) == 0) builder.Append('-');
        for (int i = 0; i < digits; i++) builder.Append((char)('0' + (i == 0 ? random.Next(1, 10) : random.Next(10))));
        int point = random.Next(0, digits + 1);
        if (point < digits) builder.Insert(builder.Length - (digits - point), '.');
        if (exponent != 0 || random.Next(2) == 0) builder.Append('e').Append(exponent.ToString(CultureInfo.InvariantCulture));
        return builder.ToString();
    }

    // The exact decimal halfway between a positive value and the next one up, sometimes cut to fewer digits or moved
    // by one in its last kept digit: the inputs that need exact rounding.
    private static string Midpoint(double value, Random random, bool single) {
        BigInteger mantissa;
        int exponent;
        if (single) {
            int bits = SingleBits((float)value);
            int biased = (bits >> 23) & 0xFF;
            mantissa = biased == 0 ? bits & 0x7FFFFF : (bits & 0x7FFFFF) | 0x800000;
            exponent = biased == 0 ? -149 : biased - 150;
        } else {
            long bits = BitConverter.DoubleToInt64Bits(value);
            int biased = (int)((bits >> 52) & 0x7FF);
            mantissa = biased == 0 ? bits & 0xFFFFFFFFFFFFFL : (bits & 0xFFFFFFFFFFFFFL) | (1L << 52);
            exponent = biased == 0 ? -1074 : biased - 1075;
        }
        // (2 * mantissa + 1) * 2^(exponent - 1)
        BigInteger numerator = 2 * mantissa + 1;
        int e = exponent - 1;
        string digits;
        int pointShift;
        if (e >= 0) {
            digits = (numerator << e).ToString(CultureInfo.InvariantCulture);
            pointShift = 0;
        } else {
            digits = (numerator * BigInteger.Pow(5, -e)).ToString(CultureInfo.InvariantCulture);
            pointShift = -e;
        }
        int keep = random.Next(3) == 0 ? digits.Length : Math.Min(digits.Length, random.Next(16, 30));
        string kept = digits.Substring(0, keep);
        int scale = pointShift - (digits.Length - keep);
        int adjust = random.Next(3) - 1;
        if (adjust != 0 && keep < digits.Length) kept = (BigInteger.Parse(kept, CultureInfo.InvariantCulture) + adjust).ToString(CultureInfo.InvariantCulture);
        return kept + "e" + (-scale).ToString(CultureInfo.InvariantCulture);
    }

    private static int CheckCorpus(string path) {
        var lines = File.ReadAllLines(path);
        int wrong = 0, checkedCount = 0;
        var inputs = new List<(string Text, NumberStyles Style, CultureInfo Culture)>(lines.Length);
        foreach (string line in lines) {
            string[] parts = line.Split('\t');
            var style = (NumberStyles)int.Parse(parts[0], CultureInfo.InvariantCulture);
            CultureInfo culture = parts[1] == "inv" ? CultureInfo.InvariantCulture : CultureInfo.GetCultureInfo(parts[1]);
            string text = Unescape(parts[2]);
            inputs.Add((text, style, culture));
            checkedCount++;
            bool doubleOk = OfficeIMO.OfficeRoundTripNumberText.TryParse(text, style, culture, out double d);
            string doubleResult = doubleOk ? BitConverter.DoubleToInt64Bits(d).ToString(CultureInfo.InvariantCulture) : "-";
            bool floatOk = OfficeIMO.OfficeRoundTripNumberText.TryParse(text, style, culture, out float f);
            string floatResult = floatOk ? SingleBits(f).ToString(CultureInfo.InvariantCulture) : "-";
            if (doubleResult != parts[3] || floatResult != parts[4]) {
                if (wrong++ < 20) Console.WriteLine($"'{parts[2]}' ({style}, {parts[1]}): OfficeIMO {doubleResult}/{floatResult}, .NET 10 {parts[3]}/{parts[4]}");
            }
        }
        Console.WriteLine($"parse corpus: {checkedCount - wrong} of {checkedCount} strings match .NET 10, {wrong} differ.");

        // Timing on the same strings: .NET Framework's parser against OfficeIMO's.
        var stopwatch = Stopwatch.StartNew();
        double sink = 0;
        for (int pass = 0; pass < 3; pass++) foreach (var input in inputs) if (double.TryParse(input.Text, input.Style, input.Culture, out double v)) sink += v;
        double framework = stopwatch.Elapsed.TotalMilliseconds / 3;
        stopwatch.Restart();
        for (int pass = 0; pass < 3; pass++) foreach (var input in inputs) if (OfficeIMO.OfficeRoundTripNumberText.TryParse(input.Text, input.Style, input.Culture, out double v)) sink += v;
        double ours = stopwatch.Elapsed.TotalMilliseconds / 3;
        Console.WriteLine($"parse time per pass over {inputs.Count} strings: .NET Framework {framework:F0} ms, OfficeIMO {ours:F0} ms.");
        GC.KeepAlive(sink);
        return wrong;
    }

    private static string Escape(string text) {
        var builder = new StringBuilder();
        foreach (char c in text) {
            if (c < 32 || c > 126 || c == '\\') builder.Append("\\u").Append(((int)c).ToString("X4", CultureInfo.InvariantCulture));
            else builder.Append(c);
        }
        return builder.ToString();
    }

    private static string Unescape(string text) {
        var builder = new StringBuilder();
        for (int i = 0; i < text.Length; i++) {
            if (text[i] == '\\' && i + 5 < text.Length + 0 && text[i + 1] == 'u') {
                builder.Append((char)int.Parse(text.Substring(i + 2, 4), NumberStyles.HexNumber, CultureInfo.InvariantCulture));
                i += 5;
            } else {
                builder.Append(text[i]);
            }
        }
        return builder.ToString();
    }

    private static int SingleBits(float value) => BitConverter.ToInt32(BitConverter.GetBytes(value), 0);

    private static string Name(CultureInfo culture) => culture.Name.Length == 0 ? "inv" : culture.Name;
}
