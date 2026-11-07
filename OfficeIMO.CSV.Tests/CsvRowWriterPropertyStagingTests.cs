using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using Xunit;

namespace OfficeIMO.CSV.Tests;

public class CsvRowWriterPropertyStagingTests
{
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void WriteObject_PreservesCapturedCultureAndFormatterOrderAcrossRowTypes(bool useStreamFactory)
    {
        using var output = new MemoryStream();
        using var text = new StreamWriter(output, new UTF8Encoding(false), 1024, leaveOpen: true);
        var options = new CsvSaveOptions { NewLine = "\n", Culture = CultureInfo.InvariantCulture };
        var context = new RowContext(options);
        using (var writer = useStreamFactory
            ? CsvRowWriter.CreateStream(output, options, leaveOpen: true)
            : new CsvRowWriter(text, options, leaveOpen: true))
        {
            context.Label = "first";
            writer.WriteObject(new ScalarRow(context));
            options.Culture = CultureInfo.InvariantCulture;
            context.Label = "initial-b";
            writer.WriteObject(new FormattableRow(context));
            options.Culture = CultureInfo.InvariantCulture;
            context.Label = "third";
            writer.WriteObject(new ScalarRow(context));
            context.Label = "fourth";
            writer.WriteObject(new ScalarRow(context));
        }
        text.Flush();
        string actual = Encoding.UTF8.GetString(output.ToArray());
        const string expected = "Amount,Label\n1234.5,first\n1234.5,formatted-b\n1234.5,third\n\"1234,5\",fourth\n";
        Assert.Equal(expected, actual);
        string[] expectedEvents = ["amount-a", "label-a", "amount-b", "format-b", "label-b",
            "amount-a", "label-a", "amount-a", "label-a"];
        Assert.Equal(expectedEvents, context.Events);
    }

    [Theory]
    [InlineData(false, 0)]
    [InlineData(false, 1)]
    [InlineData(false, 2)]
    [InlineData(true, 0)]
    [InlineData(true, 1)]
    [InlineData(true, 2)]
    public void WriteObject_PreservesLargeFieldsAndGetterFailureRecovery(bool useStreamFactory, int largePosition)
    {
        using var output = new MemoryStream();
        using var text = new StreamWriter(output, new UTF8Encoding(false), 1024, leaveOpen: true);
        var options = new CsvSaveOptions { NewLine = "\n", Culture = CultureInfo.InvariantCulture };
        var events = new List<string>();
        string large = new string('"', 7000) + ",\nzażółć🙂";
        string prefix = largePosition == 0 ? large : "prefix";
        string payload = largePosition == 1 ? large : "payload";
        string suffix = largePosition == 2 ? large : "suffix";
        using (var writer = useStreamFactory
            ? CsvRowWriter.CreateStream(output, options, leaveOpen: true)
            : new CsvRowWriter(text, options, leaveOpen: true))
        {
            writer.WriteObject(new FallbackRow(prefix, 1234.50m, payload, 42, 0.2500m, suffix, events));
            Exception? failure = Record.Exception(() => writer.WriteObject(
                new FallbackRow("failed-prefix", 99m, large, 43, 0.5m, "unwritten", events, failAtSuffix: true)));
            Assert.IsType<InvalidOperationException>(failure?.GetBaseException());
            writer.WriteObject(new FallbackRow("next-prefix", -0.25m, "short", 17, null, "next-suffix", events));
        }
        text.Flush();
        string actual = Encoding.UTF8.GetString(output.ToArray());
        string escaped = "\"" + large.Replace("\"", "\"\"") + "\"";
        string expected = "Prefix,Amount,Payload,Count,Optional,Suffix\n"
            + (largePosition == 0 ? escaped : prefix) + ",1234.50,"
            + (largePosition == 1 ? escaped : payload) + ",42,0.2500,"
            + (largePosition == 2 ? escaped : suffix) + "\n"
            + "next-prefix,-0.25,short,17,,next-suffix\n";
        Assert.Equal(expected, actual);
        string[] expectedEvents = ["prefix", "amount", "payload", "count", "optional", "suffix",
            "prefix", "amount", "payload", "count", "optional", "suffix",
            "prefix", "amount", "payload", "count", "optional", "suffix"];
        Assert.Equal(expectedEvents, events);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void WriteObject_PreservesLongTextPrefixAndScalarTailAcrossRows(bool useStreamFactory)
    {
        using var output = new MemoryStream();
        using var text = new StreamWriter(output, new UTF8Encoding(false), 1024, leaveOpen: true);
        var options = new CsvSaveOptions { NewLine = "\n", Culture = CultureInfo.InvariantCulture };
        var wide = new
        {
            A = new string('a', 500), B = new string('b', 250), C = new string('c', 125),
            D = new string('d', 62), E = new string('e', 30), F = new string('f', 14),
            G = new string('g', 6), H = new string('h', 2), Amount = decimal.MaxValue
        };
        var next = new { A = "next", B = "b", C = "c", D = "d", E = "e", F = "f", G = "g", H = "h", Amount = 1m };
        using (var writer = useStreamFactory
            ? CsvRowWriter.CreateStream(output, options, leaveOpen: true)
            : new CsvRowWriter(text, options, leaveOpen: true))
        {
            writer.WriteObject(wide);
            writer.WriteObject(next);
        }
        text.Flush();
        string expected = "A,B,C,D,E,F,G,H,Amount\n"
            + string.Join(",", new[] { wide.A, wide.B, wide.C, wide.D, wide.E, wide.F, wide.G, wide.H })
            + ",79228162514264337593543950335\nnext,b,c,d,e,f,g,h,1\n";
        Assert.Equal(expected, Encoding.UTF8.GetString(output.ToArray()));
    }

    private sealed class RowContext(CsvSaveOptions options)
    {
        internal CsvSaveOptions Options { get; } = options;
        internal CultureInfo French { get; } = CultureInfo.GetCultureInfo("fr-FR");
        internal List<string> Events { get; } = [];
        internal string Label { get; set; } = "";
    }

    private sealed class ScalarRow(RowContext context)
    {
        public decimal Amount
        {
            get
            {
                context.Events.Add("amount-a");
                context.Options.Culture = ReferenceEquals(context.Options.Culture, CultureInfo.InvariantCulture)
                    ? context.French : CultureInfo.InvariantCulture;
                return 1234.5m;
            }
        }

        public string Label
        {
            get
            {
                context.Events.Add("label-a");
                return context.Label;
            }
        }
    }

    private sealed class FormattableRow(RowContext context)
    {
        public FormattableAmount Amount
        {
            get
            {
                context.Events.Add("amount-b");
                return new FormattableAmount(context);
            }
        }

        public string Label
        {
            get
            {
                context.Events.Add("label-b");
                return context.Label;
            }
        }
    }

    private sealed class FormattableAmount(RowContext context) : IFormattable
    {
        public string ToString(string? format, IFormatProvider? provider)
        {
            context.Events.Add("format-b");
            context.Label = "formatted-b";
            return 1234.5m.ToString(format, provider);
        }
    }

    private sealed class FallbackRow(string prefix, decimal amount, string payload, int count,
        decimal? optional, string suffix, List<string> events, bool failAtSuffix = false)
    {
        public string Prefix
        {
            get
            {
                events.Add("prefix");
                return prefix;
            }
        }
        public decimal Amount
        {
            get
            {
                events.Add("amount");
                return amount;
            }
        }
        public string Payload
        {
            get
            {
                events.Add("payload");
                return payload;
            }
        }
        public int Count
        {
            get
            {
                events.Add("count");
                return count;
            }
        }
        public decimal? Optional
        {
            get
            {
                events.Add("optional");
                return optional;
            }
        }
        public string Suffix
        {
            get
            {
                events.Add("suffix");
                if (failAtSuffix) throw new InvalidOperationException("Expected getter failure.");
                return suffix;
            }
        }
    }
}
