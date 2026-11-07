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
    [InlineData(false)]
    [InlineData(true)]
    public void WriteObject_PreservesLargeQuotedFieldsAndNullableValuesWithoutRepeatingGetters(bool useStreamFactory)
    {
        using var output = new MemoryStream();
        using var text = new StreamWriter(output, new UTF8Encoding(false), 1024, leaveOpen: true);
        var options = new CsvSaveOptions { NewLine = "\n", Culture = CultureInfo.InvariantCulture };
        var events = new List<string>();
        string large = new string('"', 7000) + ",\nzażółć🙂";
        using (var writer = useStreamFactory
            ? CsvRowWriter.CreateStream(output, options, leaveOpen: true)
            : new CsvRowWriter(text, options, leaveOpen: true))
        {
            writer.WriteObject(new FallbackRow(1234.50m, large, 42, 0.2500m, events));
            writer.WriteObject(new FallbackRow(-0.25m, "short", 17, null, events));
        }
        text.Flush();
        string actual = Encoding.UTF8.GetString(output.ToArray());
        string escaped = "\"" + large.Replace("\"", "\"\"") + "\"";
        string expected = "Amount,Payload,Count,Optional\n1234.50," + escaped + ",42,0.2500\n-0.25,short,17,\n";
        Assert.Equal(expected, actual);
        string[] expectedEvents = ["amount", "payload", "count", "optional", "amount", "payload", "count", "optional"];
        Assert.Equal(expectedEvents, events);
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

    private sealed class FallbackRow(decimal amount, string payload, int count, decimal? optional, List<string> events)
    {
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
    }
}
