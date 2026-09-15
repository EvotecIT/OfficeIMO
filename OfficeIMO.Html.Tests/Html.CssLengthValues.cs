using System.Text.Json;
using System.Threading;
using OfficeIMO.Html;
using OfficeIMO.Html.Css;
using Xunit;

namespace OfficeIMO.Tests;

[Collection(HtmlCssPropertyGrammarCollection.Name)]
public sealed class HtmlCssLengthValuesTests {
    [Fact]
    public void IndependentLengthMathCorpusMatchesTypedAndContextualContracts() {
        LengthMathCase[] corpus = Read<LengthMathCase>("css-length-math-corpus.json");
        Assert.Equal(28, corpus.Length);
        var context = new HtmlCssLengthResolutionContext {
            PercentageReference = 200D,
            FontSize = 20D,
            RootFontSize = 16D,
            ViewportWidth = 600D,
            ViewportHeight = 400D,
            SmallViewportWidth = 500D,
            SmallViewportHeight = 300D,
            DynamicViewportWidth = 550D,
            DynamicViewportHeight = 200D,
            ContainerWidth = 200D,
            ContainerHeight = 100D,
            ContainerInlineSize = 200D,
            ContainerBlockSize = 100D
        };

        foreach (LengthMathCase item in corpus) {
            HtmlCssMathParseResult parsed = HtmlCssMathParser.ParseLengthPercentage(item.Value);
            Assert.Equal(ParseEnum<HtmlCssMathParseStatus>(item.ParseStatus), parsed.Status);
            if (!parsed.IsParsed) continue;
            Assert.Equal(ParseEnum<HtmlCssNumericType>(item.Type!), parsed.Expression!.Type);
            Assert.Equal(item.Canonical, parsed.Expression.CanonicalText);
            HtmlCssLengthResolutionResult resolved = HtmlCssMathResolver.ResolveLength(parsed.Expression, context);
            HtmlCssLengthResolutionStatus expectedStatus = item.ResolutionStatus == null
                ? HtmlCssLengthResolutionStatus.Resolved
                : ParseEnum<HtmlCssLengthResolutionStatus>(item.ResolutionStatus);
            Assert.Equal(expectedStatus, resolved.Status);
            if (item.Resolved.HasValue) Assert.Equal(item.Resolved.Value, resolved.Value!.Value, 10);
        }
    }

    [Fact]
    public void ResolutionReportsTheSpecificMissingContextAndUsesContainerFallback() {
        HtmlCssMathExpression percentage = Parse("25%");
        HtmlCssMathExpression em = Parse("2em");
        HtmlCssMathExpression rem = Parse("2rem");
        HtmlCssMathExpression viewport = Parse("2vmin");
        HtmlCssMathExpression container = Parse("10cqmin");

        Assert.Equal(HtmlCssLengthResolutionStatus.MissingPercentageReference,
            HtmlCssMathResolver.ResolveLength(percentage, new HtmlCssLengthResolutionContext()).Status);
        Assert.Equal(HtmlCssLengthResolutionStatus.MissingFontSize,
            HtmlCssMathResolver.ResolveLength(em, new HtmlCssLengthResolutionContext()).Status);
        Assert.Equal(HtmlCssLengthResolutionStatus.MissingRootFontSize,
            HtmlCssMathResolver.ResolveLength(rem, new HtmlCssLengthResolutionContext()).Status);
        Assert.Equal(HtmlCssLengthResolutionStatus.MissingViewportWidth,
            HtmlCssMathResolver.ResolveLength(viewport, new HtmlCssLengthResolutionContext()).Status);
        Assert.Equal(30D, HtmlCssMathResolver.ResolveLength(container, new HtmlCssLengthResolutionContext {
            SmallViewportWidth = 400D,
            SmallViewportHeight = 300D
        }).Value);
        Assert.Equal(HtmlCssLengthResolutionStatus.NonFiniteValue,
            HtmlCssMathResolver.ResolveLength(Parse("1vw"), new HtmlCssLengthResolutionContext {
                ViewportWidth = double.NaN
            }).Status);

        using var canceled = new CancellationTokenSource();
        canceled.Cancel();
        Assert.Throws<OperationCanceledException>(() => HtmlCssMathResolver.ResolveLength(
            Parse("calc(1px + 2%)"), context: new HtmlCssLengthResolutionContext {
                PercentageReference = 100D
            }, cancellationToken: canceled.Token));
    }

    [Fact]
    public void MathParsingHonorsCancellationAndEveryDeclaredBudget() {
        using var canceled = new CancellationTokenSource();
        canceled.Cancel();
        Assert.Throws<OperationCanceledException>(() =>
            HtmlCssMathParser.ParseLengthPercentage("1px", cancellationToken: canceled.Token));
        Assert.Equal(nameof(HtmlCssMathOptions.MaxInputCharacters), Assert.Throws<HtmlCssMathLimitException>(() =>
            HtmlCssMathParser.ParseLengthPercentage("123px", new HtmlCssMathOptions { MaxInputCharacters = 4 })).LimitName);
        Assert.Equal(nameof(HtmlCssMathOptions.MaxInputCharacters), Assert.Throws<HtmlCssMathLimitException>(() =>
            HtmlCssMathParser.ParseLengthPercentage("    1px", new HtmlCssMathOptions { MaxInputCharacters = 4 })).LimitName);
        Assert.Equal(nameof(HtmlCssMathOptions.MaxInputCharacters), Assert.Throws<HtmlCssMathLimitException>(() =>
            HtmlCssMathParser.ParseLengthPercentage("     ", new HtmlCssMathOptions { MaxInputCharacters = 4 })).LimitName);
        Assert.Equal(nameof(HtmlCssMathOptions.MaxTokens), Assert.Throws<HtmlCssMathLimitException>(() =>
            HtmlCssMathParser.ParseLengthPercentage("calc(1px + 2px)", new HtmlCssMathOptions { MaxTokens = 2 })).LimitName);
        Assert.Equal(nameof(HtmlCssMathOptions.MaxNestingDepth), Assert.Throws<HtmlCssMathLimitException>(() =>
            HtmlCssMathParser.ParseLengthPercentage("calc(((1px)))", new HtmlCssMathOptions { MaxNestingDepth = 2 })).LimitName);
        Assert.Equal(nameof(HtmlCssMathOptions.MaxOperations), Assert.Throws<HtmlCssMathLimitException>(() =>
            HtmlCssMathParser.ParseLengthPercentage("calc(1px + 2px)", new HtmlCssMathOptions { MaxOperations = 1 })).LimitName);
        Assert.Equal(nameof(HtmlCssMathOptions.MaxArguments), Assert.Throws<HtmlCssMathLimitException>(() =>
            HtmlCssMathParser.ParseLengthPercentage("min(1px,2px,3px)", new HtmlCssMathOptions { MaxArguments = 2 })).LimitName);
    }

    [Fact]
    public void SelectedSizingAndSpacingPropertiesExposeTypedComputedValues() {
        HtmlConversionDocument document = HtmlConversionDocument.Parse("""
            <style>
              #target { width:calc(20px + 10%); height:40px; margin-left:-5%; padding-top:2em; }
            </style>
            <div id="target">Typed layout</div>
            """);
        OfficeIMO.Html.Dom.HtmlElement target = document.Document.QuerySelector("#target")!;
        HtmlComputedStyle style = HtmlComputedStyleEngine.Compute(document,
            new HtmlComputedStyleOptions { IncludeCascadeTraces = true })[target];

        Assert.True(style.TryGetTypedValue("width", out HtmlCssPropertyValue? width));
        Assert.Equal(HtmlCssNumericType.LengthPercentage, width!.MathExpression!.Type);
        Assert.Equal(40D, HtmlCssMathResolver.ResolveLength(width.MathExpression,
            new HtmlCssLengthResolutionContext { PercentageReference = 200D }).Value);
        Assert.Equal(HtmlCssPropertyParseStatus.Parsed, style.GetCascadeTrace("width")!.Candidates.Single().GrammarStatus);
        Assert.True(style.TryGetTypedValue("margin-left", out HtmlCssPropertyValue? margin));
        Assert.Equal(-5D, margin!.MathExpression!.Value);
        Assert.True(style.TryGetTypedValue("padding-top", out HtmlCssPropertyValue? padding));
        Assert.Equal(HtmlCssLengthUnit.Em, padding!.MathExpression!.Unit);

        Assert.Equal(HtmlCssPropertyParseStatus.UnsupportedValue,
            HtmlCssPropertyParser.Parse("width", "-1px").Status);
        Assert.Equal(HtmlCssPropertyParseStatus.Parsed,
            HtmlCssPropertyParser.Parse("margin-left", "-1px").Status);
        Assert.Equal(HtmlCssPropertyParseStatus.UnsupportedValue,
            HtmlCssPropertyParser.Parse("padding-left", "auto").Status);
    }

    [Fact]
    public void ParsedExpressionTreeIsStableAndReadOnly() {
        HtmlCssMathExpression expression = Parse("calc(10px + 25%)");

        Assert.Equal(HtmlCssMathExpressionKind.Calc, expression.Kind);
        HtmlCssMathExpression sum = Assert.Single(expression.Children);
        Assert.Equal(HtmlCssMathExpressionKind.Add, sum.Kind);
        Assert.Equal(new[] { HtmlCssNumericType.Length, HtmlCssNumericType.Percentage },
            sum.Children.Select(child => child.Type).ToArray());
        Assert.IsAssignableFrom<System.Collections.IList>(sum.Children);
        Assert.True(((System.Collections.IList)sum.Children).IsReadOnly);
    }

    [Fact]
    public void EveryDeclaredLengthUnitMapsToOneTypedLiteral() {
        var names = new Dictionary<HtmlCssLengthUnit, string> {
            [HtmlCssLengthUnit.Px] = "px", [HtmlCssLengthUnit.Pt] = "pt",
            [HtmlCssLengthUnit.Pc] = "pc", [HtmlCssLengthUnit.In] = "in",
            [HtmlCssLengthUnit.Cm] = "cm", [HtmlCssLengthUnit.Mm] = "mm",
            [HtmlCssLengthUnit.Q] = "q", [HtmlCssLengthUnit.Em] = "em",
            [HtmlCssLengthUnit.Rem] = "rem", [HtmlCssLengthUnit.Vw] = "vw",
            [HtmlCssLengthUnit.Vh] = "vh", [HtmlCssLengthUnit.Vmin] = "vmin",
            [HtmlCssLengthUnit.Vmax] = "vmax", [HtmlCssLengthUnit.Svw] = "svw",
            [HtmlCssLengthUnit.Svh] = "svh", [HtmlCssLengthUnit.Svmin] = "svmin",
            [HtmlCssLengthUnit.Svmax] = "svmax", [HtmlCssLengthUnit.Lvw] = "lvw",
            [HtmlCssLengthUnit.Lvh] = "lvh", [HtmlCssLengthUnit.Lvmin] = "lvmin",
            [HtmlCssLengthUnit.Lvmax] = "lvmax", [HtmlCssLengthUnit.Dvw] = "dvw",
            [HtmlCssLengthUnit.Dvh] = "dvh", [HtmlCssLengthUnit.Dvmin] = "dvmin",
            [HtmlCssLengthUnit.Dvmax] = "dvmax", [HtmlCssLengthUnit.Cqw] = "cqw",
            [HtmlCssLengthUnit.Cqh] = "cqh", [HtmlCssLengthUnit.Cqi] = "cqi",
            [HtmlCssLengthUnit.Cqb] = "cqb", [HtmlCssLengthUnit.Cqmin] = "cqmin",
            [HtmlCssLengthUnit.Cqmax] = "cqmax"
        };

        Assert.Equal(Enum.GetValues(typeof(HtmlCssLengthUnit)).Length, names.Count);
        foreach (KeyValuePair<HtmlCssLengthUnit, string> item in names) {
            HtmlCssMathExpression expression = Parse("1" + item.Value);
            Assert.Equal(item.Key, expression.Unit);
            Assert.Equal(HtmlCssNumericType.Length, expression.Type);
        }
    }

    private static HtmlCssMathExpression Parse(string value) {
        HtmlCssMathParseResult parsed = HtmlCssMathParser.ParseLengthPercentage(value);
        Assert.Equal(HtmlCssMathParseStatus.Parsed, parsed.Status);
        return parsed.Expression!;
    }

    private static T[] Read<T>(string name) => JsonSerializer.Deserialize<T[]>(File.ReadAllText(
        Path.Combine(AppContext.BaseDirectory, "Documents", "Html", "Css", name)),
        new JsonSerializerOptions { PropertyNameCaseInsensitive = true })!;

    private static T ParseEnum<T>(string value) where T : struct => (T)Enum.Parse(typeof(T), value);

    private sealed class LengthMathCase {
        public string Name { get; set; } = string.Empty;
        public string Value { get; set; } = string.Empty;
        public string ParseStatus { get; set; } = string.Empty;
        public string? Type { get; set; }
        public string? Canonical { get; set; }
        public double? Resolved { get; set; }
        public string? ResolutionStatus { get; set; }
    }
}
