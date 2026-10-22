using System;
using OfficeIMO.Html;
using OfficeIMO.Html.Dom;
using Xunit;

namespace OfficeIMO.Html.Tests;

public sealed class HtmlNormalizedResourceCharacterLimitTests {
    [Theory]
    [InlineData("img")]
    [InlineData("srcset")]
    [InlineData("css-inline")]
    [InlineData("css-style")]
    [InlineData("css-overlap")]
    [InlineData("css-mixed-shrink")]
    public void Transformed_resources_observe_the_exact_shared_serialized_character_boundary(string scenario) {
        string source = scenario switch {
            "img" => "<img src='cid:logo'><img src='cid:logo'>",
            "srcset" => "<img srcset='cid:logo 1x, cid:logo 2x'>",
            "css-inline" => "<p style='background:url(cid:logo),url(cid:logo)'>text</p>",
            "css-style" => "<div><style>.x{background:url(cid:logo),url(cid:logo)}</style><p class='x'>text</p></div>",
            "css-overlap" => "<div><style>@import url(cid:logo);</style></div>",
            "css-mixed-shrink" => "<div><style>.x{background:url(cid:shrinkabcdefghijklmnopqrstuvwxyz),url(cid:logo)}</style></div>",
            _ => throw new ArgumentOutOfRangeException(nameof(scenario))
        };
        string resource = "data:image/png;base64," + new string('A', 64);
        HtmlNormalizationOptions Options(int? maximum) {
            HtmlUrlPolicy policy = HtmlUrlPolicy.CreateOfficeIMOProfile();
            policy.ResolvedUrlTransform = value => value.StartsWith("cid:shrink", StringComparison.OrdinalIgnoreCase) ? string.Empty :
                value.StartsWith("cid:", StringComparison.OrdinalIgnoreCase) ? resource : value;
            return new HtmlNormalizationOptions {
                ResourceUrlPolicy = policy,
                PreserveStyleElements = true,
                Limits = new HtmlConversionLimits { MaxInputCharacters = maximum }
            };
        }
        string expected = HtmlNormalizer.Normalize(source, Options(null));
        Assert.True(expected.Length > source.Length);
        Assert.Contains(resource, expected);

        // At the actual serialized boundary the same page fits. The overlapping
        // CSS import/url discoveries must not be charged as two emitted values.
        Assert.Equal(expected, HtmlNormalizer.Normalize(source, Options(expected.Length)));
        HtmlDomLimitException exception = Assert.Throws<HtmlDomLimitException>(() =>
            HtmlNormalizer.Normalize(source, Options(expected.Length - 1)));
        Assert.Equal(nameof(HtmlConversionLimits.MaxInputCharacters), exception.LimitSource);
        Assert.True(exception.Actual > exception.Limit);
    }

    [Theory]
    [InlineData("document-trailing")]
    [InlineData("inline-css-outer")]
    public void Discarded_outer_whitespace_does_not_reject_a_normalized_value_that_fits(string scenario) {
        string source = scenario == "document-trailing"
            ? "<img src='cid:logo'> \t"
            : "<p style='" + new string(' ', 30) + "background:url(cid:logo)" + new string(' ', 30) + "'>text</p>";
        string resource = "data:image/png;base64," + new string('A', 64);
        HtmlNormalizationOptions Options(int? maximum) {
            HtmlUrlPolicy policy = HtmlUrlPolicy.CreateOfficeIMOProfile();
            policy.ResolvedUrlTransform = value => value.StartsWith("cid:", StringComparison.OrdinalIgnoreCase) ? resource : value;
            return new HtmlNormalizationOptions {
                ResourceUrlPolicy = policy,
                Limits = new HtmlConversionLimits { MaxInputCharacters = maximum }
            };
        }
        string expected = HtmlNormalizer.Normalize(source, Options(null));
        Assert.True(expected.Length > source.Length);
        Assert.Contains(resource, expected);

        Assert.Equal(expected, HtmlNormalizer.Normalize(source, Options(expected.Length)));
    }
}

public sealed class HtmlNormalizationElementCorrelationTests {
    [Theory]
    [InlineData("input")]
    [InlineData("textarea")]
    [InlineData("select")]
    [InlineData("option")]
    public void Live_form_state_does_not_consume_the_canonical_character_boundary(string control) {
        string html = control switch {
            "input" => "<input value='seed'>",
            "textarea" => "<textarea>seed</textarea>",
            _ => "<select><option value='one'>One</option><option value='two'>Two</option></select>"
        };
        HtmlDocument authored = HtmlConversionDocument.Parse(html).Document.CloneAttached();
        HtmlElement element = authored.QuerySelector(control)!;
        HtmlFormControlState state = control switch {
            "input" => new HtmlFormControlState(HtmlFormControlStateKind.Input, "live"),
            "textarea" => new HtmlFormControlState(HtmlFormControlStateKind.TextArea, "live"),
            "select" => new HtmlFormControlState(HtmlFormControlStateKind.Select),
            _ => new HtmlFormControlState(HtmlFormControlStateKind.Option, isSelected: true)
        };
        element.FormState = state;
        if (control == "select") {
            authored.QuerySelectorAll("option")[0].FormState = new HtmlFormControlState(HtmlFormControlStateKind.Option);
            authored.QuerySelectorAll("option")[1].FormState = new HtmlFormControlState(HtmlFormControlStateKind.Option, isSelected: true);
        }
        string original = authored.DocumentElement!.OuterHtml;
        HtmlConversionDocument unlimited = HtmlConversionDocument.FromDocument(authored);
        int maximum = control == "input" ? 80 : unlimited.SourceHtml.Length;
        Assert.True(unlimited.SourceHtml.Length <= maximum);
        HtmlConversionDocument bounded = HtmlConversionDocument.FromDocument(authored, new HtmlConversionDocumentOptions {
            Limits = new HtmlConversionLimits { MaxInputCharacters = maximum }
        });

        HtmlDocument converted = bounded.CreateDocumentForConversion();

        HtmlFormControlState actual = Assert.IsType<HtmlFormControlState>(converted.QuerySelector(control)!.FormState);
        Assert.Equal(state.Kind, actual.Kind);
        Assert.Equal(state.Value, actual.Value);
        Assert.Equal(state.IsSelected, actual.IsSelected);
        if (control == "select") {
            Assert.False(converted.QuerySelectorAll("option")[0].FormState!.IsSelected);
            Assert.True(converted.QuerySelectorAll("option")[1].FormState!.IsSelected);
        }
        Assert.Equal(original, authored.DocumentElement.OuterHtml);
        Assert.DoesNotContain("data-form-state-", converted.DocumentElement!.OuterHtml);
    }

    [Theory]
    [InlineData("region")]
    [InlineData("image")]
    [InlineData("region-image")]
    public void Editable_layout_identity_does_not_consume_the_canonical_character_boundary(string scenario) {
        const string image = "<img alt='Logo' src='data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jf0sAAAAASUVORK5CYII='>";
        string html = scenario == "image" ? image :
            "<div style='position:absolute;width:40px;height:20px'>" +
            (scenario == "region-image" ? image : "Region") + "</div>";
        HtmlConversionDocument unlimited = HtmlConversionDocument.Parse(html);
        int maximum = unlimited.CreateDocumentForConversion().DocumentElement!.OuterHtml.Length;
        HtmlConversionDocument bounded = HtmlConversionDocument.Parse(html, new HtmlConversionDocumentOptions {
            Limits = new HtmlConversionLimits { MaxInputCharacters = maximum }
        });

        HtmlEditableLayoutProjection projection = HtmlEditableLayoutProjector.Project(bounded);

        if (scenario == "image") {
            Assert.Empty(projection.Regions);
            Assert.Equal("Logo", projection.RemainingDocument.QuerySelector("img")!.GetAttribute("alt"));
        } else {
            HtmlRenderLayoutRegion region = Assert.Single(projection.Regions);
            if (scenario == "region-image") Assert.Equal("Logo", Assert.Single(projection.GetSourceImages(region)).GetAttribute("alt"));
        }
        string remaining = projection.RemainingDocument.DocumentElement!.OuterHtml;
        Assert.DoesNotContain("data-officeimo-editable-layout-", remaining);
        Assert.Equal(html, bounded.SourceHtml);
    }

    [Fact]
    public void Authored_correlation_lookalikes_remain_real_bounded_attributes() {
        const string form = "data-form-state-0123456789abcdef0123456789abcdef";
        const string region = "data-officeimo-editable-layout-region";
        const string image = "data-officeimo-editable-layout-image";
        HtmlDocument authored = HtmlConversionDocument.Parse("<input value='seed' " + form + "='form' " +
            region + "='region' " + image + "='image'>").Document.CloneAttached();
        authored.QuerySelector("input")!.FormState = new HtmlFormControlState(HtmlFormControlStateKind.Input, "live");
        HtmlConversionDocument unlimited = HtmlConversionDocument.FromDocument(authored);
        int maximum = unlimited.SourceHtml.Length;
        HtmlConversionDocument bounded = HtmlConversionDocument.FromDocument(authored, new HtmlConversionDocumentOptions {
            Limits = new HtmlConversionLimits { MaxInputCharacters = maximum }
        });

        HtmlElement converted = bounded.CreateDocumentForConversion().QuerySelector("input")!;

        Assert.Equal("form", converted.GetAttribute(form));
        Assert.Equal("region", converted.GetAttribute(region));
        Assert.Equal("image", converted.GetAttribute(image));
        Assert.Equal("live", converted.FormState!.Value);
        Assert.Throws<HtmlDomLimitException>(() => HtmlConversionDocument.FromDocument(authored, new HtmlConversionDocumentOptions {
            Limits = new HtmlConversionLimits { MaxInputCharacters = maximum - 1 }
        }));
    }

    [Theory]
    [InlineData("unicode-prefix")]
    [InlineData("dropped-content")]
    [InlineData("fostered-table")]
    [InlineData("foreign-namespace")]
    public void Element_identity_uses_exact_normalized_opening_offsets_after_structural_parsing(string scenario) {
        string html = scenario switch {
            "unicode-prefix" => "<p>Résumé \U0001f642</p><input value='seed'>",
            "dropped-content" => "<template><p>Hidden</p></template><script>removed()</script><input value='seed'>",
            "fostered-table" => "<table><input value='seed'><tr><td>Cell</td></tr></table>",
            _ => "<svg><g><foreignObject><input value='seed'></foreignObject></g></svg>"
        };
        HtmlDocument authored = HtmlConversionDocument.Parse(html).Document.CloneAttached();
        authored.QuerySelector("input")!.FormState = new HtmlFormControlState(HtmlFormControlStateKind.Input, "live");
        HtmlConversionDocument baseline = HtmlConversionDocument.FromDocument(authored);
        string before = authored.DocumentElement!.OuterHtml;
        var native = baseline.CreateSourceDocumentForConversion();
        var normalization = new HtmlNormalizationOptions {
            UseBodyContentsOnly = false, PreserveSkippedElementMarkers = true,
            CollapseTextWhitespace = false, Limits = HtmlConversionLimits.CreateTrustedProfile()
        };
        string normalized = HtmlNormalizer.Normalize(native, normalization);
        int maximum = Math.Max(baseline.SourceHtml.Length, normalized.Length);
        HtmlConversionDocument bounded = HtmlConversionDocument.FromDocument(authored, new HtmlConversionDocumentOptions {
            Limits = new HtmlConversionLimits { MaxInputCharacters = maximum }
        });

        HtmlDocument converted = bounded.CreateDocumentForConversion();

        Assert.Equal("live", converted.QuerySelector("input")!.FormState!.Value);
        // The public adapter clone intentionally clears provider source references.
        // Observe the production correlation callback before that later clone.
        normalization.Limits = bounded.Limits.Clone();
        int matches = 0;
        HtmlNormalizer.NormalizeToDocument(native, normalization, elementNormalized: (original, target) => {
            if (target.LocalName != "input") return;
            matches++;
            Assert.Same(native.QuerySelector("input"), original);
            Assert.Equal(original.NamespaceUri, target.NamespaceUri);
            Assert.Equal(normalized.IndexOf("<input", StringComparison.OrdinalIgnoreCase), target.SourceReference!.Position.Index);
        });
        Assert.Equal(1, matches);
        Assert.Equal(before, authored.DocumentElement.OuterHtml);
        if (scenario == "dropped-content") {
            Assert.DoesNotContain("Hidden", converted.DocumentElement.OuterHtml);
            Assert.DoesNotContain("removed()", converted.DocumentElement.OuterHtml);
        }
    }

    [Fact]
    public void Live_identity_does_not_relax_real_normalized_output_limits() {
        HtmlDocument authored = HtmlConversionDocument.Parse("<input value='seed'><a href='next'>Next</a>").Document.CloneAttached();
        authored.QuerySelector("input")!.FormState = new HtmlFormControlState(HtmlFormControlStateKind.Input, "live");
        int maximum = HtmlConversionDocument.FromDocument(authored).SourceHtml.Length;
        HtmlUrlPolicy policy = HtmlUrlPolicy.CreateOfficeIMOProfile();
        policy.ResolvedUrlTransform = _ => "https://example.test/" + new string('x', maximum);
        HtmlConversionDocument bounded = HtmlConversionDocument.FromDocument(authored, new HtmlConversionDocumentOptions {
            BaseUri = new Uri("https://example.test/"), UrlPolicy = policy,
            Limits = new HtmlConversionLimits { MaxInputCharacters = maximum }
        });

        HtmlDomLimitException error = Assert.Throws<HtmlDomLimitException>(() => bounded.CreateDocumentForConversion());

        Assert.Equal(nameof(HtmlConversionLimits.MaxInputCharacters), error.LimitSource);
        Assert.True(error.Actual > error.Limit);
        Assert.Equal("live", authored.QuerySelector("input")!.FormState!.Value);
    }
}
