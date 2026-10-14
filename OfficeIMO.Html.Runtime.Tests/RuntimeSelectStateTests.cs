using OfficeIMO.Html.Providers;
using OfficeIMO.Html.Runtime;
using Xunit;

namespace OfficeIMO.Tests
{
    public sealed class RuntimeSelectStateTests {
        private static HtmlProcessRuntimeProvider Runtime() => new(
            Path.Combine(AppContext.BaseDirectory, "RuntimeWorker", "OfficeIMO.Html.Runtime.Worker.dll"), AngleSharpDomServices.Instance);

        [Theory]
        [InlineData("select.value = 'b'")]
        [InlineData("select.selectedIndex = 1")]
        [InlineData("select.options.selectedIndex = 1")]
        public async Task SelectingAPeerKeepsCleanOptionsResponsiveToDefaultChanges(string selection) {
            var capture = await Runtime().CaptureTrustedAsync(new HtmlScriptRequest {
                Html = "<select id='choice' size='2'><option id='a' value='a' selected>Alpha</option><option value='b'>Beta</option></select>",
                Scripts = new[] { "const select = document.querySelector('select'); " + selection + "; " +
                    "const first = document.querySelector('#a'); first.defaultSelected = false; first.defaultSelected = true;" }
            });

            Assert.Equal("a", Assert.Single(capture.Document.QuerySelectorAll("option:checked")).GetAttribute("value"));
        }

        [Theory]
        [InlineData("select.innerHTML = '<option value=a>Alpha</option>'")]
        [InlineData("const option = document.createElement('option'); option.value = 'a'; select.appendChild(option)")]
        [InlineData("const option = document.createElement('option'); option.value = 'a'; select.options.add(option)")]
        public async Task DynamicSingleSelectOptionsExposeTheDefaultSelectionToCapture(string insertion) {
            var capture = await Runtime().CaptureTrustedAsync(new HtmlScriptRequest {
                Html = "<select name='choice'></select>",
                Scripts = new[] { "const select = document.querySelector('select'); " + insertion }
            });

            Assert.Equal("a", Assert.Single(capture.Document.QuerySelectorAll("option:checked")).GetAttribute("value"));
        }

        [Fact]
        public async Task ExplicitlyClearedSelectionSurvivesCapture() {
            var capture = await Runtime().CaptureTrustedAsync(new HtmlScriptRequest {
                Html = "<select><option value='a'>Alpha</option></select>",
                Scripts = new[] { "document.querySelector('select').value = 'missing'" }
            });

            Assert.Empty(capture.Document.QuerySelectorAll("option:checked"));
        }

        [Fact]
        public async Task ExplicitlyChangedOptionIgnoresLaterDefaultChanges() {
            var capture = await Runtime().CaptureTrustedAsync(new HtmlScriptRequest {
                Html = "<select size='2'><option selected>Alpha</option></select>",
                Scripts = new[] { "const option = document.querySelector('option'); option.selected = false; " +
                    "option.defaultSelected = false; option.defaultSelected = true;" }
            });

            Assert.Empty(capture.Document.QuerySelectorAll("option:checked"));
        }

        [Fact]
        public async Task DeselectingTheOnlyDropdownOptionRestoresItsAutomaticSelection() {
            var capture = await Runtime().CaptureTrustedAsync(new HtmlScriptRequest {
                Html = "<select><option value='a' selected>Alpha</option></select>",
                Scripts = new[] { "document.querySelector('option').selected = false" }
            });

            Assert.Equal("a", Assert.Single(capture.Document.QuerySelectorAll("option:checked")).GetAttribute("value"));
        }

        [Theory]
        [InlineData("const target = document.querySelector('option')")]
        [InlineData("const target = document.querySelector('optgroup')")]
        public async Task DisabledChangesPreserveAnExplicitlyClearedSelection(string target) {
            var capture = await Runtime().CaptureTrustedAsync(new HtmlScriptRequest {
                Html = "<select><optgroup><option value='a'>Alpha</option><option value='b'>Beta</option></optgroup></select>",
                Scripts = new[] { "document.querySelector('select').selectedIndex = -1; " + target +
                    "; target.disabled = true; target.disabled = false;" }
            });

            Assert.Empty(capture.Document.QuerySelectorAll("option:checked"));
        }
    }
}
