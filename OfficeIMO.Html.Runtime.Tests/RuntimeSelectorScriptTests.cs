using System.Text.Json;
using OfficeIMO.Html.Providers;
using OfficeIMO.Html.Runtime;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class RuntimeSelectorScriptTests {
    private static HtmlProcessRuntimeProvider Runtime() => new(
        Path.Combine(AppContext.BaseDirectory, "RuntimeWorker", "OfficeIMO.Html.Runtime.Worker.dll"), AngleSharpDomServices.Instance);

    [Theory]
    [InlineData("", "BackCompat", true)]
    [InlineData("<!doctype html>", "CSS1Compat", false)]
    [InlineData("<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.01 Transitional//EN\">", "BackCompat", true)]
    [InlineData("<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.01 Frameset//EN\">", "BackCompat", true)]
    [InlineData("<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.01 Transitional//EN\" \"\">", "BackCompat", true)]
    [InlineData("<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.01 Frameset//EN\" \"\">", "BackCompat", true)]
    [InlineData("<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.01 Transitional//EN\" \"http://www.w3.org/TR/html4/loose.dtd\">", "CSS1Compat", false)]
    [InlineData("<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.01 Frameset//EN\" \"http://www.w3.org/TR/html4/frameset.dtd\">", "CSS1Compat", false)]
    public async Task ScriptIdentitySelectorsUseDocumentMode(string doctype, string mode, bool folded) {
        await using var page = await Runtime().OpenTrustedAsync(new HtmlScriptRequest {
            Html = doctype + "<div id='SAMPLE' class='SAMPLE'></div>"
        });
        JsonElement result = await page.EvaluateAsync("""
            (()=>{const element=document.querySelector('div');return {mode:document.compatMode,
                id:document.querySelector('#sample')===element,class:document.querySelector('.sample')===element,
                idAll:document.querySelectorAll('#sample').length,classAll:document.querySelectorAll('.sample').length,
                matchesId:element.matches('#sample'),matchesClass:element.matches('.sample'),
                exactId:element.matches('#SAMPLE'),exactClass:element.matches('.SAMPLE')};})()
            """);
        Assert.Equal(mode, result.GetProperty("mode").GetString());
        foreach (string name in new[] { "id", "class", "matchesId", "matchesClass" })
            Assert.Equal(folded, result.GetProperty(name).GetBoolean());
        Assert.Equal(folded ? 1 : 0, result.GetProperty("idAll").GetInt32());
        Assert.Equal(folded ? 1 : 0, result.GetProperty("classAll").GetInt32());
        Assert.True(result.GetProperty("exactId").GetBoolean());
        Assert.True(result.GetProperty("exactClass").GetBoolean());
    }

    [Theory]
    [InlineData("", "BackCompat", "p")]
    [InlineData("<!doctype html>", "CSS1Compat", "div")]
    [InlineData("<!DOCTYPE HTML PUBLIC \"-//W3C//DTD XHTML 1.0 Transitional//EN\" \"http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd\">", "CSS1Compat", "div")]
    public async Task ScriptMarkupSettersUseContextDocumentMode(string doctype, string mode, string tableParent) {
        await using var page = await Runtime().OpenTrustedAsync(new HtmlScriptRequest {
            Html = doctype + "<div id='inner'></div><div id='outer'><span></span></div>"
        });
        JsonElement result = await page.EvaluateAsync("""
            (()=>{const source='<p>Paragraph<table><tr><td>Cell</td></tr></table>';
                const inner=document.getElementById('inner'),outer=document.getElementById('outer');
                inner.innerHTML=source;outer.firstElementChild.outerHTML=source;
                function collect(context){const table=context.querySelector('table');return {
                    tableParent:table.parentElement.localName,adopted:table.ownerDocument===document};}
                return {mode:document.compatMode,inner:collect(inner),outer:collect(outer)};})()
            """);
        Assert.Equal(mode, result.GetProperty("mode").GetString());
        foreach (string name in new[] { "inner", "outer" }) {
            Assert.Equal(tableParent, result.GetProperty(name).GetProperty("tableParent").GetString());
            Assert.True(result.GetProperty(name).GetProperty("adopted").GetBoolean());
        }
    }

    [Fact]
    public async Task ScriptLanguageSelectorsRespectInheritanceNamespaceAndLanguageBoundary() {
        await using var page = await Runtime().OpenTrustedAsync(new HtmlScriptRequest {
            Html = "<!doctype html><section lang='en-US'><div id='inherited'></div>"
                + "<p id='unknown' lang=''></p><p id='otherNamespace'></p></section>"
                + "<p id='declared' lang='en-US'></p><p id='prefix' lang='english'></p><p id='xml' lang='de'></p>"
        });
        JsonElement result = await page.EvaluateAsync("""
            (()=>{const xml='http://www.w3.org/XML/1998/namespace';
                document.getElementById('xml').setAttributeNS(xml,'xml:lang','fr-CA');
                document.getElementById('otherNamespace').setAttributeNS('urn:unrelated','u:lang','de');
                const root=document.createElementNS('urn:custom','root');root.setAttributeNS(xml,'xml:lang','it-IT');
                const child=document.createElementNS('urn:custom','child');root.appendChild(child);document.body.appendChild(root);
                function collect(selector,id){const element=document.getElementById(id);return {
                    matches:element.matches(selector),query:[...document.querySelectorAll(selector)].some(x=>x===element)};}
                return {declared:collect(':lang(en)','declared'),inherited:collect(':lang(en)','inherited'),
                    prefix:collect(':lang(en)','prefix'),unknown:collect(':lang(en)','unknown'),
                    unrelated:collect(':lang(en)','otherNamespace'),xml:collect(':lang(fr)','xml'),
                    htmlShadowed:collect(':lang(de)','xml'),foreignInherited:child.matches(':lang(it)'),
                    foreignQuery:[...document.querySelectorAll(':lang(it)')].some(x=>x===child)};})()
            """);
        foreach (string name in new[] { "declared", "inherited", "unrelated", "xml" }) {
            Assert.True(result.GetProperty(name).GetProperty("matches").GetBoolean(), name);
            Assert.True(result.GetProperty(name).GetProperty("query").GetBoolean(), name);
        }
        foreach (string name in new[] { "prefix", "unknown", "htmlShadowed" }) {
            Assert.False(result.GetProperty(name).GetProperty("matches").GetBoolean(), name);
            Assert.False(result.GetProperty(name).GetProperty("query").GetBoolean(), name);
        }
        Assert.True(result.GetProperty("foreignInherited").GetBoolean());
        Assert.True(result.GetProperty("foreignQuery").GetBoolean());
    }
}
