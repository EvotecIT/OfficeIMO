using System;
using System.Collections;
using System.IO;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using OfficeIMO.Word.Html;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Html {
        private static IDictionary GetCache() {
            var assembly = typeof(HtmlToWordOptions).Assembly;
            var converterType = assembly.GetType("OfficeIMO.Word.Html.Converters.HtmlToWordConverter", true);
            var field = converterType.GetField("_stylesheetCache", BindingFlags.NonPublic | BindingFlags.Static);
            return (IDictionary)field!.GetValue(null)!;
        }

        private static string ComputeHash(string css) {
            using var sha = SHA256.Create();
            var bytes = sha.ComputeHash(Encoding.UTF8.GetBytes(css));
            return BitConverter.ToString(bytes).Replace("-", "");
        }

        [Fact]
        public void HtmlToWord_StylesheetCache_Reused_ForPath() {
            var path = Path.GetTempFileName();
            File.WriteAllText(path, "p { color:#111111; }");
            try {
                var cache = GetCache();
                cache.Remove(path);
                var html = $"<link rel=\"stylesheet\" href=\"{path}\" /><p>Test</p>";
                html.LoadFromHtml(new HtmlToWordOptions());
                var first = cache[path];
                html.LoadFromHtml(new HtmlToWordOptions());
                Assert.Same(first, cache[path]);
            } finally {
                File.Delete(path);
            }
        }

        [Fact]
        public void HtmlToWord_StylesheetCache_Reused_ForContent() {
            const string css = "p { color:#222222; }";
            var key = ComputeHash(css);
            var cache = GetCache();
            cache.Remove(key);
            var html = $"<style>{css}</style><p>Test</p>";
            html.LoadFromHtml(new HtmlToWordOptions());
            var first = cache[key];
            html.LoadFromHtml(new HtmlToWordOptions());
            Assert.Same(first, cache[key]);
        }
    }
}
