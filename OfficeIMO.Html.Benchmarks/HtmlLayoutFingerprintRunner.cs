using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using OfficeIMO.Drawing;

namespace OfficeIMO.Html.Benchmarks;

/// <summary>All-page output evidence collected separately from timed layout measurements.</summary>
internal static class HtmlLayoutFingerprintRunner {
    internal static int Run(string[] args) {
        if (args.Length != 1) {
            Console.Error.WriteLine("Usage: --layout-fingerprint <workload>");
            return 2;
        }
        try {
            HtmlLayoutScenario scenario = HtmlLayoutScenario.Create(args[0]);
            HtmlRenderDocument rendered = scenario.Render();
            scenario.Validate(rendered);
            var pages = rendered.Pages.Select(page => new {
                page.PageNumber, page.Width, page.Height,
                SvgSha256 = Hash(OfficeDrawingSvgExporter.ToSvg(page.CreateDrawing()))
            }).ToArray();
            Console.WriteLine(JsonSerializer.Serialize(new {
                Workload = args[0], TextSha256 = Hash(rendered.Text), Pages = pages,
                Diagnostics = rendered.Diagnostics.Select(item => new { item.Code, item.Severity, item.LossKind, item.Message }).ToArray()
            }, new JsonSerializerOptions { WriteIndented = true }));
            return 0;
        } catch (Exception exception) {
            Console.Error.WriteLine(exception);
            return 1;
        }
    }

    private static string Hash(string value) => Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(value))).ToLowerInvariant();
}
