using HtmlForgeX;
using HtmlForgeX.Tags;

namespace SyntheticReportProducer;

internal static class ReportComposer {
    internal static string Compose(ReportSpec report, VariantSpec variant, string css) {
        var links = report.Links.ToDictionary(link => link.Id, StringComparer.Ordinal);
        var tables = report.Tables.ToDictionary(table => table.Key, StringComparer.Ordinal);
        var main = new HtmlTag("main").Id(report.Key).Class("synthetic-report");
        main.AddChild(new HtmlTag("h1").ValueText(report.Title));
        main.AddChild(new HtmlTag("p").Class("evidence-note").ValueText(
            "New synthetic independent-producer report. This fixture does not represent the unavailable historical report corpus."));
        main.AddChild(Section(report.Intro));

        var navigation = new HtmlTag("nav").Class("report-navigation").Attribute("aria-label", "Report sections");
        foreach (var linkId in report.NavigationLinkIds) {
            navigation.AddChild(Link(links[linkId]));
        }
        main.AddChild(navigation);

        foreach (var region in report.Regions) {
            var section = new HtmlTag("section").Id(region.Id).Class($"report-region parallel-{region.Layout}");
            section.AddChild(new HtmlTag("h2").ValueText($"{region.Label} — {region.Title}"));
            var grid = new MonitoringGrid();
            foreach (var tableKey in region.TableKeys) {
                var table = tables[tableKey];
                var panelLink = links[table.PanelLinkId];
                grid.Panel(table.PanelTitle, panel => panel
                    .Link(panelLink.Label, panelLink.Href)
                    .Content(Flow(table, links)), span: 6);
            }
            section.AddChild(grid);
            section.AddChild(Paragraph(region.After, links));
            main.AddChild(section);
        }

        if (report.DetailsSection is { } detailsSpec) {
            var detailsSection = new HtmlTag("section").Id(detailsSpec.Id).Class("details-assessment");
            detailsSection.AddChild(new HtmlTag("h2").ValueText(detailsSpec.Label));
            detailsSection.AddChild(new HtmlTag("p").ValueText(detailsSpec.Text));
            foreach (var item in detailsSpec.Cases) {
                var details = new HtmlTag("details").Id(item.Id)
                    .Attribute(HtmlBooleanAttribute.Open, item.Open);
                details.AddChild(new HtmlTag("summary").ValueText(item.SummaryLabel));
                details.AddChild(Paragraph(item.Body, links));
                detailsSection.AddChild(details);
            }
            main.AddChild(detailsSection);
        }

        main.AddChild(Section(report.Summary));
        main.AddChild(Link(links[report.SummaryLinkId]));

        // Only the trusted local stylesheet is emitted as raw text. Tables, panels,
        // grids, queues, links, and all dynamic text use the inspected typed APIs.
        var head = new HtmlTag("head")
            .AddChild(new HtmlTag("meta", TagMode.NoClosing).Attribute("charset", "utf-8"))
            .AddChild(new HtmlTag("meta", TagMode.NoClosing).Attribute("name", "viewport").Attribute("content", "width=device-width, initial-scale=1"))
            .AddChild(new HtmlTag("meta", TagMode.NoClosing).Attribute("name", "fixture-classification").Attribute("content", "synthetic-independent-producer"))
            .AddChild(new HtmlTag("meta", TagMode.NoClosing).Attribute("name", "fixture-variant").Attribute("content", variant.Key))
            .AddChild(new HtmlTag("title").ValueText($"{report.Title} ({variant.Key})"))
            .AddChild(new HtmlTag("style").ValueRaw(TrustedHtml.FromTrustedSource(css)));
        var html = new HtmlTag("html").Attribute("lang", "en")
            .AddChild(head)
            .AddChild(new HtmlTag("body").AddChild(main));
        return "<!DOCTYPE html>\n" + html + "\n";
    }

    private static HtmlTag Flow(TableSpec spec, IReadOnlyDictionary<string, LinkSpec> links) {
        var flow = new HtmlTag("div").Id(spec.SurfaceId).Class("report-flow");
        flow.AddChild(spec.Component switch {
            "Table" => Register(spec, links),
            "MonitoringTable" => MonitoringRegister(spec, links),
            "MonitoringAlertQueue" => Queue(spec),
            _ => throw new InvalidOperationException($"Unknown typed component: {spec.Component}")
        });
        if (spec.Component == "MonitoringAlertQueue" && spec.TerminalSummary is { } queueSummary) {
            flow.AddChild(new HtmlTag("p").Class("queue-terminal-summary").ValueText(queueSummary));
        }
        flow.AddChild(Paragraph(spec.After, links));
        return flow;
    }

    private static Element Register(TableSpec spec, IReadOnlyDictionary<string, LinkSpec> links) {
        var table = new Table().AddHeaders(spec.Headers).AddFooters(spec.FooterLabels);
        foreach (var row in spec.Rows) {
            table.AddRow(
                TableCell.FromElement(RowMarker(row)),
                TableCell.TextCell(row.Detail),
                TableCell.TextCell(row.Value),
                TableCell.FromElement(Link(links[RequiredLinkId(row)])));
        }
        // Table owns the full row-group markup; its public output contract is a
        // fragment, so this single semantic table shell supplies its container.
        return new HtmlTag("table").Id(spec.Key + "-table").Class("report-register")
            .Attribute("aria-label", spec.AccessibleLabel).AddChild(table);
    }

    private static Element MonitoringRegister(TableSpec spec, IReadOnlyDictionary<string, LinkSpec> links) {
        var table = new MonitoringTable().Headers(spec.Headers)
            .Settings(settings => settings.AccessibleLabel(spec.AccessibleLabel)
                .Caption(spec.Caption ?? spec.AccessibleLabel, visible: true).Target(spec.Key).End());
        foreach (var row in spec.Rows) {
            table.Row(RowMarker(row), row.Detail, row.Value, Link(links[RequiredLinkId(row)]));
        }
        if (spec.TerminalSummary is { } summary) table.Footer(summary);
        return table;
    }

    private static Element Queue(TableSpec spec) {
        var queue = new MonitoringAlertQueue().Settings(settings => settings
            .AccessibleLabel(spec.AccessibleLabel)
            .Caption(spec.Caption ?? spec.AccessibleLabel, visible: true).Target(spec.Key).End());
        foreach (var row in spec.Rows) {
            queue.AddAlert(row.Label, row.Affected ?? string.Empty,
                Enum.Parse<MonitoringHealthState>(row.Severity ?? "Unknown"),
                row.Status ?? string.Empty, row.Owner ?? string.Empty,
                row.FirstSeen ?? string.Empty, row.LastSeen ?? string.Empty,
                row.Occurrences, row.Detail, selected: false);
        }
        return queue;
    }

    private static HtmlTag RowMarker(RowSpec row) => new HtmlTag("span")
        .Id(row.DomId ?? throw new InvalidOperationException($"No DOM anchor declared for {row.Id}"))
        .Class("row-marker").ValueText(row.Label);

    private static HtmlTag Section(MarkerSpec marker) => new HtmlTag("section").Id(marker.Id).Class("report-summary")
        .AddChild(new HtmlTag("h2").ValueText(marker.Label))
        .AddChild(new HtmlTag("p").ValueText(marker.Text));

    private static HtmlTag Paragraph(MarkerSpec marker, IReadOnlyDictionary<string, LinkSpec> links) {
        var paragraph = new HtmlTag("p").Id(marker.Id).Class("flow-following-content")
            .AddChild(new HtmlTag("strong").ValueText(marker.Label)).ValueText(" — " + marker.Text);
        if (marker.LinkId is { } linkId) {
            paragraph.ValueText(" ").AddChild(Link(links[linkId]));
        }
        return paragraph;
    }

    private static HtmlTag Link(LinkSpec link) => new Anchor(link.Href, link.Label).Id(link.Id);

    private static string RequiredLinkId(RowSpec row) => row.LinkId
        ?? throw new InvalidOperationException($"No row link declared for {row.Id}");
}
