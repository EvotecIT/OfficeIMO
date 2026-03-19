# OfficeIMO.MarkdownRenderer.Wpf

`OfficeIMO.MarkdownRenderer.Wpf` provides a reusable WPF `MarkdownView` control built on WebView2 and `OfficeIMO.MarkdownRenderer`.

## What It Solves

- renders Markdown through the existing OfficeIMO HTML shell/update pipeline
- keeps host apps in control of presets, CSS overrides, and link handling
- supports clipboard copy messages from the shell for code/table actions
- stays safe for the main OfficeIMO cross-platform CI by only compiling the WPF surface on Windows targets
- can optionally host pre-rendered HTML bodies for advanced app-specific layouts

## Requirements

- Windows desktop application using WPF
- WebView2 runtime available on the target machine
- `OfficeIMO.MarkdownRenderer.Wpf` plus its `OfficeIMO.MarkdownRenderer` dependency

If you deploy to machines that may not already have WebView2 installed, ship or bootstrap the Evergreen WebView2 Runtime as part of your installer strategy.

## Quick Start

```xml
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    xmlns:omd="clr-namespace:OfficeIMO.MarkdownRenderer.Wpf;assembly=OfficeIMO.MarkdownRenderer.Wpf">
    <Grid>
        <omd:MarkdownView x:Name="Preview"
                          Preset="Relaxed"
                          DocumentTitle="README"
                          ShellCss=":root { color-scheme: dark; }" />
    </Grid>
</Window>
```

```csharp
Preview.Markdown = """
# Hello

This content is rendered by OfficeIMO.MarkdownRenderer.Wpf.
""";

Preview.ConfigureRendererOptions = options =>
{
    options.EnableCodeCopyButtons = true;
    options.EnableTableCopyButtons = true;
};

Preview.ErrorOccurred += (_, args) =>
{
    Debug.WriteLine($"Markdown preview error during {args.Context}: {args.Exception}");
};
```

## Advanced Hosts

For hosts that need to compose richer layouts, such as timelines or card-based discussion views, set `BodyHtml` directly instead of `Markdown`. This keeps WebView shell hosting inside `OfficeIMO.MarkdownRenderer.Wpf` while still allowing the app to build custom HTML from multiple rendered markdown fragments.

## Host Theming

The control is intentionally host-driven:

- use `Preset` to choose the baseline renderer profile
- use `ShellCss` to append host-specific CSS variables and overrides
- use `ConfigureRendererOptions` for advanced renderer customization
- handle `NavigationRequested` to intercept external links before the default shell launch behavior
- handle `ErrorOccurred` if the host wants telemetry, fallback UI, or structured diagnostics

Example:

```csharp
Preview.ShellCss = """
:root {
  color-scheme: dark;
  --omd-bg: #12141d;
  --omd-fg: #e5e9f0;
  --omd-link: #6cb6ff;
}

body {
  background: var(--omd-bg);
  color: var(--omd-fg);
}

a {
  color: var(--omd-link);
}
""";
```

## Notes

- `BaseHref` expects an absolute URI and is ignored when invalid
- `Markdown` is the normal content path; `BodyHtml` is meant for advanced hosts that render multiple markdown fragments into one composed surface
- the control implements `IDisposable` so long-lived hosts can release WebView2 resources explicitly when the view is no longer needed
