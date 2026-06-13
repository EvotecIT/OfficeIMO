# OfficeIMO.MarkdownRenderer.Wpf - WPF MarkdownView host

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.MarkdownRenderer.Wpf)](https://www.nuget.org/packages/OfficeIMO.MarkdownRenderer.Wpf)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.MarkdownRenderer.Wpf?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.MarkdownRenderer.Wpf)

`OfficeIMO.MarkdownRenderer.Wpf` provides a reusable WPF `MarkdownView` control built on WebView2 and `OfficeIMO.MarkdownRenderer`.

## Install

```powershell
dotnet add package OfficeIMO.MarkdownRenderer.Wpf
```

## Requirements

- Windows desktop application using WPF.
- WebView2 runtime available on the target machine.
- `OfficeIMO.MarkdownRenderer.Wpf` and its `OfficeIMO.MarkdownRenderer` dependency.

If target machines may not already have WebView2 installed, ship or bootstrap the Evergreen WebView2 Runtime as part of your installer strategy.

## Quick start

```xml
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    xmlns:omd="clr-namespace:OfficeIMO.MarkdownRenderer.Wpf;assembly=OfficeIMO.MarkdownRenderer.Wpf">
    <Grid>
        <omd:MarkdownView x:Name="Preview"
                          Preset="Relaxed"
                          DocumentTitle="README" />
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
```

## Examples

### Handle external navigation

```csharp
using System.Diagnostics;

Preview.NavigationRequested += (_, args) =>
{
    Process.Start(new ProcessStartInfo(args.Uri.AbsoluteUri) {
        UseShellExecute = true
    });
    args.Handled = true;
};
```

### Use pre-rendered body HTML

```csharp
using OfficeIMO.MarkdownRenderer;

var options = MarkdownRendererPresets.CreateStrict();
Preview.BodyHtml = MarkdownRenderer.RenderBodyHtml(markdownText, options);
Preview.Markdown = string.Empty;
```

## What it does

- Hosts the OfficeIMO Markdown shell in a WPF/WebView2 control.
- Lets host apps choose presets, CSS overrides, renderer options, and link handling.
- Supports clipboard copy messages from shell actions.
- Allows advanced hosts to set pre-rendered `BodyHtml` instead of `Markdown`.
- Implements `IDisposable` so long-lived hosts can release WebView2 resources explicitly.

## Boundaries

- This package owns the WPF control surface.
- Generic rendering behavior belongs in `OfficeIMO.MarkdownRenderer`.
- Application-specific navigation, telemetry, fallback UI, and installer/runtime policy stay in the host app.

## Targets and license

- Targets: Windows WPF/WebView2 package targets.
- License: MIT.
- Repository: [EvotecIT/OfficeIMO](https://github.com/EvotecIT/OfficeIMO)
