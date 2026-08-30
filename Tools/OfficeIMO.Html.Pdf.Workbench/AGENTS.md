# OfficeIMO.Html.Pdf.Workbench

| Setting | Value |
|---------|-------|
| **Interactivity Mode** | Server |
| **Interactivity Scope** | Per-page |

## Rendering configuration

This project uses per-page Interactive Server with prerendering. It was created with `dotnet new blazor -int Server`.

Pages are static SSR by default. Only components that explicitly add `@rendermode InteractiveServer` become interactive.

## Adding new components

- Create routable pages in `Components/Pages/` and shared components in `Components/`.
- Keep read-only pages static. Add `@rendermode InteractiveServer` only where live editing or interaction requires it.
- Static pages may use standard HTML forms with `[SupplyParameterFromForm]`.

## Data access

- Components may inject server-side services directly. Do not add an HTTP API layer for in-process OfficeIMO or HtmlTinkerX calls.
- Keep conversion policy and rendering behavior in the owning libraries; this project is a thin local operator surface.

## Environment constraints

- Interactive components run through a server-side SignalR circuit.
- Do not inject `HttpContext` into interactive components.
- Browser APIs require explicit JS interop. Prefer ordinary HTML behavior and loopback artifact endpoints when sufficient.
- The host must remain loopback-only and Chromium capture must default to `HtmlBrowserNetworkPolicy.Offline`.

## Don'ts

- Do not make all routes globally interactive.
- Do not add `@rendermode` to `Routes` in `App.razor`.
- Do not turn this local tool into the public WebAssembly converter; the public browser-local surface has a separate static deployment contract.
