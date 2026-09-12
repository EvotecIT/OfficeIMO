# Retained AngleSharp JavaScript bindings

This internal build retains AngleSharp.Js 1.0.1 DOM bindings while allowing the
OfficeIMO runtime to configure its interpreter host and module loader. It builds
as `AngleSharp.Js.OfficeIMO.dll` and is deployed with the runtime worker. It is
not a separate NuGet package or an OfficeIMO-owned JavaScript interpreter.

`Upstream/` contains the 44 C# source files from
[AngleSharp.Js commit 7161da4](https://github.com/AngleSharp/AngleSharp.Js/tree/7161da4cec7a57ad76d7426bbcf981e0c97900e1/src/AngleSharp.Js).
The original Git blob hashes and paths are recorded in `upstream.json`.
The MIT license is retained in `AngleSharp.Js.LICENSE.txt` and copied to worker
build and publish outputs.

The source changes are limited to an engine-configuration callback:

- `JsScriptingOptions.ConfigureEngine` accepts the window and Jint options; the
  option snapshot retains that callback.
- `EngineInstance` invokes it after configuring the existing DOM wrapper and
  stack-depth limit, before constructing the interpreter.

Module URL resolution, resource authority, import maps and asynchronous document
execution belong to `OfficeIMO.Html.Runtime.Worker`. They are not embedded in
these retained bindings. If a compatible upstream package exposes the required
configuration callback, this project can be replaced by that dependency and the
runtime integration adapted at its configuration boundary.

When refreshing these sources, compare against the recorded upstream commit,
preserve attribution, review the callback patch, and run the runtime integration
and deployed-worker tests. A package-version label alone does not identify this
modified assembly.
