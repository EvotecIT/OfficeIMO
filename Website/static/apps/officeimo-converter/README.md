# OfficeIMO Converter Static App Mount

This folder is reserved for the published Blazor WebAssembly conversion app.

The app should be published here by the website build or a dedicated publish step:

```powershell
dotnet publish <BlazorWasmProject>.csproj -c Release -o Website\static\apps\officeimo-converter
```

The app must remain static-host compatible. It should use OfficeIMO byte and stream APIs in the browser runtime and must not require a server process, Office, LibreOffice, Redis, queues, or private storage.
