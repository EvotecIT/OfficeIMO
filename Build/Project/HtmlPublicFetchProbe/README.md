# Optional public HTML fetch probe

This opt-in tool checks the host-side byte-acquisition boundary for one public
HTML URL. It saves the exact response bytes and a manifest with fetch time,
final URL, redirect hops, connected IP address, media type and SHA-256 digest.
It does not parse the untrusted HTML on the host. The destination directory
must be new so an existing corpus cannot be overwritten.

```bash
dotnet run --project Build/Project/HtmlPublicFetchProbe/OfficeIMO.Html.PublicFetchProbe.csproj -c Release -- \
  https://wpt.live/css/css-pseudo/first-letter-001-ref.html /tmp/officeimo-public-fetch-wpt
```

Additional allowed DNS hosts may follow the output path when a known redirect
requires one. The broker rejects loopback and private IPv4,
nonstandard ports, HTTP downgrades, unlisted hosts, compressed responses,
non-UTF-8 HTML and responses beyond its request and byte budgets. It connects
to the validated IPv4 address directly while preserving the URL's TLS host.
This probe only acquires bytes. It does not discover or prefetch page assets,
execute the page, or qualify the separate untrusted runtime profile.
