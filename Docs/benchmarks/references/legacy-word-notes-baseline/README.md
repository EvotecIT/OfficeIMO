# Reconstruct the first-correct note reader

This packet reconstructs the frozen Word reader source used by the matched note comparison. Start with OfficeIMO commit `f9d02fda5e27d16051d3e16ba27143da17647ecd`, take exactly the 12 paths in `source-manifest.json` from initial note commit `c4072bd0d1a5801da0acfa1a547d1e0025104c69`, then apply `first-correct.patch`. Ten files retain the initial commit's normalized Git blobs. The patch changes only `LegacyDocFootnoteReader.cs` and `LegacyWord69Notes.cs`; c407 alone has different content in those two files.

The manifest records normalized Git blob IDs and frozen byte SHA-256 values separately. Eleven frozen files use CRLF; `LegacyWord69Notes.cs` uses LF. All are UTF-8 without a BOM. The patch uses LF.

Use a disposable clone containing both commits. Keep this packet outside that clone. From the clone root, check out the manifest's `baseCommit`, then restore exactly the manifest's 12 paths from `initialNoteCommit` into the working tree. Do not replace the whole checkout with the initial note commit. These are source-reconstruction instructions, not commands executed by this packet.

Before applying the patch, normalize just those 12 files to LF with UTF-8 without a BOM. Check and apply `first-correct.patch` with ordinary `git apply --check` and `git apply`, respectively. After applying it, restore each file's recorded line endings and verify both content and bytes. This PowerShell 7 fragment performs the last verification; set `$packet` to this packet's folder and run from the clone root:

```powershell
$manifest = Get-Content (Join-Path $packet 'source-manifest.json') -Raw | ConvertFrom-Json
$utf8 = [Text.UTF8Encoding]::new($false)
foreach ($file in $manifest.files) {
    $path = Join-Path (Get-Location).Path $file.path
    $text = [IO.File]::ReadAllText($path).Replace("`r`n", "`n")
    $body = $utf8.GetBytes($text)
    $header = [Text.Encoding]::ASCII.GetBytes("blob $($body.Length)`0")
    $blobBytes = [byte[]]::new($header.Length + $body.Length)
    [Buffer]::BlockCopy($header, 0, $blobBytes, 0, $header.Length)
    [Buffer]::BlockCopy($body, 0, $blobBytes, $header.Length, $body.Length)
    $blob = [Convert]::ToHexString([Security.Cryptography.SHA1]::HashData($blobBytes)).ToLowerInvariant()
    if ($blob -cne $file.frozenNormalizedGitBlobSha1) { throw "Git blob differs: $($file.path)" }
    if ($file.lineEndings -eq 'CRLF') { $text = $text.Replace("`n", "`r`n") }
    [IO.File]::WriteAllText($path, $text, $utf8)
    if ((Get-FileHash $path -Algorithm SHA256).Hash -cne $file.frozenSha256) {
        throw "Frozen bytes differ: $($file.path)"
    }
}
```

Reconstruction was checked by independently applying the patch hunks in memory to the initial Git blobs, then checking all 12 normalized blob IDs and exact frozen byte hashes. No duplicated source files or binaries are included. `SHA256SUMS` binds this packet's four payload files.

Core, common runtime libraries and the benchmark workload are bound separately to final source commit `3bcbdb49d74fb31b46aba197025acec98f352d19`. The manifest keeps those identities separate from historical Word source. Seed that final Core DLL/PDB at the evaluated Core target under the reference artifacts directory, then restore and build only Word with the recorded SDK, arguments and environment. `BuildProjectReferences=false` prevents building historical Core. Verify the seeded Core hash before and after, and use identical common runtime/workload bytes in every compared reader root. The selected recorded binary hashes identify the measured assembly; this packet does not contain those binaries.

Recompilation is not promised to produce a byte-identical DLL: compiler, SDK, path and source metadata can affect it. Timing admission still requires complete matching workload validation. Synthetic note scale inputs do not establish independent-producer compatibility, and the original independent FIB104 endnote fixture remains a separate qualification gap.

The source changes and patch are covered by the included MIT license, Copyright (c) 2022 Evotec. Keep its complete notice with copies.
