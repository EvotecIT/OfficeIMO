# Inspecting the captured measurements

`raw-packets.zip` contains the large JSON captures under their original filenames. `raw-packets-manifest.json` records the archive hash and every member's length and SHA-256. The uncompressed captures retain their original bytes, including line endings. Summaries, small captures, fixtures and runners remain in this directory.

Extract the archive to a scratch directory before inspecting a packet. With PowerShell, use `Expand-Archive -LiteralPath ./raw-packets.zip -DestinationPath <scratch-directory>`. With an installed ZIP utility, use `unzip raw-packets.zip -d <scratch-directory>`. Replace the placeholder with a real task output directory.

Each existing evidence manifest still uses the original packet filenames. Resolve those names in the extracted directory. Follow the report and captured runner for runtime, producer, source-head and assembly requirements; unpacking measurements alone does not rebuild the original benchmark snapshots.
