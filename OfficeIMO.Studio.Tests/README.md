# Studio acceptance tests

Run the regular suite with:

```powershell
dotnet test OfficeIMO.Studio.Tests/OfficeIMO.Studio.Tests.csproj
```

The process-recovery tests launch this test assembly in child processes with temporary profiles. They terminate a writer after its edits and session record have been stored, then verify recovery or privacy settings in another process. This covers abrupt process termination, not power-loss durability or installed application startup.

To save rendered screenshots, set `OFFICEIMO_STUDIO_VISUAL_OUTPUT` to a task-owned output directory before running the suite.

## Real storage failures on Linux or WSL

After building the test project, run:

```bash
bash OfficeIMO.Studio.Tests/Run-StorageAcceptance.sh
```

An optional first argument selects a different built `OfficeIMO.Studio.Tests.dll`. The runner requires .NET 10, `unshare`, `mount`, `umount`, and permission to create user and mount namespaces.

The runner creates an isolated mount namespace and an 8 MiB temporary filesystem. It exhausts that filesystem to check failed edit, save, and session writes; verifies that source bytes, the existing output, and the previous recovery snapshot survive; then frees space and retries. A second case detaches the source path and recovers its saved edits to another filesystem, using a separate backing mount to verify the original bytes.

The mounts and temporary files are removed when the runner exits, including on failure. It does not fill a normal disk or detach a system mount. These checks establish Linux filesystem behavior; physical removable devices, network outages, and native desktop dialogs need separate platform acceptance.
