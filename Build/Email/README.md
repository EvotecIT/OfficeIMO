# Email interoperability qualification

Run selected lanes from the repository root with PowerShell:

```powershell
./Build/Email/Test-EmailInteroperability.ps1 -Lane Managed -Framework net8.0

# Windows host using libpff already installed in a WSL distribution:
$env:OFFICEIMO_EMAIL_STORE_LIBPFF_WSL = 'Ubuntu'
./Build/Email/Test-EmailInteroperability.ps1 -Lane Managed,LibPff -OutputPath ./artifacts/email-qualification
```

`OutputPath` must be new or empty. The default creates a unique temporary directory.
`NoBuild` and `NoRestore` are available for an already built test assembly. Supported
frameworks are `net8.0`, `net10.0`, and Windows `net472`.

| Lane | Evidence | Existing opt-in prerequisite |
| --- | --- | --- |
| Managed | Ordinary email artifact, store, address-book, Reader and HTML regression tests, including generated independent-producer fixtures | None |
| MsgReader | MSGReader sample corpus with per-artifact semantic comparison | `OFFICEIMO_EMAIL_CORPUS_ROOT` or `EVOTEC_GITHUB_ROOT` containing `MSGReader` |
| MimeKit | MIME, TNEF and mbox sample corpus, adversarial scenarios and duplicate-header order | The same corpus root containing `MimeKit` |
| LibPff | Generated Unicode PST inspection and semantic export through libpff | `OFFICEIMO_EMAIL_STORE_PFFINFO` and the corresponding export tool, or Windows `OFFICEIMO_EMAIL_STORE_LIBPFF_WSL` |
| Outlook | Artifact exchange, PST scan and classic Outlook mount/read/remove | Installed classic Outlook on Windows; both `OFFICEIMO_EMAIL_OUTLOOK_INTEROP=1` and `OFFICEIMO_EMAIL_STORE_OUTLOOK_INTEROP=1` |
| Smime | Caller-provided real Outlook signed/encrypted artifacts | `OFFICEIMO_EMAIL_SMIME_CORPUS`; follow the existing corpus test's certificate/expected-result contract |
| PrivateStores | Bounded read of caller-owned PST/OST stores | `OFFICEIMO_EMAIL_STORE_CORPUS` |
| PrivateStoreConversion | Bounded conversion, strict semantic verification and source preservation | The store corpus plus `OFFICEIMO_EMAIL_STORE_CORPUS_CONVERT=1` |

The runner does not download corpora or enable Outlook automation. Select external
lanes only after their prerequisites are configured. Outlook tests may launch the
installed application and use its configured profile. Corpus provenance and licenses
are described in [the producer catalog](../../OfficeIMO.Email.Tests/Corpora/producer-corpora.json).

Each requested lane must produce a TRX result with at least one test, no skips or
failures, and every named test executed. A missing prerequisite, zero-test selection,
or partially selected external lane fails qualification even when `dotnet test`
returns success. Unrequested lanes appear as `disabled` in `qualification.json`.
The report records source commit, modified-source status, SDK, framework, filters,
counts and timing. TRX files live in each lane's subdirectory.

A successful managed lane does not qualify disabled external producers. The report
describes the selected operation and host; it does not claim coverage of every mail
producer, installed Outlook version or platform. Private corpus TRX files can contain
local paths and test diagnostics. Keep that evidence local and review it before sharing.
