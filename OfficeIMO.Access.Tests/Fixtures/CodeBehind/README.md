# Native VBA application fixtures

These files contain synthetic project-owned data. Microsoft Access created the application carrier and the form/report class modules; no VBA or event procedure was executed.

| Fixture | Producer and contract |
| --- | --- |
| `empty-application-jet4.mdb` | Access `NewCurrentDatabase` format 9, before adding its first VBA project |
| `empty-application-ace12.accdb` | Access `NewCurrentDatabase` format 12, before adding its first VBA project |
| `code-behind-jet4.mdb` | Existing Jet 4 designer corpus with native `Form_BoundForm1` and `Report_BoundReport` classes |
| `code-behind-ace12.accdb` | Existing rich ACE designer corpus with the same native classes and expanded designer records; physical profile is ACE 14 despite the historical filename |

The form and report have inert Open handlers. `GroupChoice1` has an inert AfterUpdate handler. Existing designer data and the embedded macro remain in the files. PROJECT declares the bound modules with `DocClass`; their source carries the native `VB_Base` identity.

`New-NativeCodeBehind.ps1 -FixtureRoot <fixtures> -OutputRoot <new-output-folder>` regenerates code-behind from the `Designer` corpus. `New-NativeEmptyApplications.ps1 -OutputRoot <new-output-folder>` creates empty applications and separate copies containing an independently imported first module. Both scripts require Windows Microsoft Access only for fixture generation, force macros disabled, and close their owned application. They are outside the library runtime and ordinary test execution.
