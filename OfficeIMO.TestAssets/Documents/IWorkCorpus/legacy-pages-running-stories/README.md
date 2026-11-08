# Pages 09 default running stories

`header+footer.pages` and `header+footer.pages.fodt` are unchanged files from
[the libetonyek regression corpus](https://git.libreoffice.org/libetonyek-test/+/a548a9569c18cdffc24d3ba5d549f6b423c449d3/testset/Pages/4/),
at revision `a548a9569c18cdffc24d3ba5d549f6b423c449d3`.
The corpus credits David Tardon and Anurag Kanungo as test-file contributors in
[its CREDITS file](https://git.libreoffice.org/libetonyek-test/+/a548a9569c18cdffc24d3ba5d549f6b423c449d3/CREDITS).

Both files remain licensed under
[Creative Commons Attribution-ShareAlike 4.0 International](https://creativecommons.org/licenses/by-sa/4.0/),
as stated in the pinned [upstream README](https://git.libreoffice.org/libetonyek-test/+/a548a9569c18cdffc24d3ba5d549f6b423c449d3/README.md).
No changes were made to either distributed fixture. Test-created derivatives
change only selected XML controls and remain subject to the same fixture license;
they are synthetic parser controls rather than independent producer evidence.

| File | Bytes | SHA-256 |
|---|---:|---|
| `header+footer.pages` | 39692 | `65fc2f9997671f3aab6cdd11ac1eb180fb4854efcf4c4a7fcb37939f628d6b8f` |
| `header+footer.pages.fodt` | 8275 | `3b945ba2dc0f79b9da20df2f11fa9048b656df3112e7e5be4ae07c68e76068d1` |

The native package declares XML grammar `92008102400`. Its selected single body
section uses the default page template, with `Header.` and `Footer.` in the named
odd/default storages. First-page and even-page alternatives are stored blank and
not selected. The FODT provides an independent semantic reference for both
stories. The package's embedded `QuickLook/Preview.pdf` provides the native visual
reference; it does not establish the appearance of converted DOCX output.

The focused tests check typed story recovery and saved/reopened editable DOCX,
including the selected 10-point black text, while keeping body content separate.
They also exercise inactive templates, unqualified section selection, unresolved
default storages and the shared text budget. Other section layouts, first/even
selection, fields, wrapping, tab stops and cross-application appearance remain
outside this fixture's qualification.
