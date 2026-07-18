# OneNote fixture provenance

The five `testOneNote*.one` files in this folder are copied unchanged from the Apache Tika test corpus at commit
`63e22d08ef249cc73a6d02da7bc199fc3623a607`. Apache Tika is licensed under the Apache License 2.0.
The accompanying `LICENSE-APACHE-2.0.txt` and `NOTICE.txt` retain the applicable redistribution terms
and Apache Tika attribution for these fixtures.

Upstream path:

`tika-parsers/tika-parsers-standard/tika-parsers-standard-modules/tika-parser-microsoft-module/src/test/resources/test-documents/`

These files are test inputs only. OfficeIMO.OneNote does not copy or depend on Apache Tika parser code.

`handwriting_recognition.one` is copied unchanged from the `onenote.rs` parser corpus at commit
`5138a39a3f4e72b840932f9872fecde52fa9da60`, path
`crates/parser/tests/samples/handwriting_recognition.one`. The upstream project distributes this fixture under
the Mozilla Public License 2.0; `LICENSE-MPL-2.0.txt` retains that license. This fixture is test input only.
OfficeIMO.OneNote does not copy or depend on `onenote.rs` parser code. Samples in the upstream `joplin/`
subdirectory are intentionally excluded because they carry separate AGPL-3.0-or-later terms.

The `makecab-lzx-*.cab` files were generated from the corresponding unchanged
`.one` fixtures with the Windows `makecab` utility using LZX compression and a
21-bit window. They are independent compressed-artifact oracles for aligned,
verbatim, and multi-CFDATA Cabinet LZX decoding; expanding them reproduces the
source fixtures byte-for-byte.

The `makecab-lzx*-e8.cab` files were generated from a deterministic 4,096-byte
OfficeIMO test pattern containing three x86 `E8` relative-call sequences, using
every Cabinet window size from 15 through 21 bits. They verify compressed
matches and optional LZX E8 postprocessing across the complete supported range
against `makecab`; the test reconstructs the expected bytes independently
rather than storing a second payload copy.

`makecab-lzx-notebook.onepkg` was generated from a small notebook authored by
`OneNoteNotebookWriter`, then packaged with `makecab` using LZX compression and
a 21-bit window. It exercises the complete `.onetoc2` plus `.one` package path
through `OneNotePackageReader`.

| File | Size | SHA-256 |
|---|---:|---|
| `testOneNote.one` | 30,288 | `B614DC94B890B53DB7CB2D3053382CB398C59385533C256E2509850CC3247270` |
| `testOneNote2016.one` | 14,744 | `FCFC3C2E65482DC6F70F6A613B058E908F67DB2EBB16A343BC2367E02BBB471C` |
| `testOneNoteEmbeddedWordDoc.one` | 33,096 | `CF38E39CB5CED46F377C832E5FF0FA5E789945930F77C294BA5E866429A2A028` |
| `testOneNoteFromOffice365.one` | 29,387 | `093F20ECB2196F8E6C07CFA6D7C7ACB65A50AD3126A95444FE33086A37AAA4D5` |
| `testOneNoteFromOffice365-2.one` | 69,986 | `8CD245ED549043534118A00CE29715147C880C38CA88C3481ACC19AE28E980C2` |
| `handwriting_recognition.one` | 180,020 | `2CFF8769CCF0AF6209D96D5E0650661077EDBA2D2BAE4E4AA691F06CAEA35456` |
| `makecab-lzx-testOneNote2016.cab` | 2,974 | `531F6465AB1AE92E5011D6CDB614130A91358E714F5EEF995AF50799EFAD5BEC` |
| `makecab-lzx-testOneNoteFromOffice365-2.cab` | 8,027 | `F79E91DE0D23296D602AD50FF15DC1947A95C12AED5A84B7FA45EDAC4ABA3E02` |
| `makecab-lzx-e8.cab` | 223 | `4510A5D1579FA302372A8671BF480190CD90ED669E3B87C9A55801467E932F48` |
| `makecab-lzx15-e8.cab` | 221 | `F8C625153CC4404B66B9440EE3744514C7DFBC1913210E4CC4E2FD912C11958E` |
| `makecab-lzx16-e8.cab` | 221 | `D0A30AB570809CC27CD6E0AE93866FF15BC0DB94474A1320DC4B16916FA7AE1E` |
| `makecab-lzx17-e8.cab` | 221 | `E1943D518932CF43F3C79BB5E79F911BB337EF8CBA29782E703430D3A1587801` |
| `makecab-lzx18-e8.cab` | 221 | `E96D428F4CC2229A05A0EAFED4A80EA498DBA16EEF7632E0E0E1F99BE83BD6CC` |
| `makecab-lzx19-e8.cab` | 221 | `5E18D1107F8E5CA44098D64A4EF4B379E1B85B7F5A7489430E9C1C597B58CD49` |
| `makecab-lzx20-e8.cab` | 223 | `ECF8F47CA8DAAFE7E0CDA284F8262235A79FB740FBB9AAD5029CDB8FB9398637` |
| `makecab-lzx-notebook.onepkg` | 1,943 | `69DBA4B5157447F8EFB0DDA07F7315ED2F1534189C9E45CAECD0258AD71EF27A` |

Source: https://github.com/apache/tika/tree/63e22d08ef249cc73a6d02da7bc199fc3623a607/tika-parsers/tika-parsers-standard/tika-parsers-standard-modules/tika-parser-microsoft-module/src/test/resources/test-documents

Ink fixture source: https://github.com/msiemens/onenote.rs/blob/5138a39a3f4e72b840932f9872fecde52fa9da60/crates/parser/tests/samples/handwriting_recognition.one
