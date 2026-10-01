# Pinned FA(3) schema evidence

The `Schemas` directory contains unchanged official CRD schema bytes used by the
offline validation tests. The production package does not embed or download them;
applications supply a local directory to `Fa3SchemaBundle.LoadDirectory`.

| Local file | Official source | SHA-256 |
| --- | --- | --- |
| FA3.xsd | https://crd.gov.pl/wzor/2025/06/25/13775/schemat.xsd | B646B6B525F51ADF1BB2545F111FC8CA6E7AA6DD2F98948F1667D3695C06D958 |
| StrukturyDanych_v10-0E.xsd | https://crd.gov.pl/xml/schematy/dziedzinowe/mf/2022/01/05/eD/DefinicjeTypy/StrukturyDanych_v10-0E.xsd | 1137CE6E3C11C2B9EF3F05E4E72D6DCD6B4FA94908EA558F2BA15DE0259BB2AA |
| ElementarneTypyDanych_v10-0E.xsd | https://crd.gov.pl/xml/schematy/dziedzinowe/mf/2022/01/05/eD/DefinicjeTypy/ElementarneTypyDanych_v10-0E.xsd | 8A531CB181D3E298D11B28766655AE91FEE2D7851440095932FFC82137ED2BE1 |
| KodyKrajow_v10-0E.xsd | https://crd.gov.pl/xml/schematy/dziedzinowe/mf/2022/01/05/eD/DefinicjeTypy/KodyKrajow_v10-0E.xsd | 1D41A1B3184188F2D20A51D3AFDE26204DDA182EC5DACF018204DCC9870DC644 |

The 26 independent document fixtures are linked from
`OfficeIMO.Invoicing.Tests/Fixtures/FA3`. Original schema imports remain intact and
resolve to verified local snapshots. Tests reject modified imports, malformed
inputs and attempts to replace the pinned schema through document attributes.
