# iWork reader corpus

These package fixtures prove the bounded Pages, Numbers, and Keynote reader against files produced by multiple iWork generations. They remain read-only test inputs; OfficeIMO does not rewrite them.

| Folder | Upstream | Revision | Producer evidence | License |
|---|---|---|---|---|
| `nim-iwork` | [nim-iwork](https://github.com/halcyon-oss/nim-iwork) | `60a8e875692ac934956cdb5b39f88e557c97cca4` | Pages, Numbers, and Keynote 14.5 fixtures | MIT, copyright Alfred |
| `iwork-converter` | [iwork-converter](https://github.com/obriensp/iwork-converter) | `be4828260466c9c022ec12f9dc9dbfeb15ab1dea` | Pages 14.1, Numbers 11.1, and Keynote 8.1 fixtures | MIT, copyright Steve Dunham |
| `numbers-parser` | [numbers-parser](https://github.com/masaccio/numbers-parser) | `1c6c5c3d2e29a9abb601596678089f0a6c85d64c` | Numbers 15.1 fixture plus independently produced formula and merged-range workbooks | MIT, copyright Jon Connell |
| `picodocs` | [PicoDocs](https://github.com/PicoMLX/PicoDocs) | `5c18743d3d8120a76da124bd512a3cf5bcc28e82` | Pages 14.4.1 package produced by Pages 14.5 with sections, headers, footers, an image, hyperlinks, and three editable tables | MIT, copyright Pico MLX |
| `keynotekit` | [KeynoteKit](https://github.com/memfrag/KeynoteKit) | `5e01f1e061b608e16c4480444d8e04790a625b34` | Keynote 15.2.1 image and editable-table fixtures independently maintained as parser/writer regressions | 0BSD, copyright Martin Johannesson |

The complete upstream license notices are reproduced in `OfficeIMO.IWork/THIRD-PARTY-NOTICES.md`. Fixture provenance and expected semantic assertions live beside the executable corpus tests in `OfficeIMO.IWork.Tests`.

## Fixture checksums

| Fixture | SHA-256 |
|---|---|
| `iwork-converter/a.key` | `929347827a7478c123dd3e3828e9751b5cf2ae977d2edd2a5f0774014fc4fefc` |
| `iwork-converter/a.numbers` | `43afd5a01cff9283fea5a11716d96f696b5853b75f82a3fe4469f240dfb82947` |
| `iwork-converter/a.pages` | `8481e3071c8ea1cc9543354bcd1ff66f79e6c2c8686096fa3ae60ab96fd49ebf` |
| `keynotekit/imagedeck-v15.2.1.key` | `a9af589197588e04ee52388b0aa6c2dad1110e5d6db814b58afe543831cf2128` |
| `keynotekit/tabledeck-v15.2.1.key` | `384962b1fff18abc5a901b59dc5f8820c2a959977f18f90dc9cd10095bdd0a56` |
| `nim-iwork/simple.key` | `ba95755df82ceb0ca834e1e03e2777c34fad906320d8336b4f3fefc6b48607eb` |
| `nim-iwork/simple.numbers` | `d0b00d9cae5985cccaa3b2fb251fae92eb0e38360fb4b5df8b4350eb658f752b` |
| `nim-iwork/simple.pages` | `5aee6d03277d2db2104f593e64afe081dec539f0117b97124b6f99158124c93e` |
| `numbers-parser/issue-102-v15.1.numbers` | `88a9fa7be095d03004478393a87a4a97602d7468f839d067ec9118c524c55176` |
| `numbers-parser/test-10-formulas.numbers` | `dd85bad68898ce5b065f277c0b9be1f3c32d696e3baa6b09d3614bbd35a5249f` |
| `numbers-parser/test-9-merges.numbers` | `d640c0012d629834161827cb2f564d0966d24e159586f82426a69c12a8f334cf` |
| `picodocs/sample-v14.4.pages` | `4714477138d0a4090fc2ee2ba2ebb6adcd0fb6ce20a28897a6247a8e17d1ddce` |
