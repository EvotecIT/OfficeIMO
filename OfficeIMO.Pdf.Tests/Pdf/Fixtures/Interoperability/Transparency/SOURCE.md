# Isolated Form group fixtures

Cairo 1.18.4 produces these four PDFs independently of OfficeIMO. Each page is
100 by 80 points, with two overlapping red rectangles inside an isolated RGB
transparency group. The group opacity is one half. `child-alpha` gives each
rectangle one-half opacity, and `nested` adds another one-half-opacity group.
The fixtures exercise group compositing rather than per-child attenuation.

The generator also writes reference PNGs. Cairo is an optional fixture-production
tool; ordinary tests consume the checked-in PDFs and do not require it.

```sh
cc generate_cairo_groups.c $(pkg-config --cflags --libs cairo) -o generate-cairo
./generate-cairo /path/to/task-output
```

The generator sets the PDF creation date to a fixed value. Reproduction can still
differ with the Cairo version or compression implementation.

| File | SHA-256 |
| --- | --- |
| cairo-child-alpha-nested.pdf | `469c1ff1ee4656cbb01b1127b3d8e252fd1a52e7e4a0b0112d442a19bf887b00` |
| cairo-child-alpha.pdf | `9548662e6ae977cd5e035d44ad64a1ddfc7c89fba761a547f50f518f4a6063ed` |
| cairo-opaque-nested.pdf | `8743273613fa44e1a4483c2eeba2e1e86092039bbd890db7894800d6ae9ee392` |
| cairo-opaque.pdf | `840b042a7b1968557d6eb93ed482b8b4a0c01e92660b3a7af58b93c087fe657a` |
