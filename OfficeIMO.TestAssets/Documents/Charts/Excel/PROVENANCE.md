# Excel-produced exploded radial charts

`exploded-slices.xlsx` was created with Microsoft Excel desktop
16.0.20326.20158 on Windows. Its worksheet contains Complete = 7 and
Pending = 3, with one pie chart and one doughnut chart. Excel's point API
sets the Complete slice explosion to 25 percent and Pending to zero.
Both native chart parts contain a point-level `c:explosion val="25"` record.
The PNGs were exported by Excel from those charts and show the selected
slice moved away from the center.

| File | Role | SHA-256 |
| --- | --- | --- |
| `exploded-slices.xlsx` | Editable independent-producer package | `1EAE6F34F4E7CE2C18E2B4F53EB75EF2680EC1159B1BE38EC8D5B1164DAEE0D3` |
| `exploded-pie-reference.png` | Excel chart export | `E97A622B5FD565F16DDCA78E3E99E3123EF686303791882AB672D74A250987F6` |
| `exploded-doughnut-reference.png` | Excel chart export | `93490C60CAA857318C1396E9D5833C781EC0DDBBDCA9282380083C2CC39E9E2A` |

On a Windows host with Excel installed, run
`powershell.exe -NoProfile -STA -File .\generate-exploded-slices.ps1 -OutputPath .\exploded-slices.xlsx`
from this directory to recreate the package and chart exports. Excel may
regenerate package metadata and images differently across builds; inspect
the native chart records and visible references before changing the fixture
or its hashes. The references qualify slice direction and separation, but
not OfficeIMO's surrounding worksheet layout or Excel typography.
