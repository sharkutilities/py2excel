<h1 align = "center">Python Native Function(s) in Excel</h1>

<div align = "center">

[![GitHub Issues](https://img.shields.io/github/issues/sharkutilities/py2excel?style=plastic)](https://github.com/sharkutilities/py2excel/issues)
[![GitHub Forks](https://img.shields.io/github/forks/sharkutilities/py2excel?style=plastic)](https://github.com/sharkutilities/py2excel/network)
[![GitHub Stars](https://img.shields.io/github/stars/sharkutilities/py2excel?style=plastic)](https://github.com/sharkutilities/py2excel/stargazers)
[![LICENSE File](https://img.shields.io/github/license/sharkutilities/py2excel?style=plastic)](https://github.com/sharkutilities/py2excel/blob/master/LICENSE)

</div>

<div align = "justify">

[**Python Native Functions (`py2excel`)**](https://github.com/sharkutilities/py2excel) is a set of native MS Excel functions derived/inspired
from [PyPI](https://pypi.org/) modules. The functions are written in pure macros/VBA to be used inside a macro-enabled workbook or can be used as
add-ins by saving the files and importing them from the File > Options > Add-Ins tab.

## Modules / Add-Ins

The modules/add-ins are available in two formats - [`scripts`](./scripts/) and [`add-ins`](./addins/). The scripts contain scripts and
functions in `*.vb` file format which can be directly added to a macro-enabled workbook/worksheet as per preference. However, it is
recommended to either add the contents of the `*.vb` in the "Personal Workspace" or directly import the codes using add-ins from the
created files.

### Fiscal Year ([`fiscalyear`](https://pypi.org/project/fiscalyear/))

<div align = "center">

| Parameter Name | Accepted Type | Optional Parameter | Default Value | Parameter Definition |
| :---: | :---: | :---: | :---: | --- |
| **`value`** | `DATE` | | | Current Year |
| **`prefix`** | `STRING` | ✔ | "F.Y. " | Prefix to be added at the beginning of the resolved finanicial year. |
| **`fmt`** | `STRING` | ✔ | "YYYY" | Returns the year in YYYY or YY format depending upon user-preference. |

</div>

</div>
