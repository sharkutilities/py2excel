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

### Fiscal Year

The project is inspired by the [`fiscalyear``](https://pypi.org/project/fiscalyear/) library hosted in PyPI. The script provides ready-made
functions to users who want to convert dates to and from the calendar to the financial year and vice-versa. The following functions/methods
are available:

#### Function: `fiscalYear` | Release Date 21-05-2024

<div align = "center">

[![function-script](https://img.shields.io/badge/👨‍💻-Script_File-blue?style=plastic)](./scripts/fiscalYear.vb)
[![ms-excel-addins](https://img.shields.io/badge/🎉-MS_Excel_AddIns-blue?style=plastic)](./addins/FiscalYear.xlam)
[![function-example](https://img.shields.io/badge/📜-Function_Example-blue?style=plastic)](./scripts/README.md#function-fiscalyear--release-date-21-05-2024)

| Parameter Name | Accepted Type | Optional Parameter | Default Value | Parameter Definition |
| :---: | :---: | :---: | :---: | --- |
| **`value`** | `DATE` | | | Current Year |
| **`prefix`** | `STRING` | ✔ | "F.Y. " | Prefix to be added at the beginning of the resolved finanicial year. |
| **`fmt`** | `STRING` | ✔ | "YYYY" | Returns the year in YYYY or YY format depending upon user-preference. |
| **`quarter`** | `BOOLEAN` | ✔ | FALSE | Returns the quarter number for the financial year. |

</div>

### Calculator Functions

The project is intended to provide easy-to-use calculations like fixed deposit, recurring deposit, etc. calculator as an MS Excel plugin.
Simply integrate the same as add-ins as mentioned. The following functions are available:

#### Function: `fixedDeposit` | Release Date 29-06-2024

<div align = "center">

[![function-script](https://img.shields.io/badge/👨‍💻-Script_File-blue?style=plastic)](./scripts/calculators.vb)
[![ms-excel-addins](https://img.shields.io/badge/🎉-MS_Excel_AddIns-blue?style=plastic)](./addins/calculators.xlam)

| Parameter Name | Accepted Type | Optional Parameter | Default Value | Parameter Definition |
| :---: | :---: | :---: | :---: | --- |
| **`principal`** | `VARIANT` | | | Principal Amount, i.e. A/C Opening Balance |
| **`rate`** | `VARIANT` | | | The (expected) rate of interest. Rate of interest can be either between 0-1 representing actual value, or 0-100 i.e., percentage. |
| **`period`** | `VARIANT` | | | Time period, in years/ |
| **`type_`** | `STRING` | ✔ | "C" | {`C` : Compound Interest, `S` : Simple Interest} |
| **`type_`** | `STRING` | ✔ | "Q" | {`M` : Monthly Compunded, `Q` : Quarterly Compunded, `Y` : Yearly Compunded} |

</div>

</div>
