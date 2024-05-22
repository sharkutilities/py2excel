<h1 align = "center">VB Scripts</h1>

<div align = "justify">

The scripts presents the raw macro/vba functions which can be added to a macro enabled workbook (`*.xlsm`) files and can be executed when
the "Enable Content" feature is accepted. One may add the function the "Personal Workspace" scope such that the function is available
system-wide. However, add-ins is recommended and is available under the [add-ins](../addins) directory.

## Modules / Add-Ins

Module wise documentation an usage notes (accepted parameter) as available in [README](../README.md#modules--add-ins), the below document
provides usages and examples on function by function basis.

### Fiscal Year

The project is inspired from the [`fiscalyear`](https://pypi.org/project/fiscalyear/) library hosted in PyPI. The script provides ready-made
functions to users who wants to convert dates to- and from- calendar to financial year and vice-versa. The following functions/methods
are available:

#### Function: `fiscalYear` | Release Date 21-05-2024

<div align = "center">

| Given Date | Function Input | Function Output |
| :---: | :---: | :---: |
| 01-01-2024 | `=fiscalYear("01-01-2024")` | F.Y. 2023-2024 |
| 31-01-2024 | `=fiscalYear("31-01-2024","FY ")` | FY 2023-2024 |
| 01-02-2024 | `=fiscalYear("01-02-2024",,"YY")` | F.Y. 23-24 |
| 05-02-2024 | `=fiscalYear("05-02-2024","FY ", "YY")` | FY 23-24 |
| 09-08-2024 | `=fiscalYear("09-08-2024",,,TRUE)` | F.Y. 2024-2025 Q2 |
| 22-05-2024 | `=fiscalYear("22-05-2024",,"YY",TRUE)` | F.Y. 24-25 Q1 |

</div>

</div>
