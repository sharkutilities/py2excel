OPTION EXPLICIT

' A Set of Simple Functions to Mimic the `fiscalyear` Python-Library
' 
' The PyPI library [`fiscalyear`(]https://pypi.org/project/fiscalyear/)
' is a lightweight module providing utility functions for managing the
' fiscal calendar. This VBA script can be utilized in the same way to
' convert to and from from a "financial year" to "calendar year" and
' vice versa.
' 
' 
' Author: Debmalya Pramanik

FUNCTION fiscalYear(value AS DATE, OPTIONAL prefix AS STRING = "F.Y. ", OPTIONAL fmt AS STRING = "YYYY") AS STRING
    DIM START_MONTH AS INTEGER: START_MONTH = 4 ' ? defaults to the Indian subcontinent
    DIM RIGHT_SUBSTRING_LENGTH AS INTEGER: RIGHT_SUBSTRING_LENGTH = LEN(fmt)

    DIM rStartYear AS INTEGER
    DIM rFinalYear AS INTEGER

    IF YEAR(value) < START_MONTH THEN
        rStartYear = YEAR(value) - 1
        rFinalYear = YEAR(value)
    ELSE
        rStartYear = YEAR(value)
        rFinalYear = YEAR(value) + 1
    END IF

    ' truncate the resolved year address as per the format specifier
    rStartYear = RIGHT(rStartYear, RIGHT_SUBSTRING_LENGTH)
    rFinalYear = RIGHT(rFinalYear, RIGHT_SUBSTRING_LENGTH)

    fiscalYear = prefix & rStartYear & "-" & rFinalYear
END FUNCTION
