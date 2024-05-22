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

FUNCTION fiscalYear(value AS DATE, OPTIONAL prefix AS STRING = "F.Y. ", OPTIONAL fmt AS STRING = "YYYY", OPTIONAL quarter AS BOOLEAN = FALSE) AS STRING
    DIM START_MONTH AS INTEGER: START_MONTH = 4 ' ? defaults to the Indian subcontinent
    DIM RIGHT_SUBSTRING_LENGTH AS INTEGER: RIGHT_SUBSTRING_LENGTH = LEN(fmt)

    DIM retval AS STRING ' final resolved value to return

    DIM rStartYear AS INTEGER
    DIM rFinalYear AS INTEGER
    DIM rQuarterNum AS INTEGER

    ' ? convert the year into fiscal/financial year format
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

    ' ? convert the month into quarter information
    ' TODO fix this considering starting month
    IF MONTH(value) <= 3 THEN
        rQuarterNum = 4
    ELSEIF MONTH(value) <= 6 THEN
        rQuarterNum = 1
    ELSEIF MONTH(value) <= 9 THEN
        rQuarterNum = 2
    ELSE
        rQuarterNum = 3
    END IF

    ' resolved string, considering quarter parameter to true/false
    IF quarter THEN
        retval = prefix & rStartYear & "-" & rFinalYear & " Q" & rQuarterNum
    ELSE
        retval = prefix & rStartYear & "-" & rFinalYear
    END IF

    fiscalYear = retval
END FUNCTION
