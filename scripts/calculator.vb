OPTION EXPLICIT

' A Set of Simple Functions to Quickly Calculate Values
'
' The script file is designed to provide utility functions for
' quick calculations that is not limited to fixed deposit, recurring
' deposit, calculators but much more. Most of the functions are
' derived using simple mathematical based formula as mentioned below.
'
'
' Author: Debmalya Pramanik

FUNCTION fixedDeposit(principal AS VARIANT, rate AS VARIANT, period AS VARIANT, OPTIONAL type_ AS STRING = "C", OPTIONAL compunding AS STRING = "Q") AS VARIANT
    DIM maturityValue AS VARIANT ' final maturity amount

    DIM n AS INTEGER ' compounding factor, monthly = 12, quarterly = 4, etc.

    IF compunding = "Y" THEN
        n = 1
    ELSEIF compunding = "Q" THEN
        n = 4
    ELSEIF compunding = "M" THEN
        n = 12
    ELSE
        n = 0 ' invalid input
    END IF

    ' consider rate percentage as number, example rate = 6.5% = 0.065
    IF rate > 1 THEN
        rate = rate / 100
    ELSE
        rate = rate
    END IF

    IF type_ = "S" THEN
        maturityValue = principal + (principal * rate * period)
    ELSEIF type_ = "C" THEN
        maturityValue = principal * (1 + rate / n) ^ (n * period)
    ELSE
        maturityValue = 0 ' some/one parameter is invalid
    END IF

    fixedDeposit = maturityValue
END FUNCTION
