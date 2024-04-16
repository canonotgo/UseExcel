This is a library that stores VBA code.

SUBSTITUTE:
    function is used to replace occurrences of a specified substring within a text string. The syntax for the SUBSTITUTE function is:
    =SUBSTITUTE(text, old_text, new_text, [instance_num])

LEFT:
    function is used to extract a specified number of characters from the beginning (left side) of a text string. Here's a brief overview of how to use the LEFT function:
    LEFT(text, [num_chars])

Example: Get prefix of mac address
    =LEFT(SUBSTITUTE(B2,":",""),6)

Example: Whether the value is the same
    =IF(D2=F2,"Same","Different")

Example:
    =VLOOKUP(A2,'C:\Users\cc\Desktop\Excel\[all_data.xlsx]Sheet1'!$A$1:$D$100000,2,FALSE)

Example:
    =COUNTA('Sheet0'!B2:B1000)
    =COUNTA(INDIRECT("'"&B2&"'!A:A"))
    =COUNTIF('Sheet0'!G:G,B3)