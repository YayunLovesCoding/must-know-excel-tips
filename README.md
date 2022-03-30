# A collection of excel tips/functions that I find useful at work.
vlookup -> https://support.microsoft.com/en-us/office/vlookup-function-0bbc8083-26fe-4963-8ab8-93a18ad188a1

INDEX, MATCH -> lookup the value in a given range -> https://www.bing.com/search?q=index+match+in+excel&cvid=1793f1299ae44ed4b7c7253c5c3ee4c8&aqs=edge.0.0l2j69i57j0l6.4979j0j1&pglt=515&FORM=ANNAB1&PC=U531
For example, INDEX(data_area, row_num, col_num), row_num <- MATCH(row_name, row_name_included_in_this_column, 0), 0 means exact match.


sumif -> sum numbers from a range when all of the specified criteria are met, based on AND logic.
https://www.got-it.ai/solutions/excel-chat/excel-tutorial/sumif/multiple-criteria-sumif
The syntax of SUMIFS is:
SUMIFS(sum_range, criteria_range1, criteria1, criteria_range2, criteria2,...)
