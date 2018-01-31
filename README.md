# large-xlsxtocsv
Using stream, it converts xlsx files to csv files which you can define delimiter for.

I added several methods for handling excel files carefully.
Such as
addMissedDelimitersIfItIsNotACell
addMissedDelimitersForBlankCell
addMissedDelimitersAtTheEnd

References :
https://poi.apache.org/spreadsheet/limitations.html
https://svn.apache.org/repos/asf/poi/trunk/src/examples/src/org/apache/poi/xssf/eventusermodel/XLSX2CSV.java
