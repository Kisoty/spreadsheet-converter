# Spreadsheet converter

Simple library to convert google spreadsheets to xml and back.

Uses php simpleXML lib, therefore has pretty low performance. However, library was made for small sheets and kinda fun (if you can call working with google API fun,
especially via official client), so I see no troubles with that.

Converter supports a lot of google spreadsheet features like cell styles, row and column width and length, cell merges etc. 
I see no reason to list everything here since noone needs this lib and docs to it.

If there is someone who wants to use it, do whatever you want :) There are also some bugs with links to subsequent sheets in formulas and cell merges with IMPORTRANGE,
though it seems to be pretty common for google spreadsheets, and it's not so easy and interesting to fix this, so I'm not gonna do it (at least, for now).
