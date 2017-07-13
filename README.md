
# virtualSheet
This is a project to make an easy way to interact with tables in Excel VBA. It allows for addressing a table in a spreadsheet using a simple syntax, as well as complex filtering and searching. In most cases a range is returned so that all properties of the cells are available.

Some features:
*Search by word content
*Use a list of synonyms for searching
*Attempting to match while ignoring plurals
*Match using an arbitrary function
*Allows a list of words to ignore when searching

The excel spreadsheet provides documentation for these features as well as the class virtualSheets



Example simple usage:


'Creates a virtual Sheet and loads a table starting at cell A1
```vba
Set thisTable = New virtualSheet
thisTable.load ActiveSheet.Range("A1")
```



'Reurns a range of cells in column "betty" that match either "tampa" or "philly" in column "city" using user defined search parameters
```vba
Set thisCell = thisTable.rangeByMatch(Array("tampa", "philly"), Array("city"), "betty", matchType:="byParams")
If TypeOf thisCell Is Range Then thisCell.Font.Size = 14
```


