
# virtualSheet
This is a project to make an easy way to interact with tables in Excel VBA. It allows for addressing cells arranged as a table in excel using a simple VBA syntax, as well as complex filtering and searching. In most cases a range is returned so that all properties of the cells are available. The virtualSheet only exists as a reference to ranges on a spreadsheet, and thus leaves no trace on the spreadsheet once it goes out of scope.

Some features:

*Search by word content (i.e. "brown dog" matches "Jaimie's little dog is brown")

*Use a list of synonyms for searching

*Attempting to match while ignoring plurals

*Match using an arbitrary function

*Allows a list of words to ignore when matching 

The excel spreadsheet provides documentation for these features as well as the class virtualSheets. There are additional undocumented features such as column math and joining.



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


