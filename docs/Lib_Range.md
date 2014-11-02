# Range Library

* Name file: LIB_File.bas

## Function lists

* [getAddress](#getaddress-range)
* [selectXlDownRange](#selectxldownrange-range)

### getAddress (cell)

Return the address of the specified range (eg. $A$1)

#### Arguments
* **cell** as Range: the specified range

#### Specifications / limitations
* Return the absolute address of the range


### selectXlDownRange (cell)

Simulate the CTRL + SHIFT + DOWN from a specified range. Returns the range that should be selected.
For VBA users: returns the cell.end(xlDown) related range.

*Note: this could be a replacement of the [Dymanic Ranges](http://support.microsoft.com/kb/830287). Use one of the other method when best appropriated.*

#### Arguments
* **cell** as Range: the specified range

#### Specifications / limitations
* No Performance tests have been perfomed
* Will end after before the first empty cell
* If the next cell is empty, the function will return only the specified cell
* Can not be used in Name Manager

Example: =sum(selectXlDownRange(b3)) will sum all the values of the cell from b3 until b4, b5,... until Excel finds an empty cell
