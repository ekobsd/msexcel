# Excel

This is a cheat sheet repo for Excel

> Note: In this tutorial I will be using Ms Windows Keyboard layout

## Keyboard Shortcuts -- Navigation and Selection

| Shortcut        | Description                |
| --------------- | -------------------------- |
| Tab             | Go to next cell in the row |
| ctrl + arrow up         | go all the way up/nearest filled cell up        |
| ctrl + arrow down         | go all the way down/nearest filled cell down        |
| ctrl + arrow left        | go all the way left/nearest filled cell left        |
| ctrl + arrow right         | go all the way right/nearest filled cell right        |
| ctrl + space         | select all column        |
| ctrl + a         | Select entire sheet or table(if cursor is inside a table)        |
| shift + arrow         | Select cell following the arrow (up,left,down,right)       |

## Keyboard Shortcuts -- Copy Pasting

| Shortcut        | Description                |
| --------------- | -------------------------- |
| ctrl + c         | copy selected cell(s)        |
| ctrl + x         | cut selected cell(s)        |
| ctrl + v         | paste selected cell(s)        |
| ctrl + shift + v | Paste values without style |

## Keyboard Shortcuts -- Create a Table

| Shortcut        | Description                |
| --------------- | -------------------------- |
| ctrl + t             | Creating a Table |

> Note: We can remove the table(convert it as normal cell) by using "Convert to range" menu on table design tab.

## Tips

### Formulas

- You can add formulas to a cell by typing `=` and then the formula
- After typing `=` you can click on a cell to add it to the formula (e.g. `=C4*D4`)
- You can make any mathematical operation in a formula (e.g. `=A1+B1`)
- You can drag the bottom right corner of a cell to copy the formula to other cells
- Select multiple cells and drag the bottom right corner to copy the formula to multiple cells
- If you copy a cell with a formula and paste it into another cell, the formula will be updated to use the new cell's position

> Note: `A1` is a cell reference, `A` is the column and `1` is the row

### Functions
> - A function always followed with parenthesis symbol (open and closed)
> - A function needs input ("text", number, logical(True,False), or output from other function)
> - A function will produce and returning result (text, number, logical (true,false)
> - some common functions can be found below

#### Text

- `=CONCATENATE(A1, " ", B1)` - Returns the concatenation of the values in A1 and B1
- `=LEFT(A1, 5)` - Returns the first 5 characters of the value in A1
- `=RIGHT(A1, 5)` - Returns the last 5 characters of the value in A1
- `=MID(A1, 5, 10)` - Returns 10 characters of the value in A1 starting from the 5th character
- `=LEN(A1)` - Returns the length of the value in A1
- `=LOWER(A1)` - Returns the value in A1 in lower case
- `=UPPER(A1)` - Returns the value in A1 in upper case
- `=PROPER(A1)` - Returns the value in A1 in proper case (e.g. "hello world" becomes "Hello World")
- `=TRIM(A1)` - Returns the value in A1 with all leading and trailing spaces removed
- `=SUBSTITUTE(A1, " ", "")` - Returns the value in A1 with all spaces removed

> Note: Google Sheets and Excel have many more functions than the ones listed above

#### Numbers

- `=MAX(A1:A10)` - Returns the maximum value in the range
- `=MIN(A1:A10)` - Returns the minimum value in the range
- `=SUM(A1:A10)` - Returns the sum of the values in the range
- `=AVERAGE(A1:A10)` - Returns the average of the values in the range
- `=COUNT(A1:A10)` - Returns the number of values in the range
- `=COUNTA(A1:A10)` - Returns the number of non-empty values in the range
- `=COUNTBLANK(A1:A10)` - Returns the number of empty values in the range
- `=COUNTIF(A1:A10, ">10")` - Returns the number of values in the range that are greater than 10
- `=COUNTIFS(A1:A10, ">10", B1:B10, "<20")` - Returns the number of values in the range that are greater than 10 and less than 20
- `=SUMIF(A1:A10, ">10")` - Returns the sum of the values in the range that are greater than 10
- `=SUMIFS(A1:A10, ">10", B1:B10, "<20")` - Returns the sum of the values in the range that are greater than 10 and less than 20
- `=AVERAGEIF(A1:A10, ">10")` - Returns the average of the values in the range that are greater than 10
- `=AVERAGEIFS(A1:A10, ">10", B1:B10, "<20")` - Returns the average of the values in the range that are greater than 10 and less than 20
- `=IF(A1>10, "Greater than 10", "Less than 10")` - Returns "Greater than 10" if the value in A1 is greater than 10, otherwise returns "Less than 10"
- `=OR(A1>10, B1>10)` - Returns TRUE if either A1 or B1 is greater than 10
- `=SUMIF(A1:A10, ">10")` - Returns the sum of the values in the range that are greater than 10
- `=SUMIF(A1:A10, ">10", B1:B10)` - Returns the sum of the values in the range B1:B10 if the value in the range A1:A10 is greater than 10

> Note: You can also return a formula in an IF statement (e.g. `=IF(A1>10, A1*2, A1*3)`)

### Absolute References (Lock the cell position)

- You can use `$` to make a reference absolute (e.g. `$A$1`)
- If you copy a cell with an absolute reference and paste it into another cell, the reference will not be updated

> Note: It is useful to use absolute references when you are referencing a cell that you do not want to change (e.g. a constant)
>
> BEWARE: If the '$' symbol is placed before the column letter (e.g. `$A1`) then the column will not change when the formula is copied, but the row will and vice versa (e.g. `A$1`)

### Conditional Formatting

Conditional formatting allows you to change the style of a cell based on its value

- Select the cells you want to apply the conditional formatting to
- Click on the "Format" menu and select "Conditional formatting"
- Select the type of conditional formatting you want to apply (e.g. "Greater than")
- Enter the value you want to compare the cell to (e.g. 10), add the percentage if you want to compare it to a percentage (e.g. 10%, or use 0.1 for 10%)
- Select the style you want to apply to the cell if the condition is met (e.g. "Bold")
- Click on the "Done" button


### Split Column Text to Multiple Columns

- Select the column you want to split
- Click on the "Data" menu and select "Split text to columns"
- Select the separator you want to use (e.g. "Space")

> Example: "John Doe" cell will be split into "John" and "Doe" cells in new columns

### Extract Text from a Cell

- `=LEFT(A1, 5)` - Returns the first 5 characters from the `A1` cell
- `=RIGHT(A1, 5)` - Returns the last 5 characters from the `A1` cell
- `=MID(A1, 5, 10)` - Returns 10 characters from the `A1` cell starting from the 5th character

> Note: starts at 1, not 0 like many programming languages

### Lookup Table using INDEX MATCH

A lookup table allows you to lookup a value in a table and return a value from another column in the same row
Pairing the INDEX and MATCH functions together allows you to return cells to the LEFT of your match. It's a pretty handy combination and the example file holds a simple demonstration for your reference.

`INDEX` will need input:
- range for reference (change this with only column range, not the entire table)
- row position (change this with `MATCH` for dinamically access the row position value)
- column position (we can get rid of this, because the range for reference has already become one column

> - so the formula will be as simple as `=INDEX(Column range, row position)`
> - INDEX will return the value of selected selected with matching row position

`MATCH` will need input
- Value to be compared (we use @ for dinamically access each rows for table)
- Value for reference (or column for reference to be compared)
- number 1, 0, -1. We choose 0 for exactly match the value.
- `MATCH` will return cell position value of exactly match compared cell
  
> - We combine `INDEX` and `Match` as a handy tools to fill up blank cell with proper value
> - We use bracket [] to access column of table
> - the formula will be `=INDEX(table[column for reference],match([@column to be compared], table[column to be reference],0))`
> - We used an absolute reference ($ symbol) for the table if we want to be able to copy the formula to other cells

### Sorting table (passed, William has mastered it)

### Filtering table (passed, William has mastered it)

### Pivot Table

A Pivot Table allows you to summarize data from a table into a new table

- Select the cells you want to include in the pivot table
- Click on the "Insert" menu and select "Pivot table"

> Note: A new sheet will be created with the pivot table

- Go to the "Pivot table editor" section on the right
- Select either the suggested pivot table or create your own by selecting the columns, rows, values or filters you want to include in the pivot table
- Click on the "Add" button

> TIP: You can then insert a chart based on the pivot table

### Calculate the percentage of a currency value

- Add a column with the value you want to calculate the percentage of (e.g. 100) and format it as currency or accounting
- Add a column with the percentage you want to calculate (e.g. 10%)
- Add a column with the formula `=A1*B1` (e.g. `=100*10%`)

### General Tips

- You can use the `=` operator to convert a value to a number (e.g. `=A1*2`)
- You can click on a column header to select the entire column
- You can add a column by right clicking on a column header and selecting "Insert 1 left"
- You can calculate the number of days between two dates by subtracting them (e.g. `=A1-A2`)
- Double click on the column right border to auto resize the column to fit the content
- You can rotate text in a cell by selecting the cell and then clicking on the "Text Rotation" button in the toolbar
- You can either use `.5` or `50%` to represent 50%

#### Other Function about Dates

- `=YEAR(TODAY())` - Returns the current year
- `=TEXT(YEAR(TODAY()),"YY")` - Returns the current year in two digits (YY) format
- `=TEXT(TODAY(),"DD/MM/YYYY")` - Returns the current date in DD/MM/YYYY format
- `=MONTH(TODAY())` - Returns the current month
- `=DAY(TODAY())` - Returns the current day
- `=WEEKDAY(TODAY())` - Returns the current day of the week (e.g. 1 for Sunday, 2 for Monday, etc.)
- `=TEXT(TODAY(),"dddd")` - Returns the current day of the week (e.g. Sunday, Monday, etc.)
- `=DATE(YEAR(TODAY()), MONTH(TODAY()), DAY(TODAY())+1)` - Returns the date of tomorrow
