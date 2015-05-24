# ExcelLink
Easy binding links from WPF applications to Excel

This project uses Microsoft Office Interop for Excel and allows to retrieve cells from any worksheet and easily
bind them to WPF properties. Cells Values support INotifyPropertyChanged to have auto feedback.

Cells will be automatically updated when Excel changes even if the Cell is a formula (dependent cell) of the cell
that was actually changed.

Excel can be open, or will be opened in the background. If the file is already open in Excel ending the 
application will leave Excel running. As per design of the Interop services from Microsoft Excel will automatically 
close when the app exists if Excel wasn't currently running with the sheet used.

## Example bindings in WPF

```
// the datacontext for the WPF application
class DataContext
{
    private ExcelLink.Workbook _workbook = ExcelLink.Workbook("testdata.xlsx", true);
    public ExcelLink.Workbook Workbook { get { return _workbook; } }
}
```

Then you can use in XAML:
```
<TextBox Text="{Binding 'workbook.Sheets[Sheet1].Cells[2][3].Value'}"  />
```
This will bind to worksheet 'Sheet1' (when it exists) and the cell on row 2 column 3 or in Excel speak 'C3'.
(note that row, column and sheet indices all start from 1 and not 0!)

## Binding code properties to Excel
Next to binding Cell values from Excel to WPF controls. The cell itself can also be seen as a WPF control. 
It is a DependencyObject. This allows you to maintain a local property set of data and bind that back to Excel
causing updates on your model to be reflected in Excel and the other way around.

```
// for simplification no INotifyPropertyChanged is implemented on these examples
class MyAddress
{
    public String Streetname { get; set; }
    public String City { get; set; }
}
class MyPerson
{
    public String Firstname { get; set; }
    public String Lastname {get; set; }
    public MyAddress Address { get; set; }
}
```

Now to setup the binding from an Excel workbook use standard bindings.
```
MyPerson personA = new MyPerson();
personA.Address = new MyAddress();
personA.Street = "New Street";
ExcelLink.Cell cell = workbook.Sheet["Sheet1"].Cells[1][2];

Binding b = new Binding("Address.Streetname");
b.Source = personA;
b.Mode = BindingMode.TwoWay;
BindingOperations.SetBinding(cell, ExcelLink.Cell.ValueProperty, b);
```

Note that this will update the value in Excel with the current value that was already setup in the person A variable.
Using ExcelLink.ExcelBinding.Bind() allows one to first read the current value from the Excel sheet into the property
and then revert to the TwoWay binding. Example of that:
```
ExcelLink.ExcelBinding.Bind(cell, personA, "Address.Streetname", BindingMode.TwoWay, true);
```

# TODO
- support (named) cell ranges on sheets
- get size of ranges
- check COM cleanup

# License
Distributed under MIT license.

Copyright 2015 Sander Cox / Parallel Dimension