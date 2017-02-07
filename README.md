# Dexiom.EPPlusExporter
[![Build status](https://ci.appveyor.com/api/projects/status/pbnru8yvomkpov5u?svg=true)](https://ci.appveyor.com/project/jpare/dexiom-epplusexporter)
[![codecov](https://codecov.io/gh/Dexiom/Dexiom.EPPlusExporter/branch/master/graph/badge.svg)](https://codecov.io/gh/Dexiom/Dexiom.EPPlusExporter)
[![NuGet](https://img.shields.io/nuget/v/Dexiom.EPPlusExporter.svg)](https://www.nuget.org/packages/Dexiom.EPPlusExporter/)

## Download & Install

```
Install-Package Dexiom.EPPlusExporter
```

## Wiki

Please review the [Wiki](https://github.com/Dexiom/Dexiom.EPPlusExporter/wiki) pages on how to use Dexiom.EPPlusExporter.

## Quick Usage Preview 

### Basic example
Let's say you want to dump an array or a list of objects to Excel (without any specific formatting).  
This is what you would do:
```csharp
//create the exporter
var exporter = EnumerableExporter.Create(employees);

//generate the document
var excelPackage = exporter.CreateExcelPackage(); 

//save the document
excelPackage.SaveAs(new FileInfo("C:\\example1.xlsx")); 
```

### Quick Customizations (using fluent interface)
Quick customization can be accomplished by using the fluent interface like this:

```csharp
var excelPackage = EnumerableExporter.Create(employees) //new EnumerableExporter<Employee>(employees)
	.DefaultNumberFormat(typeof(DateTime), "yyyy-MM-dd") //set default format for all DateTime columns
	.NumberFormatFor(n => n.DateOfBirth, "{0:yyyy-MMM-dd}") //set a specific format for the "DateOfBirth"
	.Ignore(n => n.UserName) //do not show the "UserName" column in the output
	.Ignore(n => n.Email) //do not show the "Column" column in the output
	.TextFormatFor(n => n.Phone, "Cell: {0}") //add the "Cell: " prefix to the value
	.StyleFor(n => n.DateContractEnd, style =>
	{
	    style.Fill.Gradient.Color1.SetColor(Color.Yellow);
	    style.Fill.Gradient.Color2.SetColor(Color.Green);
	}) //the cells in this columns now have a gradiant background
	.CreateExcelPackage();
```

* Available customizations:
 * **Ignore** is used to skip a column when generating the document
 * **DefaultNumberFormat** is used to specify a default display format for a specific type
 * **NumberFormatFor** is used to set a specific format (just like you would using Excel)
 * **TextFormatFor** is used to convert a value to text
 * **StyleFor** is used to alter the style for a specific column

## License
Copyright (c) Consultation Dexiom. All rights reserved.

Licensed under the [MIT](LICENSE.txt) License.
