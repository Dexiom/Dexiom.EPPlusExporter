# Dexiom.EPPlusExporter
[![Build status](https://ci.appveyor.com/api/projects/status/pbnru8yvomkpov5u?svg=true)](https://ci.appveyor.com/project/jpare/dexiom-epplusexporter)

Project Description
-------------------
A very simple, yet incredibly powerfull library to generate Excel documents out of objects, arrays, lists, collections, etc.

### Download & Install

```
Install-Package Dexiom.EPPlusExporter
```

#Usage

## Creating a new exporter
There are two ways to instantiate the exporter.
```csharp
var exporter1 = new EnumerableExporter<Employee>(employees) //#1 with the standard contructor
var exporter2 = EnumerableExporter.Create(employees) //#2 with the create method using type inference
```
Both gives the same output however, **we recommend #2** because it will make things much easier when working with anonymous types (*the type inference is important when working with the fluent interface*). It's also shorter ;-)

## Basic example
Let's say you want to dump an array or a list of objects to Excel (without any specific formatting).  
One line of code is all you need.

```csharp
var excelPackage = EnumerableExporter.Create(employees).CreateExcelPackage();
```

## Exporting an Anonymous Enumerable

```csharp
var employees = TestHelper.GetEmployees().Select(n => new
{
	Login = n.UserName,
	Mail = n.Email
});

var exporter = EnumerableExporter.Create(employees);
var excelPackage = exporter.CreateExcelPackage();
```

## Quick Customizations (using fluent interface)
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
