**_Not ready for production_**
---

# Dexiom.EPPlusExporter


Project Description
-------------------
A very simple, yet incredibly powerfull library to generate Excel documents with objects, arrays, lists, collections, etc.

### Download & Install

```
Install-Package Dexiom.EPPlusExporter
```

#Usage

## Creating a new exporter
There are two ways to instantiate the exporter.
```cs
var exporter1 = new EnumerableExporter<Employee>(employees) //#1 with the standard contructor
var exporter2 = EnumerableExporter.Create(employees) //#2 with the create method using type inference
```
Both gives the same output however, **we recommend #2** because it will make things much easier when working with anonymous types (*the type inference is important when working with the fluent interface*). It's also shorter ;-)

## Basic example
Let's say you want to dump an array or a list of objects to Excel (without any specific formatting).  
One line of code is all you need.

```cs
var excelPackage = EnumerableExporter.Create(employees).CreateExcelPackage();
```

## Exporting an Anonymous Enumerable

```cs
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

```cs
var excelPackage = EnumerableExporter.Create(employees) //new EnumerableExporter<Employee>(employees)
	.Ignore(n => n.UserName) //do not show the "UserName" column in the output
	.Ignore(n => n.Phone) //do not show the "Phone" column in the output
	.DisplayFormatFor(n => n.DateOfBirth, "{0:yyyy-MM-dd}") //Set a specific format for the "DateOfBirth"
	.CreateExcelPackage();
```

* Available customizations:
 * **Ignore** is used to skip a column when generating the document
 * **DisplayFormatFor** is used to specify a custom display format
