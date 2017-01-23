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


## The basic example
Let's say you want to dump an array or a list of objects to Excel (without any specific formatting).
One line of code is all you need.

```cs
var excelPackage = EnumerableExporter.Create(employees).CreateExcelPackage();
```

## Exporting anonymous enumerables

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
var excelPackage = new EnumerableExporter<Employee>(data)
	.Ignore(n => n.UserName) //do not show the "UserName" column in the output
	.Ignore(n => n.Phone) //do not show the "Phone" column in the output
	.DisplayFormatFor(n => n.DateOfBirth, "{0:yyyy-MM-dd}") //Set a specific format for the "DateOfBirth"
	.CreateExcelPackage();
```

* Available customizations:
 * Ignore: to ignore a column when generating the Excel document
 * DisplayFormatFor is used to specify a custom display format
