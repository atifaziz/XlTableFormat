# XlTableFormat

## Introduction

The [XlTable format](https://github.com/atifaziz/XlTableFormat/wiki/XlTable:-Fast-Table-Format) is a fast table format when using [Dynamic Data Exchange (DDE)](http://en.wikipedia.org/wiki/Dynamic_Data_Exchange) to get data from Microsoft Excel. This project provides a single C# file called `XlTableFormat.cs` that adds the neccessary types and code to read and decode the XlTable format.

## Installing

There is a [NuGet package](https://www.nuget.org/packages/XlTableFormat.Source/) that embeds `XlTableFormat.cs` in a C# project.

There is no stand-alone class library version. If you need one, create a C# class library project, add the `XlTableFormat.Source` package from NuGet and declare the following types public:

    public partial class XlTableFormat {}
    public partial interface IXlTableDataFactory<out T> { }
    public partial class XlTableDataFactory<T> { }

## Usage

The simplest way to read the XlTable format is to use ones of the following `XlTableFormat.Read` overloads, each of which takes the data as input (either as a byte array or as a stream) and lazily yields a sequence of decoded objects:

    public static IEnumerable<object> Read(byte[] data)
    public static IEnumerable<object> Read(Stream stream)

The XlTable format defines the following data block types:

- Table
- Float
- String
- Bool
- Error
- Blank
- Int
- Skip

The type of each object in the sequence returned from the `Read` method is mapped from XlTable data blocks to C# and runtime types as follows:

- Table  = array of `int` with two integers, the row and column count respectively
- Float  = `double`
- String = `string`
- Bool   = `bool`
- Error  = [`ErrorWrapper`](http://msdn.microsoft.com/en-us/library/system.runtime.interopservices.errorwrapper.aspx)
- Blank  = `null`
- Int    = `int`
- Skip   = [`Missing.Value`](http://msdn.microsoft.com/en-us/library/system.reflection.missing.value.aspx)

The first object returned in the sequence always defines the table dimensions as an array of two integers with the row count as the first integer and column count as the second. The remaining objects are sequential content of table as cells values.
