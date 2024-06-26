<!-- default badges list -->
![](https://img.shields.io/endpoint?url=https://codecentral.devexpress.com/api/v1/VersionRange/128613095/19.2.3%2B)
[![](https://img.shields.io/badge/Open_in_DevExpress_Support_Center-FF7200?style=flat-square&logo=DevExpress&logoColor=white)](https://supportcenter.devexpress.com/ticket/details/T220356)
[![](https://img.shields.io/badge/📖_How_to_use_DevExpress_Examples-e9f6fc?style=flat-square)](https://docs.devexpress.com/GeneralInformation/403183)
[![](https://img.shields.io/badge/💬_Leave_Feedback-feecdd?style=flat-square)](#does-this-example-address-your-development-requirementsobjectives)
<!-- default badges end -->

# Spreadsheet Document API - Insert a Reference to an External Workbook

This example demonstrates how to insert an external reference link from a workbook to another workbook. 

An external workbook is created and populated with random data by importing a data table at runtime. Subsequently, the workbook is added to the [ExternalWorkbookCollection](https://docs.devexpress.com/OfficeFileAPI/DevExpress.Spreadsheet.ExternalWorkbookCollection). A cell formula with a reference to an external workbook is inserted in the current worksheet. The worksheet is saved to .XLSX file and opened with an application registered for that file format.

> [!important]
> The **Universal Subscription** or an additional **Office File API Subscription** is required to use this example in production code. Please refer to the [DevExpress Subscription](https://www.devexpress.com/Buy/NET/) page for pricing information.

## Files to Review

* [Form1.cs](./CS/DocServerExternalWorkbookSample/Form1.cs) (VB: [Form1.vb](./VB/DocServerExternalWorkbookSample/Form1.vb))

## Documentation

* [Cell Referencing](https://docs.devexpress.com/OfficeFileAPI/14916/spreadsheet-document-api/cell-basics/cell-referencing)
<!-- feedback -->
## Does this example address your development requirements/objectives?

[<img src="https://www.devexpress.com/support/examples/i/yes-button.svg"/>](https://www.devexpress.com/support/examples/survey.xml?utm_source=github&utm_campaign=spreadsheet-document-api-external-workbook&~~~was_helpful=yes) [<img src="https://www.devexpress.com/support/examples/i/no-button.svg"/>](https://www.devexpress.com/support/examples/survey.xml?utm_source=github&utm_campaign=spreadsheet-document-api-external-workbook&~~~was_helpful=no)

(you will be redirected to DevExpress.com to submit your response)
<!-- feedback end -->
