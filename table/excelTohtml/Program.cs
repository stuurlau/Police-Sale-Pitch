using Aspose.Cells;
// See https://aka.ms/new-console-template for more information
Workbook workbook = new Workbook("../feature-comparison.xlsx");

Worksheet worksheet = workbook.Worksheets[0];

// workbook.Worksheets = worksheet;

// Create HTML save options
HtmlSaveOptions saveOptions = new HtmlSaveOptions();
saveOptions.ExportImagesAsBase64 = true;
saveOptions.ExportGridLines = true;
saveOptions.ExportRowColumnHeadings = false;
saveOptions.IsExportComments = true;

// // Set options for a nicer HTML table
saveOptions.TableCssId = "ExcelTable";
saveOptions.AddTooltipText = true;
saveOptions.ExportWorksheetCSSSeparately = true;
saveOptions.ExportSimilarBorderStyle = true;
saveOptions.ExportSingleTab = true; // Embed all resources into a single HTML file


// HtmlSaveOptions options = new HtmlSaveOptions();
saveOptions.AddTooltipText = true; // Ensures comments are converted to tooltips

workbook.Save("table.html", saveOptions);