﻿using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using DinkToPdf;
using DinkToPdf.Contracts;
using System.Data;
using System.Globalization;
using System.Text;


namespace TelegramBotAPI.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class FileController : ControllerBase
    {
        private readonly string _folderPath;
        private readonly IConverter _pdfConverter;
        public FileController(IConfiguration configuration, IConverter pdfConverter)
        {
            _folderPath = configuration["FolderPath"];
            _pdfConverter = pdfConverter;
        }

        [HttpGet("getfile")]
        public IActionResult GetFile([FromQuery] string parameter1, [FromQuery] string parameter2)
        {
            try
            {
                var folderPath = Path.Combine(_folderPath, "");
                if (!Directory.Exists(folderPath))
                {
                    return NotFound("Folder not found.");
                }
                bool isLandscape = false;
                // Get all files in the folder
                var files = Directory.GetFiles(folderPath);
                var paramsDt = parameter1.Split('_');
                var paramText = paramsDt[0].Trim();
                var FromDate = DateTime.Parse(paramsDt[1].Trim());
                FromDate = FromDate.AddMinutes(-2);
                var ToDate = DateTime.Parse(paramsDt[2].Trim());
                ToDate = ToDate.AddMinutes(2);
                DataTable dt = new DataTable();
                dt.Columns.Add("FileName");
                dt.Columns.Add("CreatedDateTime", typeof(DateTime));

                foreach (string filePath in files)
                {
                    // Get the file name
                    string fileName = System.IO.Path.GetFileName(filePath);
                    // Get the creation date of the file
                    DateTime creationDate = System.IO.File.GetCreationTime(filePath);
                    string extenstion = System.IO.Path.GetExtension(filePath);
                    if (extenstion.Contains("xls"))
                    {
                        dt.Rows.Add();
                        dt.Rows[dt.Rows.Count - 1][0] = fileName;
                        dt.Rows[dt.Rows.Count - 1][1] = creationDate;
                    }
                }

                if (files.Length == 0)
                {
                    return NotFound("No files found in the folder.");
                }

                DataView view = dt.DefaultView;
                view.Sort = "CreatedDateTime DESC";
                string formattedDateFrom = FromDate.ToString("MM/dd/yyyy HH:mm:ss", CultureInfo.InvariantCulture);
                string formattedDateTo = ToDate.ToString("MM/dd/yyyy HH:mm:ss", CultureInfo.InvariantCulture);
                //view.RowFilter = $"CreatedDateTime >= #{formattedDateFrom}# AND CreatedDateTime <= #{formattedDateTo}#";

                //DataTable sortedDate = view.ToTable();

                if (paramText == "M2M EXCEL")
                {
                    view.RowFilter = $"CreatedDateTime >= #{formattedDateFrom}# AND CreatedDateTime <= #{formattedDateTo}# AND FileName like '%M2M%'";
                    dt = view.ToTable();

                    //dt.DefaultView.RowFilter = "FileName like '%M2M%'";
                    //dt = dt.DefaultView.ToTable();

                    if (dt.Rows.Count > 1)
                    {
                        isLandscape = true;
                        string FileName1 = dt.Rows[1][0].ToString();
                        string FileName2 = dt.Rows[0][0].ToString();
                        dt = new DataTable();
                        dt = M2M(FileName1, FileName2);

                    }

                }
                else if (paramText == "UPDATE EXCEL")
                {
                    view.RowFilter = $"CreatedDateTime >= #{formattedDateFrom}# AND CreatedDateTime <= #{formattedDateTo}# AND FileName like '%UPDATE ALL%'";
                    dt = view.ToTable();

                    //dt.DefaultView.RowFilter = "FileName like '%UPDATE ALL%'";
                    //dt = dt.DefaultView.ToTable();

                    if (dt.Rows.Count > 1)
                    {
                        string FileName1 = dt.Rows[1][0].ToString();
                        string FileName2 = dt.Rows[0][0].ToString();
                        dt = new DataTable();
                        dt = FileUpdateExcel(FileName1, FileName2);
                    }
                }
                else if (paramText == "COM POS")
                {
                    view.RowFilter = $"CreatedDateTime >= #{formattedDateFrom}# AND CreatedDateTime <= #{formattedDateTo}# AND FileName like '%COM POS%'";
                    //dt.DefaultView.RowFilter = "FileName like '%COM POS%'";
                    dt = view.ToTable();
                    //dt = dt.DefaultView.ToTable();

                    if (dt.Rows.Count > 1)
                    {
                        string FileName1 = dt.Rows[1][0].ToString();
                        string FileName2 = dt.Rows[0][0].ToString();
                        dt = new DataTable();
                        dt = COMPOS(FileName1, FileName2);

                    }

                }
                else
                {
                    return NotFound($"No file found for your input {parameter1.Trim()}");
                }
                // Process parameters and retrieve the file
                var pdfContent = "";
                pdfContent += ConvertDataTableToHtml(dt);



                // Convert the accumulated content to PDF
                var pdfBytes = ConvertHtmlToPdf(pdfContent, isLandscape);
                parameter1 = paramText.Replace(" ", "") + "_" + DateTime.Now.ToString("ddMMyyyyHHmmss");
                // Return PDF as bytes
                return File(pdfBytes, "application/pdf", $"{parameter1}.pdf");
            }
            catch (Exception ex)
            {
                throw ex;
            }

        }
        private DataTable LoadXlsxIntoDataTable(string xlsxFilePath)
        {
            try
            {
                using (var package = new ExcelPackage(new FileInfo(xlsxFilePath)))
                {
                    var worksheet = package.Workbook.Worksheets[0]; // Get the first worksheet
                    var dataTable = new DataTable();

                    // Load columns
                    foreach (var firstRowCell in worksheet.Cells[1, 1, 1, worksheet.Dimension.End.Column])
                    {
                        dataTable.Columns.Add(firstRowCell.Text);
                    }

                    // Load rows
                    for (var rowNum = 2; rowNum <= worksheet.Dimension.End.Row; rowNum++)
                    {
                        var wsRow = worksheet.Cells[rowNum, 1, rowNum, worksheet.Dimension.End.Column];
                        var row = dataTable.NewRow();
                        foreach (var cell in wsRow)
                        {
                            row[cell.Start.Column - 1] = cell.Text;
                        }
                        dataTable.Rows.Add(row);
                    }

                    return dataTable;
                }
            }
            catch (Exception ex)
            {

                throw ex;
            }

        }

        private string ConvertDataTableToHtml(DataTable dataTable)
        {
            try
            {
                var html = @"
    <style>
        table {
            page-break-inside: auto;
            width: 100%;
            border-collapse: collapse;
        }

        tr {
            page-break-inside: avoid;
            page-break-after: auto;
        }
 thead {
        display: table-header-group;
    }
    tbody {
        display: table-row-group;
    }


        td, th {
            word-wrap: break-word;
            padding: 5px;
            border: 1px solid black;
        }
    </style>";
                html += "<table border='1' cellpadding='5' cellspacing='0' style='width: 100%;'> <thead>";
                html += "<tr>";
                foreach (DataColumn column in dataTable.Columns)
                {
                    html += $"<th>{column.ColumnName}</th>";
                }
                html += "</tr> </thead>";
                html += "<tbody>";

                foreach (DataRow row in dataTable.Rows)
                {
                    html += "<tr>";
                    foreach (var cell in row.ItemArray)
                    {
                        html += $"<td>{cell}</td>";
                    }
                    html += "</tr>";
                }

                html += "</tbody>";
                html += "</table>";
                return html;
            }
            catch (Exception ex)
            {

                throw ex;
            }

        }

        private byte[] ConvertHtmlToPdf(string htmlContent, bool islandscap = false)
        {
            try
            {
                var globalSettings = new GlobalSettings
                {
                    ColorMode = ColorMode.Color,
                    PaperSize = PaperKind.A4,
                    Orientation = islandscap ? Orientation.Landscape : Orientation.Portrait,
                };
                var dt = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");
                var objectSettings = new ObjectSettings
                {
                    WebSettings = { DefaultEncoding = "utf-8" },
                    PagesCount = true,
                    HtmlContent = htmlContent,
                    HeaderSettings = new HeaderSettings
                    {
                        FontSize = 6,
                        FontName = "Times New Roman",
                        Right = $"{dt}",
                        Spacing = 5,
                    },

                };

                var pdf = new HtmlToPdfDocument()
                {
                    GlobalSettings = globalSettings,
                    Objects = { objectSettings },
                };

                return _pdfConverter.Convert(pdf);
            }
            catch (Exception ex)
            {

                throw ex;
            }

        }

        private DataTable M2M(string fileName1, string fileName2)
        {
            try
            {
                DataTable dt = new DataTable();
                dt.Columns.Add("Login");
                dt.Columns.Add("Name");
                dt.Columns.Add("Per (%)-1");
                dt.Columns.Add("M2M-1");
                dt.Columns.Add("Partner-1");
                dt.Columns.Add("Net Amount-1");
                dt.Columns.Add("Per (%)-2");
                dt.Columns.Add("M2M-2");
                dt.Columns.Add("Partner-2");
                dt.Columns.Add("Net Amount-2");
                dt.Columns.Add("Diff. Per (%)");
                dt.Columns.Add("Diff. M2M");
                dt.Columns.Add("Diff. Partner");
                dt.Columns.Add("Diff. Net Amount");

                DataTable dt1 = new DataTable();
                dt1.Columns.Add("Login");
                dt1.Columns.Add("Name");
                dt1.Columns.Add("Per (%)");
                dt1.Columns.Add("M2M");
                dt1.Columns.Add("Partner");
                dt1.Columns.Add("Net Amount");


                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                string destinationFilePath = Path.Combine(_folderPath, "");
                string FilePath1 = destinationFilePath + "\\" + fileName1;
                string FilePath2 = destinationFilePath + "\\" + fileName2;

                using (var package = new ExcelPackage(new FileInfo(FilePath1)))
                {
                    // Get the first worksheet in the workbook
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                    // Iterate through the rows and columns
                    for (int row = 3; row <= worksheet.Dimension.Rows - 1; row++)
                    {
                        dt.Rows.Add();
                        for (int col = 1; col <= worksheet.Dimension.Columns; col++)
                        {
                            //if (col == 4 || col == 5) continue;
                            var dl = worksheet.Cells[row, col].Value.ToString();
                            var pr = dl.Replace("%", "").ToString();
                            if (decimal.TryParse(pr, out decimal decimalValue))
                            {
                                decimal intValue = Convert.ToDecimal(decimalValue);
                                if (dl.Contains("%"))
                                {
                                    dt.Rows[dt.Rows.Count - 1][col - 1] = intValue + " %";
                                }
                                else
                                {
                                    if (col == 4 || col == 5)
                                    {

                                    }
                                    else if (col == 6)
                                    {
                                        dt.Rows[dt.Rows.Count - 1][3] = intValue.ToString("N2");
                                    }
                                    else if (col == 7)
                                    {
                                        dt.Rows[dt.Rows.Count - 1][4] = intValue.ToString("N2");
                                    }
                                    else if (col == 8)
                                    {
                                        dt.Rows[dt.Rows.Count - 1][5] = intValue.ToString("N2");
                                    }
                                    else
                                    {
                                        dt.Rows[dt.Rows.Count - 1][col - 1] = intValue;
                                    }

                                }
                            }
                            else
                            {
                                dt.Rows[dt.Rows.Count - 1][col - 1] = worksheet.Cells[row, col].Value;
                            }
                        }
                    }
                }

                using (var package = new ExcelPackage(new FileInfo(FilePath2)))
                {
                    // Get the first worksheet in the workbook
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                    // Iterate through the rows and columns
                    for (int row = 3; row <= worksheet.Dimension.Rows - 1; row++)
                    {
                        dt1.Rows.Add();
                        for (int col = 1; col <= worksheet.Dimension.Columns; col++)
                        {
                            var dl = worksheet.Cells[row, col].Value.ToString();
                            var pr = dl.Replace("%", "").ToString();
                            if (decimal.TryParse(pr, out decimal decimalValue))
                            {
                                decimal intValue = Convert.ToDecimal(decimalValue);
                                if (dl.Contains("%"))
                                {
                                    dt1.Rows[dt1.Rows.Count - 1][col - 1] = intValue + " %";
                                }
                                else
                                {
                                    if (col == 4 || col == 5)
                                    {

                                    }
                                    else if (col == 6)
                                    {
                                        dt1.Rows[dt1.Rows.Count - 1][3] = intValue.ToString("N2");
                                    }
                                    else if (col == 7)
                                    {
                                        dt1.Rows[dt1.Rows.Count - 1][4] = intValue.ToString("N2");
                                    }
                                    else if (col == 8)
                                    {
                                        dt1.Rows[dt1.Rows.Count - 1][5] = intValue.ToString("N2");
                                    }
                                    else
                                    {
                                        dt1.Rows[dt1.Rows.Count - 1][col - 1] = intValue;
                                    }

                                }
                            }
                            //if (decimal.TryParse(dl, out decimal decimalValue))
                            //{
                            //    int intValue = Convert.ToInt32(decimalValue);
                            //    dt1.Rows[dt1.Rows.Count - 1][col - 1] = intValue;
                            //}
                            else
                            {
                                dt1.Rows[dt1.Rows.Count - 1][col - 1] = worksheet.Cells[row, col].Value;
                            }
                        }
                    }
                }

                for (int i = 0; i <= dt.Rows.Count - 1; i++)
                {
                    dt1.DefaultView.RowFilter = "Login=" + dt.Rows[i][0];

                    if (dt1.DefaultView.ToTable().Rows.Count > 0)
                    {
                        DataRow[] rows = dt1.Select("Login='" + dt.Rows[i][0] + "'");

                        // Update the specified column for each selected row
                        dt.Rows[i][6] = rows[0].ItemArray[2];
                        dt.Rows[i][7] = rows[0].ItemArray[3];
                        dt.Rows[i][8] = rows[0].ItemArray[4];
                        dt.Rows[i][9] = rows[0].ItemArray[5];

                        dt.Rows[i][10] = Convert.ToDecimal(dt.Rows[i][2].ToString().Replace("%", "")) - Convert.ToDecimal(dt.Rows[i][6].ToString().Replace("%", "")) + "%";
                        dt.Rows[i][11] = (Convert.ToDecimal(dt.Rows[i][3]) - Convert.ToDecimal(dt.Rows[i][7])).ToString("N2");
                        dt.Rows[i][12] = (Convert.ToDecimal(dt.Rows[i][4]) - Convert.ToDecimal(dt.Rows[i][8])).ToString("N2");
                        dt.Rows[i][13] = (Convert.ToDecimal(dt.Rows[i][5]) - Convert.ToDecimal(dt.Rows[i][9])).ToString("N2");


                        //if (!string.IsNullOrEmpty(dt.Rows[i][3].ToString()))
                        //{
                        //    dt.Rows[i][9] = Convert.ToInt32(dt.Rows[i][3]) - Convert.ToInt32(dt.Rows[i][6]);
                        //}


                        //UpdateDataTable(dt, "Login='" + Convert.ToInt64(dt.Rows[i][0]) + "'", 7, dt1.DefaultView.Table.Rows[0][7]);
                    }
                }
                return dt;
            }
            catch (Exception ex)
            {

                throw ex;
            }


        }
        private DataTable File11000(string fileName1, string fileName2)
        {
            try
            {
                DataTable dt = new DataTable();
                dt.Columns.Add("Login");
                dt.Columns.Add("Name");
                dt.Columns.Add("Per(%)");
                dt.Columns.Add("Credit");
                dt.Columns.Add("Shared Credit");
                dt.Columns.Add("M2M");
                dt.Columns.Add("Partner");
                dt.Columns.Add("Net Amount");
                dt.Columns.Add("Diff in Net Amt");

                DataTable dt1 = new DataTable();
                dt1.Columns.Add("Login");
                dt1.Columns.Add("Name");
                dt1.Columns.Add("Per(%)");
                dt1.Columns.Add("Credit");
                dt1.Columns.Add("Shared Credit");
                dt1.Columns.Add("M2M");
                dt1.Columns.Add("Partner");
                dt1.Columns.Add("Net Amount");

                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                string destinationFilePath = Path.Combine(_folderPath, "");
                string FilePath1 = destinationFilePath + "\\" + fileName1;
                string FilePath2 = destinationFilePath + "\\" + fileName2;

                using (var package = new ExcelPackage(new FileInfo(FilePath1)))
                {
                    // Get the first worksheet in the workbook
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                    // Iterate through the rows and columns
                    for (int row = 3; row <= worksheet.Dimension.Rows - 1; row++)
                    {
                        dt.Rows.Add();
                        for (int col = 1; col <= worksheet.Dimension.Columns; col++)
                        {
                            var dl = worksheet.Cells[row, col].Value.ToString();
                            var pr = dl.Replace("%", "").ToString();
                            if (decimal.TryParse(pr, out decimal decimalValue))
                            {
                                int intValue = Convert.ToInt32(decimalValue);
                                if (dl.Contains("%"))
                                {
                                    dt.Rows[dt.Rows.Count - 1][col - 1] = intValue + " %";
                                }
                                else
                                {
                                    dt.Rows[dt.Rows.Count - 1][col - 1] = intValue;
                                }
                            }
                            else
                            {
                                dt.Rows[dt.Rows.Count - 1][col - 1] = worksheet.Cells[row, col].Value;
                            }
                        }
                    }
                }

                using (var package = new ExcelPackage(new FileInfo(FilePath2)))
                {
                    // Get the first worksheet in the workbook
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                    // Iterate through the rows and columns
                    for (int row = 3; row <= worksheet.Dimension.Rows - 1; row++)
                    {
                        dt1.Rows.Add();
                        for (int col = 1; col <= worksheet.Dimension.Columns; col++)
                        {
                            var dl = worksheet.Cells[row, col].Value.ToString();
                            if (decimal.TryParse(dl, out decimal decimalValue))
                            {
                                int intValue = Convert.ToInt32(decimalValue);
                                dt1.Rows[dt1.Rows.Count - 1][col - 1] = intValue;
                            }
                            else
                            {
                                dt1.Rows[dt1.Rows.Count - 1][col - 1] = worksheet.Cells[row, col].Value;
                            }
                        }
                    }
                }

                for (int i = 0; i <= dt.Rows.Count - 1; i++)
                {
                    dt1.DefaultView.RowFilter = "Login=" + dt.Rows[i][0];
                    UpdateDataTable(dt, "Login='" + Convert.ToInt64(dt.Rows[i][0]) + "'", 7, dt1.DefaultView.Table.Rows[0][7]);
                }
                return dt;
            }
            catch (Exception ex)
            {

                throw ex;
            }


        }
        private DataTable COMPOS(string fileName1, string fileName2)
        {
            try
            {
                DataTable dtCMX = new DataTable();
                dtCMX.Columns.Add("Symbol");
                dtCMX.Columns.Add("Type-1");
                dtCMX.Columns.Add("Volume-1");
                dtCMX.Columns.Add("Holding Volume-1");
                dtCMX.Columns.Add("Type-2");
                dtCMX.Columns.Add("Volume-2");
                dtCMX.Columns.Add("Holding Volume-2");
                dtCMX.Columns.Add("Type-Diff");
                dtCMX.Columns.Add("Volume-Diff");
                dtCMX.Columns.Add("Holding Volume-Diff");


                DataTable dtCMX1 = new DataTable();
                dtCMX1.Columns.Add("Symbol");
                dtCMX1.Columns.Add("Type");
                dtCMX1.Columns.Add("Volume");
                dtCMX1.Columns.Add("Holding Volume");

                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                string destinationFilePath = Path.Combine(_folderPath, "");
                string FilePath1 = destinationFilePath + "\\" + fileName1;
                string FilePath2 = destinationFilePath + "\\" + fileName2;

                using (var package = new ExcelPackage(new FileInfo(FilePath1)))
                {
                    // Get the first worksheet in the workbook
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                    // Iterate through the rows and columns
                    for (int row = 3; row <= worksheet.Dimension.Rows - 1; row++)
                    {
                        dtCMX.Rows.Add();
                        for (int col = 1; col <= worksheet.Dimension.Columns; col++)
                        {
                            if (col < 5)
                            {
                                if (Convert.ToString(worksheet.Cells[row, col].Value) == "MCX" || Convert.ToString(worksheet.Cells[row, col].Value) == "STOCKS")
                                {
                                    dtCMX.Rows[dtCMX.Rows.Count - 1][col - 1] = Convert.ToString(worksheet.Cells[row, col].Value);
                                    row = row + 1;
                                    col = worksheet.Dimension.Columns + 1;
                                    continue;
                                }

                                var dl = Convert.ToString(worksheet.Cells[row, col].Value);
                                if (decimal.TryParse(dl, out decimal decimalValue))
                                {
                                    int intValue = Convert.ToInt32(decimalValue);
                                    dtCMX.Rows[dtCMX.Rows.Count - 1][col - 1] = intValue;
                                }
                                else
                                {
                                    dtCMX.Rows[dtCMX.Rows.Count - 1][col - 1] = worksheet.Cells[row, col].Value;
                                }
                            }

                        }
                    }
                }

                using (var package = new ExcelPackage(new FileInfo(FilePath2)))
                {
                    // Get the first worksheet in the workbook
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                    // Iterate through the rows and columns
                    for (int row = 3; row <= worksheet.Dimension.Rows - 1; row++)
                    {
                        dtCMX1.Rows.Add();
                        for (int col = 1; col <= worksheet.Dimension.Columns; col++)
                        {
                            if (col < 5)
                            {
                                if (Convert.ToString(worksheet.Cells[row, col].Value) == "MCX" || Convert.ToString(worksheet.Cells[row, col].Value) == "STOCKS")
                                {
                                    dtCMX.Rows[dtCMX.Rows.Count - 1][col - 1] = Convert.ToString(worksheet.Cells[row, col].Value);
                                    row = row + 1;
                                    col = worksheet.Dimension.Columns + 1;
                                    continue;
                                }
                                var dl = Convert.ToString(worksheet.Cells[row, col].Value);
                                //var dl = worksheet.Cells[row, col].Value.ToString();
                                if (decimal.TryParse(dl, out decimal decimalValue))
                                {
                                    int intValue = Convert.ToInt32(decimalValue);
                                    dtCMX1.Rows[dtCMX1.Rows.Count - 1][col - 1] = intValue;
                                }
                                else
                                {
                                    dtCMX1.Rows[dtCMX1.Rows.Count - 1][col - 1] = worksheet.Cells[row, col].Value;
                                }
                            }
                        }
                    }
                }

                for (int i = 0; i <= dtCMX.Rows.Count - 1; i++)
                {
                    dtCMX1.DefaultView.RowFilter = "Symbol='" + dtCMX.Rows[i][0] + "'";
                    if (dtCMX1.DefaultView.ToTable().Rows.Count > 0)
                    {
                        DataRow[] rows = dtCMX1.Select("Symbol='" + dtCMX.Rows[i][0] + "'");

                        // Update the specified column for each selected row
                        dtCMX.Rows[i][4] = rows[0].ItemArray[1];
                        dtCMX.Rows[i][5] = rows[0].ItemArray[2];
                        dtCMX.Rows[i][6] = rows[0].ItemArray[3];
                        dtCMX.Rows[i][8] = Convert.ToInt32(dtCMX.Rows[i][2]) - Convert.ToInt32(dtCMX.Rows[i][5]);
                        if (!string.IsNullOrEmpty(dtCMX.Rows[i][3].ToString()))
                        {
                            dtCMX.Rows[i][9] = Convert.ToInt32(dtCMX.Rows[i][3]) - Convert.ToInt32(dtCMX.Rows[i][6]);
                        }

                        if (Convert.ToString(dtCMX.Rows[i][8]) != "")
                        {
                            if (Convert.ToInt32(dtCMX.Rows[i][8]) < 0)
                            {
                                dtCMX.Rows[i][7] = "SELL";
                            }
                            else
                            {
                                dtCMX.Rows[i][7] = "BUY";
                            }
                        }

                        //foreach (DataRow row in rows)
                        //{
                        //    row[columnIndex + 1] = Convert.ToInt32(newValue) - Convert.ToInt32(row[columnIndex]);
                        //}
                    }
                    //UpdateDataTable(dtCMX, "Symbol='" + dtCMX.Rows[i][0] + "'", 3, dtCMX1.DefaultView.Table.Rows[0][2], "COMPOS");
                }
                return dtCMX;
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }
        private DataTable FileUpdateExcel(string fileName1, string fileName2)
        {
            try
            {
                DataTable dt = new DataTable();
                dt.Columns.Add("Master");
                dt.Columns.Add("Name");
                dt.Columns.Add("Per & Loss - 1");
                dt.Columns.Add("Per & Loss - 2");
                dt.Columns.Add("Difference in Per&Loss");

                DataTable dt1 = new DataTable();
                dt1.Columns.Add("Master");
                dt1.Columns.Add("Name");
                dt1.Columns.Add("Per & Loss");

                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                string destinationFilePath = Path.Combine(_folderPath, "");
                string FilePath1 = destinationFilePath + "\\" + fileName1;
                string FilePath2 = destinationFilePath + "\\" + fileName2;

                using (var package = new ExcelPackage(new FileInfo(FilePath1)))
                {
                    // Get the first worksheet in the workbook
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                    // Iterate through the rows and columns
                    for (int row = 3; row <= worksheet.Dimension.Rows - 1; row++)
                    {
                        dt.Rows.Add();
                        for (int col = 1; col <= worksheet.Dimension.Columns; col++)
                        {
                            var dl = worksheet.Cells[row, col].Value.ToString();
                            if (decimal.TryParse(dl, out decimal decimalValue))
                            {
                                int intValue = Convert.ToInt32(decimalValue);
                                dt.Rows[dt.Rows.Count - 1][col - 1] = intValue;
                            }
                            else
                            {
                                dt.Rows[dt.Rows.Count - 1][col - 1] = worksheet.Cells[row, col].Value;
                            }
                        }
                    }
                }

                using (var package = new ExcelPackage(new FileInfo(FilePath2)))
                {
                    // Get the first worksheet in the workbook
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                    // Iterate through the rows and columns
                    for (int row = 3; row <= worksheet.Dimension.Rows - 1; row++)
                    {
                        dt1.Rows.Add();
                        for (int col = 1; col <= worksheet.Dimension.Columns; col++)
                        {
                            var dl = worksheet.Cells[row, col].Value.ToString();
                            if (decimal.TryParse(dl, out decimal decimalValue))
                            {
                                int intValue = Convert.ToInt32(decimalValue);
                                dt1.Rows[dt1.Rows.Count - 1][col - 1] = intValue;
                            }
                            else
                            {
                                dt1.Rows[dt1.Rows.Count - 1][col - 1] = worksheet.Cells[row, col].Value;
                            }

                        }
                    }
                }

                for (int i = 0; i <= dt.Rows.Count - 1; i++)
                {
                    dt1.DefaultView.RowFilter = "Master='" + dt.Rows[i][0] + "'";

                    if (dt1.DefaultView.ToTable().Rows.Count > 0)
                    {
                        DataRow[] rows = dt1.Select("Master='" + dt.Rows[i][0] + "'");

                        // Update the specified column for each selected row
                        dt.Rows[i][3] = rows[0].ItemArray[2];
                        if (dt.Rows[i][2] == "") dt.Rows[i][2] = "0";
                        if (dt.Rows[i][3] == "") dt.Rows[i][3] = "0";
                        dt.Rows[i][4] = Convert.ToInt32(dt.Rows[i][3]) - Convert.ToInt32(dt.Rows[i][2]); ;
                    }
                    //UpdateDataTable(dt, "Master='" + dt.Rows[i][0] + "'", 2, dt1.DefaultView.Table.Rows[0][2]);
                }
                return dt;
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }
        static void UpdateDataTable(DataTable table, string filterExpression, int columnIndex, object newValue, string FileType = "")
        {
            // Select rows that match the filter expression
            DataRow[] rows = table.Select(filterExpression);

            // Update the specified column for each selected row
            foreach (DataRow row in rows)
            {
                if (row[columnIndex] == "") row[columnIndex] = "0";
                row[columnIndex + 1] = Convert.ToInt32(newValue) - Convert.ToInt32(row[columnIndex]);
                if (FileType == "COMPOS")
                {

                    if (Convert.ToString(row[1]) != "")
                    {
                        if (Convert.ToInt32(row[columnIndex + 1]) < 0)
                        {
                            row[columnIndex + 1] = "<b>" + row[columnIndex + 1] + "</b>";
                            row[1] = "SELL";
                        }
                        else
                        {
                            row[1] = "BUY";
                        }
                    }
                }
            }
        }
    }

}
