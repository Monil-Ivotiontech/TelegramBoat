using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using DinkToPdf;
using DinkToPdf.Contracts;
using System.Data;
using static System.Runtime.InteropServices.JavaScript.JSType;
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
		public IActionResult GetFile([FromQuery] string parameter1, [FromQuery] int parameter2)
		{
			var folderPath = Path.Combine(_folderPath, "");
			if (!Directory.Exists(folderPath))
			{
				return NotFound("Folder not found.");
			}

			// Get all files in the folder
			var files = Directory.GetFiles(folderPath);

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
			string formattedDateFrom = DateTime.Today.ToString("MM/dd/yyyy HH:mm:ss", CultureInfo.InvariantCulture);
			view.RowFilter = $"CreatedDateTime >= #{formattedDateFrom}#";

			DataTable sortedDate = view.ToTable();

			if (parameter1.Trim() == "11000")
			{
				dt.DefaultView.RowFilter = "FileName like '%11000%'";
				dt = dt.DefaultView.ToTable();

				if (dt.Rows.Count > 1)
				{
					string FileName1 = dt.Rows[0][0].ToString();
					string FileName2 = dt.Rows[1][0].ToString();
					dt = new DataTable();
					dt = File11000(FileName1, FileName2);

				}

			}
			else if (parameter1.Trim() == "UPDATE EXCEL")
			{
				dt.DefaultView.RowFilter = "FileName like '%UPDATE EXCEL%'";
				dt = dt.DefaultView.ToTable();

				if (dt.Rows.Count > 1)
				{
					string FileName1 = dt.Rows[0][0].ToString();
					string FileName2 = dt.Rows[1][0].ToString();
					dt = new DataTable();
					dt = FileUpdateExcel(FileName1, FileName2);

				}

			}



			// Process parameters and retrieve the file
			var pdfContent = "";
			pdfContent += ConvertDataTableToHtml(dt);



			// Convert the accumulated content to PDF
			var pdfBytes = ConvertHtmlToPdf(pdfContent);
			parameter1 = parameter1.Replace(" ", "") + "_" + DateTime.Now.ToString("ddMMyyyyHHmmss");
			// Return PDF as bytes
			return File(pdfBytes, "application/pdf", $"{parameter1}.pdf");
		}
		private DataTable LoadXlsxIntoDataTable(string xlsxFilePath)
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

		private string ConvertDataTableToHtml(DataTable dataTable)
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

		private byte[] ConvertHtmlToPdf(string htmlContent)
		{
			var globalSettings = new GlobalSettings
			{
				ColorMode = ColorMode.Color,
				PaperSize = PaperKind.A4,
				Orientation = Orientation.Portrait,
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


		private DataTable File11000(string fileName1, string fileName2)
		{
			DataTable dt = new DataTable();
			dt.Columns.Add("Login");
			dt.Columns.Add("Name");
			dt.Columns.Add("Per(%)");
			dt.Columns.Add("Credit");
			dt.Columns.Add("Credit &");
			dt.Columns.Add("M2M");
			dt.Columns.Add("Partner");
			dt.Columns.Add("Net Amount");
			dt.Columns.Add("Difference in Net Amount");

			DataTable dt1 = new DataTable();
			dt1.Columns.Add("Login");
			dt1.Columns.Add("Name");
			dt1.Columns.Add("Per(%)");
			dt1.Columns.Add("Credit");
			dt1.Columns.Add("Credit &");
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
				dt1.DefaultView.RowFilter = "Login=" + dt.Rows[i][0];
				UpdateDataTable(dt, "Login='" + Convert.ToInt64(dt.Rows[i][0]) + "'", 7, dt1.DefaultView.Table.Rows[0][7]);
			}
			return dt;

		}

		private DataTable FileUpdateExcel(string fileName1, string fileName2)
		{
			DataTable dt = new DataTable();
			dt.Columns.Add("Master");
			dt.Columns.Add("Name");
			dt.Columns.Add("Per & Loss");
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
				UpdateDataTable(dt, "Master='" + dt.Rows[i][0] + "'", 2, dt1.DefaultView.Table.Rows[0][2]);
			}
			return dt;
		}
		static void UpdateDataTable(DataTable table, string filterExpression, int columnIndex, object newValue)
		{
			// Select rows that match the filter expression
			DataRow[] rows = table.Select(filterExpression);

			// Update the specified column for each selected row
			foreach (DataRow row in rows)
			{
				row[columnIndex + 1] = Convert.ToInt32(newValue) - Convert.ToInt32(row[columnIndex]);
			}
		}
	}

}
