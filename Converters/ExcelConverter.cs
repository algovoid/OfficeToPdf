using Microsoft.Office.Interop.Excel;
using OfficeToPdf.Models;
using OfficeToPdf.Services;

namespace OfficeToPdf.Converters
{
	public class ExcelConverter : IOfficeConverter
	{
		private Microsoft.Office.Interop.Excel.Application? _excelApp;

		public OfficeDocumentType Type => OfficeDocumentType.Excel;

		public void Convert(string inputPath, string outputPath)
		{
			_excelApp ??= ApplicationSpawner.SpawnExcel();
			Workbook? workbook = null;
			try
			{
				workbook = _excelApp.Workbooks.Open(inputPath, ReadOnly: true);
				workbook.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, outputPath);
			}
			finally
			{
				workbook?.Close(false);
				ApplicationSpawner.ReleaseComObject(workbook);
			}
		}

		public void Quit()
		{
			if (_excelApp != null)
			{
				ApplicationSpawner.QuitAndRelease(_excelApp, () => _excelApp.Quit());
				_excelApp = null;
			}
		}
	}
}
