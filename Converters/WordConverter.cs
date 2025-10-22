using Microsoft.Office.Interop.Word;
using OfficeToPdf.Models;
using OfficeToPdf.Services;

namespace OfficeToPdf.Converters
{
	public class WordConverter : IOfficeConverter
	{
		private Application? _wordApp;

		public OfficeDocumentType Type => OfficeDocumentType.Word;

		public void Convert(string inputPath, string outputPath)
		{
			_wordApp ??= ApplicationSpawner.SpawnWord();
			object inputFile = inputPath;
			object outputFile = outputPath;
			object missing = System.Type.Missing;

			Document? doc = null;
			try
			{
				doc = _wordApp.Documents.Open(ref inputFile, ref missing, ReadOnly: true);
				doc.SaveAs2(outputFile, WdSaveFormat.wdFormatPDF);
			}
			finally
			{
				doc?.Close(false);
				ApplicationSpawner.ReleaseComObject(doc);
			}
		}

		public void Quit()
		{
			if (_wordApp != null)
			{
				ApplicationSpawner.QuitAndRelease(_wordApp, () => _wordApp.Quit(false));
				_wordApp = null;
			}
		}
	}
}

