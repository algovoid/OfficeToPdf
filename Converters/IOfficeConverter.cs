using OfficeToPdf.Models;

namespace OfficeToPdf.Converters
{
	public interface IOfficeConverter
	{
		OfficeDocumentType Type { get; }
		void Convert(string inputPath, string outputPath);
	}
}
