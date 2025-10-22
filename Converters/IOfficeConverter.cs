using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using OfficeToPdf.Models;

namespace OfficeToPdf.Converters
{
	public interface IOfficeConverter
	{
		OfficeDocumentType Type { get; }
		void Convert(string inputPath, string outputPath);
	}
}
