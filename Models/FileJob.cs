using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeToPdf.Models
{
	public record FileJob(string InputPath, string OutputPath, int Attempts = 0);
}
