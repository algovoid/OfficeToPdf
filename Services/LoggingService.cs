using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Serilog;
using System.IO;

namespace OfficeToPdf.Services
{
	public static class LoggingService
	{
		public static void Initialize()
		{
			Directory.CreateDirectory("logs");
			Log.Logger = new LoggerConfiguration()
				.WriteTo.Console()
				.WriteTo.File("logs/log.txt", rollingInterval: RollingInterval.Day)
				.MinimumLevel.Debug()
				.CreateLogger();
		}

		public static void Shutdown()
		{
			Log.CloseAndFlush();
		}
	}
}

