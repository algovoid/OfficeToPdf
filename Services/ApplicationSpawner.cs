using System.Runtime.InteropServices;

using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Action = System.Action;

namespace OfficeToPdf.Services
{
	public static class ApplicationSpawner
	{
		public static Word.Application SpawnWord()
		{
			var app = new Word.Application { Visible = false };
			return app;
		}

		public static Excel.Application SpawnExcel()
		{
			var app = new Excel.Application
			{
				Visible = false,
				DisplayAlerts = false
			};
			return app;
		}

		public static PowerPoint.Application SpawnPowerPoint()
		{
			var app = new PowerPoint.Application();
			return app;
		}

		public static void ReleaseComObject(object? obj)
		{
			if (obj == null) return;
			try { Marshal.FinalReleaseComObject(obj); }
			catch { }
		}

		public static void QuitAndRelease(object? app, Action? quitAction = null)
		{
			try { quitAction?.Invoke(); }
			catch { }
			ReleaseComObject(app);
			GC.Collect();
			GC.WaitForPendingFinalizers();
		}
	}
}

