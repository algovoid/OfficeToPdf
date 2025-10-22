using Serilog;
using OfficeToPdf.Models;
using OfficeToPdf.Services;
using OfficeToPdf.Workers;

namespace OfficeToPdf
{
	class Program
	{
		static int Main(string[] args)
		{
			if (args.Length < 2)
			{
				Console.WriteLine("Usage: OfficeToPdf <input-folder> <output-folder> [maxWorkers]");
				return 1;
			}

			LoggingService.Initialize();

			var input = args[0];
			var output = args[1];
			int maxWorkers = args.Length > 2 ? int.Parse(args[2]) : Math.Max(1, Environment.ProcessorCount / 2);

			var queue = new JobQueue();
			var files = Directory.EnumerateFiles(input, "*.*", SearchOption.TopDirectoryOnly)
				.Where(f => new[] { ".doc", ".docx", ".xls", ".xlsx", ".xlsm" }
				.Contains(Path.GetExtension(f).ToLowerInvariant()));

			foreach (var file in files)
			{
				var outPath = Path.Combine(output, Path.GetFileNameWithoutExtension(file) + ".pdf");
				queue.Enqueue(new FileJob(file, outPath));
			}

			Log.Information("Queued {Count} files. Starting {Workers} workers...", files.Count(), maxWorkers);

			using var cts = new CancellationTokenSource();
			var workers = Enumerable.Range(1, maxWorkers)
				.Select(i => new OfficeWorker(queue, cts.Token, i, 2, 60))
				.ToList();

			foreach (var w in workers) w.Start();

			while (!queue.IsEmpty)
			{
				Thread.Sleep(1000);
				if (Console.KeyAvailable && Console.ReadKey(true).Key == ConsoleKey.Q)
				{
					Log.Warning("Manual cancel received.");
					cts.Cancel();
					break;
				}
			}

			workers.ForEach(w => w.StopAndWait());
			LoggingService.Shutdown();
			return 0;
		}
	}
}
