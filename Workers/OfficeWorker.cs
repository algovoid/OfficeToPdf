using Serilog;
using OfficeToPdf.Models;
using OfficeToPdf.Converters;
using OfficeToPdf.Services;

namespace OfficeToPdf.Workers
{
	public class OfficeWorker
	{
		private readonly JobQueue _queue;
		private readonly CancellationToken _token;
		private readonly int _id;
		private readonly int _maxRetries;
		private readonly int _timeoutSeconds;

		private readonly IOfficeConverter _word = new WordConverter();
		private readonly IOfficeConverter _excel = new ExcelConverter();
		//private readonly IOfficeConverter _ppt = new PowerPointConverter();

		private Thread? _thread;
		private volatile bool _running = false;
		private readonly ManualResetEventSlim _stopped = new(false);

		public OfficeWorker(JobQueue queue, CancellationToken token, int id, int maxRetries, int timeoutSeconds)
		{
			_queue = queue;
			_token = token;
			_id = id;
			_maxRetries = maxRetries;
			_timeoutSeconds = timeoutSeconds;
		}

		public void Start()
		{
			_thread = ThreadFactory.CreateSTAThread(Run, $"Worker-{_id}");
			_running = true;
			_thread.Start();
		}

		public void StopAndWait()
		{
			_running = false;
			_stopped.Wait();
		}

		private void Run()
		{
			try
			{
				while (_running && !_token.IsCancellationRequested)
				{
					if (!_queue.TryDequeue(out var job)) continue;

					Log.Information("[Worker {Id}] Processing {File}", _id, job.InputPath);

					bool success = false;
					for (int attempt = 0; attempt <= _maxRetries && !success; attempt++)
					{
						try
						{
							var type = DetectType(job.InputPath);
							var converter = GetConverter(type);

							if (converter == null)
							{
								Log.Warning("[Worker {Id}] Unsupported file: {File}", _id, job.InputPath);
								break;
							}

							var task = Task.Run(() => converter.Convert(job.InputPath, job.OutputPath));
							if (task.Wait(_timeoutSeconds * 1000))
							{
								success = true;
								Log.Information("[Worker {Id}] Success: {File}", _id, job.InputPath);
							}
							else
							{
								Log.Warning("[Worker {Id}] Timeout: {File}", _id, job.InputPath);
								RestartConverters();
							}
						}
						catch (Exception ex)
						{
							Log.Error(ex, "[Worker {Id}] Error converting {File}", _id, job.InputPath);
							RestartConverters();
							Thread.Sleep(2000);
						}
					}
				}
			}
			finally
			{
				RestartConverters();
				_stopped.Set();
				Log.Information("[Worker {Id}] Stopped.", _id);
			}
		}

		private IOfficeConverter? GetConverter(OfficeDocumentType type) =>
			type switch
			{
				OfficeDocumentType.Word => _word,
				OfficeDocumentType.Excel => _excel,
				//OfficeDocumentType.PowerPoint => _ppt,
				_ => null
			};

		private static OfficeDocumentType DetectType(string path)
		{
			var ext = System.IO.Path.GetExtension(path).ToLowerInvariant();
			return ext switch
			{
				".doc" or ".docx" => OfficeDocumentType.Word,
				".xls" or ".xlsx" or ".xlsm" => OfficeDocumentType.Excel,
				".ppt" or ".pptx" => OfficeDocumentType.PowerPoint,
				_ => OfficeDocumentType.Unknown
			};
		}

		private void RestartConverters()
		{
			(_word as WordConverter)?.Quit();
			(_excel as ExcelConverter)?.Quit();
			//(_ppt as PowerPointConverter)?.Quit();
		}
	}
}
