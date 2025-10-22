using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Collections.Concurrent;
using OfficeToPdf.Models;

namespace OfficeToPdf.Workers
{
	public class JobQueue
	{
		private readonly BlockingCollection<FileJob> _queue = new();

		public void Enqueue(FileJob job) => _queue.Add(job);
		public bool TryDequeue(out FileJob job) => _queue.TryTake(out job, 500);
		public bool IsEmpty => _queue.Count == 0;
	}
}
