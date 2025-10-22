namespace OfficeToPdf.Services
{
	public static class ThreadFactory
	{
		public static Thread CreateSTAThread(ThreadStart action, string name)
		{
			var thread = new Thread(action)
			{
				IsBackground = true,
				Name = name
			};
			thread.SetApartmentState(ApartmentState.STA);
			return thread;
		}
	}
}

