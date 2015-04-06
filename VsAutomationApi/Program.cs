namespace VsAutomationApi
{
	using System;

	using EnvDTE;

	internal class Program
	{
		private readonly DTE _dte;

		public Program()
		{
			// Visual Studio 2013 -> VisualStudio.DTE.12.0
			_dte = (DTE)System.Runtime.InteropServices.Marshal.GetActiveObject("VisualStudio.DTE.12.0");
		}

		[STAThread]
		static void Main(string[] args)
		{
			Program prog = new Program();

			// Fixing 'Application is Busy' and 'Call was Rejected By Callee' Errors  
			// MSDN: http://msdn.microsoft.com/en-us/library/ms228772%28v=vs.80%29.aspx
			MessageFilter.Register();

			try
			{
				Console.WriteLine(prog._dte.Solution.FullName);
			}
			catch (Exception ex)
			{
				Console.WriteLine(ex.Message);
			}
			finally
			{
				MessageFilter.Revoke();
			}

			Console.ReadKey(true);
		}
	}
}
