using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Permissions;
using System.Text;
using System.Threading.Tasks;

namespace ReadOnlyRecommendWatcher
{
	class Program
	{
		static void Main(string[] args) => Run(args);

		[PermissionSet(SecurityAction.Demand, Name = "FullTrust")]
		static void Run(string[] args)
		{
			var input = new DirectoryInfo(args.Length > 0 && !string.IsNullOrEmpty(args[0]) ? args[0] : ".");
			if (!input.Exists)
			{
				Console.WriteLine("Directory not found");
				return;
			}
			Console.WriteLine($"Watch {input.FullName}");

			var fsw = new FileSystemWatcher
			{
				Path = input.FullName,
				EnableRaisingEvents = true,
			};

			void handler(object _, FileSystemEventArgs e)
			{
				if (e.ChangeType == WatcherChangeTypes.Deleted)
				{
					return;
				}
				var file = new FileInfo(e.FullPath);
				if (new[] { ".docx", ".xlsx", ".pptx" }.Contains(file.Extension) && !file.Name.StartsWith("~$"))
				{
					ReadOnlyRecommend.Program.Main(new[] { file.FullName, "/y" });
				}
				Console.WriteLine(e.FullPath + " Created");
			}
			fsw.Created += handler;
			fsw.Changed += handler;
			

			while (true)
			{
				Console.ReadLine();
			}
		}
	}
}
