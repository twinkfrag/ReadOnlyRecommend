using System;
using System.IO;
using System.IO.Compression;
using System.Xml;
using System.Xml.XPath;

namespace ReadOnlyRecommend
{
	public class Program
	{
		public static void Main(string[] args)
		{
			bool isYes = false;
			string inputfile = null;

			foreach (var arg in args)
			{
				if (arg.ToLower().Trim('-', '/') == "y")
				{
					isYes = true;
				}
				else
				{
					inputfile = arg;
				}
			}

			inputfile = string.IsNullOrEmpty(inputfile) ? Console.ReadLine().Trim('"') : inputfile;

			var path = new FileInfo(inputfile);
			if (!path.Exists)
			{
				path = new FileInfo(Path.Combine(Environment.GetEnvironmentVariable("TEMP"), inputfile));
			}
			if (!path.Exists)
			{
				Console.WriteLine($"File {inputfile} is not found.");
				return;
			}
			if (path.Length == 0L)
			{
				Console.WriteLine($"File {inputfile} is empty.");
				return;
			}

			switch (path.Extension)
			{
				case ".docx":
					Console.Write($"Edit Word File: {path.FullName}");
					if (isYes)
					{
						Console.WriteLine(" .");
					}
					else
					{
						Console.WriteLine(" ? [Y/n]");
						isYes = Console.ReadLine() != "n";
					}
					if (isYes)
					{
						WordOverwrite(path);
					}
					break;
				case ".xlsx":
					Console.Write($"Edit Excel File: {path.FullName}");
					if (isYes)
					{
						Console.WriteLine(" .");
					}
					else
					{
						Console.WriteLine(" ? [Y/n]");
						isYes = Console.ReadLine() != "n";
					}
					if (isYes)
					{
						ExcelOverwrite(path);
					}
					break;
				case ".pptx":
					Console.Write($"Edit PowerPoint File: {path.FullName}");
					if (isYes)
					{
						Console.WriteLine(" .");
					}
					else
					{
						Console.WriteLine(" ? [Y/n]");
						isYes = Console.ReadLine() != "n";
					}
					if (isYes)
					{
						PowerPointOverwrite(path);
					}
					break;
				default:
					break;
			}
		}

		static void WordOverwrite(FileInfo path)
		{
			using (var archive = ZipFile.Open(path.FullName, ZipArchiveMode.Update))
			{
				var entry = archive.GetEntry(@"word/settings.xml");
				using (var st = entry.Open())
				{
					var (xml, navi) = LoadXmlNavigator(st);

					navi.MoveToFollowing(XPathNodeType.Element); //w:settings
					var ns_w = navi.GetNamespace("w");

					if (navi.MoveToFollowing("writeProtection", ns_w))
					{
						Console.WriteLine("Already ReadOnlyRecommended");
						return;
					}
					else
					{
						// <w:writeProtection>
						// is first child of <w:settings>
						navi.PrependChild("<w:writeProtection w:recommended=\"1\"/>");

						OverwriteXml(st, xml);
						Console.WriteLine("Set ReadOnlyRecommended");
					}
				}
			}
		}

		static void ExcelOverwrite(FileInfo path)
		{
			using (var archive = ZipFile.Open(path.FullName, ZipArchiveMode.Update))
			{
				var entry = archive.GetEntry(@"xl/workbook.xml");
				using (var st = entry.Open())
				{
					var (xml, navi) = LoadXmlNavigator(st);

					navi.MoveToFollowing(XPathNodeType.Element); //workbook
					navi.MoveToChild("fileVersion", navi.NamespaceURI);

					if (navi.MoveToFollowing("fileSharing", navi.NamespaceURI))
					{
						Console.WriteLine("Already ReadOnlyRecommended");
						return;
					}
					else
					{
						// <fileSharing>
						// is next of <fileVersion>
						navi.InsertAfter("<fileSharing readOnlyRecommended=\"1\"/>");

						OverwriteXml(st, xml);
						Console.WriteLine("Set ReadOnlyRecommended");
					}
				}
			}
		}


		static void PowerPointOverwrite(FileInfo path)
		{
			using (var archive = ZipFile.Open(path.FullName, ZipArchiveMode.Update))
			{
				var entry = archive.GetEntry(@"ppt/presProps.xml");
				using (var st = entry.Open())
				{
					var (xml, navi) = LoadXmlNavigator(st);

					navi.MoveToFollowing(XPathNodeType.Element); //p:presentationPr
					var ns_p = navi.GetNamespace("p");

					navi.MoveToChild("extLst", ns_p);

					navi.MoveToChild("ext", ns_p);
					do
					{
						navi.MoveToFirstAttribute();
						if (navi.Value == "{1BD7E111-0CB8-44D6-8891-C1BB2F81B7CC}")
						{
							Console.WriteLine("Already ReadOnlyRecommended");
							return;
						}
					}
					while (navi.MoveToFollowing("ext", ns_p));

					// <p:ext><p1710:readonlyRecommended /></p:ext>
					// is last child of <p:extLst>
					navi.MoveToParent();
					navi.InsertAfter("<p:ext uri=\"{1BD7E111-0CB8-44D6-8891-C1BB2F81B7CC}\"/>");
					navi.MoveToNext();
					navi.AppendChild("<p1710:readonlyRecommended xmlns:p1710=\"http://schemas.microsoft.com/office/powerpoint/2017/10/main\" val=\"1\"/>");

					OverwriteXml(st, xml);
					Console.WriteLine("Set ReadOnlyRecommended");
				}
			}
		}

		static (XmlDocument, XPathNavigator) LoadXmlNavigator(Stream st)
		{
			var xml = new XmlDocument
			{
				PreserveWhitespace = true
			};
			xml.Load(st);
			var navi = xml.CreateNavigator();
			return (xml, navi);
		}

		static void OverwriteXml(Stream stream, XmlDocument xml)
		{
			stream.Seek(0, SeekOrigin.End);
			stream.SetLength(0);
			using (var writer = new StreamWriter(stream))
			{
				xml.Save(writer);
			}
		}
	}
}
