using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TestReadExcel.Model;

namespace TestReadExcel
{
	class Program
	{
		static void Main(string[] args)
		{
			var apps = LoadDataApplications(@".\Config\data.xlsx");
			foreach (var item in apps)
			{
				Console.WriteLine("Code: {0}", item.Code);
				Console.WriteLine("Name: {0}\n", item.Name);
			}

			Console.ReadKey();
		}
		private static List<MApplication> LoadDataApplications(string configFile)
		{
			DataTable excelDT = SpreadSheet.GetSheetData(configFile, "APPLICATIONS");
			List<MApplication> apps = new List<MApplication>();

			foreach (DataRow row in excelDT.Rows)
			{
				try
				{
					if (!string.IsNullOrEmpty(row["CODE"].ToString()))
					{
						var app = new MApplication
						{
							Code = Convert.ToInt32(row["CODE"]),
							Name = row["NAME"].ToString()
						};
						apps.Add(app);
					}
				}
				catch (Exception ex)
				{
					throw ex;
				}
			}
			return apps;
		}
	}
}
