using ExcelTest.Model;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTest
{
	public class Program
	{
		static void Main(string[] args)
		{
			if (args.Count() == 0 )
			{
				Console.WriteLine("ExcelTest.exe XlsxFilePathA.xls XlsxFilePathB.xls");
				return;
			}

			var argsFiles = args;
			foreach (var record in argsFiles)
			{
				Console.WriteLine("\r\n");
				var XlsFilePath = record;

				var xlsxsheetmodel = new XlsxSheetModel();

				///bookよみこみ
				var book = WorkbookFactory.Create(XlsFilePath);
				xlsxsheetmodel.SheetName = "SetPrint";

				List<string> sheetNameList = new List<string>();
				for (int i = 0; i < book.NumberOfSheets; i++)
				{
					// シートオブジェクトを取得
					ISheet sht = book.GetSheetAt(i);
					// シート名を取得して、リストに保存
					sheetNameList.Add(sht.SheetName);
					var isHidden = book.IsSheetHidden(i);
					Console.WriteLine($"シート名:{sht.SheetName} 隠れているか：{isHidden.ToString()} index:{i}");
				}
				Console.WriteLine(string.Join(",", sheetNameList));
				var sheet = book.GetSheet(xlsxsheetmodel.SheetName);
				
				//sheet.ForceFormulaRecalculation = true;

				//印刷させたいシートを設定
				xlsxsheetmodel.SheetIndex = book.GetSheetIndex(xlsxsheetmodel.SheetName);

				Console.WriteLine($"index:{xlsxsheetmodel.SheetIndex} XlsxFileName{XlsFilePath} ");
			}
			
			return;
		}
	}
}
