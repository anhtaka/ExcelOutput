using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using ClosedXML.Excel;
using System.IO;
namespace ExcelOutput
{
	public partial class Form1 : Form
	{
		public Form1()
		{
			InitializeComponent();
		}

		private void button1_Click(object sender, EventArgs e)
		{
			//Excelのパス
			string fileName = @"C:\c_sharp\test.xlsx";
			Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

			//Excelが開かないようにする
			xlApp.Visible = false;

			//指定したパスのExcelを起動
			Workbook wb = xlApp.Workbooks.Open(Filename: fileName);

			try
			{
				//Sheetを指定
				((Worksheet)wb.Sheets[1]).Select();
			}
			catch (Exception ex)
			{
				//Sheetがなかった場合のエラー処理

				//Appを閉じる
				wb.Close(false);
				xlApp.Quit();

				//Errorメッセージ
				Console.WriteLine("指定したSheetは存在しません．");
				Console.ReadLine();

				//実行を終了
				System.Environment.Exit(0);
			}

			//変数宣言
			Range CellRange;

			for (int i = 1; i <= 5; i++)
			{
				//書き込む場所を指定
				CellRange = xlApp.Cells[i, 1] as Range;

				//書き込む内容
				CellRange.Value2 = "繰り返し" + i + "回目";
			}

			//Appを閉じる
			wb.Close(true);
			xlApp.Quit();

		}

		private void button2_Click(object sender, EventArgs e)
		{
			excel_OutPutEx();
		}
		private void excel_OutPutEx()
		{

			//Excelオブジェクトの初期化
			Excel.Application ExcelApp = null;
			Excel.Workbooks wbs = null;
			Excel.Workbook wb = null;
			Excel.Sheets shs = null;
			Excel.Worksheet ws = null;

			try
			{
			//Excelシートのインスタンスを作る
			ExcelApp = new Excel.Application();
			wbs = ExcelApp.Workbooks;
			wb = wbs.Add();

			shs = wb.Sheets;
			ws = shs[1];
			ws.Select(Type.Missing);

			ExcelApp.Visible = false;

			// エクセルファイルにデータをセットする
			for ( int i = 1; i < 10; i++ )
			{
			// Excelのcell指定
			Excel.Range w_rgn = ws.Cells;
			Excel.Range rgn = w_rgn[i, 1];

			try
			{
			// Excelにデータをセット
			rgn.Value2 = "hoge";
			}
			finally
			{
			// Excelのオブジェクトはループごとに開放する
			Marshal.ReleaseComObject(w_rgn);
			Marshal.ReleaseComObject(rgn);
			w_rgn = null;
			rgn = null;
			}
			}

			//excelファイルの保存
			wb.SaveAs(@"C:\c_sharp\test1.xlsx");
			wb.Close(false);
			ExcelApp.Quit();
			}
			finally
			{
				//Excelのオブジェクトを開放し忘れているとプロセスが落ちないため注意
				Marshal.ReleaseComObject(ws);
				Marshal.ReleaseComObject(shs);
				//Marshal.RelesaeComObject(wb);
				Marshal.ReleaseComObject(wbs);
				Marshal.ReleaseComObject(ExcelApp);
				ws = null;
				shs = null;
				wb = null;
				wbs = null;
				ExcelApp = null;

				GC.Collect();
			}

		}

		private void button3_Click(object sender, EventArgs e)
		{
			var workbook = new XLWorkbook();
			var worksheet = workbook.Worksheets.Add("Sample Sheet");
			worksheet.Cell("A1").Value = "Hello World!";
			workbook.SaveAs("HelloWorld.xlsx");
		}
	}
}
