using System;
using System.Text;
using System.Collections.Generic;
using System.IO.Compression;
using System.Collections;
using System.Linq;
using System.Threading.Tasks;
using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HPSF;
using NPOI.OpenXmlFormats.Wordprocessing;
using System.Runtime.Serialization;
using Debug = System.Diagnostics.Debug;
using System.Diagnostics.Metrics;
using static NPOI.HSSF.Util.HSSFColor;


namespace xlsx2csv
{
	internal class CConvert
	{
		/// <summary>
		/// セルのint値を取得する
		/// </summary>
		int GetCellInt(ICell cell)
		{
			int ret = 0;
			if (cell != null)
			{
				int.TryParse(cell.ToString(), out ret);
			}
			return ret;
		}

		/// <summary>
		/// セルのstring値を取得する
		/// </summary>
		string GetCellString(ICell cell)
		{
			string ret = "";
			if (cell != null)
			{
				ret = cell.ToString().Replace(",", "%2C");
			}
			return ret;
		}

		/// <summary>
		/// 32bitのエンディアンを変換する
		/// </summary>
		int SwapEndian32(int value)
		{
			byte[] bytes = BitConverter.GetBytes(value);

			if (BitConverter.IsLittleEndian)
			{
				System.Array.Reverse(bytes);
			}

			return BitConverter.ToInt32(bytes, 0);
		}

		/// <summary>
		/// 16bitのエンディアンを変換する
		/// </summary>
		short SwapEndian16(short value)
		{
			byte[] bytes = BitConverter.GetBytes(value);

			if (BitConverter.IsLittleEndian)
			{
				System.Array.Reverse(bytes);
			}

			return BitConverter.ToInt16(bytes, 0);
		}

		/// <summary>
		/// パスセパレータをプラットフォーム対応に置換する
		/// </summary>
		/// <param name="path">パス</param>
		/// <returns>置換後のパス</returns>
		private string ReplacePathSeparator(string path)
		{
			char ch = Path.DirectorySeparatorChar;
			if (ch == '/')
			{
				return path.Replace('\\', ch);
			}
			else
			if (ch == '\\')
			{
				return path.Replace('/', ch);
			}

			return path;
		}

		/// <summary>
		/// 変換処理
		/// </summary>
		/// <param name="infile">入力ファイル</param>
		/// <param name="outfile">出力ファイル</param>
		internal void DoConvert(string infile, string outfile)
		{
			int columnCount = 0;
			List<string> csvlist = new List<string>();

			infile = ReplacePathSeparator(infile);
			outfile = ReplacePathSeparator(outfile);

			// 入力ファイル存在チェック
			if (!File.Exists(infile))
			{
				Console.WriteLine($"'{infile}' not found.");
				return;
			}

			// ファイルを開く
			// 他のプロセスでファイルを開いているときでも、読み取り可能にする
			using (FileStream file = new FileStream(infile, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				IWorkbook workbook = new XSSFWorkbook(file); // ワークブックを読み込む
				ISheet sheet = workbook.GetSheetAt(0); // 最初のシートを取得

				// 行と列をループしてセルの値を取得
				for (int i = sheet.FirstRowNum; i <= sheet.LastRowNum; i++)
				{
					IRow row = sheet.GetRow(i);
					if (i == sheet.FirstRowNum)
					{
						// ヘッダー行を読み飛ばす
						// カラム行数の取得
						columnCount = row.LastCellNum - row.FirstCellNum;
						Debug.WriteLine($"colomnCount={columnCount}");
						continue;
					}

					if (row != null)
					{
						string csv = String.Empty;
						for (int j=0; j<columnCount; j++)
						{
							string tmp = GetCellString(row.GetCell(j));
							csv += tmp;
							if (j < (columnCount - 1))
							{
								csv += ",";
							}
						}
						if (csv == String.Empty) { break; } // IDが未指定であれば、終端マークとみなし終了
						Debug.WriteLine(csv);
						csvlist.Add(csv);
					}
				} // 行と列をループしてセルの値を取得の終わり
			} // usingの終わり

			// UTF-8 エンコーディングでファイルに書き込み
			using (StreamWriter writer = new StreamWriter(outfile, false, Encoding.UTF8))
			{
				foreach (var line in csvlist)
				{
					writer.WriteLine(line);
				}
			}
			Console.WriteLine($"'{outfile}' wrote.");
		}
	}
}
