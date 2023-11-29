//---------------------------------------------------------------------------
/** @file Program.cs
 * xlsx から、UTF-8のcsvに変換するコンバータ
 *
 * $Id: Program.cs 2023-11-29 18:25:00Z maru $
 *
 *	@author $Author: maru $
 *  @date $Date:: 2023-11-19 18:10:27 +0900#$
 *  @version $Revision: n $
 *
 **/
//---------------------------------------------------------------------------
using System;
using System.Text;
using System.IO;
using System.Collections;
using Debug = System.Diagnostics.Debug;
using System.Globalization;

namespace xlsx2csv
{
	internal class Program
	{
		static void Main(string[] args)
		{
			Console.WriteLine("xlsx2csv by Pirota");
			string ext = null;
			string infile = null;   // 入力ファイル名
			string outfile = null;   // 出力ファイル名

			// オプション解析
			for (int i = 0; i < args.Length; i++)
			{
				char ch = args[i][0];
				// スイッチ判定
				if ((ch == '-' || ch == '/') && args[i].Length >= 2)
				{
					char opt = args[i][1];

					switch (opt)
					{
						case '?':
							usage();
							break;

						default:
							Console.WriteLine("Missing option '{0}'.", opt);
							Environment.Exit(1);
							break;
					}

				}
				else
				{
					// 入力ファイル名
					if (infile == null)
					{
						infile = new string(args[i]);
					}
					else
					{
						// 出力ファイル
						if (outfile == null)
						{
							outfile = new string(args[i]);
						}
					}

				}
			}
			// 使用法
			if (args.Length == 0 || infile == null)
			{
				usage();

			}

			// 出力ファイル名が未指定だったら、入力からコピー
			if (outfile == null)
			{
				outfile = Path.GetFileNameWithoutExtension(infile);
			}

			// 拡張子チェック
			ext = Path.GetExtension(infile);
			if (ext == "")
			{
				infile += ".xlsx";
			}
			ext = Path.GetExtension(outfile);
			if (ext == "")
			{
				outfile += ".csv";
			}

			//
#if DEBUG
			Debug.WriteLine("infile={0} outfile={1}", infile, outfile);
#endif

#if NETCOREAPP
			// コードページ エンコーディング プロバイダーを登録
			Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
#endif

			// 変換開始
			CConvert cmp = new CConvert();

			cmp.DoConvert(infile, outfile);

			cmp = null;
			GC.Collect();

		}

		/// <summary>
		/// 使用法の表示
		/// </summary>
		private static void usage()
		{
			Console.WriteLine("usage: xlsx2csv [infile].xlsx [outfile].csv");

			Environment.Exit(0);
		}

	}
}