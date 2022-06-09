// See https://aka.ms/new-console-template for more information
using GrapeCity.Documents.Excel;

Console.WriteLine("Excelワークブックからデータをインポート");

// ---------- 従来の方法 ----------
var sw1 = new System.Diagnostics.Stopwatch();
var workbook1 = new Workbook();

sw1.Start();

// Excelファイルを読み込む
workbook1.Open("importdata-table.xlsx");

// ワークシート追加
var worksheet1 = workbook1.Worksheets.Add();

// 新規シートにデータをコピー
workbook1.Worksheets[0].Range[0, 0, 21, 7].Copy(worksheet1.Range[0, 0, 21, 7]);

// 既存シートを削除
workbook1.Worksheets[0].Delete();

sw1.Stop();

// 結果の表示
Console.WriteLine("処理時間（Open、Add、Copy、Delete）");
Console.WriteLine($"　{sw1.ElapsedMilliseconds}ミリ秒");

workbook1.Save("result1.xlsx");

// ---------- ImportDataメソッドを使う方法（V5J） ----------
var sw2 = new System.Diagnostics.Stopwatch();
var workbook2 = new Workbook();

sw2.Start();

// Excelファイルからデータを読み込む
var data1 = Workbook.ImportData("importdata-table.xlsx", "売り上げデータ", 0, 0, 21, 7);

// シートにデータをコピー
workbook2.Worksheets[0].Range[0, 0, 21, 7].Value = data1;

sw2.Stop();

// 結果の表示
Console.WriteLine("処理時間（ImportData、V5J）");
Console.WriteLine($"　{sw2.ElapsedMilliseconds}ミリ秒");

workbook2.Save("result2.xlsx");

// ---------- ImportDataメソッドを使う方法（V5J SP1） ----------
var sw3 = new System.Diagnostics.Stopwatch();
var workbook3 = new Workbook();

sw3.Start();

// Excelファイルからデータを読み込む
var data2 = Workbook.ImportData("importdata-table.xlsx", "売り上げデータ!Sales2022Q1");

// シートにデータをコピー
workbook3.Worksheets[0].Range[0, 0, data2.GetLength(0), data2.GetLength(1)].Value = data2;


sw3.Stop();

// 結果の表示
Console.WriteLine("処理時間（ImportData、V5J SP1）");
Console.WriteLine($"　{sw2.ElapsedMilliseconds}ミリ秒");

workbook3.Save("result3.xlsx");




