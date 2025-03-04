using System;
using System.IO;

class Program
{
    static void Main(string[] args)
    {
        // 引数が指定されているかチェック
        if (args.Length == 0)
        {
            Console.WriteLine("使用方法: hello2.exe <出力ファイルパス>");
            return;
        }

        string filePath = args[0];

        try
        {
            // "Hello World" をファイルに書き込む
            File.AppendAllText(filePath, "Hello World" + Environment.NewLine);
            Console.WriteLine("Hello World を" + filePath + " に書き込みました。");
        }
        catch (Exception ex)
        {
            Console.WriteLine("エラー: " + ex.Message);
        }
    }
}
