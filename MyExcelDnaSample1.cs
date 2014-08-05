// --------------------------------------------------------------------------------------------------------------------
// <copyright file="MyExcelDnaSample1.cs" company="Tadahiro Ishisaka">
//   Copyright 2014 Tadahiro Ishisaka
// </copyright>
// <summary>
//   Excel DNAを使用してExcelのユーザー定義関数(UDF)を作成するサンプル
//   UDFはstaticなクラスにstaticなメソッドとして記述する。
// </summary>
// --------------------------------------------------------------------------------------------------------------------
namespace MyExcelDnaSample1
{
    using ExcelDna.Integration;

    /// <summary>
    ///     Excel DNAを使用してExcelのユーザー定義関数(UDF)を作成するサンプル
    ///     UDFはstaticなクラスにstaticなメソッドとして記述する。
    /// </summary>
    public static class MyExcelDnaSample1
    {
        /// <summary>引数の名前にHelloを言う</summary>
        /// <param name="name">あいさつの相手</param>
        /// <returns>あいさつ</returns>
        [ExcelFunction(Description = "引数にHelloを言う。", Category = "Tadahiro Lib.")]
        public static string SayHello(string name)
        {
            return "こんにちは " + name + " さん。";
        }

        /// <summary>配列を返す関数の例</summary>
        /// <param name="name">あいさつの相手</param>
        /// <returns>あいさつ</returns>
        [ExcelFunction(Description = "1列4行の値を返す", Category = "Tadahiro Lib.")]
        public static object[,] SayHellos(string name)
        {
            return new object[,] { { 1 }, { 2 }, { 3 }, { 4 } };
        }
    }
}