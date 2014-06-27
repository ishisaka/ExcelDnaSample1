namespace MyExcelDnaSample1
{
    using ExcelDna.Integration;

    /// <summary>
    /// Excel DNAを使用してExcelのユーザー定義関数(UDF)を作成するサンプル
    /// UDFはstaticなクラスにstaticなメソッドとして記述する。
    /// </summary>
    public static class MyExcelDnaSample1
    {
        /// <summary>
        /// 引数の名前にHelloを言う
        /// </summary>
        /// <param name="name">あいさつの相手</param>
        /// <returns>あいさつ</returns>
        [ExcelFunction(Description = "引数にHelloを言う。", Category = "Tadahiro Lib.")]
        public static string SayHello(string name)
        {
            return "Hello " + name;
        }
    }
}