using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;


namespace ConsoleMappingData
{
    class Program
    {
        static void Main(string[] args)
        {
            // Step 1.檔案清單列出

            //1.1 讀取檔案名稱存成List("FileName":filename)
            List<string> filename = new List<string>();
            DirectoryInfo readfile = new DirectoryInfo(@"D:\微軟MCS\CSV檔");
            foreach (var file in readfile.GetFiles())
            {
                var name = file.Name.ToString().Replace(".csv", "");
                filename.Add(name);
            }
            Console.WriteLine(filename.Count);
            //Step 2.DB Table List
            var datasource = @"10.1.225.17";
            var database = "SSISDB";
            var username = "msdba";
            var password = "1qaz@wsx";
            string connString =  @"Data Source=" + datasource + ";Initial Catalog=" + database + ";Persist Security Info=True;User ID=" + username + ";Password=" + password;

            //2.1 連接DB
            SqlConnection conn = new SqlConnection(connString);
            try
            {
                Console.WriteLine("Open Connection");
                conn.Open();
                Console.WriteLine("Success!!!");

            }
            catch (Exception e) {
                Console.WriteLine("Fail  " + e.Message);
            }
            
            //2.2 讀取相對應DB table 欄位名稱([TableName],[InitialDataFiles])存成List("FileName":TableName, "CSVName":InitialDataFiles)

            //Step 3. 1 + 2 Mapping

            //3.1 將兩個List利用迴圈比對

            //3.2 印出結果

            //Step 4.Export EXCE




        }
    }
}
