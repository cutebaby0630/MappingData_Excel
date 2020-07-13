using DocumentFormat.OpenXml.Drawing;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;

namespace ConsoleMappingData
{
    public class Data
    {
        public string TableName { get; set; }
        //public string CSVName { get; set; }
        public string SchemaName { get; set; }
        public int ColumnCount { get; set; }
        public string TableDescription { get; set; }
        public string CreateDate { get; set; }
        public string ModifyDate { get; set; }
        public string FileName { get; set; }
        public string CSVName { get; set; }
        public override string ToString()
        {
            return base.ToString();
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Step 1.檔案清單列出

            //1.1 讀取檔案名稱存成List("FileName":filename)
            List<Data> filename = new List<Data>();
            DirectoryInfo readfile = new DirectoryInfo(@"D:\微軟MCS\CSV檔");
            foreach (var file in readfile.GetFiles())
            {
                filename.Add(new Data()
                {
                    CSVName = file.Name.ToString(),
                    FileName = file.Name.ToString().Replace(".csv", "")
                });
            }
            Console.WriteLine(filename.Count);

            //Step 2.DB Table List
            var datasource = @"10.1.225.17";
            var database = "HISDB";
            var username = "msdba";
            var password = "1qaz@wsx";
            string connString = @"Data Source=" + datasource + ";Initial Catalog=" + database + ";Persist Security Info=True;User ID=" + username + ";Password=" + password;
            string sql = @"SELECT S1.NAME TableName, --AS 資料表名稱,
                                  schema_name(s1.schema_id) SchemaName,--AS 結構名稱,
                                  S1.MAX_COLUMN_ID_USED ColumnCount,--AS 欄位數,
                                  ISNULL(S3.VALUE,'') TableDescription, --AS 資料表描述,
                                  s1.CREATE_DATE CreateDate,--AS 建立時間,
                                  S1.MODIFY_DATE ModifyDate --AS 修改時間
                           FROM SYS.TABLES S1
                           LEFT JOIN (
                                        SELECT * FROM SYS.OBJECTS S WHERE TYPE = 'PK'
                                     ) S2 ON S1.OBJECT_ID = S2.parent_object_id 
                           LEFT JOIN ( 
                                        SELECT T.Name , Convert(Varchar(500),P.Value) as Value
                                        FROM SYS.EXTENDED_PROPERTIES P
                                        INNER JOIN SYS.objects T  ON P.MAJOR_ID = T.OBJECT_ID 
                                        LEFT JOIN SYS.TABLES O ON T.parent_object_id = O.object_id
                                        INNER JOIN SYS.schemas S ON T.schema_id = S.schema_id 
                                        LEFT JOIN SYS.COLUMNS C ON T.object_id = c.object_id and P.MINOR_ID = C.column_id  
                                        LEFT JOIN SYS.indexes I  ON T.object_id = I.object_id and P.MINOR_ID = I.INDEX_id 
                                        WHERE P.CLASS = 1 AND T.TYPE = 'U' AND C.Name IS NULL
                                     ) S3  ON S1.NAME = S3.NAME
                           WHERE s1.is_ms_shipped = 0
                           AND (s1.NAME NOT LIKE 'DDSC[_]%'
                           AND s1.NAME NOT LIKE 'STG[_]%'
                           AND s1.NAME NOT LIKE 'ERR[_]%'
                           AND s1.NAME NOT LIKE 'TMP%'
                           AND s1.NAME NOT LIKE 'TEMP%'
                           )
                           ORDER BY 1";
            DataTable dt = new DataTable();

            //2.1 連接DB取得相對應資料表資料
            SqlConnection conn = new SqlConnection(connString);
            try
            {
                Console.WriteLine("Open Connection");
                conn.Open();
                SqlCommand cmd = new SqlCommand(sql, conn);
                SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(cmd);
                sqlDataAdapter.Fill(dt);
                Console.WriteLine("Success!!!");
                conn.Close();
            }
            catch (Exception e)
            {
                Console.WriteLine("Fail  " + e.Message);
            }

            //2.2 讀取相對應DB table 欄位名稱([TableName],[InitialDataFiles])存成List("FileName":TableName, "CSVName":InitialDataFiles)
            List<Data> dbfilename = new List<Data>();
            foreach (DataRow row in dt.Rows)
            {
                dbfilename.Add(new Data()
                {
                    TableName = row["TableName"].ToString(),
                    SchemaName = row["SchemaName"].ToString(),
                    ColumnCount = (int)row["ColumnCount"],
                    TableDescription = row["TableDescription"].ToString(),
                    CreateDate = row["CreateDate"].ToString(),
                    ModifyDate = row["ModifyDate"].ToString()
                });
            }


            //Step 3. 1 + 2 Mapping

            //3.1 Map two List
            //var result = dbfilename.MapFrom(a => filename.Select(b => b.FileName).Contains(a.TableName)).ToList();
            var result1 = from db in dbfilename
                          join local in filename
                          on db.TableName equals local.FileName into map
                          from allresult in map.DefaultIfEmpty()
                          select new { db.TableName, db.TableDescription, db.ModifyDate, Filename = allresult?.CSVName ?? String.Empty };



            //3.2 印出結果
            /*foreach (var v in result1)
            {
                Console.WriteLine(v.TableName + "\t" + v.TableDescription + "\t" + v.ModifyDate + "\t" + v.Filename);
            }*/

            //Step 4.Export EXCEL
            //4.1 產生EXCEL
            var excelname = "MapData" + DateTime.Now.ToString("yyyyMMddhhmm") + ".xlsx";
            var excel = new FileInfo(excelname);
            ExcelPackage.LicenseContext = LicenseContext.Commercial;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var finish = new ExcelPackage(excel))
            {
                finish.Workbook.Worksheets.Add("結果");
                Byte[] bin = finish.GetAsByteArray();
                File.WriteAllBytes(@"D:\微軟MCS\" + excelname, bin);

            }
            //4.2 讀取EXCEL
            FileInfo excel_new = new FileInfo(@"D:\微軟MCS\" + excelname);
            using (ExcelPackage package = new ExcelPackage(excel_new))
            {
                //4.3將值塞到EXCEL裡

                ExcelWorksheet firstsheet = package.Workbook.Worksheets[0];
                int rowIndex = 1;
                int colIndex = 1;
                //4.3.1塞資料到某一格
                firstsheet.Cells[rowIndex, colIndex++].Value = "TableName";
                firstsheet.Cells[rowIndex, colIndex++].Value = "TableDescription";
                firstsheet.Cells[rowIndex, colIndex++].Value = "ModifyDate";
                firstsheet.Cells[rowIndex, colIndex++].Value = "Filename";

                foreach (var v in result1)
                {
                    rowIndex++;
                    colIndex = 1;
                    firstsheet.Column(rowIndex).AutoFit();
                    firstsheet.Cells[rowIndex, colIndex++].Value = v.TableName;
                    firstsheet.Cells[rowIndex, colIndex++].Value = v.TableDescription;
                    firstsheet.Cells[rowIndex, colIndex++].Value = v.ModifyDate;
                    firstsheet.Cells[rowIndex, colIndex++].Value = v.Filename;
                    
                }
                
                Byte[] bin = package.GetAsByteArray();
                File.WriteAllBytes(@"D:\微軟MCS\" + excelname, bin);

            }
        }
    }
}
