using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
//-------Excel
using System.Runtime.InteropServices;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;
using LottoApp.DB;
//-----------
namespace LottoApp.Logic
{
    internal class ExcelModule
    {
        Excel.Application oXL = null;
        Excel._Workbook oWB = null;
        Excel._Worksheet oSheet = null;
        public void Init()
        {
            oXL = new Excel.Application(); //Excel application create
            oXL.Visible = true; //true 설정하면 엑셀 작업되는 내용이 보인다.
            oXL.Interactive = false; //false로 설정하면 사용자의 조작에 방해받지 않음.﻿
            oXL.Interactive = true; //사용자의 조작 허용

            oWB = (Excel._Workbook)(oXL.Workbooks.Add(Missing.Value));//워크북생성
            oSheet = (Excel._Worksheet)oWB.ActiveSheet;//시트 가져오기
            Excel.Range oRng = null; //각종 셀 처리 담당할 변수﻿
        }
        public void Dispose()
        {
            oWB.Close(false);
            oXL.Quit();

            if (oSheet != null)
                Marshal.ReleaseComObject(oSheet);
            if (oWB != null)
                Marshal.ReleaseComObject(oWB);
            if (oXL != null)
                Marshal.ReleaseComObject(oXL);

            //가비지 컬렉터
            GC.Collect();
        }

        public void ReadExcelFile(string path)
        {
            Excel.Application oXL = null;
            Excel.Workbook oWB = null;
            Excel.Worksheet oSheet = null;
            try
            {
                // Excel 시작 & Application Object얻어오기
                oXL = new Excel.Application();
                oWB = oXL.Workbooks.Open(path);



                // 첫 번째 Worksheet 선택
                oSheet = oWB.Worksheets.get_Item(1) as Excel.Worksheet;

                // Used 영역 선택
                Excel.Range oRng = oSheet.UsedRange;

                // Data를 배열로 받아옴       
                object[,] data = oRng.Value;

                // 0 base가 아니라 1 base인 것 주의
                for (int row = 2; row <= oRng.Rows.Count; row++)
                {
                    int id = Int32.Parse(data[row, 1].ToString());
                    int N1 = Int32.Parse(data[row, 2].ToString());
                    int N2 = Int32.Parse(data[row, 3].ToString());
                    int N3 = Int32.Parse(data[row, 4].ToString());
                    int N4 = Int32.Parse(data[row, 5].ToString());
                    int N5 = Int32.Parse(data[row, 6].ToString());
                    int N6 = Int32.Parse(data[row, 7].ToString());
                    int Bonus = Int32.Parse(data[row, 8].ToString());
                    /*for (int col = 1; col <= oRng.Columns.Count; col++)
                    {
                        if (data[row, col] != null)
                        {
                            Console.WriteLine("[" + row + "," + col + "] " + data[row, col]);
                        }
                    }*/
                }
                oWB.Close(false);

                //oWB.Close(true); // save 

                oXL.Quit();
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                ReleaseExcelObject(oSheet);

                ReleaseExcelObject(oWB);

                ReleaseExcelObject(oXL);

            }

        }
        public void ReadExcelToDB(string path)
        {
            Excel.Application oXL = null;
            Excel.Workbook oWB = null;
            Excel.Worksheet oSheet = null;

            SQLiteDB db = new SQLiteDB();
            db.Init();
            db.ConnectDB();
            Lotto lotto = new Lotto();
            try
            {
                // Excel 시작 & Application Object얻어오기
                oXL = new Excel.Application();
                oWB = oXL.Workbooks.Open(path);



                // 첫 번째 Worksheet 선택
                oSheet = oWB.Worksheets.get_Item(1) as Excel.Worksheet;

                // Used 영역 선택
                Excel.Range oRng = oSheet.UsedRange;

                // Data를 배열로 받아옴       
                object[,] data = oRng.Value;

                // 0 base가 아니라 1 base인 것 주의
                for (int row = 2; row <= oRng.Rows.Count; row++)
                {
                    int id = Int32.Parse(data[row, 1].ToString());
                    int N1 = Int32.Parse(data[row, 2].ToString());
                    int N2 = Int32.Parse(data[row, 3].ToString());
                    int N3 = Int32.Parse(data[row, 4].ToString());
                    int N4 = Int32.Parse(data[row, 5].ToString());
                    int N5 = Int32.Parse(data[row, 6].ToString());
                    int N6 = Int32.Parse(data[row, 7].ToString());
                    int Bonus = Int32.Parse(data[row, 8].ToString());
                    int[] numbers = { N1, N2, N3, N4, N5, N6};
                    
                    

                    StringBuilder sb = new StringBuilder();
                    sb.Append("INSERT INTO LottoHistory(ID, N1, N2, N3, N4, N5, N6, Bonus, BA)");
                    sb.Append("VALUES(");
                    sb.Append(id); sb.Append(", ");
                    sb.Append(N1); sb.Append(", ");
                    sb.Append(N2); sb.Append(", ");
                    sb.Append(N3); sb.Append(", ");
                    sb.Append(N4); sb.Append(", ");
                    sb.Append(N5); sb.Append(", ");
                    sb.Append(N6); sb.Append(", ");
                    sb.Append(Bonus); sb.Append(",");
                    sb.Append("'"); sb.Append(lotto.BitToString(lotto.ToBitArray(numbers))); sb.Append("'"); sb.Append(")");
                    db.Query(sb.ToString());

                    sb.Clear();
                    sb = null;
                    /*for (int col = 1; col <= oRng.Columns.Count; col++)
                    {
                        if (data[row, col] != null)
                        {
                            Console.WriteLine("[" + row + "," + col + "] " + data[row, col]);
                        }
                    }*/
                }
                oWB.Close(false);

                //oWB.Close(true); // save 

                oXL.Quit();
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                ReleaseExcelObject(oSheet);

                ReleaseExcelObject(oWB);

                ReleaseExcelObject(oXL);

                db.DisconnectDB();
            }

        }
        private void ReleaseExcelObject(object obj)
        {
            try
            {
                if (obj != null)
                {
                    Marshal.ReleaseComObject(obj);

                    obj = null;
                }
            }
            catch (Exception ex)
            {
                obj = null;
                throw ex;
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}
