using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Data;
using Microsoft.Office.Interop.Excel;
using MsExcel = Microsoft.Office.Interop.Excel;
using MYTU;

namespace Shalaopo
{
    class shashade
    {
        static void Main(string[] args)
        {
            
           
                string WorkDir = @"D:\气象\";

            string filename = WorkDir + "河北.xlsx";
           
           //  string[] RawDataStr_A = File.ReadAllLines(i+ ".xlsx" + ".csv", Encoding.UTF8);
         //   RawDataStr_A = File.ReadAllLines(i + ".xlsx" + ".csv", Encoding.Default);

            Microsoft.Office.Interop.Excel.Application appExcel = new Application();
            Microsoft.Office.Interop.Excel._Workbook ExcelBooks = null;


            ExcelBooks = Shashade(appExcel);
            if (File.Exists((string)filename))//判断文件已经是否存在
            {
                File.Delete((string)filename);//若已存在，则删除
            }

            ExcelBooks.SaveAs(filename);

            //stopWatch.Stop();
            //TimeSpan ts3 = stopWatch.Elapsed;
            //string elapsedTime3 = String.Format("{0:00}:{1:00}:{2:00}.{3:00}", ts3.Hours, ts3.Minutes, ts3.Seconds, ts3.Milliseconds / 10);
            //Console.WriteLine("RunTime " + elapsedTime3);
            ExcelBooks.Close();
            //appExcel.Quit();
            ExcelBooks = null;
            appExcel = null;


        }
        public static MsExcel._Workbook Shashade(MsExcel.Application appExcel)
        {

           
            _Workbook ExcelBooks = null;
            appExcel.Visible = false;
            ExcelBooks = appExcel.Workbooks.Add();

            string Sheet_Name = "站点数据";
            _Worksheet ExcelSheets = TU.Add_Sheets(ExcelBooks, Sheet_Name);//添加一个sheet
            ExcelSheets.Name = Sheet_Name;
           
            System.Data.DataTable dalaopo = new System.Data.DataTable("dalaopo");
            dalaopo.Columns.Add("站点", typeof(double )); // 
            dalaopo.Columns.Add("降水数据", typeof(double ));


            System.Data.DataTable jieguoTB = new System.Data.DataTable();
            
            double[] n = new double[] {129,135,128,132,99,113,157,81,70,111,78,138,103,88,110,91};
            double[] T = new double[] {5,10,20,50};
            for(int i=0;i< n.Length;i++)
            {
                List<double> gailvL = new List<double>();
                for(int j=0;j< T.Length;j++)
                {
                    double gailv = 1-50/(n[i]*T[j]);
                    gailvL.Add(gailv);
                }
                double[] gailvd = gailvL.ToArray();
                string[] gailvS = MYTUW.TUW.doublearrTOstring(gailvd);
                System.Data.DataTable jieguo = MYTUW.TUW.ArToDT1(gailvS);
                jieguoTB.Merge(jieguo);
            }
            #region
            /*   for (int i = 1; i < RawDataStr_A.Length; i++)
            {
                char[] seperators = { ',' };
                string[] R_str1 = TU.ParseStringTo_Array<string>(RawDataStr_A[i], seperators);
               
                double zhandian = Convert.ToDouble(R_str1[0]);
                string jiangshuistr = Convert.ToString(R_str1[R_str1.Length -1]);
               double jiangshui= Convert.ToDouble(R_str1[R_str1.Length - 1]);
                if (Convert .ToDouble ( jiangshuistr) > 30000)
                {
                    string housanwei = Convert.ToString(R_str1[R_str1.Length - 1]).Substring(Convert.ToString(R_str1[R_str1.Length - 1]).Length - 3);
                   // Console.WriteLine(housanwei);
                    if (Convert .ToDouble (housanwei)>500&& Convert.ToDouble(housanwei)!=700)
                    {
                        jiangshui = Convert.ToDouble(housanwei);
                    }
                   else
                    {
                        jiangshui = 1;
                    }
                       
                }
                dalaopo.Rows.Add(zhandian, jiangshui * 0.1);


            }
           var groupNew = from row in dalaopo.AsEnumerable()
                           group row by new
                           {
                               zhandian = (row.Field<double>("站点"))
                           }
                           into groupRes
                           orderby groupRes.Key.zhandian
                          select groupRes;
            //把分完组的每个表格放入DS中
            DataSet RTB_Grp = new DataSet();
            foreach (var eachGroup in groupNew)
            {
                System.Data.DataTable TBMid1 = eachGroup.CopyToDataTable();

                RTB_Grp.Tables.Add(TBMid1);
            }
            //对每天的数据进行处理并放入 Traffic_TB中RTB_Grp.Tables .Count
            for (int i=0;i< RTB_Grp.Tables.Count; i++)
                {
                System.Data.DataTable meigezhandian = RTB_Grp.Tables[i];
                List<double> baoyurishu = new List<double>();
                List<double> quannianyuliang = new List<double>();
                for (int j=0;j < meigezhandian.Rows.Count;j++)
                {
                    double yuliang = Convert.ToDouble(meigezhandian.Rows[j][1]);
                    quannianyuliang.Add(yuliang);
                    if (yuliang>50)
                    {
                        baoyurishu.Add(yuliang);
                    }
                }
                double baoyutianshu = baoyurishu.Count;
                double baoyuliang = baoyurishu.Sum();
                double baoyuqiangdu = baoyuliang / baoyutianshu;
                double zongyuliang = quannianyuliang.Sum();
                double baoyugongxianlv = baoyuliang / zongyuliang;
                jieguoTB.Rows.Add(Convert.ToDouble(meigezhandian.Rows[0][0]), baoyutianshu,
                   baoyuliang, baoyuqiangdu, baoyugongxianlv);
              }
          */
            #endregion
            MYTUW.TUW.DTToExcel(jieguoTB , ExcelSheets, 2, 1);

            MYTUW.TUW.Delete_sheet(appExcel, ExcelBooks);
            return ExcelBooks;
        }
    }
}
