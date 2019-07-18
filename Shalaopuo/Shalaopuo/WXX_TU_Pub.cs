using Microsoft.Office.Interop.Excel;
using MYTU;

using System;
using System.Data;
using System.IO;
using System.Linq;
using MsExcel = Microsoft.Office.Interop.Excel;
using MathNet.Numerics;
using MathNet.Numerics.LinearRegression;
using System.Collections.Generic;

namespace MYTUW
{

    /// <summary>
    /// Excel基本操作命令, 是一些通用命令；
    /// </summary>
    public static partial class TUW
    {
        

        public static double Sta_str2Double_impl(string str)
        {
            str = str.Trim();
            str = str.TrimStart('0'); // 把开头的 '0'都去掉
            if (str.Length == 0)
            {
                return 0.0;
            }

            double R_v = Convert.ToDouble(str);
            // Console.WriteLine(str);
            // Console.WriteLine(R_v);

            return R_v;
        }


        /// <summary>
        /// 将用户输入桩号(string)转换为Double, 若是抛出 Exception， 表明用户输入字符串错误
        /// </summary>        
        /// <returns></returns>
        public static Double Station_Str2Double(string StationStr)
        {
            bool IFyouK = false;
            bool IFyoujia = false;
            //判断字符串是否有+或者K
            StationStr = StationStr.Trim();
            char[] charIf = StationStr.ToCharArray();
            for (int i = 0; i < charIf.Length; i++)
            {
                if (charIf[i] == 'K' || charIf[i] == 'k')
                {
                    IFyouK = true;
                }
                if (charIf[i] == '+')
                {
                    IFyoujia = true;
                }
            }
            if (IFyouK == true && IFyoujia == true)
            {
                StationStr = StationStr.Trim();
                string[] str = StationStr.Split(new char[] { '+' });
                string string1 = str[0];
                string1 = string1.Substring(1); // 删去第一个字母 ‘K'  // string string1 = str[0].Replace("K", "");
                string string2 = str[1];
                //判断str1中是否有小数点
                char[] char1 = string1.ToCharArray();
                for (int i = 0; i < char1.Length; i++)
                {
                    if (char1[i] == '.')
                    {
                        throw new Exception("用户输入公里桩号错误，不能有小数点" + StationStr);

                        // Console.WriteLine("用户输入公里桩号错误，不能有小数点" + StationStr);
                    }
                }
                //判断str2中是否只有一个小数点
                char[] char2 = string2.ToCharArray();
                int num = 0;
                for (int i = 0; i < char2.Length; i++)
                {
                    if (char2[i] == '.')
                    {
                        num = num + 1;
                    }
                }
                if (num > 1)
                {
                    throw new Exception("用户输入米桩号错误" + StationStr);
                }


                //对str1和str2中全是0的情况进行处理

                double double1 = Sta_str2Double_impl(string1);
                double double2 = Sta_str2Double_impl(string2);

                double reture_v = double1 * 1000 + double2;

                return reture_v;
            }
            else if (IFyouK == false && IFyoujia == true)
            {
                StationStr = StationStr.Trim();
                string[] str = StationStr.Split(new char[] { '+' });
                string string1 = str[0];
                string string2 = str[1];
                //判断str1中是否有小数点
                char[] char1 = string1.ToCharArray();
                for (int i = 0; i < char1.Length; i++)
                {
                    if (char1[i] == '.')
                    {
                        throw new Exception("用户输入公里桩号错误，不能有小数点" + StationStr);

                        // Console.WriteLine("用户输入公里桩号错误，不能有小数点" + StationStr);
                    }
                }
                //判断str2中是否只有一个小数点
                char[] char2 = string2.ToCharArray();
                int num = 0;
                for (int i = 0; i < char2.Length; i++)
                {
                    if (char2[i] == '.')
                    {
                        num = num + 1;
                    }
                }
                if (num > 1)
                {
                    throw new Exception("用户输入米桩号错误" + StationStr);
                }


                //对str1和str2中全是0的情况进行处理

                double double1 = Sta_str2Double_impl(string1);
                double double2 = Sta_str2Double_impl(string2);

                double reture_v = double1 * 1000 + double2;

                return reture_v;
            }
            else if (IFyouK == false && IFyoujia == false)
            {
                StationStr = StationStr.Trim();

                double reture_v = Convert.ToDouble(StationStr);
                return reture_v;
            }
            throw new Exception("桩号格式错误" + StationStr);
        } // End of public static Double Station_Str2Double


        /// <summary>
        /// 将桩号（Double）转换为String
        /// </summary>
        /// <param name="StationDouble"></param>
        /// <returns></returns>
        public static String Station_Double2Str(double StationDouble)
        {
            double x = Math.Floor(StationDouble / 1000);
            double y = StationDouble % 1000;
            string Y = y.ToString("000.0");
            string Station = "K" + x + "+" + Y;
            return Station;
        }
        /// <summary>
        /// 把double数组转换为String数组
        /// </summary>
        /// <param name="doubles"></param>
        /// <returns></returns>
        public static string[] doublearrTOstring(double[] doubles)
        {
            string[] strings = new string[doubles.Length];//空的string数组，假定长度为3（string数组的长度>=double数组的长度）
            for (int i = 0; i < strings.Length; i++)
            {
                strings[i] = doubles[i].ToString();//将double数组中的元素转换为string，插入string数组中
            }
            return strings;
        }
        /// <summary>
        /// DataTable转换为一维字符串数组
        /// </summary>
        /// <returns></returns>
        public static string[] DtToArr1Str(System.Data.DataTable dt)
        {
            if (dt == null || dt.Rows.Count == 0) return new string[0];
            string[] sr = new string[dt.Rows.Count];
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                if (Convert.IsDBNull(dt.Rows[i][0])) sr[i] = "";
                else sr[i] = dt.Rows[i][0] + "";
            }
            return sr;
        }
        /// <summary>
        /// 把一个一维数组转换为DataTable 的一行
        /// </summary>
        /// <param name="arr"></param>
        /// <returns></returns>
        public static System.Data.DataTable ArToDT1(string[] arr)
        {
            System.Data.DataTable dataSouce = new System.Data.DataTable();
            for (int i = 0; i < arr.Length; i++)
            {
                DataColumn newColumn = new DataColumn(i.ToString(), arr[i].GetType());
                dataSouce.Columns.Add(newColumn);
            }

            DataRow newRow = dataSouce.NewRow();
            for (int j = 0; j < arr.Length; j++)
            {
                newRow[j.ToString()] = arr[j];
            }
            dataSouce.Rows.Add(newRow);

            return dataSouce;
        }
        /// <summary>  
        /// 把一个一维数组转换为DataTable 的一列
        /// </summary>  
        /// <param name="ColumnName">列名</param>  
        /// <param name="Array">一维数组</param>  
        /// <returns>返回DataTable</returns>  
        public static System.Data.DataTable ArToDT(string[] arr)
        {
            System.Data.DataTable dataSouce = new System.Data.DataTable();
            /* for (int i = 0; i < arr.Length; i++)
             {*/
            // DataColumn newColumn = new DataColumn(i.ToString(), arr[i].GetType());
            dataSouce.Columns.Add("string", typeof(string));
            //   }*/


            for (int j = 0; j < arr.Length; j++)
            {
                dataSouce.Rows.Add(arr[j]);
            }


            return dataSouce;
        }
        /// <summary>
        /// 将datatable转化为二维数组
        /// </summary>
        /// <param name="dt"></param>
        /// <returns></returns>
        public static string[,] Turnarray(System.Data.DataTable dt)
        {
            string[,] array = new string[dt.Rows.Count, dt.Columns.Count];
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    array[i, j] = dt.Rows[i][j].ToString().ToUpper().Trim();
                }
            }
            return array;
        }
        /// <summary>
        /// 将二维数组转换为DataTable
        /// </summary>
        /// <param name="arr"></param>
        /// <returns></returns>
        public static System.Data.DataTable ConvertToDataTable(string[,] arr)
        {

            System.Data.DataTable dataSouce = new System.Data.DataTable();
            for (int i = 0; i < arr.GetLength(1); i++)
            {
                DataColumn newColumn = new DataColumn(i.ToString(), arr[0, 0].GetType());
                dataSouce.Columns.Add(newColumn);
            }
            for (int i = 0; i < arr.GetLength(0); i++)
            {
                DataRow newRow = dataSouce.NewRow();
                for (int j = 0; j < arr.GetLength(1); j++)
                {
                    newRow[j.ToString()] = arr[i, j];
                }
                dataSouce.Rows.Add(newRow);
            }
            return dataSouce;

        }
        /// <summary>
        ////// <summary>
        /// 查找
        /// </summary>
        /// <returns></returns>
        public static void Select(System.Data.DataTable EnvirDataTB,
            DateTime Begin_Time, DateTime End_Time, int Env_TempID)

        {
            var MySelRes = from line in EnvirDataTB.AsEnumerable()
                           where (line.Field<int>("DataTypeID")) == Env_TempID &&
                           (line.Field<DateTime>("MeasureTime")) > Begin_Time &&
                           (line.Field<DateTime>("MeasureTime")) < End_Time
                           select line;
            int numMySelRec = MySelRes.Count();

        }
        /// </summary>
        /// <summary>
        /// 将CSV文件中内容读取到DataTable中
        /// </summary>
        /// <param name="path">CSV文件路径</param>
        /// <param name="hasTitle">是否将CSV文件的第一行读取为DataTable的列名</param>
        /// <returns></returns>
        public static System.Data.DataTable ReadFromCSV(string path, bool hasTitle = false)
        {
            System.Data.DataTable dt = new System.Data.DataTable();           //要输出的数据表
            StreamReader sr = new StreamReader(path); //文件读入流
            bool bFirst = true;                       //指示是否第一次读取数据

            //逐行读取
            string line;
            while ((line = sr.ReadLine()) != null)
            {
                string[] elements = line.Split(',');

                //第一次读取数据时，要创建数据列
                if (bFirst)
                {
                    for (int i = 0; i < elements.Length; i++)
                    {
                        dt.Columns.Add();
                    }
                    bFirst = false;
                }

                //有标题行时，第一行当做标题行处理
                if (hasTitle)
                {
                    for (int i = 0; i < dt.Columns.Count && i < elements.Length; i++)
                    {
                        dt.Columns[i].ColumnName = elements[i];
                    }
                    hasTitle = false;
                }
                else //读取一行数据
                {
                    if (elements.Length == dt.Columns.Count)
                    {
                        dt.Rows.Add(elements);
                    }
                    else
                    {
                        //throw new Exception("CSV格式错误：表格各行列数不一致");
                    }
                }
            }
            sr.Close();

            return dt;
        }

        /// <summary>
        /// 将DataTable内容保存到CSV文件中
        /// </summary>
        /// <param name="dt">数据表</param>
        /// <param name="path">CSV文件地址</param>
        /// <param name="hasTitle">是否要输出数据表各列列名作为CSV文件第一行</param>
        public static void SaveToCSV(System.Data.DataTable dt, string path, bool hasTitle = false)
        {
            StreamWriter sw = new StreamWriter(path);

            //输出标题行（如果有）
            if (hasTitle)
            {
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    sw.Write(dt.Columns[i].ColumnName);
                    if (i != dt.Columns.Count - 1)
                    {
                        sw.Write(",");
                    }
                }
                sw.WriteLine();
            }

            //输出文件内容
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    sw.Write(dt.Rows[i][j].ToString());
                    if (j != dt.Columns.Count - 1)
                    {
                        sw.Write(",");
                    }
                }
                sw.WriteLine();
            }

            sw.Close();
        }
        /// <summary>
        /// 删除空的sheet
        /// </summary>
        /// <param name="appExcel"></param>
        /// <param name="ExcelBooks"></param>
        public static void Delete_sheet(MsExcel.Application appExcel, MsExcel._Workbook ExcelBooks)
        {
           // Console.WriteLine(ExcelBooks.Worksheets.Count);
            for (int i = ExcelBooks.Worksheets.Count; i >0 ; i--)
            {
             //   Console.WriteLine(i);
              //  Console.WriteLine(((MsExcel.Worksheet)ExcelBooks.Worksheets[i]).UsedRange.Count);
                if (((MsExcel.Worksheet)ExcelBooks.Worksheets[i]).UsedRange.Count<2)  
                {
                   
                    appExcel.DisplayAlerts = false;
                    ((MsExcel.Worksheet)ExcelBooks.Worksheets[i]).Delete();
                   appExcel.DisplayAlerts = true;
                }
            }
        }
    
         
        
        

        public static string Move_excel_active_cell(int mode, string Old_Cell_Position, int number_of_Cell_to_move)
        {
            string Colword = Old_Cell_Position.Substring(0, 1);
            string ColwordSec = Old_Cell_Position.Substring(1, 1);
            string Rownumb = Old_Cell_Position.Remove(0, 1);
            int lengthfirst = Asc(Colword) + number_of_Cell_to_move - 90;///首字母进位差值
            int lengthsecond = Asc(ColwordSec) + number_of_Cell_to_move - 90;///第二字母进位差值
            int Colwordnum = Asc(Colword);///首字母ASC码
            int ColwordSecnum = Asc(ColwordSec);///次字ASC码


            if (mode == 1 && Asc(ColwordSec) < 65)
            {
                if (lengthfirst <= 0)
                {
                    int Movedwordasc = Asc(Colword) + number_of_Cell_to_move;
                    string Movedcol = Chr(Movedwordasc);
                    string Movedcell = Movedcol + Rownumb;
                    return Movedcell;
                }
                ///判断是否进位
                else
                {
                    string secondletter = Chr(lengthfirst + 64);
                    string Movedcell = "A" + secondletter + Rownumb;
                    return Movedcell;
                }



            }
            ///列向移动，且行字符只有一个字母 （例如B12）
            if (mode == 1 && Asc(ColwordSec) >= 65)
            {

                if (lengthsecond <= 0)
                {
                    int Movedwordasc = Asc(ColwordSec) + number_of_Cell_to_move;
                    string Movedcol = Chr(Movedwordasc);
                    string Movedcell = Colword + Movedcol + Rownumb;
                    return Movedcell;
                }
                else
                {
                    string firstletter = Chr(Colwordnum + 1);
                    string secondletter = Chr(lengthsecond + 65);
                    string Movedcell = firstletter + secondletter + Rownumb;
                    return Movedcell;
                }


            }
            ///列向移动，且行字符有两个字母 （例如AB12）
            if (mode == 0)
            {
                int Movednumb = Convert.ToInt32(Rownumb) + number_of_Cell_to_move;
                string Movedcell = Colword + Convert.ToString(Movednumb);
                return Movedcell;
            }
            else
            {
                return null;
                throw new Exception(" mode is not valid.");
            }

        }//0行向移动，1列向移动
         /// <summary>
         /// Excel单元格的移动 列数 小于676
         /// </summary>
         /// <param name="character"></param>
         /// <returns></returns>
        public static int Asc(string character)
        {
            if (character.Length == 1)
            {
                System.Text.ASCIIEncoding asciiEncoding = new System.Text.ASCIIEncoding();
                int intAsciiCode = (int)asciiEncoding.GetBytes(character)[0];
                return (intAsciiCode);
            }
            else
            {
                throw new Exception("Character is not valid.");
            }

        }
        ///转ASC码

        public static string Chr(int asciiCode)
        {
            if (asciiCode >= 0 && asciiCode <= 255)
            {
                System.Text.ASCIIEncoding asciiEncoding = new System.Text.ASCIIEncoding();
                byte[] byteArray = new byte[] { (byte)asciiCode };
                string strCharacter = asciiEncoding.GetString(byteArray);
                return (strCharacter);
            }
            else
            {
                throw new Exception("ASCII Code is not valid.");
            }
        }
        /// <summary>
        /// 为Excel的指定单元格赋值并设置格式
        /// </summary>
        /// <param name="DistressData_sheet"></param>
        /// <param name="aaa"></param>
        /// <param name="Highwayname"></param>
        public static void Setcell(_Worksheet DistressData_sheet,string aaa,string Highwayname)
        {
            MsExcel.Range name_PC = DistressData_sheet.get_Range(aaa);
            name_PC.Value = Highwayname;//写入公路名称
            name_PC.NumberFormatLocal = "@";
            name_PC.ColumnWidth = 20;
        }
        /// <summary>
        /// 为导入EXcel的表格设置自适应边框
        /// </summary>
        /// <param name="PCIoldFinDT"></param>
        /// <param name="PCi_Sheet"></param>
        /// <param name="begincell"></param>
        public static void Frame(System.Data.DataTable PCIoldFinDT, _Worksheet PCi_Sheet, string begincell)
        {
            int moveothcn = PCIoldFinDT.Columns.Count - 1;
            string moveoth1 = MYTUW.TUW.Move_excel_active_cell(1, begincell, moveothcn);
            int moveothro = PCIoldFinDT.Rows.Count;
            string moveoth2 = MYTUW.TUW.Move_excel_active_cell(0, moveoth1, moveothro);
            TU.set_Table_Format(PCi_Sheet, begincell, moveoth2);
        }
        /// <summary>
        /// 把DataTable写入Excel
        /// </summary>
        /// <param name="PCIoldFinDT"></param>
        /// <param name="PCi_Sheet"></param>
        /// <param name="row"></param>
        /// <param name="column"></param>
        public static void DTToExcel(System.Data.DataTable PCIoldFinDT, _Worksheet PCi_Sheet,int row,int column)
        {
            for (int i = 0; i < PCIoldFinDT.Rows.Count; i++)//循环赋值
            {
                for (int f = 0; f < PCIoldFinDT.Columns.Count; f++)
                {
                    PCi_Sheet.Cells[i + row, f + column] = PCIoldFinDT.Rows[i][f];
                }
            }
        }
        /// <summary>
        /// 重复合并指定单元格后面的单元格
        /// </summary>
        public static void Mergecell1(string begincell, System.Data.DataTable ChedaoDT,_Worksheet PCi_Sheet)
        {
            //string begincell = "N7";
            for (int i = 0; i < ChedaoDT.Columns.Count / 2; i++)
            {
                string moveoth1 = MYTUW.TUW.Move_excel_active_cell(1, begincell, 1);
                MsExcel.Range cells_place = PCi_Sheet.Range[begincell, moveoth1];
                cells_place.Merge();
                begincell = MYTUW.TUW.Move_excel_active_cell(1, moveoth1, 1);
            }//合并单元格

        }
        /// <summary>
        ///  求数组的元素的n次方的和
        /// </summary>
        /// <param name="arr"></param>
        /// <param name="n"></param>
        /// <param name="length"></param>
        /// <returns></returns>
        public static double SumArr(double[] arr, int n, int length) //求数组的元素的n次方的和
        {
            double s = 0;
            for (int i = 0; i < length; i++)
            {
                if (arr[i] != 0 || n != 0)
                    s = s + Math.Pow(arr[i], n);
                else
                    s = s + 1;
            }
            return s;
        }
        public static double SumArr(double[] arr1, int n1, double[] arr2, int n2, int length)
        {
            double s = 0;
            for (int i = 0; i < length; i++)
            {
                if ((arr1[i] != 0 || n1 != 0) && (arr2[i] != 0 || n2 != 0))
                    s = s + Math.Pow(arr1[i], n1) * Math.Pow(arr2[i], n2);
                else
                    s = s + 1;
            }
            return s;

        }

        public static double[] ComputGauss(double[,] Guass, int n)
        {
            int i, j;
            int k, m;
            double temp;
            double max;
            double s;
            double[] x = new double[n];
            for (i = 0; i < n; i++) x[i] = 0.0;//初始化

            for (j = 0; j < n; j++)
            {
                max = 0;
                k = j;
                for (i = j; i < n; i++)
                {
                    if (Math.Abs(Guass[i, j]) > max)
                    {
                        max = Guass[i, j];
                        k = i;
                    }
                }


                if (k != j)
                {
                    for (m = j; m < n + 1; m++)
                    {
                        temp = Guass[j, m];
                        Guass[j, m] = Guass[k, m];
                        Guass[k, m] = temp;
                    }
                }
                if (0 == max)
                {
                    // "此线性方程为奇异线性方程" 
                    return x;
                }

                for (i = j + 1; i < n; i++)
                {
                    s = Guass[i, j];
                    for (m = j; m < n + 1; m++)
                    {
                        Guass[i, m] = Guass[i, m] - Guass[j, m] * s / (Guass[j, j]);
                    }
                }

            }//结束for (j=0;j<n;j++)

            for (i = n - 1; i >= 0; i--)
            {
                s = 0;
                for (j = i + 1; j < n; j++)
                {
                    s = s + Guass[i, j] * x[j];
                }
                x[i] = (Guass[i, n] - s) / Guass[i, i];
            }
            return x;
        }//返回值是函数的系数
    }
}