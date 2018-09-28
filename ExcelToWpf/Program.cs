using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelToWpf
{
    class Program
    {
        static void Main(string[] args)
        {

            string targetFile = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), @"20180605検査員の作業範囲.xlsx");

            Excel.Application xlApp = new Excel.Application();

            var workbook = xlApp.Workbooks.Open(targetFile , ReadOnly:true);

            foreach (Excel.Worksheet sheet in workbook.Worksheets)
            {
                var range = sheet.UsedRange;
                int col = range.Columns.Count;
                int row = range.Rows.Count;

                if (sheet.Name == "print")
                {
                    Border[,] rawborder = new Border[row + 1, col + 1];

                    if (col >= 1 && row >= 1)
                    {
                        for (int i = 1; i <= row + 1; i++)
                        {
                            for (int j = 1; j <= col + 1; j++)
                            {
                                var rg = sheet.Cells[i, j] as Excel.Range;
                                readBorderInfo(rg, i, j, rawborder);
                            }
                        }
                    }

                    Dictionary<Tuple<int, int>, List<Border>> res = new Dictionary<Tuple<int, int>, List<Border>>();

                    for (int i = 0; i < rawborder.GetLength(0); i++)
                    {
                        for (int j = 0; j < rawborder.GetLength(1); j++)
                        {
                            //Console.Write(rawborder[i, j].ToString() + " ");

                            findMap(i, j, rawborder, res);
                        }
                        //Console.WriteLine();
                    }

                    if(res.Count > 0)
                    {
                        var filter1 = res.Values.Where(r => r.Count >= 5 && r[0].Equals(r[r.Count -1]));

                        if (filter1.Count() > 0)
                        {
                            var max = filter1.Select(r => r.Count).Max();
                            var target = filter1.FirstOrDefault(r => r.Count == max);

                            var leftmost = target.Select(r => r.Y).Min();
                            var rightmost = target.Select(r => r.Y).Max();
                            var topmost = target.Select(r => r.X).Min();
                            var bottommost = target.Select(r => r.X).Max();
                            //Console.WriteLine(String.Join(" ", target.Select(r => r.Position())));

                            Console.WriteLine(leftmost + " " + rightmost + " " + topmost + " " + bottommost);
                        }
                    }
                    
                }

            }

            xlApp.Quit();

            if (xlApp != null)
            {
                int excelProcessId = -1;
                GetWindowThreadProcessId(new IntPtr(xlApp.Hwnd), out excelProcessId);

                System.Diagnostics.Process ExcelProc = System.Diagnostics.Process.GetProcessById(excelProcessId);
                if (ExcelProc != null)
                {
                    ExcelProc.Kill();
                }
            }

            Console.ReadKey();
        }

        private static void findMap(int i, int j, Border[,] rawborder, Dictionary<Tuple<int, int>, List<Border>> res)
        {
            var d = canIniStart(i, j, rawborder);
            if (rawborder[i, j].L > 0 && rawborder[i, j].T > 0 && d != Direction.None)
            {
                Tuple<int, int> tmp = new Tuple<int, int>(i, j);

                if (!res.ContainsKey(tmp)) res.Add(tmp, new List<Border>());

                res[tmp].Add(rawborder[i, j]);

                if (d == Direction.Down)
                {
                    goDown(i + 1, j, rawborder, res[tmp]);
                }
            }
        }

        private static void goDown(int ii, int jj, Border[,] rawborder, List<Border> list)
        {
            //current box has left border
            //Console.WriteLine("goDown " + rawborder[ii, jj].Position() + " " + rawborder[ii, jj].ToString());

            if (rawborder[ii, jj].L > 0)
            {
                //check can go left
                if (jj - 1 <= 0)
                {
                    //for later purpose
                }

                //check can go down
                if (ii + 1 < rawborder.GetLength(0) && rawborder[ii + 1, jj].L > 0)
                {
                    list.Add(rawborder[ii, jj]);
                    goDown(ii + 1, jj, rawborder, list);
                }
                else if (rawborder[ii, jj].B > 0)
                {
                    if (jj + 1 < rawborder.GetLength(1) && rawborder[ii, jj + 1].B > 0)
                    {
                        list.Add(rawborder[ii, jj]);
                        goRight(ii, jj + 1, rawborder, list);
                    }
                }

            }
            else if (rawborder[ii, jj].R > 0)
            {

            }

        }

        private static void goRight(int i, int j, Border[,] rawborder, List<Border> list)
        {
            //current box has left border
            //Console.WriteLine("GoRight " + rawborder[i, j].Position() + " " + rawborder[i, j].ToString());
            if (rawborder[i, j].B > 0)
            {
                //check can go down
                if (i + 1 < rawborder.GetLength(0) && rawborder[i + 1, j].L > 0)
                {
                    list.Add(rawborder[i, j]);
                    goDown(i + 1, j, rawborder, list);
                }
                else if (j + 1 < rawborder.GetLength(1) && rawborder[i, j + 1].B > 0)
                {
                    list.Add(rawborder[i, j]);
                    goRight(i, j + 1, rawborder, list);
                }
                else if (rawborder[i, j].R > 0)
                {
                    if (i - 1 >= 0 && rawborder[i - 1, j].R > 0)
                    {
                        list.Add(rawborder[i, j]);
                        goTop(i - 1, j, rawborder, list);
                    }
                }

            }
            else if (rawborder[i, j].T > 0)
            {

            }
        }

        private static void goTop(int i, int j, Border[,] rawborder, List<Border> list)
        {
            //Console.WriteLine("GoTop " + rawborder[i, j].Position() + " " + rawborder[i, j].ToString());
            //current box has left border
            if (rawborder[i, j].R > 0)
            {
                //check right top left

                //top
                if (i - 1 >= 0 && rawborder[i - 1, j].R > 0)
                {
                    list.Add(rawborder[i, j]);
                    goTop(i - 1, j, rawborder, list);
                }
                else
                {
                    if (rawborder[i, j].T > 0)
                    {
                        //left
                        if (j - 1 >= 0 && rawborder[i, j - 1].T > 0)
                        {
                            list.Add(rawborder[i, j]);
                            goLeft(i, j - 1, rawborder, list);
                        }
                    }
                }
            }
            else if (rawborder[i, j].L > 0)
            {

            }
        }

        private static void goLeft(int i, int j, Border[,] rawborder, List<Border> list)
        {
            //Console.WriteLine("GoLeft " + rawborder[i, j].Position() + " " + rawborder[i, j].ToString());
            //something wrong here
            //current box has left border

            if (list[0].Equals(rawborder[i, j]))
            {
                list.Add(rawborder[i, j]);
                return;
            }

            if (rawborder[i, j].T > 0)
            {
                //check can go top left down
                if (j - 1 >= 0 && rawborder[i, j - 1].T > 0)
                {
                    list.Add(rawborder[i, j]);
                    goLeft(i, j - 1, rawborder, list);
                }
            }
            else if (rawborder[i, j].B > 0)
            {

            }
        }

        private static Direction canIniStart(int i, int j, Border[,] rawborder)
        {
            //go left
            if (j - 1 <= 0)
            {
                //for later purpose
            }

            //go down
            if (i + 1 < rawborder.GetLength(0))
            {
                if (rawborder[i + 1, j].L > 0) return Direction.Down;
            }

            //go right
            if (j + 1 < rawborder.GetLength(1))
            {
                if (rawborder[i, j + 1].B > 0) return Direction.Right;
            }

            return Direction.None;
        }

        private static void readBorderInfo(Excel.Range rg, int i, int j, Border[,] rawborder)
        {
            var border = rg.Borders;
            rawborder[i - 1, j - 1].L = getborder(border[Excel.XlBordersIndex.xlEdgeLeft]);
            rawborder[i - 1, j - 1].T = getborder(border[Excel.XlBordersIndex.xlEdgeTop]);
            rawborder[i - 1, j - 1].R = getborder(border[Excel.XlBordersIndex.xlEdgeRight]);
            rawborder[i - 1, j - 1].B = getborder(border[Excel.XlBordersIndex.xlEdgeBottom]);

            rawborder[i - 1, j - 1].X = i - 1;
            rawborder[i - 1, j - 1].Y = j - 1;

            rawborder[i - 1, j - 1].Content = "" + rg.Value;

            if (rg.MergeCells)
            {
                var xx = rg.MergeArea;

                var yy = xx.Address;

                var zz = xx.AddressLocal;
                
                rawborder[i - 1, j - 1].Merged = true;
            }
        }

        private static int getborder(Excel.Border border)
        {
            var style = border.LineStyle;
            var weight = border.Weight;

            if (style == Excel.XlLineStyle.xlLineStyleNone.GetHashCode())
                return 0;
            else if (style == Excel.XlLineStyle.xlContinuous.GetHashCode())
            {
                if (weight == Excel.XlBorderWeight.xlHairline.GetHashCode())
                    return 1;
                else if (weight == Excel.XlBorderWeight.xlThin.GetHashCode())
                    return 1;
                else if (weight == Excel.XlBorderWeight.xlThick.GetHashCode())
                    return 4;
                else if (weight == Excel.XlBorderWeight.xlMedium.GetHashCode())
                    return 2;
            }
            else if (style == Excel.XlLineStyle.xlDouble.GetHashCode())
                return 20;

            return 0;
        }


        [DllImport("user32.dll")]
        public static extern IntPtr GetWindowThreadProcessId(IntPtr hWnd, out int ProcessId);
    }

    enum Direction
    {
        Left, Down, Right, Up, None
    }

    struct Border
    {
        /// <summary>
        /// Left border
        /// </summary>
        public int L;

        /// <summary>
        /// Top border
        /// </summary>
        public int T;

        /// <summary>
        /// Right border
        /// </summary>
        public int R;

        /// <summary>
        /// Bottom border
        /// </summary>
        public int B;

        /// <summary>
        /// row number start from 0
        /// </summary>
        public int X;

        /// <summary>
        /// column number start from 0
        /// </summary>
        public int Y;

        /// <summary>
        /// content
        /// </summary>
        public string Content;

        /// <summary>
        /// is this cell a merged cell
        /// </summary>
        public bool Merged;

        public string ToString()
        {
            return L + "," + T + "," + R + "," + B;
        }

        public string Position()
        {
            return X + "," + Y;
        }
    }
}
