using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

using App = Microsoft.Office.Interop.Excel.Application;

namespace OptimizationMethods
{
    public enum ExtremumType
    {
        Minimum,
        Maximum
    }
    static class OneDimensionalSearches
    {
        private static readonly double sqrt5 = Math.Sqrt(5.0);
        public static double Epsilon = 1E-7;
        public static double FastPow(double x, uint power)
        {
            double result = 1.0;
            uint bit = ((uint)1) << 31;
            while(bit > 0)
            {
                result *= result;
                if ((power & bit) != 0)
                    result *= x;
                bit = bit >> 1;
            }
            return result;
        }
        public static void DichotomyMethod(ExtremumType type, double leftBorder, double rightBorder, Func<double, double> function, out double left, out double right)
        {
            double delta = Epsilon / 2.0;
            double x1, x2;

            switch(type)
            {
                case ExtremumType.Minimum:
                    do
                    {
                        x1 = (leftBorder + rightBorder - delta) / 2.0;
                        x2 = (leftBorder + rightBorder + delta) / 2.0;
                        if (function(x1) > function(x2))
                            leftBorder = x1;
                        else rightBorder = x2;
                    } while (Math.Abs(leftBorder - rightBorder) > Epsilon);
                    break;
                case ExtremumType.Maximum:
                    do
                    {
                        x1 = (leftBorder + rightBorder - delta) / 2.0;
                        x2 = (leftBorder + rightBorder + delta) / 2.0;
                        if (function(x1) < function(x2))
                            leftBorder = x1;
                        else rightBorder = x2;
                    } while (Math.Abs(leftBorder - rightBorder) > Epsilon);
                    break;
            }

            left = leftBorder;
            right = rightBorder;
        }
        public static void GoldenRatioMethod(ExtremumType type, double leftBorder, double rightBorder, Func<double, double> function, out double left, out double right)
        {
            double goldNumber = (3.0 - sqrt5) / 2.0;

            double x1 = leftBorder + goldNumber * (rightBorder - leftBorder);
            double x2 = rightBorder - goldNumber * (rightBorder - leftBorder);

            double f1 = function(x1);
            double f2 = function(x2);

            switch(type)
            {
                case ExtremumType.Minimum:
                    do
                    {                        
                        if (f1 > f2)
                        {
                            leftBorder = x1;
                            x1 = x2;
                            f1 = f2;
                            x2 = rightBorder - goldNumber * (rightBorder - leftBorder);
                            f2 = function(x2);
                        }
                        else
                        {
                            rightBorder = x2;
                            x2 = x1;
                            f2 = f1;
                            x1 = leftBorder + goldNumber * (rightBorder - leftBorder);
                            f1 = function(x1);
                        }                       
                    } while (Math.Abs(leftBorder - rightBorder) > Epsilon);
                    break;
                case ExtremumType.Maximum:
                    do
                    {
                        if (f1 < f2)
                        {
                            leftBorder = x1;
                            x1 = x2;
                            f1 = f2;
                            x2 = leftBorder + (2.0 - goldNumber) * (rightBorder - leftBorder);
                            f2 = function(x2);
                        }
                        else
                        {
                            rightBorder = x2;
                            x2 = x1;
                            f2 = f1;
                            x1 = leftBorder - goldNumber * (rightBorder - leftBorder);
                            f1 = function(x1);
                        }
                    } while (Math.Abs(leftBorder - rightBorder) > Epsilon);
                    break;
            }
            left = leftBorder;
            right = rightBorder;
        }
        public static void FindInterval(ExtremumType type, double x0, Func<double, double> function, out double left, out double right)
        {
            double xk = 0.0;
            uint i = 1;
            double delta = Epsilon;

            switch(type)
            {
                case ExtremumType.Minimum:
                    if (function(x0) > function(x0 - delta))
                        delta *= -1;
                    xk = x0 + delta;
                    do
                    {
                        i++;
                        x0 = xk;
                        xk = x0 + (FastPow(2, i) - 1.0) * delta;
                    } while (function(x0) > function(xk));
                    break;
                case ExtremumType.Maximum:
                    if (function(x0) < function(x0 - delta))
                        delta *= -1;
                    xk = x0 + delta;
                    do
                    {
                        i++;
                        x0 = xk;
                        xk = x0 + (FastPow(2, i) - 1.0) * delta;
                    } while (function(x0) < function(xk));
                    break;
            }

            if(delta > 0)
            {
                left = x0 + (FastPow(2, i - 2) - 1.0) * delta;
                right = xk;
            }
            else
            {
                left = xk;
                right = x0 + (FastPow(2, i - 2) - 1.0) * delta;
            }
        }
        public static void FibonacciMethod(ExtremumType type, double leftBorder, double rightBorder, Func<double, double> function, out double left, out double right)
        {
            int n = 1;
            int k = 2;
            long Fn2 = 1;
            double x1, x2;
            double f1, f2;

            left = leftBorder;
            right = rightBorder;

            List<long> fibonacciNumbers = new List<long>();
            fibonacciNumbers.Add(1);
            fibonacciNumbers.Add(1);

            while(Fn2 < (rightBorder - leftBorder) / Epsilon)
            {
                Fn2 = fibonacciNumbers[n] + fibonacciNumbers[n - 1];
                fibonacciNumbers.Add(Fn2);
                n++;
            }
            n -= 2;

            Fn2 = fibonacciNumbers[n + 2];
            x1 = leftBorder + fibonacciNumbers[n] / (double)Fn2 * (rightBorder - leftBorder);
            x2 = leftBorder + fibonacciNumbers[n + 1] / (double)Fn2 * (rightBorder - leftBorder);
            f1 = function(x1);
            f2 = function(x2);

            switch(type)
            {
                case ExtremumType.Minimum:
                    while (k <= n)
                    {
                        if (f1 > f2)
                        {
                            left = x1;
                            x1 = x2;
                            f1 = f2;
                            x2 = left + fibonacciNumbers[n - k + 2] / (double) Fn2 * (rightBorder - leftBorder);
                            f2 = function(x2);
                        }
                        else
                        {
                            right = x2;
                            x2 = x1;
                            f2 = f1;
                            x1 = left + fibonacciNumbers[n - k + 1] / (double) Fn2 * (rightBorder - leftBorder);
                            f1 = function(x1);
                        }
                        k++;
                    }
                    break;
                case ExtremumType.Maximum:
                    while (k <= n)
                    {
                        if (f1 < f2)
                        {
                            left = x1;
                            x1 = x2;
                            f1 = f2;
                            x2 = left + fibonacciNumbers[n - k + 2] / Fn2 * (rightBorder - leftBorder);
                            f2 = function(x2);
                        }
                        else
                        {
                            right = x2;
                            x2 = x1;
                            f2 = f1;
                            x1 = left + fibonacciNumbers[n - k + 1] / Fn2 * (rightBorder - leftBorder);
                            f1 = function(x1);
                        }
                        k++;
                    }
                    break;
            }
        }
        public static int DichotomyMethodResearch(double leftBorder, double rightBorder, Func<double, double> function, out double left, out double right)
        {
            double delta = Epsilon / 2.0;
            double x1, x2;
            double f1, f2;
            int count = 0;

            Application app = new Application();
            Workbook wb = app.Workbooks.Add();
            Worksheet ws = wb.ActiveSheet;

            ws.Range["A1"].Value = "i";
            ws.Range["B1"].Value = "x1";
            ws.Range["C1"].Value = "x2";
            ws.Range["D1"].Value = "f(x1)";
            ws.Range["E1"].Value = "f(x2)";
            ws.Range["F1"].Value = "ai";
            ws.Range["G1"].Value = "bi";
            ws.Range["H1"].Value = "bi - ai";
            ws.Range["I1"].Value = "(b(i-1) - a(i-1))/(bi - ai)";

            int i = 2;
            double prev = rightBorder - leftBorder;
            do
            {
                x1 = (leftBorder + rightBorder - delta) / 2.0;
                x2 = (leftBorder + rightBorder + delta) / 2.0;
                f1 = function(x1);
                f2 = function(x2);

                if (f1 > f2)
                    leftBorder = x1;
                else rightBorder = x2;
                count += 2;

                ws.Cells[i, 1].Value = i - 1;
                ws.Cells[i, 2].Value = x1;
                ws.Cells[i, 3].Value = x2;
                ws.Cells[i, 4].Value = f1;
                ws.Cells[i, 5].Value = f2;
                ws.Cells[i, 6].Value = leftBorder;
                ws.Cells[i, 7].Value = rightBorder;
                ws.Cells[i, 8].Value = (rightBorder - leftBorder);
                ws.Cells[i, 9].Value = (prev / (rightBorder - leftBorder));

                i++;
                prev = rightBorder - leftBorder;
            } while (Math.Abs(leftBorder - rightBorder) > Epsilon);

            wb.SaveAs(Path.Combine(Environment.CurrentDirectory, "dichotomy" + ((int)Math.Log10(Epsilon)) + ".xlsx"));
            wb.Close();

            left = leftBorder;
            right = rightBorder;
            return count;
        }
        public static int GoldenRatioMethodResearch(double leftBorder, double rightBorder, Func<double, double> function, out double left, out double right)
        {
            double goldNumber = (3.0 - sqrt5) / 2.0;

            double x1 = leftBorder + goldNumber * (rightBorder - leftBorder);
            double x2 = rightBorder - goldNumber * (rightBorder - leftBorder);

            double f1 = function(x1);
            double f2 = function(x2);

            int count = 2;

            Application app = new Application();
            Workbook wb = app.Workbooks.Add();
            Worksheet ws = wb.ActiveSheet;

            ws.Range["A1"].Value = "i";
            ws.Range["B1"].Value = "x1";
            ws.Range["C1"].Value = "x2";
            ws.Range["D1"].Value = "f(x1)";
            ws.Range["E1"].Value = "f(x2)";
            ws.Range["F1"].Value = "ai";
            ws.Range["G1"].Value = "bi";
            ws.Range["H1"].Value = "bi - ai";
            ws.Range["I1"].Value = "(b(i-1) - a(i-1))/(bi - ai)";

            int i = 2;
            double prev = rightBorder - leftBorder;

            do
            {
                ws.Cells[i, 1].Value = i - 1;
                ws.Cells[i, 2].Value = x1;
                ws.Cells[i, 3].Value = x2;
                ws.Cells[i, 4].Value = f1;
                ws.Cells[i, 5].Value = f2;

                if (f1 > f2)
                {
                    leftBorder = x1;
                    x1 = x2;
                    f1 = f2;
                    x2 = rightBorder - goldNumber * (rightBorder - leftBorder);
                    f2 = function(x2);
                }
                else
                {
                    rightBorder = x2;
                    x2 = x1;
                    f2 = f1;
                    x1 = leftBorder + goldNumber * (rightBorder - leftBorder);
                    f1 = function(x1);
                }

                count++;

                ws.Cells[i, 6].Value = leftBorder;
                ws.Cells[i, 7].Value = rightBorder;
                ws.Cells[i, 8].Value = (rightBorder - leftBorder);
                ws.Cells[i, 9].Value = (prev / (rightBorder - leftBorder));
                prev = rightBorder - leftBorder;

                i++;
            } while (Math.Abs(leftBorder - rightBorder) > Epsilon);

            wb.SaveAs(Path.Combine(Environment.CurrentDirectory, "goldenRatio" + ((int)Math.Log10(Epsilon)) + ".xlsx"));
            wb.Close();

            left = leftBorder;
            right = rightBorder;
            return count;
        }
        public static int FibonacciMethodResearch(double leftBorder, double rightBorder, Func<double, double> function, out double left, out double right)
        {
            int n = -1;
            int k = 1;
            long Fn2 = 1;
            double x1, x2;
            double f1, f2;
            int count = 2;

            Application app = new Application();
            Workbook wb = app.Workbooks.Add();
            Worksheet ws = wb.ActiveSheet;

            ws.Range["A1"].Value = "i";
            ws.Range["B1"].Value = "x1";
            ws.Range["C1"].Value = "x2";
            ws.Range["D1"].Value = "f(x1)";
            ws.Range["E1"].Value = "f(x2)";
            ws.Range["F1"].Value = "ai";
            ws.Range["G1"].Value = "bi";
            ws.Range["H1"].Value = "bi - ai";
            ws.Range["I1"].Value = "(b(i-1) - a(i-1))/(bi - ai)";

            left = leftBorder;
            right = rightBorder;

            List<long> fibonacciNumbers = new List<long>();
            fibonacciNumbers.Add(1);
            fibonacciNumbers.Add(1);

            while (Fn2 <= (rightBorder - leftBorder) / Epsilon)
            {
                Fn2 = fibonacciNumbers[n + 2] + fibonacciNumbers[n + 1];
                fibonacciNumbers.Add(Fn2);
                n++;
            }

            x1 = leftBorder + fibonacciNumbers[n] / (double)Fn2 * (rightBorder - leftBorder);
            x2 = leftBorder + fibonacciNumbers[n + 1] / (double)Fn2 * (rightBorder - leftBorder);
            f1 = function(x1);
            f2 = function(x2);

            double prev = rightBorder - leftBorder;

            while (k - 1 <= n)
            {
                ws.Cells[k + 1, 1].Value = k;
                ws.Cells[k + 1, 2].Value = x1;
                ws.Cells[k + 1, 3].Value = x2;
                ws.Cells[k + 1, 4].Value = f1;
                ws.Cells[k + 1, 5].Value = f2;

                if (f1 > f2)
                {
                    left = x1;
                    x1 = x2;
                    f1 = f2;
                    x2 = left + fibonacciNumbers[n - k + 2] / (double)Fn2 * (rightBorder - leftBorder);
                    f2 = function(x2);
                }
                else
                {
                    right = x2;
                    x2 = x1;
                    f2 = f1;
                    x1 = left + fibonacciNumbers[n - k + 1] / (double)Fn2 * (rightBorder - leftBorder);
                    f1 = function(x1);
                }

                ws.Cells[k + 1, 6].Value = left;
                ws.Cells[k + 1, 7].Value = right;
                ws.Cells[k + 1, 8].Value = (right - left);
                ws.Cells[k + 1, 9].Value = (prev / (right - left));

                count++;
                prev = right - left;

                k++;
            }

            wb.SaveAs(Path.Combine(Environment.CurrentDirectory, "fibonacci" + ((int)Math.Log10(Epsilon)) + ".xlsx"));
            wb.Close();
            return count;
        }
        public static void FindIntervalResearch(double x0, Func<double, double> function, out double left, out double right)
        {
            double xk = 0.0;
            uint i = 1;
            double delta = Epsilon;

            Application app = new Application();
            Workbook wb = app.Workbooks.Add();
            Worksheet ws = wb.ActiveSheet;

            ws.Range["A1"].Value = "i";
            ws.Range["B1"].Value = "xi";
            ws.Range["C1"].Value = "f(xi)";

            if (function(x0) > function(x0 - delta))
                delta *= -1;
            else if(function(x0) < function(x0 + delta))
            {
                left = x0 - delta;
                right = x0 + delta;

                ws.Cells[2, 1] = 1;
                ws.Cells[2, 2] = left;
                ws.Cells[2, 3] = function(left);

                ws.Cells[3, 1] = 2;
                ws.Cells[3, 2] = right;
                ws.Cells[3, 3] = function(right);

                wb.SaveAs(Path.Combine(Environment.CurrentDirectory, "findInterval" + ((int)Math.Log10(Epsilon)) + ".xlsx"));
                wb.Close();

                return;
            }

            xk = x0 + delta;
            ws.Cells[2, 1] = 1;
            ws.Cells[2, 2] = xk;
            ws.Cells[2, 3] = function(x0 - delta);

            double xk1;
            do
            {
                i++;
                xk1 = xk;
                double power = FastPow(2, i);
                xk = x0 + (power - 1.0) * delta;

                ws.Cells[i + 1, 1].Value = i;
                ws.Cells[i + 1, 2].Value = xk;
                ws.Cells[i + 1, 3].Value = function(xk);
            } while (function(xk) < function(xk1));

            if (delta > 0)
            {
                left = x0 + (FastPow(2, i - 2) - 1.0) * delta;
                right = xk;
            }
            else
            {
                left = xk;
                right = x0 + (FastPow(2, i - 2) - 1.0) * delta;
            }

            wb.SaveAs(Path.Combine(Environment.CurrentDirectory, "findInterval" + ((int)Math.Log10(Epsilon)) + ".xlsx"));
            wb.Close();
        }
    }
}
