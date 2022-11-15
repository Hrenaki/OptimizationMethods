using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NumMath;
using EParser;
using Microsoft.Office.Interop.Excel;
using Accord.Statistics;

using App = Microsoft.Office.Interop.Excel.Application;
using System.IO;
using Accord.Statistics.Distributions.Univariate;
using Accord.Math.Random;

namespace OptimizationMethods
{
    class Program
    {
        static void Main(string[] args)
        {
            double[] C = new double[] { 6, 2, 4, 2, 8, 8 };
            double[] a = new double[] { -3, 4, -8, -6, 3, -6 };
            double[] b = new double[] { 9, -7, 3, -9, -2, -8 };

            Func function = t =>
            {
                double result = 0;
                for (int i = 0; i < C.Length; i++)
                    result += C[i] / (1 + (t[0] - a[i]) * (t[0] - a[i]) + (t[1] - b[i]) * (t[1] - b[i]));
                return result;
            };

            //Rectangle rectangle = new Rectangle(-10, 10, -10, 10);

            //StatisticalSearches.Epsilon = 1E-3;
            //StatisticalSearches.P = 1E-3;
            //DescentMethods.Epsilon = 1E-3;
            //OneDimensionalSearches.Epsilon = 1E-3;
            //Vector point = StatisticalSearches.Alg3(ExtremumType.Maximum, rectangle, function, 100, out _);

            //MethodComparison(ExtremumType.Maximum, rectangle, function);
        }

        //static void RandomSearch_table(ExtremumType type, Area area, Func function)
        //{
        //    Application app = new Application();
        //    Workbook wb = app.Workbooks.Add();
        //    Worksheet ws = wb.ActiveSheet;
        //
        //    ws.Range["A1"].Value = "EPS";
        //    ws.Range["B1"].Value = "P";
        //    ws.Range["C1"].Value = "N";
        //    ws.Range["D1"].Value = "(x, y)";
        //    ws.Range["E1"].Value = "f(x, y)";
        //
        //    double[] epsilons = new double[] { 1E-1, 1E-2, 1E-3 };
        //    double[] P = new double[] { 1E-1, 1E-2, 1E-3, 1E-4, 1E-5 };
        //
        //    Vector point;
        //    long N;
        //    int pos = 2;
        //    for(int i = 0; i < epsilons.Length; i++)
        //    {
        //        StatisticalSearches.Epsilon = epsilons[i];
        //
        //        for(int j = 0; j < P.Length; j++, pos++)
        //        {
        //            StatisticalSearches.P = P[j];
        //            point = StatisticalSearches.RandomSearch(type, area, function, out N);
        //
        //            ws.Cells[pos, 1].Value = StatisticalSearches.Epsilon.ToString("E4");
        //            ws.Cells[pos, 2].Value = StatisticalSearches.P.ToString("E4");
        //            ws.Cells[pos, 3].Value = N;
        //            ws.Cells[pos, 4].Value = point.ToString("E4");
        //            ws.Cells[pos, 5].Value = function(point.values).ToString("E4");
        //        }
        //    }
        //
        //    ws.Columns.AutoFit();
        //    wb.SaveAs(Path.Combine(Environment.CurrentDirectory, "RandomSearch_table.xlsx"));
        //    wb.Close();
        //}
        //static void MethodComparison(ExtremumType type, Area area, Func function)
        //{
        //    StatisticalSearches.Epsilon = 1E-3;
        //    StatisticalSearches.P = 1E-3;
        //    DescentMethods.Epsilon = 1E-3;
        //    OneDimensionalSearches.Epsilon = 1E-3;
        //
        //    Application app = new Application();
        //    Workbook wb = app.Workbooks.Add();
        //    Worksheet ws = wb.ActiveSheet;
        //
        //    ws.Range["A1"].Value = "seed";
        //    ws.Range["B1"].Value = "m";
        //    ws.Range["C1"].Value = "Method";
        //    ws.Range["D1"].Value = "(x, y)";
        //    ws.Range["E1"].Value = "f(x, y)";
        //    ws.Range["F1"].Value = "calc count";
        //
        //    int calcCount;
        //    int pos = 2;
        //    int[] m = new int[] { 2, 4, 6, 8 };
        //    int[] seeds = new int[] { 876, 98796, 3665 };
        //
        //    for (int j = 0; j < seeds.Length; j++)
        //    {
        //        
        //        ws.Cells[pos, 1].Value = seeds[j];
        //
        //        for (int i = 0; i < m.Length; i++)
        //        {
        //            ws.Cells[pos, 2].Value = m[i];
        //
        //            Generator.Seed = seeds[j];
        //            Vector point = StatisticalSearches.Alg1(type, area, function, m[i], out calcCount);
        //            ws.Cells[pos, 3].Value = "Alg1";
        //            ws.Cells[pos, 4].Value = point.ToString("E4");
        //            ws.Cells[pos, 5].Value = function(point.values).ToString("E4");
        //            ws.Cells[pos, 6].Value = calcCount;
        //            pos++;
        //
        //            Generator.Seed = seeds[j];
        //            point = StatisticalSearches.Alg2(type, area, function, m[i], out calcCount);
        //            ws.Cells[pos, 3].Value = "Alg2";
        //            ws.Cells[pos, 4].Value = point.ToString("E4");
        //            ws.Cells[pos, 5].Value = function(point.values).ToString("E4");
        //            ws.Cells[pos, 6].Value = calcCount;
        //            pos++;
        //
        //            Generator.Seed = seeds[j];
        //            point = StatisticalSearches.Alg3(type, area, function, m[i], out calcCount);
        //            ws.Cells[pos, 3].Value = "Alg3";
        //            ws.Cells[pos, 4].Value = point.ToString("E4");
        //            ws.Cells[pos, 5].Value = function(point.values).ToString("E4");
        //            ws.Cells[pos, 6].Value = calcCount;
        //            pos++;
        //        }
        //    }
        //
        //    ws.Columns.AutoFit();
        //    wb.SaveAs(Path.Combine(Environment.CurrentDirectory, "MethodComparison.xlsx"));
        //    wb.Close();
        //}

        static void CGM_research_x0(Func f, int num, Vector result_point, ExtremumType type)
        {
            Application app = new Application();
            Workbook wb = app.Workbooks.Add();
            Worksheet ws = wb.ActiveSheet;

            Vector[] points = new Vector[] { new Vector(1, -1), new Vector(1, 2), new Vector(10, 10), new Vector(0, 0) };
            Vector start;

            ws.Range["A1"].Value = "x0";
            ws.Range["B1"].Value = "xk";
            ws.Range["C1"].Value = "f(xk)";
            ws.Range["D1"].Value = "кол-во итераций";
            ws.Range["E1"].Value = "кол-во выч. функции";
            ws.Range["F1"].Value = "норма s_k+1";

            int calc_count = 0;
            int step;
            double norm_s;

            for (int i = 0; i < points.Length; i++)
            {
                start = points[i];
                ws.Cells[i + 2, 1].Value = start[0].ToString("E5") + " " + start[1].ToString("E5");
                step = DescentMethods.CGM_FletcherReeves(type, f, start, out calc_count, out norm_s);
                ws.Cells[i + 2, 2].Value = start[0].ToString("E5") + " " + start[1].ToString("E5");
                ws.Cells[i + 2, 3].Value = f(start.values).ToString("E5");
                ws.Cells[i + 2, 4].Value = step;
                ws.Cells[i + 2, 5].Value = calc_count;
                ws.Cells[i + 2, 6].Value = Math.Sqrt(norm_s).ToString("E5");
            }

            ws.Columns.AutoFit();
            wb.SaveAs(Path.Combine(Environment.CurrentDirectory, "CGM" + num + "_x0_research.xlsx"));
            wb.Close();
        }
        static void Pierson_research_x0(Func f, int num, Vector result_point, ExtremumType type)
        {
            Application app = new Application();
            Workbook wb = app.Workbooks.Add();
            Worksheet ws = wb.ActiveSheet;

            Vector[] points = new Vector[] { new Vector(1, -1), new Vector(1, 2), new Vector(10, 10), new Vector(0, 0) };
            Vector start;

            ws.Range["A1"].Value = "x0";
            ws.Range["B1"].Value = "xk";
            ws.Range["C1"].Value = "f(xk)";
            ws.Range["D1"].Value = "кол-во итераций";
            ws.Range["E1"].Value = "кол-во выч. функции";
            ws.Range["F1"].Value = "норма s_k+1";

            int calc_count = 0;
            int step;
            double norm_s;

            for (int i = 0; i < points.Length; i++)
            {
                start = points[i];
                ws.Cells[i + 2, 1].Value = start[0].ToString("E5") + " " + start[1].ToString("E5");
                step = DescentMethods.PiersonAlgorithm(type, f, start, out calc_count, out norm_s);
                ws.Cells[i + 2, 2].Value = start[0].ToString("E5") + " " + start[1].ToString("E5");
                ws.Cells[i + 2, 3].Value = f(start.values).ToString("E5");
                ws.Cells[i + 2, 4].Value = step;
                ws.Cells[i + 2, 5].Value = calc_count;
                ws.Cells[i + 2, 6].Value = Math.Sqrt(norm_s).ToString("E5");
            }

            ws.Columns.AutoFit();
            wb.SaveAs(Path.Combine(Environment.CurrentDirectory, "Pierson" + num + "_x0_research.xlsx"));
            wb.Close();
        }

        static void CGM_research_eps(Func f, int num, ExtremumType type)
        {
            Application app = new Application();
            Workbook wb = app.Workbooks.Add();
            Worksheet ws = wb.ActiveSheet;

            Vector start = new Vector(1, -1);
            Vector temp;

            ws.Range["A1"].Value = "eps";
            ws.Range["B1"].Value = "x0";
            ws.Range["C1"].Value = "xk";
            ws.Range["D1"].Value = "f(xk)";
            ws.Range["E1"].Value = "кол-во итераций";
            ws.Range["F1"].Value = "кол-во выч. функции";
            ws.Range["G1"].Value = "норма (M_k+1 - M_k)";

            int calc_count = 0;
            int step;
            int pos = 2;
            double norm_s;

            double[] eps = new double[] { 1E-3, 1E-4, 1E-5, 1E-6, 1E-7};
            for(int i = 0; i < eps.Length; i++, pos++)
            {
                DescentMethods.Epsilon = eps[i];
                temp = new Vector(start[0], start[1]);

                ws.Cells[pos, 1].Value = eps[i].ToString("E5");
                ws.Cells[pos, 2].Value = temp[0].ToString("E5") + " " + temp[1].ToString("E5");
                step = DescentMethods.CGM_FletcherReeves(type, f, temp, out calc_count, out norm_s);
                ws.Cells[pos, 3].Value = temp[0].ToString("E5") + " " + temp[1].ToString("E5");
                ws.Cells[pos, 4].Value = f(temp.values).ToString("E5");
                ws.Cells[pos, 5].Value = step;
                ws.Cells[pos, 6].Value = calc_count;
                ws.Cells[pos, 7].Value = Math.Sqrt(norm_s).ToString("E5");
            }

            ws.Columns.AutoFit();
            wb.SaveAs(Path.Combine(Environment.CurrentDirectory, "CGM" + num + "_eps_research.xlsx"));
            wb.Close();
        }
        static void Pierson_research_eps(Func f, int num, ExtremumType type)
        {
            Application app = new Application();
            Workbook wb = app.Workbooks.Add();
            Worksheet ws = wb.ActiveSheet;

            Vector start = new Vector(1, -1);
            Vector temp;

            ws.Range["A1"].Value = "eps";
            ws.Range["B1"].Value = "x0";
            ws.Range["C1"].Value = "xk";
            ws.Range["D1"].Value = "f(xk)";
            ws.Range["E1"].Value = "кол-во итераций";
            ws.Range["F1"].Value = "кол-во выч. функции";
            ws.Range["G1"].Value = "норма s_k+1";

            int calc_count = 0;
            int step;
            int pos = 2;
            double norm_s;

            double[] eps = new double[] { 1E-3, 1E-4, 1E-5, 1E-6, 1E-7 };
            for (int i = 0; i < eps.Length; i++, pos++)
            {
                DescentMethods.Epsilon = eps[i];
                temp = new Vector(start[0], start[1]);

                ws.Cells[pos, 1].Value = eps[i].ToString("E5");
                ws.Cells[pos, 2].Value = temp[0].ToString("E5") + " " + temp[1].ToString("E5");
                step = DescentMethods.PiersonAlgorithm(type, f, temp, out calc_count, out norm_s);
                ws.Cells[pos, 3].Value = temp[0].ToString("E5") + " " + temp[1].ToString("E5");
                ws.Cells[pos, 4].Value = f(temp.values).ToString("E5");
                ws.Cells[pos, 5].Value = step;
                ws.Cells[pos, 6].Value = calc_count;
                ws.Cells[pos, 7].Value = Math.Sqrt(norm_s).ToString("E5");
            }

            ws.Columns.AutoFit();
            wb.SaveAs(Path.Combine(Environment.CurrentDirectory, "Pierson" + num + "_eps_research.xlsx"));
            wb.Close();
        }

        /*
        static void PenaltyMethod_research_func_H()
        {          
            Func<double, double> H_f1 = t => Math.Abs(t);
            Func<double, double> H_f2 = t => Math.Abs(t) * Math.Abs(t);
            Func<double, double> H_f3 = t => Math.Abs(t) * Math.Abs(t) * Math.Abs(t) * Math.Abs(t);
            Func<double, double>[] funcs = new Func<double, double>[] { H_f1, H_f2, H_f3 };

            Func<double, double> G = t => (t + Math.Abs(t)) / 2.0;

            Func func = t => (t[0] - t[1]) * (t[0] - t[1]) + 10.0 * (t[0] + 5.0) * (t[0] + 5.0);
            Func<double, double> func_r = t => 2.0 * t;

            Func[] equations = new Func[] { t => 1.0 - t[0] - t[1] };
            Func[] inequations = new Func[] { };

            int step;
            int calc_count;
            int pos = 2;

            Application app = new Application();
            Workbook wb = app.Workbooks.Add();
            Worksheet ws = wb.ActiveSheet;

            ws.Range["A1"].Value = "N";
            ws.Range["B1"].Value = "r0";
            ws.Range["C1"].Value = "g(r)";
            ws.Range["D1"].Value = "x0";
            ws.Range["E1"].Value = "eps";
            ws.Range["F1"].Value = "кол-во итераций";
            ws.Range["G1"].Value = "кол-во выч. функции";
            ws.Range["H1"].Value = "xk";
            ws.Range["I1"].Value = "f(xk)";

            for (int i = 0; i < funcs.Length; i++, pos++)
            {
                Vector start = new Vector(-1, -1);
                Vector r = new Vector(1.0);

                step = PenaltyMethods.PenaltyMethod(func, start, r, func_r, funcs[i], G, equations, inequations, out calc_count, out _);

                ws.Cells[pos, 1].Value = (i + 1).ToString();
                ws.Cells[pos, 2].Value = 1.0.ToString();
                ws.Cells[pos, 3].Value = "2.0 * r";
                ws.Cells[pos, 4].Value = "(-1, -1)";
                ws.Cells[pos, 5].Value = PenaltyMethods.Epsilon.ToString("E5");
                ws.Cells[pos, 6].Value = step;
                ws.Cells[pos, 7].Value = calc_count;
                ws.Cells[pos, 8].Value = start.ToString("E5");
                ws.Cells[pos, 9].Value = func(start.values).ToString("E5");
            }
            ws.Columns.AutoFit();
            wb.SaveAs(Path.Combine(Environment.CurrentDirectory, "PenaltyMethod_funcH.xlsx"));
            wb.Close();
        }
        static void PenaltyMethod_research_func_G()
        {
            Func<double, double> G_f1 = t => 0.5 * (t + Math.Abs(t));
            Func<double, double> G_f2 = t => (0.5 * (t + Math.Abs(t))) * (0.5 * (t + Math.Abs(t)));
            Func<double, double> G_f3 = t => (0.5 * (t + Math.Abs(t))) * (0.5 * (t + Math.Abs(t))) * (0.5 * (t + Math.Abs(t))) * (0.5 * (t + Math.Abs(t)));
            Func<double, double>[] funcs = new Func<double, double>[] { G_f1, G_f2, G_f3 };

            Func<double, double> H = t => Math.Abs(t);

            Func func = t => (t[0] - t[1]) * (t[0] - t[1]) + 10.0 * (t[0] + 5.0) * (t[0] + 5.0);
            Func<double, double> func_r = t => 2.0 * t;

            Func[] equations = new Func[] { };
            Func[] inequations = new Func[] { t => -t[0] - t[1]};

            int step;
            int calc_count;
            int pos = 2;

            Application app = new Application();
            Workbook wb = app.Workbooks.Add();
            Worksheet ws = wb.ActiveSheet;

            ws.Range["A1"].Value = "N";
            ws.Range["B1"].Value = "r0";
            ws.Range["C1"].Value = "g(r)";
            ws.Range["D1"].Value = "x0";
            ws.Range["E1"].Value = "eps";
            ws.Range["F1"].Value = "кол-во итераций";
            ws.Range["G1"].Value = "кол-во выч. функции";
            ws.Range["H1"].Value = "xk";
            ws.Range["I1"].Value = "f(xk)";

            for (int i = 0; i < funcs.Length; i++, pos++)
            {
                Vector start = new Vector(-1, -1);
                Vector r = new Vector(1.0);

                step = PenaltyMethods.PenaltyMethod(func, start, r, func_r, H, funcs[i], equations, inequations, out calc_count, out _);

                ws.Cells[pos, 1].Value = (i + 1).ToString();
                ws.Cells[pos, 2].Value = 1.0.ToString();
                ws.Cells[pos, 3].Value = "2.0 * r";
                ws.Cells[pos, 4].Value = "(-1, -1)";
                ws.Cells[pos, 5].Value = PenaltyMethods.Epsilon.ToString("E5");
                ws.Cells[pos, 6].Value = step;
                ws.Cells[pos, 7].Value = calc_count;
                ws.Cells[pos, 8].Value = start.ToString("E5");
                ws.Cells[pos, 9].Value = func(start.values).ToString("E5");
            }
            ws.Columns.AutoFit();
            wb.SaveAs(Path.Combine(Environment.CurrentDirectory, "PenaltyMethod_funcG.xlsx"));
            wb.Close();
        }

        static void PenaltyMethod_research_func_1_r0()
        {
            Func<double, double> G = t => 0.5 * (t + Math.Abs(t));
            Func<double, double> H = t => Math.Abs(t);

            Func func = t => (t[0] - t[1]) * (t[0] - t[1]) + 10.0 * (t[0] + 5.0) * (t[0] + 5.0);
            Func<double, double> func_r = t => 2.0 * t;

            Func[] equations = new Func[] { };
            Func[] inequations = new Func[] { t => -t[0] - t[1] };

            double[] r_arr = new double[] { 1.0, 10.0, 100.0 };

            int step;
            int calc_count;
            int pos = 2;

            Application app = new Application();
            Workbook wb = app.Workbooks.Add();
            Worksheet ws = wb.ActiveSheet;

            ws.Range["A1"].Value = "N";
            ws.Range["B1"].Value = "r0";
            ws.Range["C1"].Value = "g(r)";
            ws.Range["D1"].Value = "x0";
            ws.Range["E1"].Value = "eps";
            ws.Range["F1"].Value = "кол-во итераций";
            ws.Range["G1"].Value = "кол-во выч. функции";
            ws.Range["H1"].Value = "xk";
            ws.Range["I1"].Value = "f(xk)";

            for (int i = 0; i < r_arr.Length; i++, pos++)
            {
                Vector start = new Vector(-1, -1);
                Vector r = new Vector(r_arr[i]);

                step = PenaltyMethods.PenaltyMethod(func, start, r, func_r, H, G, equations, inequations, out calc_count, out _);

                ws.Cells[pos, 1].Value = 1.ToString();
                ws.Cells[pos, 2].Value = r_arr[i].ToString();
                ws.Cells[pos, 3].Value = "2.0 * r";
                ws.Cells[pos, 4].Value = "(-1, -1)";
                ws.Cells[pos, 5].Value = PenaltyMethods.Epsilon.ToString("E5");
                ws.Cells[pos, 6].Value = step;
                ws.Cells[pos, 7].Value = calc_count;
                ws.Cells[pos, 8].Value = start.ToString("E5");
                ws.Cells[pos, 9].Value = func(start.values).ToString("E5");
            }
            ws.Columns.AutoFit();
            wb.SaveAs(Path.Combine(Environment.CurrentDirectory, "PenaltyMethod_func_1_r0.xlsx"));
            wb.Close();
        }
        static void PenaltyMethod_research_func_2_r0()
        {
            Func<double, double> G = t => 0.5 * (t + Math.Abs(t));
            Func<double, double> H = t => Math.Abs(t);

            Func func = t => (t[0] - t[1]) * (t[0] - t[1]) + 10.0 * (t[0] + 5.0) * (t[0] + 5.0);
            Func<double, double> func_r = t => 2.0 * t;

            Func[] equations = new Func[] { t => 1.0 - t[0] - t[1] };
            Func[] inequations = new Func[] { };

            double[] r_arr = new double[] { 1.0, 10.0, 100.0 };

            int step;
            int calc_count;
            int pos = 2;

            Application app = new Application();
            Workbook wb = app.Workbooks.Add();
            Worksheet ws = wb.ActiveSheet;

            ws.Range["A1"].Value = "N";
            ws.Range["B1"].Value = "r0";
            ws.Range["C1"].Value = "g(r)";
            ws.Range["D1"].Value = "x0";
            ws.Range["E1"].Value = "eps";
            ws.Range["F1"].Value = "кол-во итераций";
            ws.Range["G1"].Value = "кол-во выч. функции";
            ws.Range["H1"].Value = "xk";
            ws.Range["I1"].Value = "f(xk)";

            for (int i = 0; i < r_arr.Length; i++, pos++)
            {
                Vector start = new Vector(-1, -1);
                Vector r = new Vector(r_arr[i]);

                step = PenaltyMethods.PenaltyMethod(func, start, r, func_r, H, G, equations, inequations, out calc_count, out _);

                ws.Cells[pos, 1].Value = 1.ToString();
                ws.Cells[pos, 2].Value = r_arr[i].ToString();
                ws.Cells[pos, 3].Value = "2.0 * r";
                ws.Cells[pos, 4].Value = "(-1, -1)";
                ws.Cells[pos, 5].Value = PenaltyMethods.Epsilon.ToString("E5");
                ws.Cells[pos, 6].Value = step;
                ws.Cells[pos, 7].Value = calc_count;
                ws.Cells[pos, 8].Value = start.ToString("E5");
                ws.Cells[pos, 9].Value = func(start.values).ToString("E5");
            }
            ws.Columns.AutoFit();
            wb.SaveAs(Path.Combine(Environment.CurrentDirectory, "PenaltyMethod_func_2_r0.xlsx"));
            wb.Close();
        }

        static void PenaltyMethod_research_func_1_r()
        {
            Func<double, double> G = t => 0.5 * (t + Math.Abs(t));
            Func<double, double> H = t => Math.Abs(t);

            Func func = t => (t[0] - t[1]) * (t[0] - t[1]) + 10.0 * (t[0] + 5.0) * (t[0] + 5.0);
            Func<double, double> func_r = t => 2.0 * t;

            Func[] equations = new Func[] { };
            Func[] inequations = new Func[] { t => -t[0] - t[1] };

            Func<double, double>[] r_funcs = new Func<double, double>[] { t => 2.0 * t, t => 10 * t, t => t * t };

            int step;
            int calc_count;
            int pos = 2;

            Application app = new Application();
            Workbook wb = app.Workbooks.Add();
            Worksheet ws = wb.ActiveSheet;

            ws.Range["A1"].Value = "N";
            ws.Range["B1"].Value = "r0";
            ws.Range["C1"].Value = "g(r)";
            ws.Range["D1"].Value = "x0";
            ws.Range["E1"].Value = "eps";
            ws.Range["F1"].Value = "кол-во итераций";
            ws.Range["G1"].Value = "кол-во выч. функции";
            ws.Range["H1"].Value = "xk";
            ws.Range["I1"].Value = "f(xk)";

            for (int i = 0; i < r_funcs.Length; i++, pos++)
            {
                Vector start = new Vector(-1, -1);
                Vector r = new Vector(2.0);

                step = PenaltyMethods.PenaltyMethod(func, start, r, r_funcs[i], H, G, equations, inequations, out calc_count, out _);

                ws.Cells[pos, 1].Value = 1.ToString();
                ws.Cells[pos, 2].Value = 2.ToString();
                ws.Cells[pos, 3].Value = i.ToString();
                ws.Cells[pos, 4].Value = "(-1, -1)";
                ws.Cells[pos, 5].Value = PenaltyMethods.Epsilon.ToString("E5");
                ws.Cells[pos, 6].Value = step;
                ws.Cells[pos, 7].Value = calc_count;
                ws.Cells[pos, 8].Value = start.ToString("E5");
                ws.Cells[pos, 9].Value = func(start.values).ToString("E5");
            }
            ws.Columns.AutoFit();
            wb.SaveAs(Path.Combine(Environment.CurrentDirectory, "PenaltyMethod_func_1_r.xlsx"));
            wb.Close();
        }
        static void PenaltyMethod_research_func_2_r()
        {
            Func<double, double> G = t => 0.5 * (t + Math.Abs(t));
            Func<double, double> H = t => Math.Abs(t);

            Func func = t => (t[0] - t[1]) * (t[0] - t[1]) + 10.0 * (t[0] + 5.0) * (t[0] + 5.0);
            Func<double, double> func_r = t => 2.0 * t;

            Func[] equations = new Func[] { t => 1.0 - t[0] - t[1] };
            Func[] inequations = new Func[] { };

            Func<double, double>[] r_funcs = new Func<double, double>[] { t => 2.0 * t, t => 10 * t, t => t * t };

            int step;
            int calc_count;
            int pos = 2;

            Application app = new Application();
            Workbook wb = app.Workbooks.Add();
            Worksheet ws = wb.ActiveSheet;

            ws.Range["A1"].Value = "N";
            ws.Range["B1"].Value = "r0";
            ws.Range["C1"].Value = "g(r)";
            ws.Range["D1"].Value = "x0";
            ws.Range["E1"].Value = "eps";
            ws.Range["F1"].Value = "кол-во итераций";
            ws.Range["G1"].Value = "кол-во выч. функции";
            ws.Range["H1"].Value = "xk";
            ws.Range["I1"].Value = "f(xk)";

            for (int i = 0; i < r_funcs.Length; i++, pos++)
            {
                Vector start = new Vector(-1, -1);
                Vector r = new Vector(2.0);

                step = PenaltyMethods.PenaltyMethod(func, start, r, r_funcs[i], H, G, equations, inequations, out calc_count, out _);

                ws.Cells[pos, 1].Value = 1.ToString();
                ws.Cells[pos, 2].Value = 2.ToString();
                ws.Cells[pos, 3].Value = i.ToString();
                ws.Cells[pos, 4].Value = "(-1, -1)";
                ws.Cells[pos, 5].Value = PenaltyMethods.Epsilon.ToString("E5");
                ws.Cells[pos, 6].Value = step;
                ws.Cells[pos, 7].Value = calc_count;
                ws.Cells[pos, 8].Value = start.ToString("E5");
                ws.Cells[pos, 9].Value = func(start.values).ToString("E5");
            }
            ws.Columns.AutoFit();
            wb.SaveAs(Path.Combine(Environment.CurrentDirectory, "PenaltyMethod_func_2_r.xlsx"));
            wb.Close();
        }

        static void PenaltyMethod_research_func_1_x0()
        {
            Func<double, double> G = t => 0.5 * (t + Math.Abs(t));
            Func<double, double> H = t => Math.Abs(t);

            Func func = t => (t[0] - t[1]) * (t[0] - t[1]) + 10.0 * (t[0] + 5.0) * (t[0] + 5.0);
            Func<double, double> func_r = t => 2.0 * t;

            Func[] equations = new Func[] { };
            Func[] inequations = new Func[] { t => -t[0] - t[1] };

            Func<double, double> r_func = t => 2.0 * t;

            Vector[] points = new Vector[] { new Vector(-1, -1), new Vector(5, 10), new Vector(-3.8, 3.8) };

            int step;
            int calc_count;
            int pos = 2;

            Application app = new Application();
            Workbook wb = app.Workbooks.Add();
            Worksheet ws = wb.ActiveSheet;

            ws.Range["A1"].Value = "N";
            ws.Range["B1"].Value = "r0";
            ws.Range["C1"].Value = "g(r)";
            ws.Range["D1"].Value = "x0";
            ws.Range["E1"].Value = "eps";
            ws.Range["F1"].Value = "кол-во итераций";
            ws.Range["G1"].Value = "кол-во выч. функции";
            ws.Range["H1"].Value = "xk";
            ws.Range["I1"].Value = "f(xk)";

            for (int i = 0; i < points.Length; i++, pos++)
            {
                Vector start = new Vector(points[i]);
                Vector r = new Vector(2.0);
                ws.Cells[pos, 4].Value = start.ToString();

                step = PenaltyMethods.PenaltyMethod(func, start, r, r_func, H, G, equations, inequations, out calc_count, out _);

                ws.Cells[pos, 1].Value = 1.ToString();
                ws.Cells[pos, 2].Value = 2.ToString();
                ws.Cells[pos, 3].Value = 1.ToString();
                ws.Cells[pos, 5].Value = PenaltyMethods.Epsilon.ToString("E5");
                ws.Cells[pos, 6].Value = step;
                ws.Cells[pos, 7].Value = calc_count;
                ws.Cells[pos, 8].Value = start.ToString("E5");
                ws.Cells[pos, 9].Value = func(start.values).ToString("E5");
            }
            ws.Columns.AutoFit();
            wb.SaveAs(Path.Combine(Environment.CurrentDirectory, "PenaltyMethod_func_1_x0.xlsx"));
            wb.Close();
        }
        static void PenaltyMethod_research_func_2_x0()
        {
            Func<double, double> G = t => 0.5 * (t + Math.Abs(t));
            Func<double, double> H = t => Math.Abs(t);

            Func func = t => (t[0] - t[1]) * (t[0] - t[1]) + 10.0 * (t[0] + 5.0) * (t[0] + 5.0);
            Func<double, double> func_r = t => 2.0 * t;

            Func[] equations = new Func[] { t => 1.0 - t[0] - t[1] };
            Func[] inequations = new Func[] { };

            Func<double, double> r_func = t => 2.0 * t;

            Vector[] points = new Vector[] { new Vector(-1, -1), new Vector(15, 10), new Vector(-3.8, 3.8) };

            int step;
            int calc_count;
            int pos = 2;

            Application app = new Application();
            Workbook wb = app.Workbooks.Add();
            Worksheet ws = wb.ActiveSheet;

            ws.Range["A1"].Value = "N";
            ws.Range["B1"].Value = "r0";
            ws.Range["C1"].Value = "g(r)";
            ws.Range["D1"].Value = "x0";
            ws.Range["E1"].Value = "eps";
            ws.Range["F1"].Value = "кол-во итераций";
            ws.Range["G1"].Value = "кол-во выч. функции";
            ws.Range["H1"].Value = "xk";
            ws.Range["I1"].Value = "f(xk)";

            for (int i = 0; i < points.Length; i++, pos++)
            {
                Vector start = new Vector(points[i]);
                Vector r = new Vector(2.0);
                ws.Cells[pos, 4].Value = start.ToString();

                step = PenaltyMethods.PenaltyMethod(func, start, r, r_func, H, G, equations, inequations, out calc_count, out _);

                ws.Cells[pos, 1].Value = 1.ToString();
                ws.Cells[pos, 2].Value = 2.ToString();
                ws.Cells[pos, 3].Value = 1.ToString();
                ws.Cells[pos, 5].Value = PenaltyMethods.Epsilon.ToString("E5");
                ws.Cells[pos, 6].Value = step;
                ws.Cells[pos, 7].Value = calc_count;
                ws.Cells[pos, 8].Value = start.ToString("E5");
                ws.Cells[pos, 9].Value = func(start.values).ToString("E5");
            }
            ws.Columns.AutoFit();
            wb.SaveAs(Path.Combine(Environment.CurrentDirectory, "PenaltyMethod_func_2_x0.xlsx"));
            wb.Close();
        }

        static void PenaltyMethod_research_func_1_eps()
        {
            Func<double, double> G = t => 0.5 * (t + Math.Abs(t));
            Func<double, double> H = t => Math.Abs(t);

            Func func = t => (t[0] - t[1]) * (t[0] - t[1]) + 10.0 * (t[0] + 5.0) * (t[0] + 5.0);
            Func<double, double> func_r = t => 2.0 * t;

            Func[] equations = new Func[] { };
            Func[] inequations = new Func[] { t => -t[0] - t[1]};

            Func<double, double> r_func = t => 2.0 * t;

            double[] eps = new double[] { 1E-2, 1E-3, 1E-4, 1E-5 };

            int step;
            int calc_count;
            int pos = 2;

            Application app = new Application();
            Workbook wb = app.Workbooks.Add();
            Worksheet ws = wb.ActiveSheet;

            ws.Range["A1"].Value = "N";
            ws.Range["B1"].Value = "r0";
            ws.Range["C1"].Value = "g(r)";
            ws.Range["D1"].Value = "x0";
            ws.Range["E1"].Value = "eps";
            ws.Range["F1"].Value = "кол-во итераций";
            ws.Range["G1"].Value = "кол-во выч. функции";
            ws.Range["H1"].Value = "xk";
            ws.Range["I1"].Value = "f(xk)";

            for (int i = 0; i < eps.Length; i++, pos++)
            {
                Vector start = new Vector(-1.0, -1.0);
                Vector r = new Vector(2.0);
                ws.Cells[pos, 4].Value = start.ToString();

                PenaltyMethods.Epsilon = eps[i];
                step = PenaltyMethods.PenaltyMethod(func, start, r, r_func, H, G, equations, inequations, out calc_count, out _);

                ws.Cells[pos, 1].Value = 1.ToString();
                ws.Cells[pos, 2].Value = 2.ToString();
                ws.Cells[pos, 3].Value = 1.ToString();
                ws.Cells[pos, 5].Value = PenaltyMethods.Epsilon.ToString("E5");
                ws.Cells[pos, 6].Value = step;
                ws.Cells[pos, 7].Value = calc_count;
                ws.Cells[pos, 8].Value = start.ToString("E5");
                ws.Cells[pos, 9].Value = func(start.values).ToString("E5");
            }
            ws.Columns.AutoFit();
            wb.SaveAs(Path.Combine(Environment.CurrentDirectory, "PenaltyMethod_func_1_eps.xlsx"));
            wb.Close();
        }
        static void PenaltyMethod_research_func_2_eps()
        {
            Func<double, double> G = t => 0.5 * (t + Math.Abs(t));
            Func<double, double> H = t => Math.Abs(t);

            Func func = t => (t[0] - t[1]) * (t[0] - t[1]) + 10.0 * (t[0] + 5.0) * (t[0] + 5.0);
            Func<double, double> func_r = t => 2.0 * t;

            Func[] equations = new Func[] { t => 1.0 - t[0] - t[1] };
            Func[] inequations = new Func[] { };

            Func<double, double> r_func = t => 2.0 * t;

            double[] eps = new double[] { 1E-2, 1E-3, 1E-4, 1E-5 };

            int step;
            int calc_count;
            int pos = 2;

            Application app = new Application();
            Workbook wb = app.Workbooks.Add();
            Worksheet ws = wb.ActiveSheet;

            ws.Range["A1"].Value = "N";
            ws.Range["B1"].Value = "r0";
            ws.Range["C1"].Value = "g(r)";
            ws.Range["D1"].Value = "x0";
            ws.Range["E1"].Value = "eps";
            ws.Range["F1"].Value = "кол-во итераций";
            ws.Range["G1"].Value = "кол-во выч. функции";
            ws.Range["H1"].Value = "xk";
            ws.Range["I1"].Value = "f(xk)";

            for (int i = 0; i < eps.Length; i++, pos++)
            {
                Vector start = new Vector(-1.0, -1.0);
                Vector r = new Vector(2.0);
                ws.Cells[pos, 4].Value = start.ToString();

                PenaltyMethods.Epsilon = eps[i];
                step = PenaltyMethods.PenaltyMethod(func, start, r, r_func, H, G, equations, inequations, out calc_count, out _);

                ws.Cells[pos, 1].Value = 1.ToString();
                ws.Cells[pos, 2].Value = 2.ToString();
                ws.Cells[pos, 3].Value = 1.ToString();
                ws.Cells[pos, 5].Value = PenaltyMethods.Epsilon.ToString("E5");
                ws.Cells[pos, 6].Value = step;
                ws.Cells[pos, 7].Value = calc_count;
                ws.Cells[pos, 8].Value = start.ToString("E5");
                ws.Cells[pos, 9].Value = func(start.values).ToString("E5");
            }
            ws.Columns.AutoFit();
            wb.SaveAs(Path.Combine(Environment.CurrentDirectory, "PenaltyMethod_func_2_eps.xlsx"));
            wb.Close();
        }

        /*
        static void BarrierMethod_research_func_G()
        {
            Func<double, double> G_f1 = t => -1.0 / t;
            Func<double, double> G_f2 = t => -Math.Log(-t);
            Func<double, double>[] funcs = new Func<double, double>[] { G_f1, G_f2 };

            Func<double, double> H = t => Math.Abs(t);

            Func func = t => (t[0] - t[1]) * (t[0] - t[1]) + 10.0 * (t[0] + 5.0) * (t[0] + 5.0);
            Func<double, double> func_r = t => t / 2.0;

            Func[] equations = new Func[] { };
            Func[] inequations = new Func[] { t => -t[0] - t[1] };

            int step;
            int calc_count;
            int pos = 2;

            Application app = new Application();
            Workbook wb = app.Workbooks.Add();
            Worksheet ws = wb.ActiveSheet;

            ws.Range["A1"].Value = "N";
            ws.Range["B1"].Value = "r0";
            ws.Range["C1"].Value = "g(r)";
            ws.Range["D1"].Value = "x0";
            ws.Range["E1"].Value = "eps";
            ws.Range["F1"].Value = "кол-во итераций";
            ws.Range["G1"].Value = "кол-во выч. функции";
            ws.Range["H1"].Value = "xk";
            ws.Range["I1"].Value = "f(xk)";

            for (int i = 0; i < funcs.Length; i++, pos++)
            {
                Vector start = new Vector(5, 5);
                Vector r = new Vector(100.0);
                ws.Cells[pos, 4].Value = start.ToString();

                step = PenaltyMethods.BarrierMethod(func, start, r, func_r, H, funcs[i], equations, inequations, out calc_count, out _);

                ws.Cells[pos, 1].Value = (i + 1).ToString();
                ws.Cells[pos, 2].Value = 100.0.ToString();
                ws.Cells[pos, 3].Value = "r / 2.0";
                ws.Cells[pos, 5].Value = PenaltyMethods.Epsilon.ToString("E5");
                ws.Cells[pos, 6].Value = step;
                ws.Cells[pos, 7].Value = calc_count;
                ws.Cells[pos, 8].Value = start.ToString("E5");
                ws.Cells[pos, 9].Value = func(start.values).ToString("E5");
            }
            ws.Columns.AutoFit();
            wb.SaveAs(Path.Combine(Environment.CurrentDirectory, "BarrierMethod_funcG.xlsx"));
            wb.Close();
        }
        static void BarrierMethod_research_func_r0()
        {
            Func<double, double> G = t => -1.0 / t;
            Func<double, double> H = t => Math.Abs(t);

            Func func = t => (t[0] - t[1]) * (t[0] - t[1]) + 10.0 * (t[0] + 5.0) * (t[0] + 5.0);
            Func<double, double> func_r = t => t / 2.0;

            Func[] equations = new Func[] { };
            Func[] inequations = new Func[] { t => -t[0] - t[1] };

            double[] r_arr = new double[] { 0.5, 4.0, 8.0 };

            int step;
            int calc_count;
            int pos = 2;

            Application app = new Application();
            Workbook wb = app.Workbooks.Add();
            Worksheet ws = wb.ActiveSheet;

            ws.Range["A1"].Value = "N";
            ws.Range["B1"].Value = "r0";
            ws.Range["C1"].Value = "g(r)";
            ws.Range["D1"].Value = "x0";
            ws.Range["E1"].Value = "eps";
            ws.Range["F1"].Value = "кол-во итераций";
            ws.Range["G1"].Value = "кол-во выч. функции";
            ws.Range["H1"].Value = "xk";
            ws.Range["I1"].Value = "f(xk)";

            for (int i = 0; i < r_arr.Length; i++, pos++)
            {
                Vector start = new Vector(5, 5);
                Vector r = new Vector(r_arr[i]);
                ws.Cells[pos, 4].Value = start.ToString();

                step = PenaltyMethods.BarrierMethod(func, start, r, func_r, H, G, equations, inequations, out calc_count, out _);

                ws.Cells[pos, 1].Value = 1.ToString();
                ws.Cells[pos, 2].Value = r_arr[i].ToString();
                ws.Cells[pos, 3].Value = "r / 2.0";
                ws.Cells[pos, 5].Value = PenaltyMethods.Epsilon.ToString("E5");
                ws.Cells[pos, 6].Value = step;
                ws.Cells[pos, 7].Value = calc_count;
                ws.Cells[pos, 8].Value = start.ToString("E5");
                ws.Cells[pos, 9].Value = func(start.values).ToString("E5");
            }
            ws.Columns.AutoFit();
            wb.SaveAs(Path.Combine(Environment.CurrentDirectory, "BarrierMethod_func_r0_1.xlsx"));
            wb.Close();
        }
        static void BarrierMethod_research_func_r()
        {
            Func<double, double> G = t => -1.0 / t;
            Func<double, double> H = t => Math.Abs(t);

            Func func = t => (t[0] - t[1]) * (t[0] - t[1]) + 10.0 * (t[0] + 5.0) * (t[0] + 5.0);
            Func<double, double> func_r = t => 2.0 * t;

            Func[] equations = new Func[] { };
            Func[] inequations = new Func[] { t => -t[0] - t[1] };

            Func<double, double>[] r_funcs = new Func<double, double>[] { t => t / 2.0, t => t / 10.0, t => t / 100.0 };

            int step;
            int calc_count;
            int pos = 2;

            Application app = new Application();
            Workbook wb = app.Workbooks.Add();
            Worksheet ws = wb.ActiveSheet;

            ws.Range["A1"].Value = "N";
            ws.Range["B1"].Value = "r0";
            ws.Range["C1"].Value = "g(r)";
            ws.Range["D1"].Value = "x0";
            ws.Range["E1"].Value = "eps";
            ws.Range["F1"].Value = "кол-во итераций";
            ws.Range["G1"].Value = "кол-во выч. функции";
            ws.Range["H1"].Value = "xk";
            ws.Range["I1"].Value = "f(xk)";

            for (int i = 0; i < r_funcs.Length; i++, pos++)
            {
                Vector start = new Vector(5, 5);
                Vector r = new Vector(10.0);
                ws.Cells[pos, 4].Value = start.ToString();

                step = PenaltyMethods.BarrierMethod(func, start, r, r_funcs[i], H, G, equations, inequations, out calc_count, out _);

                ws.Cells[pos, 1].Value = 1.ToString();
                ws.Cells[pos, 2].Value = 100.ToString();
                ws.Cells[pos, 3].Value = i.ToString();
                ws.Cells[pos, 5].Value = PenaltyMethods.Epsilon.ToString("E5");
                ws.Cells[pos, 6].Value = step;
                ws.Cells[pos, 7].Value = calc_count;
                ws.Cells[pos, 8].Value = start.ToString("E5");
                ws.Cells[pos, 9].Value = func(start.values).ToString("E5");
            }
            ws.Columns.AutoFit();
            wb.SaveAs(Path.Combine(Environment.CurrentDirectory, "BarrierMethod_func_r.xlsx"));
            wb.Close();
        }
        static void BarrierMethod_research_func_x0()
        {
            Func<double, double> G = t => -1.0 / t;
            Func<double, double> H = t => Math.Abs(t);

            Func func = t => (t[0] - t[1]) * (t[0] - t[1]) + 10.0 * (t[0] + 5.0) * (t[0] + 5.0);

            Func[] equations = new Func[] { };
            Func[] inequations = new Func[] { t => -t[0] - t[1] };

            Func<double, double> r_func = t => t / 2.0;

            Vector[] points = new Vector[] { new Vector(5, -5), new Vector(15, 10), new Vector(-3.8, 3.9) };

            int step;
            int calc_count;
            int pos = 2;

            Application app = new Application();
            Workbook wb = app.Workbooks.Add();
            Worksheet ws = wb.ActiveSheet;

            ws.Range["A1"].Value = "N";
            ws.Range["B1"].Value = "r0";
            ws.Range["C1"].Value = "g(r)";
            ws.Range["D1"].Value = "x0";
            ws.Range["E1"].Value = "eps";
            ws.Range["F1"].Value = "кол-во итераций";
            ws.Range["G1"].Value = "кол-во выч. функции";
            ws.Range["H1"].Value = "xk";
            ws.Range["I1"].Value = "f(xk)";

            for (int i = 0; i < points.Length; i++, pos++)
            {
                Vector start = new Vector(points[i]);
                Vector r = new Vector(100.0);
                ws.Cells[pos, 4].Value = start.ToString();

                step = PenaltyMethods.BarrierMethod(func, start, r, r_func, H, G, equations, inequations, out calc_count, out _);

                ws.Cells[pos, 1].Value = 1.ToString();
                ws.Cells[pos, 2].Value = 100.ToString();
                ws.Cells[pos, 3].Value = 1.ToString();
                ws.Cells[pos, 5].Value = PenaltyMethods.Epsilon.ToString("E5");
                ws.Cells[pos, 6].Value = step;
                ws.Cells[pos, 7].Value = calc_count;
                ws.Cells[pos, 8].Value = start.ToString("E5");
                ws.Cells[pos, 9].Value = func(start.values).ToString("E5");
            }
            ws.Columns.AutoFit();
            wb.SaveAs(Path.Combine(Environment.CurrentDirectory, "BarrierMethod_func_x0.xlsx"));
            wb.Close();
        }
        static void BarrierMethod_research_func_eps()
        {
            Func<double, double> G = t => -1.0 / t;
            Func<double, double> H = t => Math.Abs(t);

            Func func = t => (t[0] - t[1]) * (t[0] - t[1]) + 10.0 * (t[0] + 5.0) * (t[0] + 5.0);
            Func<double, double> func_r = t => t / 2.0;

            List<Func> equations = new List<Func>() { };
            List<Func> inequations = new List<Func> { t => -t[0] - t[1]};

            double[] eps = new double[] { 1E-2, 1E-3, 1E-4, 1E-5 };

            int step;
            int calc_count;
            int pos = 2;

            Application app = new Application();
            Workbook wb = app.Workbooks.Add();
            Worksheet ws = wb.ActiveSheet;

            ws.Range["A1"].Value = "N";
            ws.Range["B1"].Value = "r0";
            ws.Range["C1"].Value = "g(r)";
            ws.Range["D1"].Value = "x0";
            ws.Range["E1"].Value = "eps";
            ws.Range["F1"].Value = "кол-во итераций";
            ws.Range["G1"].Value = "кол-во выч. функции";
            ws.Range["H1"].Value = "xk";
            ws.Range["I1"].Value = "f(xk)";

            for (int i = 0; i < eps.Length; i++, pos++)
            {
                Vector start = new Vector(5, 5);
                Vector r = new Vector(100.0);
                ws.Cells[pos, 4].Value = start.ToString();

                PenaltyMethods.Epsilon = eps[i];
                OneDimensionalSearches.Epsilon = eps[i];
                DescentMethods.Epsilon = eps[i];

                step = PenaltyMethods.BarrierMethod(func, start, r, func_r, H, G, equations, inequations, out calc_count, out _);

                ws.Cells[pos, 1].Value = 1.ToString();
                ws.Cells[pos, 2].Value = 100.ToString();
                ws.Cells[pos, 3].Value = 1.ToString();
                ws.Cells[pos, 5].Value = PenaltyMethods.Epsilon.ToString("E5");
                ws.Cells[pos, 6].Value = step;
                ws.Cells[pos, 7].Value = calc_count;
                ws.Cells[pos, 8].Value = start.ToString("E5");
                ws.Cells[pos, 9].Value = func(start.values).ToString("E5");
            }
            ws.Columns.AutoFit();
            wb.SaveAs(Path.Combine(Environment.CurrentDirectory, "BarrierMethod_func_eps.xlsx"));
            wb.Close();
        }
        */
    }
}