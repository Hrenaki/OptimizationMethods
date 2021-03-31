using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NumMath;
using EParser;
using Microsoft.Office.Interop.Excel;

using App = Microsoft.Office.Interop.Excel.Application;
using System.IO;

namespace OptimizationMethods
{
    class Program
    {
        static void Main(string[] args)
        {
            //OneDimensionalSearches.Epsilon = 2E-1;
            //OneDimensionalSearches.FindIntervalResearch(3.01, x => (x - 3.0) * (x - 3.0), out _, out _);

            //for (int i = 1; i <= 7; i++)
            //{
            //    OneDimensionalSearches.Epsilon = Math.Pow(10, -i);
            //    Console.WriteLine(OneDimensionalSearches.FibonacciMethodResearch(-2.0, 20.0, x => (x - 3.0) * (x - 3.0), out _, out _));
            //}
            //double left, right;
            //OneDimensionalSearches.FibonacciMethodResearch(-2.0, 20.0, x => (x - 3.0) * (x - 3.0), out left, out right);
            //Console.WriteLine((right + left) / 2.0);           
            //Console.WriteLine(Math.Abs((right + left) / 2.0 - 3.0));           
            //Console.ReadLine();    
            //Func f = t => 100.0 * (t[1] - t[0]) * (t[1] - t[0]) + (1.0 - t[0]) * (1.0 - t[0]);
            //Func f = t => t[0] * t[0] + t[1] * t[1];

            Func f = t => 100.0 * (t[1] - t[0]) * (t[1] - t[0]) + (1.0 - t[0]) * (1.0 - t[0]);
            //Func f = t => 100.0 * (t[1] - t[0] * t[0]) * (t[1] - t[0] * t[0]) + (1.0 - t[0]) * (1.0 - t[0]);
            //Func f = t => Math.Exp(-Math.Pow(t[0] - 3.0, 2) - Math.Pow((t[1] - 1.0) / 3.0, 2)) + 
            //    2.0 * Math.Exp(-Math.Pow((t[0] - 2.0) / 2.0, 2) - Math.Pow(t[1] - 2.0, 2));
            Vector start = new Vector(30, 0);
            int step = DescentMethods.CGM_FletcherReeves(ExtremumType.Minimum, f, start);
            //int step = DescentMethods.PiersonAlgorithm(ExtremumType.Maximum, f, start);
            CGM_research_x0(f);
            Pierson_research_x0(f);

            //DescentMethods.PiersonAlgorithm(f, start);
        }

        static void CGM_research_x0(Func f)
        {
            Application app = new Application();
            Workbook wb = app.Workbooks.Add();
            Worksheet ws = wb.ActiveSheet;

            Vector[] points = new Vector[] { new Vector(1, -1), new Vector(10, 100), new Vector(30, 0) };
            Vector start;

            ws.Range["A1"].Value = "x0";
            ws.Range["B1"].Value = "xk";
            ws.Range["C1"].Value = "f(xk)";
            ws.Range["D1"].Value = "кол-во итераций";
            ws.Range["E1"].Value = "кол-во выч. функции";

            int calc_count = 0;
            int step;

            for (int i = 0; i < points.Length; i++)
            {
                start = points[i];
                ws.Cells[i + 2, 1].Value = start[0].ToString("E5") + " " + start[1].ToString("E5");
                step = DescentMethods.CGM_FletcherReevesResearch(ExtremumType.Maximum, f, start, out calc_count);
                ws.Cells[i + 2, 2].Value = start[0].ToString("E5") + " " + start[1].ToString("E5");
                ws.Cells[i + 2, 3].Value = f(start.values).ToString("E5");
                ws.Cells[i + 2, 4].Value = step;
                ws.Cells[i + 2, 5].Value = calc_count;
            }

            wb.SaveAs(Path.Combine(Environment.CurrentDirectory, "CGM2_x0_research.xlsx"));
            wb.Close();
        }
        static void Pierson_research_x0(Func f)
        {
            Application app = new Application();
            Workbook wb = app.Workbooks.Add();
            Worksheet ws = wb.ActiveSheet;

            Vector[] points = new Vector[] { new Vector(1, -1), new Vector(10, 100), new Vector(30, 0) };
            Vector start;

            ws.Range["A1"].Value = "x0";
            ws.Range["B1"].Value = "xk";
            ws.Range["C1"].Value = "f(xk)";
            ws.Range["D1"].Value = "кол-во итераций";
            ws.Range["E1"].Value = "кол-во выч. функции";

            int calc_count = 0;
            int step;

            for (int i = 0; i < points.Length; i++)
            {
                start = points[i];
                ws.Cells[i + 2, 1].Value = start[0].ToString("E5") + " " + start[1].ToString("E5");
                step = DescentMethods.PiersonAlgorithm(ExtremumType.Maximum, f, start, out calc_count);
                ws.Cells[i + 2, 2].Value = start[0].ToString("E5") + " " + start[1].ToString("E5");
                ws.Cells[i + 2, 3].Value = f(start.values).ToString("E5");
                ws.Cells[i + 2, 4].Value = step;
                ws.Cells[i + 2, 5].Value = calc_count;
            }

            wb.SaveAs(Path.Combine(Environment.CurrentDirectory, "Pierson2_x0_research.xlsx"));
            wb.Close();
        }
    }
}