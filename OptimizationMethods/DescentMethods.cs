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
   public static class VectorExtensions
   {
      public static double InnerMult(this Vector a, Vector b)
      {
         double res = 0;
         for (int i = 0; i < a.size; i++)
            res += a[i] * b[i];
         return res;
      }
   }
   public static class DescentMethods
   {
      public static double Epsilon = 1E-7;
      public static double Delta = 1E-7;
      public static int MaxInterationCount = 1000;
      private static void getGradient(Func function, Vector point, Vector gradient)
      {
         double h = 1E-5;
         double temp;
         double f_right, f_left;
         double[] x = point.values;

         for (int i = 0; i < point.size; i++)
         {
            temp = x[i];
            x[i] = -(h - x[i]);
            f_left = function(x);

            x[i] = h + temp;
            f_right = function(x);

            gradient[i] = (f_right - f_left) / (2.0 * h);
         }
      }
      private static void getAntigradient(Func function, Vector point, Vector gradient)
      {
         double h = 1E-10;
         double temp;
         double f_right, f_left;
         double[] x = point.values;

         for (int i = 0; i < point.size; i++)
         {
            temp = x[i];
            x[i] = temp - h;
            f_left = function(x);

            x[i] = temp + h;
            f_right = function(x);

            gradient[i] = (f_left - f_right) / (2.0 * h);
         }
      }
      private static double findLambda(ExtremumType type, Func<double, double> function, double start, out int calc_count)
      {
         double v_left, v_right;
         double left, right;
         int calc;
         OneDimensionalSearches.FindInterval(type, start, function, out v_left, out v_right, out calc_count);
         OneDimensionalSearches.DichotomyMethod(type, v_left, v_right, function, out left, out right, out calc);
         calc_count += calc;
         return (right + left) / 2.0;
      }
      public static int GaussAlgorithm(ExtremumType type, Func function, Vector start, out int calc_count, out double norm_s)
      {
         int size = start.size;
         int step = 0;
         int curVariable = 0;
         int calc = 0;
         double lambda;
         double curNorm = 0;
         double[] arg = new double[size];

         Func<double, double> g = v =>
         {
            for (int i = 0; i < size; i++)
               arg[i] = start[i];
            arg[curVariable] = start[curVariable] + v;
            return function(arg);
         };

         norm_s = double.MaxValue;
         calc_count = 0;
         do
         {
            lambda = findLambda(type, g, 0, out calc);
            calc_count += calc;
            start[curVariable] += lambda;
            curNorm += lambda * lambda;

            if (++curVariable == size)
            {
               if (curNorm < Epsilon * Epsilon)
               {
                  norm_s = curNorm;
                  return step;
               }
               curNorm = 0;
               step++;
            }

            curVariable %= size;
         } while (step <= MaxInterationCount);
         return step;
      }
      public static int CGM_FletcherReeves(ExtremumType type, Func function, Vector start, out int calc_count, out double norm_s)
      {
         int i;
         int size = start.size;
         int step = 0;
         int iterations = 0;
         calc_count = 0;

         double[] arg = new double[size];
         double lambda;

         double prev_innerMult, cur_innerMult;
         double wk;

         Vector s = new Vector(size);
         Vector r = new Vector(size);

         Func<double, double> g = v =>
         {
            for (int j = 0; j < size; j++)
               arg[j] = start[j] + v * s[j];
            return function(arg);
         };

         switch (type)
         {
            case ExtremumType.Minimum:
               while (true)
               {
                  getAntigradient(function, start, r);
                  calc_count += 2 * size;

                  for (i = 0; i < size; i++)
                     s[i] = r[i];
                  cur_innerMult = r.InnerMult(r);

                  while (true)
                  {
                     step++;
                     // finding minimum of g(v)
                     lambda = findLambda(type, g, 0.0, out _);

                     // x_k+1 = x_k + v * s_k
                     for (i = 0; i < size; i++)
                        start[i] += lambda * s[i];

                     getAntigradient(function, start, r);
                     calc_count += 2 * size;

                     prev_innerMult = cur_innerMult;
                     cur_innerMult = r.InnerMult(r);

                     // w_k+1 = (r_k+1, r_k+1) / (r_k, r_k)
                     wk = cur_innerMult / prev_innerMult;
                     // s_k+1 = r_k+1 + w_k+1 * s_k
                     for (i = 0; i < size; i++)
                        s[i] = r[i] + wk * s[i];

                     if (step > size + 1)
                     {
                        iterations += step;
                        step = 0;
                        break;
                     }
                     else if ((norm_s = s.sqrMagnitude) < Epsilon * Epsilon)
                     {
                        iterations += step;
                        return iterations;
                     }
                  }
               }
               break;
            case ExtremumType.Maximum:
               while (true)
               {
                  getGradient(function, start, r);
                  calc_count += 2 * size;

                  for (i = 0; i < size; i++)
                     s[i] = r[i];
                  cur_innerMult = r.InnerMult(r);

                  while (true)
                  {
                     step++;

                     // finding minimum of g(v)
                     lambda = findLambda(type, g, 0.0, out _);

                     // x_k+1 = x_k + v * s_k
                     for (i = 0; i < size; i++)
                        start[i] += lambda * s[i];

                     getGradient(function, start, r);

                     prev_innerMult = cur_innerMult;
                     cur_innerMult = r.InnerMult(r);

                     // w_k+1 = (r_k+1, r_k+1) / (r_k, r_k)
                     wk = cur_innerMult / prev_innerMult;
                     // s_k+1 = r_k+1 + w_k+1 * s_k
                     for (i = 0; i < size; i++)
                        s[i] = r[i] + wk * s[i];

                     if ((norm_s = s.sqrMagnitude) < Epsilon * Epsilon)
                     {
                        iterations += step;
                        return iterations;
                     }
                     else if (step > size + 1)
                     {
                        iterations += step;
                        step = 0;
                        break;
                     }
                  }
               }
               break;
         }
         norm_s = -1;
         return -1;
      }
      public static int PiersonAlgorithm(ExtremumType type, Func function, Vector start, out int calc_count, out double norm_s)
      {
         int i, j;
         int size = start.size;
         int step = 0;
         calc_count = 0;

         Vector gradient = new Vector(size);
         Vector dgradient = new Vector(size);
         double[] temp = new double[size];
         double[] dx = new double[size];

         double lambda;
         double distance;
         Func<double, double> g = v =>
         {
            for (int s = 0; s < size; s++)
               temp[s] = start[s] + v * dx[s];
            return function(temp);
         };

         double[,] approxMatrix = new double[size, size];
         for (i = 0; i < size; i++)
            approxMatrix[i, i] = 1.0;

         switch (type)
         {
            case ExtremumType.Minimum:
               do
               {
                  getGradient(function, start, gradient);
                  calc_count += 2 * size;

                  if (step % 2 == 0)
                  {
                     for (i = 0; i < size; i++)
                     {
                        for (j = 0; j < size; j++)
                           approxMatrix[i, j] = 0.0;
                        approxMatrix[i, i] = 1.0;
                     }
                  }

                  for (i = 0; i < size; i++)
                  {
                     dx[i] = 0;
                     for (j = 0; j < size; j++)
                        dx[i] += approxMatrix[i, j] * gradient[j];
                     dx[i] *= -1.0;
                  }

                  lambda = findLambda(type, g, 0.0, out _);
                  distance = 0;
                  for (i = 0; i < size; i++)
                  {
                     dx[i] *= lambda;
                     start[i] += dx[i];
                     distance += dx[i] * dx[i];
                  }

                  // dgradient = gradient(x_k+1)
                  getGradient(function, start, dgradient);
                  calc_count += 2 * size;

                  step++;
                  if ((norm_s = distance) < Epsilon * Epsilon)
                     break;

                  // dgradient = gradient(x_k+1) - gradient(x_k)
                  for (i = 0; i < size; i++)
                     dgradient[i] -= gradient[i];

                  // min_v = (dg)^T * approxMat * dg
                  // gradient = approxMat * dg
                  // dx = dx - approxMat * dg
                  lambda = 0;
                  for (i = 0; i < size; i++)
                  {
                     gradient[i] = 0;
                     for (j = 0; j < size; j++)
                        gradient[i] += approxMatrix[i, j] * dgradient[j];
                     dx[i] -= gradient[i];
                     lambda += dgradient[i] * gradient[i];
                  }
                  for (i = 0; i < size; i++)
                     for (j = 0; j < size; j++)
                        approxMatrix[i, j] += dx[i] * gradient[j] / lambda;
               } while (true);
               break;
            case ExtremumType.Maximum:
               do
               {
                  getAntigradient(function, start, gradient);
                  calc_count += 2 * size;

                  if (step % 2 == 0)
                  {
                     for (i = 0; i < size; i++)
                     {
                        for (j = 0; j < size; j++)
                           approxMatrix[i, j] = 0.0;
                        approxMatrix[i, i] = 1.0;
                     }
                  }

                  for (i = 0; i < size; i++)
                  {
                     dx[i] = 0;
                     for (j = 0; j < size; j++)
                        dx[i] += approxMatrix[i, j] * gradient[j];
                     dx[i] *= -1.0;
                  }

                  lambda = findLambda(type, g, 0.0, out _);
                  distance = 0;
                  for (i = 0; i < size; i++)
                  {
                     dx[i] *= lambda;
                     start[i] += dx[i];
                     distance += dx[i] * dx[i];
                  }

                  // dgradient = gradient(x_k+1)
                  getAntigradient(function, start, dgradient);
                  calc_count += 2 * size;

                  step++;
                  if ((norm_s = distance) < Epsilon * Epsilon)
                     break;

                  // dgradient = gradient(x_k+1) - gradient(x_k)
                  for (i = 0; i < size; i++)
                     dgradient[i] -= gradient[i];

                  // min_v = (dg)^T * approxMat * dg
                  // gradient = approxMat * dg
                  // dx = dx - approxMat * dg
                  lambda = 0;
                  for (i = 0; i < size; i++)
                  {
                     gradient[i] = 0;
                     for (j = 0; j < size; j++)
                        gradient[i] += approxMatrix[i, j] * dgradient[j];
                     dx[i] -= gradient[i];
                     lambda += dgradient[i] * gradient[i];
                  }
                  for (i = 0; i < size; i++)
                     for (j = 0; j < size; j++)
                        approxMatrix[i, j] += dx[i] * gradient[j] / lambda;
               } while (true);
               break;
            default:
               norm_s = -1;
               return -1;
               break;
         }
         return step;
      }
      public static int CGM_FletcherReeves_table(ExtremumType type, int num, Func function, Vector start)
      {
         Application app = new Application();
         Workbook wb = app.Workbooks.Add();
         Worksheet ws = wb.ActiveSheet;

         int i;
         int size = start.size;
         int step = 0;
         int pos = 2;
         int iterations = 0;

         double[] arg = new double[size];
         double lambda;
         double angle;

         double prev_innerMult, cur_innerMult;
         double wk;

         Vector s = new Vector(size);
         Vector r = new Vector(size);

         Func<double, double> g = v =>
         {
            for (int j = 0; j < size; j++)
               arg[j] = start[j] + v * s[j];
            return function(arg);
         };

         ws.Range["A1"].Value = "i";
         ws.Range["B1"].Value = "(x_i, y_i)";
         ws.Range["C1"].Value = "f(x_i, y_i)";
         ws.Range["D1"].Value = "(s_1, s_2)";
         ws.Range["E1"].Value = "лямбда";
         ws.Range["F1"].Value = "норма s_i+1";
         ws.Range["G1"].Value = "угол";
         ws.Range["H1"].Value = "градиент/антиградиент";

         switch (type)
         {
            case ExtremumType.Minimum:
               while (true)
               {
                  getAntigradient(function, start, r);

                  for (i = 0; i < size; i++)
                     s[i] = r[i];
                  cur_innerMult = r.InnerMult(r);

                  while (true)
                  {
                     step++;
                     ws.Cells[pos, 1].Value = iterations + step - 1;
                     ws.Cells[pos, 2].Value = start[0].ToString("E5") + " " + start[1].ToString("E5");
                     ws.Cells[pos, 3].Value = function(start.values).ToString("E5");
                     ws.Cells[pos, 4].Value = s[0].ToString("E5") + " " + s[1].ToString("E5");
                     angle = Math.Acos((start[0] * s[0] + start[1] * s[1]) / (Math.Sqrt(s[0] * s[0] + s[1] * s[1]) * Math.Sqrt(start[0] * start[0] + start[1] * start[1])));
                     ws.Cells[pos, 7].Value = angle.ToString("E5");
                     ws.Cells[pos, 8].Value = r[0].ToString("E5") + " " + r[1].ToString("E5");

                     // finding minimum of g(v)
                     lambda = findLambda(type, g, 0.0, out _);
                     ws.Cells[pos, 5].Value = lambda.ToString("E5");

                     // x_k+1 = x_k + v * s_k
                     for (i = 0; i < size; i++)
                        start[i] += lambda * s[i];

                     getAntigradient(function, start, r);

                     prev_innerMult = cur_innerMult;
                     cur_innerMult = r.InnerMult(r);

                     // w_k+1 = (r_k+1, r_k+1) / (r_k, r_k)
                     wk = cur_innerMult / prev_innerMult;
                     // s_k+1 = r_k+1 + w_k+1 * s_k
                     for (i = 0; i < size; i++)
                        s[i] = r[i] + wk * s[i];

                     ws.Cells[pos, 6].Value = s.magnitude.ToString("E5");
                     pos++;

                     if (step > size + 1)
                     {
                        iterations += step;
                        step = 0;
                        break;
                     }
                     else if (s.sqrMagnitude < Epsilon * Epsilon)
                     {
                        iterations += step;
                        ws.Columns.AutoFit();
                        wb.SaveAs(Path.Combine(Environment.CurrentDirectory, "CGM" + num + "_steps_research.xlsx"));
                        wb.Close();
                        return iterations;
                     }
                  }
               }
               break;
            case ExtremumType.Maximum:
               while (true)
               {
                  getGradient(function, start, r);

                  for (i = 0; i < size; i++)
                     s[i] = r[i];
                  cur_innerMult = r.InnerMult(r);

                  while (true)
                  {
                     step++;
                     ws.Cells[pos, 1].Value = iterations + step - 1;
                     ws.Cells[pos, 2].Value = start[0].ToString("E5") + " " + start[1].ToString("E5");
                     ws.Cells[pos, 3].Value = function(start.values).ToString("E5");
                     ws.Cells[pos, 4].Value = s[0].ToString("E5") + " " + s[1].ToString("E5");
                     angle = Math.Acos((start[0] * s[0] + start[1] * s[1]) / (Math.Sqrt(s[0] * s[0] + s[1] * s[1]) * Math.Sqrt(start[0] * start[0] + start[1] * start[1])));
                     ws.Cells[pos, 7].Value = angle.ToString("E5");
                     ws.Cells[pos, 8].Value = r[0].ToString("E5") + " " + r[1].ToString("E5");

                     // finding minimum of g(v)
                     lambda = findLambda(type, g, 0.0, out _);
                     ws.Cells[pos, 5].Value = lambda.ToString("E5");

                     // x_k+1 = x_k + v * s_k
                     for (i = 0; i < size; i++)
                        start[i] += lambda * s[i];

                     getGradient(function, start, r);

                     prev_innerMult = cur_innerMult;
                     cur_innerMult = r.InnerMult(r);

                     // w_k+1 = (r_k+1, r_k+1) / (r_k, r_k)
                     wk = cur_innerMult / prev_innerMult;
                     // s_k+1 = r_k+1 + w_k+1 * s_k
                     for (i = 0; i < size; i++)
                        s[i] = r[i] + wk * s[i];

                     ws.Cells[pos, 6].Value = s.magnitude.ToString("E5");
                     pos++;

                     if (s.sqrMagnitude < Epsilon * Epsilon)
                     {
                        iterations += step;
                        ws.Columns.AutoFit();
                        wb.SaveAs(Path.Combine(Environment.CurrentDirectory, "CGM" + num + "_steps_research.xlsx"));
                        wb.Close();
                        return iterations;
                     }
                     else if (step > size + 1)
                     {
                        iterations += step;
                        step = 0;
                        break;
                     }
                  }
               }
               break;
         }
         return -1;
      }
      public static int PiersonAlgorithm_table(ExtremumType type, int num, Func function, Vector start)
      {
         Application app = new Application();
         Workbook wb = app.Workbooks.Add();
         Worksheet ws = wb.ActiveSheet;

         int i, j;
         int size = start.size;
         int step = 0;
         int pos = 2;

         Vector gradient = new Vector(size);
         Vector dgradient = new Vector(size);
         double[] temp = new double[size];
         double[] dx = new double[size];

         double lambda;
         double distance = 0;
         double angle;
         Func<double, double> g = v =>
         {
            for (int s = 0; s < size; s++)
               temp[s] = start[s] + v * dx[s];
            return function(temp);
         };

         double[,] approxMatrix = new double[size, size];
         for (i = 0; i < size; i++)
            approxMatrix[i, i] = 1.0;

         ws.Range["A1"].Value = "i";
         ws.Range["B1"].Value = "(x_i, y_i)";
         ws.Range["C1"].Value = "f(x_i, y_i)";
         ws.Range["D1"].Value = "(s_1, s_2)";
         ws.Range["E1"].Value = "лямбда";
         ws.Range["F1"].Value = "норма (M_i+1 - M_i)";
         ws.Range["G1"].Value = "угол";
         ws.Range["H1"].Value = "аппроксимация матрицы\nвторых производных";

         switch (type)
         {
            case ExtremumType.Minimum:
               do
               {
                  ws.Cells[pos, 1].Value = step;
                  ws.Cells[pos, 2].Value = start[0].ToString("E5") + " " + start[1].ToString("E5");
                  ws.Cells[pos, 3].Value = function(start.values).ToString("E5");

                  getGradient(function, start, gradient);

                  if (step % 2 == 0)
                  {
                     for (i = 0; i < size; i++)
                     {
                        for (j = 0; j < size; j++)
                           approxMatrix[i, j] = 0.0;
                        approxMatrix[i, i] = 1.0;
                     }
                  }

                  string matrix_str = "";
                  for (i = 0; i < size; i++)
                     for (j = 0; j < size; j++)
                        matrix_str += approxMatrix[i, j].ToString("E5") + " ";
                  ws.Cells[pos, 8].Value = matrix_str;

                  for (i = 0; i < size; i++)
                  {
                     dx[i] = 0;
                     for (j = 0; j < size; j++)
                        dx[i] += approxMatrix[i, j] * gradient[j];
                     dx[i] *= -1.0;
                  }
                  ws.Cells[pos, 4].Value = dx[0].ToString("E5") + " " + dx[1].ToString("E5");
                  angle = Math.Acos((start[0] * dx[0] + start[1] * dx[1]) / (Math.Sqrt(dx[0] * dx[0] + dx[1] * dx[1]) * Math.Sqrt(start[0] * start[0] + start[1] * start[1])));
                  ws.Cells[pos, 7].Value = angle.ToString("E5");

                  lambda = findLambda(type, g, 0.0, out _);
                  ws.Cells[pos, 5].Value = lambda.ToString("E5");

                  distance = 0;
                  for (i = 0; i < size; i++)
                  {
                     dx[i] *= lambda;
                     start[i] += dx[i];
                     distance += dx[i] * dx[i];
                  }
                  ws.Cells[pos, 6].Value = Math.Sqrt(distance).ToString("E5");
                  pos++;

                  // dgradient = gradient(x_k+1)
                  getGradient(function, start, dgradient);

                  step++;
                  if (distance < Epsilon * Epsilon)
                     break;

                  // dgradient = gradient(x_k+1) - gradient(x_k)
                  for (i = 0; i < size; i++)
                     dgradient[i] -= gradient[i];

                  // min_v = (dg)^T * approxMat * dg
                  // gradient = approxMat * dg
                  // temp = dx - approxMat * dg
                  lambda = 0;
                  for (i = 0; i < size; i++)
                  {
                     gradient[i] = 0;
                     for (j = 0; j < size; j++)
                        gradient[i] += approxMatrix[i, j] * dgradient[j];
                     temp[i] = dx[i] - gradient[i];
                     lambda += dgradient[i] * gradient[i];
                  }
                  for (i = 0; i < size; i++)
                     for (j = 0; j < size; j++)
                        approxMatrix[i, j] += temp[i] * gradient[j] / lambda;
               } while (true);
               break;
            case ExtremumType.Maximum:
               do
               {
                  ws.Cells[pos, 1].Value = step;
                  ws.Cells[pos, 2].Value = start[0].ToString("E5") + " " + start[1].ToString("E5");
                  ws.Cells[pos, 3].Value = function(start.values).ToString("E5");

                  getAntigradient(function, start, gradient);

                  if (step % 2 == 0)
                  {
                     for (i = 0; i < size; i++)
                     {
                        for (j = 0; j < size; j++)
                           approxMatrix[i, j] = 0.0;
                        approxMatrix[i, i] = 1.0;
                     }
                  }

                  string matrix_str = "";
                  for (i = 0; i < size; i++)
                     for (j = 0; j < size; j++)
                        matrix_str += approxMatrix[i, j].ToString("E5") + " ";
                  ws.Cells[pos, 8].Value = matrix_str;

                  for (i = 0; i < size; i++)
                  {
                     dx[i] = 0;
                     for (j = 0; j < size; j++)
                        dx[i] += approxMatrix[i, j] * gradient[j];
                     dx[i] *= -1.0;
                  }
                  ws.Cells[pos, 4].Value = dx[0].ToString("E5") + " " + dx[1].ToString("E5");
                  angle = Math.Acos((start[0] * dx[0] + start[1] * dx[1]) / (Math.Sqrt(dx[0] * dx[0] + dx[1] * dx[1]) * Math.Sqrt(start[0] * start[0] + start[1] * start[1])));
                  ws.Cells[pos, 7].Value = angle.ToString("E5");

                  lambda = findLambda(type, g, 0.0, out _);
                  ws.Cells[pos, 5].Value = lambda.ToString("E5");
                  distance = 0;
                  for (i = 0; i < size; i++)
                  {
                     dx[i] *= lambda;
                     start[i] += dx[i];
                     distance += dx[i] * dx[i];
                  }
                  ws.Cells[pos, 6].Value = Math.Sqrt(distance).ToString("E5");
                  pos++;

                  // dgradient = gradient(x_k+1)
                  getAntigradient(function, start, dgradient);

                  step++;
                  if (distance < Epsilon * Epsilon)
                     break;

                  // dgradient = gradient(x_k+1) - gradient(x_k)
                  for (i = 0; i < size; i++)
                     dgradient[i] -= gradient[i];

                  // min_v = (dg)^T * approxMat * dg
                  // gradient = approxMat * dg
                  // dx = dx - approxMat * dg
                  lambda = 0;
                  for (i = 0; i < size; i++)
                  {
                     gradient[i] = 0;
                     for (j = 0; j < size; j++)
                        gradient[i] += approxMatrix[i, j] * dgradient[j];
                     dx[i] -= gradient[i];
                     lambda += dgradient[i] * gradient[i];
                  }
                  for (i = 0; i < size; i++)
                     for (j = 0; j < size; j++)
                        approxMatrix[i, j] += dx[i] * gradient[j] / lambda;
               } while (true);
               break;
            default:
               return -1;
               break;
         }

         ws.Columns.AutoFit();
         wb.SaveAs(Path.Combine(Environment.CurrentDirectory, "Pierson" + num + "_steps_research.xlsx"));
         wb.Close();
         return step;
      }
   }
}