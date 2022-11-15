using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EParser;
using NumMath;

namespace OptimizationMethods
{
   public static class PenaltyMethods
   {
      public static double Epsilon = 1E-7;
      public static int MaxIterationCount = 10;

      public static int PenaltyMethod(Func function, Vector start, Vector r, Func<double, double> function_r, Func<double, double> H, Func<double, double> G, List<Func> equations, List<Func> inequations, out int calc_count, out double norm)
      {
         int step = 0;
         int calc;
         int i;
         int size = start.size;
         int restriction_size = r.size;
         int equations_size = equations.Count;
         int inequations_size = inequations.Count;
         bool flag;
         double temp;
         Vector prevPoint = new Vector(size);
         double[] arg = new double[size + restriction_size];

         Func Q = t =>
         {
            double res = function(t);
            for (i = 0; i < equations_size; i++)
               res += r[i] * H(equations[i](t));
            for (i = 0; i < inequations_size; i++)
               //res += r[equations_size + i] * G(inequations[i](t));
               res += inequations[i](t);
            return res;
         };

         calc_count = 0;

         for (i = 0; i < size; i++)
            prevPoint[i] = start[i];

         do
         {
            step++;

            DescentMethods.GaussAlgorithm(ExtremumType.Minimum, Q, start, out calc, out _);
            calc_count += calc;
            if ((norm = prevPoint.Distance(start)) < Epsilon * Epsilon)
               break;

            flag = true;
            for (i = 0; i < equations_size; i++)
               if (Math.Abs(equations[i](start.values)) >= Epsilon)
               {
                  flag = false;
                  break;
               }
            for (i = 0; i < inequations_size; i++)
               if ((temp = inequations[i](start.values)) > 0.0 && Math.Abs(temp) >= Epsilon)
               {
                  flag = false;
                  break;
               }
            if (flag)
               break;

            for (i = 0; i < size; i++)
               prevPoint[i] = start[i];

            for (i = 0; i < restriction_size; i++)
               r[i] = function_r(r[i]);
         } while (step <= MaxIterationCount);

         return step;
      }
      public static int BarrierMethod(Func function, Vector start, Vector r, Func<double, double> function_r, Func<double, double> H, Func<double, double> G, List<Func> equations, List<Func> inequations, out int calc_count, out double norm)
      {
         int step = 0;
         int calc;
         int i;
         int size = start.size;
         int restriction_size = r.size;
         int equations_size = equations.Count;
         int inequations_size = inequations.Count;
         bool flag;
         double temp;
         Vector prevPoint = new Vector(size);
         double[] arg = new double[size + restriction_size];

         Func Q = t =>
         {
            double res = function(t);
            for (i = 0; i < equations_size; i++)
               res += r[i] * H(equations[i](t));
            for (i = 0; i < inequations_size; i++)
               // res += r[equations_size + i] * G()
               //res += (100 + 1.0 / r[equations_size + i]) * (1.0 -  1.0 / (1 + Math.Exp(10.0 / r[equations_size + i] * inequations[i](t))));
               res += inequations[i](t);
            return res;
         };

         calc_count = 0;
         norm = 0;

         for (i = 0; i < size; i++)
            prevPoint[i] = start[i];

         do
         {
            step++;

            DescentMethods.GaussAlgorithm(ExtremumType.Minimum, Q, start, out calc, out _);
            calc_count += calc;

            flag = true;
            for (i = 0; i < equations_size; i++)
               if (Math.Abs(equations[i](start.values)) >= Epsilon)
               {
                  flag = false;
                  break;
               }

            for (i = 0; i < inequations_size; i++)
               if ((temp = inequations[i](start.values)) > 0.0 && Math.Abs(temp) >= Epsilon)
               {
                  flag = false;
                  break;
               }

            if (!flag)
            {
               start.values = prevPoint.values;
               break;
            }

            if ((norm = prevPoint.Distance(start)) < Epsilon * Epsilon)
               break;

            for (i = 0; i < size; i++)
               prevPoint[i] = start[i];

            for (i = 0; i < restriction_size; i++)
               r[i] = function_r(r[i]);
         } while (flag);

         return step;
      }
   }
}