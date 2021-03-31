using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NumMath;
using EParser;

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
    static class DescentMethods
    {
        public static double Epsilon = 1E-7;
        public static double Delta = 1E-7;
        private static void getGradient(Func function, Vector point, Vector gradient)
        {
            double h = 1E-10;
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

            for(int i = 0; i < point.size; i++)
            {
                temp = x[i];
                x[i] = temp - h;
                f_left = function(x);

                x[i] = temp + h;
                f_right = function(x);

                gradient[i] = (f_left - f_right) / (2.0 * h);
            }
        }
        private static double findLambda(ExtremumType type, Func<double, double> function, double start)
        {
            double v_left, v_right;
            double left, right;
            OneDimensionalSearches.FindInterval(type, start, function, out v_left, out v_right);
            OneDimensionalSearches.FibonacciMethod(type, v_left, v_right, function, out left, out right);
            return (right + left) / 2.0;
        }
        public static int CGM_FletcherReeves(ExtremumType type, Func function, Vector start)
        {
            int i;
            int size = start.size;
            int step = 0;
            int iterations = 0;

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
                        for (i = 0; i < size; i++)
                            s[i] = r[i];
                        cur_innerMult = r.InnerMult(r);

                        while (true)
                        {
                            step++;

                            // finding minimum of g(v)
                            lambda = findLambda(type, g, 0.0);

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

                            if (step > size + 1)
                            {
                                iterations += step;
                                step = 0;
                                break;
                            }
                            else if (s.sqrMagnitude < Epsilon * Epsilon)
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
                        for (i = 0; i < size; i++)
                            s[i] = r[i];
                        cur_innerMult = r.InnerMult(r);

                        while (true)
                        {
                            step++;

                            // finding maximum of g(v)
                            lambda = findLambda(type, g, 0.0);

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

                            if (step > size + 1)
                            {
                                iterations += step;
                                step = 0;
                                break;
                            }
                            else if (s.sqrMagnitude < Epsilon * Epsilon)
                            {
                                iterations += step;
                                return iterations;
                            }
                        }
                    }
                    break;
            }
            return -1;
        }
        public static int PiersonAlgorithm(ExtremumType type, Func function, Vector start)
        {
            int i, j, k;
            int size = start.size;
            int step = 0;

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
            }; ;

            double[,] approxMatrix = new double[size, size];
            for (i = 0; i < size; i++)
                approxMatrix[i, i] = 1.0;

            switch(type)
            {
                case ExtremumType.Minimum:
                    do
                    {
                        getGradient(function, start, gradient);
                        for (j = 0; j < size; j++)
                        {
                            dx[j] = 0;
                            for (k = 0; k < size; k++)
                                dx[j] += approxMatrix[j, k] * gradient[j];
                            dx[j] *= -1.0;
                        }

                        lambda = findLambda(type, g, 0.0);
                        distance = 0;
                        for (i = 0; i < size; i++)
                        {
                            dx[i] *= lambda;
                            start[i] += dx[i];
                            distance += dx[i] * dx[i];
                        }

                        // dgradient = gradient(x_k+1)
                        getGradient(function, start, dgradient);
                        if (dgradient.sqrMagnitude < Epsilon * Epsilon || distance < Delta)
                            break;
                        step++;

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
                        getGradient(function, start, gradient);
                        for (i = 0; i < size; i++)
                        {
                            dx[i] = 0;
                            for (j = 0; j < size; j++)
                                dx[i] += approxMatrix[i, j] * gradient[j];
                        }

                        lambda = findLambda(type, g, 0.0);
                        distance = 0;
                        for (i = 0; i < size; i++)
                        {
                            dx[i] *= lambda;
                            start[i] += dx[i];
                            distance += dx[i] * dx[i];
                        }

                        // dgradient = gradient(x_k+1)
                        getGradient(function, start, dgradient);
                        if (dgradient.sqrMagnitude < Epsilon * Epsilon || distance < Delta)
                            break;
                        step++;

                        // dgradient = gradient(x_k+1) - gradient(x_k)
                        for (i = 0; i < size; i++)
                            dgradient[i] -= gradient[i];

                        // lambda = (dg)^T * approxMat * dg
                        // gradient = approxMat * dg
                        // dx = dx - approxMat * dg = dx - gradient
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
            }
            return step;            
        }
        public static int CGM_FletcherReevesResearch(ExtremumType type, Func function, Vector start, out int calc_count)
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
                            lambda = findLambda(type, g, 0.0);

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
                            else if (s.sqrMagnitude < Epsilon * Epsilon)
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
                        for (i = 0; i < size; i++)
                            s[i] = r[i];
                        cur_innerMult = r.InnerMult(r);

                        while (true)
                        {
                            step++;

                            // finding minimum of g(v)
                            lambda = findLambda(type, g, 0.0);

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

                            if (step > size + 1)
                            {
                                iterations += step;
                                step = 0;
                                break;
                            }
                            else if (s.sqrMagnitude < Epsilon * Epsilon)
                            {
                                iterations += step;
                                return iterations;
                            }
                        }
                    }
                    break;
            }
            return -1;
        }
        public static int PiersonAlgorithm(ExtremumType type, Func function, Vector start, out int calc_count)
        {
            int i, j, k;
            int size = start.size;
            int step = 0;
            calc_count = 0;

            Vector gradient = new Vector(size);
            Vector dgradient = new Vector(size);
            double[] temp = new double[size];
            double[] dx = new double[size];

            double lambda;
            Func<double, double> g = v =>
            {
                for (int s = 0; s < size; s++)
                    temp[s] = start[s] + v * dx[s];
                return function(temp);
            }; ;

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

                        for (j = 0; j < size; j++)
                        {
                            dx[j] = 0;
                            for (k = 0; k < size; k++)
                                dx[j] += approxMatrix[j, k] * gradient[j];
                            dx[j] *= -1.0;
                        }

                        lambda = findLambda(type, g, 0.0);
                        for (i = 0; i < size; i++)
                        {
                            dx[i] *= lambda;
                            start[i] += dx[i];
                        }

                        // dgradient = gradient(x_k+1)
                        getGradient(function, start, dgradient);
                        calc_count += 2 * size;

                        if (dgradient.sqrMagnitude < Epsilon * Epsilon)
                            break;
                        step++;

                        // dgradient = gradient(x_k+1) - gradient(x_k)
                        for (i = 0; i < size; i++)
                            dgradient[i] -= gradient[i];

                        // lambda = (dg)^T * approxMat * dg
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

                        for (j = 0; j < size; j++)
                        {
                            dx[j] = 0;
                            for (k = 0; k < size; k++)
                                dx[j] += approxMatrix[j, k] * gradient[j];
                            dx[j] *= -1.0;
                        }

                        lambda = findLambda(type, g, 0.0);
                        for (i = 0; i < size; i++)
                        {
                            dx[i] *= lambda;
                            start[i] += dx[i];
                        }

                        // dgradient = gradient(x_k+1)
                        getAntigradient(function, start, dgradient);
                        calc_count += 2 * size;

                        if (dgradient.sqrMagnitude < Epsilon * Epsilon)
                            break;
                        step++;

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
            }
            return step;
        }
    }
}
