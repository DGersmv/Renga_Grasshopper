/*  Geometry Helper - C# wrapper for C++ geometric algorithms
 *
 *  This class provides P/Invoke interface to C++ DLL for complex geometric operations
 */

using System;
using System.Runtime.InteropServices;

namespace RengaPlugin
{
    /// <summary>
    /// Wrapper for C++ geometric algorithms DLL
    /// </summary>
    public static class GeometryHelper
    {
        // Имя DLL (будет RengaGeometry.dll)
        private const string DllName = "RengaGeometry.dll";

        // Пример: структура для передачи точек в C++
        [StructLayout(LayoutKind.Sequential)]
        public struct Point3D
        {
            public double X;
            public double Y;
            public double Z;
        }

        // Пример: структура для результата геометрических вычислений
        [StructLayout(LayoutKind.Sequential)]
        public struct GeometryResult
        {
            public int Success;
            public double Value1;
            public double Value2;
            public double Value3;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 256)]
            public string Message;
        }

        /// <summary>
        /// Пример: Вычисление сложной геометрической формы
        /// </summary>
        [DllImport(DllName, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Unicode)]
        public static extern GeometryResult CalculateComplexGeometry(
            [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 1)]
            Point3D[] points,
            int pointCount,
            double parameter1,
            double parameter2
        );

        /// <summary>
        /// Пример: Построение пути между точками (как в MepCircleConstructor)
        /// </summary>
        [DllImport(DllName, CallingConvention = CallingConvention.Cdecl)]
        public static extern int FindPathBetweenPoints(
            int startPointId,
            int endPointId,
            [Out, MarshalAs(UnmanagedType.LPArray)] int[] path,
            ref int pathLength
        );

        /// <summary>
        /// Пример: Создание сложной кривой
        /// </summary>
        [DllImport(DllName, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr CreateComplexCurve(
            [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 1)]
            Point3D[] controlPoints,
            int pointCount
        );

        /// <summary>
        /// Освобождение памяти, выделенной в C++
        /// </summary>
        [DllImport(DllName, CallingConvention = CallingConvention.Cdecl)]
        public static extern void FreeMemory(IntPtr ptr);

        /// <summary>
        /// C# обертка для удобного использования
        /// </summary>
        public static Renga.Point3D[] CalculatePath(Renga.Point3D[] inputPoints, double tolerance)
        {
            // Конвертация Renga.Point3D в Point3D для C++
            var points = new Point3D[inputPoints.Length];
            for (int i = 0; i < inputPoints.Length; i++)
            {
                points[i] = new Point3D
                {
                    X = inputPoints[i].X,
                    Y = inputPoints[i].Y,
                    Z = inputPoints[i].Z
                };
            }

            // Вызов C++ функции
            var result = CalculateComplexGeometry(points, points.Length, tolerance, 0.0);

            if (result.Success == 0)
            {
                throw new Exception($"C++ geometry calculation failed: {result.Message}");
            }

            // Конвертация результата обратно в Renga.Point3D
            // (пример - зависит от вашей реализации)
            return inputPoints; // Замените на реальную конвертацию
        }
    }
}





