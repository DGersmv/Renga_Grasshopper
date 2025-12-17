/*  Example: Using C++ geometry module from C# plugin
 *
 *  This shows how to integrate C++ geometric algorithms into your C# Renga plugin
 */

using System;
using Renga;
using System.Linq;

namespace RengaPlugin
{
    public partial class RengaPlugin
    {
        /// <summary>
        /// Example: Create complex geometry using C++ module
        /// </summary>
        private void CreateComplexGeometryExample()
        {
            try
            {
                // Получить точки из модели или от пользователя
                var model = m_app.Project.Model;
                var objects = model.GetObjects();
                
                // Пример: получить точки из выбранных объектов
                var selectedIds = (int[])m_app.Selection.GetSelectedObjects();
                if (selectedIds.Length < 2)
                {
                    m_app.UI.ShowMessageBox(
                        Renga.MessageIcon.MessageIcon_Warning,
                        "Geometry Example",
                        "Please select at least 2 objects");
                    return;
                }

                // Конвертировать в Point3D для C++ модуля
                var points = new GeometryHelper.Point3D[selectedIds.Length];
                for (int i = 0; i < selectedIds.Length; i++)
                {
                    var obj = objects.GetById(selectedIds[i]);
                    var placement = (obj as ILevelObject)?.GetPlacement();
                    if (placement != null)
                    {
                        points[i] = new GeometryHelper.Point3D
                        {
                            X = placement.Origin.X,
                            Y = placement.Origin.Y,
                            Z = placement.Origin.Z
                        };
                    }
                }

                // Вызвать C++ функцию для вычисления геометрии
                var result = GeometryHelper.CalculateComplexGeometry(
                    points, 
                    points.Length, 
                    0.1,  // parameter1
                    0.0   // parameter2
                );

                if (result.Success == 1)
                {
                    // Использовать результат для создания объектов в Renga
                    var message = $"Geometry calculated:\n" +
                                 $"Width: {result.Value1:F2}\n" +
                                 $"Depth: {result.Value2:F2}\n" +
                                 $"Height: {result.Value3:F2}";
                    
                    m_app.UI.ShowMessageBox(
                        Renga.MessageIcon.MessageIcon_Info,
                        "Geometry Result",
                        message);
                }
                else
                {
                    m_app.UI.ShowMessageBox(
                        Renga.MessageIcon.MessageIcon_Error,
                        "Geometry Error",
                        result.Message);
                }
            }
            catch (DllNotFoundException)
            {
                m_app.UI.ShowMessageBox(
                    Renga.MessageIcon.MessageIcon_Error,
                    "Missing DLL",
                    "RengaGeometry.dll not found. Please build the C++ project first.");
            }
            catch (Exception ex)
            {
                m_app.UI.ShowMessageBox(
                    Renga.MessageIcon.MessageIcon_Error,
                    "Error",
                    $"Failed to calculate geometry: {ex.Message}");
            }
        }

        /// <summary>
        /// Example: Find path between two objects (like MepCircleConstructor)
        /// </summary>
        private void FindPathExample()
        {
            var selectedIds = (int[])m_app.Selection.GetSelectedObjects();
            if (selectedIds.Length != 2)
            {
                m_app.UI.ShowMessageBox(
                    Renga.MessageIcon.MessageIcon_Warning,
                    "Path Finding",
                    "Please select exactly 2 objects");
                return;
            }

            try
            {
                int maxPathLength = 100;
                int[] path = new int[maxPathLength];
                int actualLength = maxPathLength;

                int result = GeometryHelper.FindPathBetweenPoints(
                    selectedIds[0],
                    selectedIds[1],
                    path,
                    ref actualLength
                );

                if (result == 1 && actualLength > 0)
                {
                    Array.Resize(ref path, actualLength);
                    var pathString = string.Join(" → ", path);
                    
                    m_app.UI.ShowMessageBox(
                        Renga.MessageIcon.MessageIcon_Info,
                        "Path Found",
                        $"Path: {pathString}");
                }
                else
                {
                    m_app.UI.ShowMessageBox(
                        Renga.MessageIcon.MessageIcon_Warning,
                        "Path Finding",
                        "No path found between selected objects");
                }
            }
            catch (Exception ex)
            {
                m_app.UI.ShowMessageBox(
                    Renga.MessageIcon.MessageIcon_Error,
                    "Error",
                    $"Path finding failed: {ex.Message}");
            }
        }

        /// <summary>
        /// Example: Create complex curve and use it in Renga
        /// </summary>
        private void CreateComplexCurveExample()
        {
            // Пример контрольных точек для кривой
            var controlPoints = new GeometryHelper.Point3D[]
            {
                new GeometryHelper.Point3D { X = 0, Y = 0, Z = 0 },
                new GeometryHelper.Point3D { X = 5, Y = 5, Z = 2 },
                new GeometryHelper.Point3D { X = 10, Y = 3, Z = 5 },
                new GeometryHelper.Point3D { X = 15, Y = 8, Z = 3 }
            };

            try
            {
                // Создать кривую в C++ модуле
                IntPtr curvePtr = GeometryHelper.CreateComplexCurve(
                    controlPoints,
                    controlPoints.Length
                );

                if (curvePtr != IntPtr.Zero)
                {
                    try
                    {
                        // Здесь можно использовать кривую для создания объектов в Renga
                        // Например, создать балку или трубу по этой кривой
                        
                        m_app.UI.ShowMessageBox(
                            Renga.MessageIcon.MessageIcon_Info,
                            "Curve Created",
                            $"Complex curve created with {controlPoints.Length} control points");
                    }
                    finally
                    {
                        // Важно: освободить память, выделенную в C++
                        GeometryHelper.FreeMemory(curvePtr);
                    }
                }
            }
            catch (Exception ex)
            {
                m_app.UI.ShowMessageBox(
                    Renga.MessageIcon.MessageIcon_Error,
                    "Error",
                    $"Failed to create curve: {ex.Message}");
            }
        }
    }
}



