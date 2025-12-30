/*  Geometry Algorithms - C++ DLL for complex geometric operations
 *
 *  This DLL provides high-performance geometric algorithms that can be called from C# via P/Invoke
 */

#pragma once

#ifdef RENGAGEOMETRY_EXPORTS
#define RENGAGEOMETRY_API __declspec(dllexport)
#else
#define RENGAGEOMETRY_API __declspec(dllimport)
#endif

// Структуры должны совпадать с C# версиями
struct Point3D
{
    double X;
    double Y;
    double Z;
};

struct GeometryResult
{
    int Success;
    double Value1;
    double Value2;
    double Value3;
    wchar_t Message[256];
};

// Экспортируемые функции (C-style для P/Invoke)
extern "C" {
    // Вычисление сложной геометрической формы
    RENGAGEOMETRY_API GeometryResult CalculateComplexGeometry(
        const Point3D* points,
        int pointCount,
        double parameter1,
        double parameter2
    );

    // Поиск пути между точками
    RENGAGEOMETRY_API int FindPathBetweenPoints(
        int startPointId,
        int endPointId,
        int* path,
        int* pathLength
    );

    // Создание сложной кривой
    RENGAGEOMETRY_API void* CreateComplexCurve(
        const Point3D* controlPoints,
        int pointCount
    );

    // Освобождение памяти
    RENGAGEOMETRY_API void FreeMemory(void* ptr);
}







