using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.IO;
using Renga;
using RengaPlugin.Commands;
using RengaPlugin.Connection;
using Newtonsoft.Json.Linq;

namespace RengaPlugin.Handlers
{
    /// <summary>
    /// Handler for get_walls command
    /// </summary>
    public class GetWallsHandler : ICommandHandler
    {
        private Renga.IApplication m_app;
        private static string logFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Renga", "RengaGH_GetWalls.log");

        public GetWallsHandler(Renga.IApplication app)
        {
            m_app = app;
        }

        private void LogToFile(string message)
        {
            try
            {
                var logDir = Path.GetDirectoryName(logFilePath);
                if (!Directory.Exists(logDir))
                    Directory.CreateDirectory(logDir);
                
                File.AppendAllText(logFilePath, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] {message}\n");
            }
            catch { }
        }

        public ConnectionResponse Handle(ConnectionMessage message)
        {
            try
            {
                var model = m_app.Project.Model;
                if (model == null)
                {
                    return new ConnectionResponse
                    {
                        Id = message.Id,
                        Success = false,
                        Error = "No active model"
                    };
                }

                var walls = new List<object>();
                var objects = model.GetObjects();
                int count = objects.Count;
                
                // Track wall IDs to ensure uniqueness
                var wallIds = new HashSet<int>();

                // Get exported 3D objects for mesh extraction
                var dataExporter = m_app.Project.DataExporter;
                var exportedObjects3D = dataExporter?.GetObjects3D();
                var exportedObjectsMap = new Dictionary<int, Renga.IExportedObject3D>();
                
                if (exportedObjects3D != null)
                {
                    int exportedCount = exportedObjects3D.Count;
                    for (int i = 0; i < exportedCount; i++)
                    {
                        try
                        {
                            var exportedObj = exportedObjects3D.Get(i);
                            exportedObjectsMap[exportedObj.ModelObjectId] = exportedObj;
                        }
                        catch { }
                    }
                }

                for (int i = 0; i < count; i++)
                {
                    try
                    {
                        var obj = objects.GetByIndex(i);
                        if (obj.ObjectType == Renga.ObjectTypes.Wall)
                        {
                            var wall = obj as Renga.ILevelObject;
                            if (wall != null)
                            {
                                var modelObject = obj as Renga.IModelObject;
                                int wallId = modelObject.Id;
                                
                                // Check for duplicate IDs (should not happen, but safety check)
                                if (wallIds.Contains(wallId))
                                {
                                    LogToFile($"WARNING: Duplicate wall ID detected: {wallId}. Skipping duplicate.");
                                    continue;
                                }
                                wallIds.Add(wallId);
                                
                                var placement = wall.GetPlacement();
                                
                                // Get wall parameters
                                double height = 0;
                                double thickness = 0;
                                try
                                {
                                    var parameters = modelObject?.GetParameters();
                                    if (parameters != null)
                                    {
                                        try
                                        {
                                            var heightParam = parameters.Get(Renga.ParameterIds.WallHeight);
                                            if (heightParam != null)
                                                height = heightParam.GetDoubleValue();
                                        }
                                        catch { }
                                        
                                        try
                                        {
                                            var thicknessParam = parameters.Get(Renga.ParameterIds.WallThickness);
                                            if (thicknessParam != null)
                                                thickness = thicknessParam.GetDoubleValue();
                                        }
                                        catch { }
                                    }
                                }
                                catch { }
                                
                                // Get baseline curve using IWallParams and IWallContour
                                object baselineData = ExtractBaselineData(obj, placement);
                                
                                // Get mesh geometry from exported 3D object
                                object meshData = null;
                                if (exportedObjectsMap.TryGetValue(obj.Id, out var exportedObj3D))
                                {
                                    meshData = ExtractMeshDataFromExported(exportedObj3D);
                                }
                                
                                var wallData = new
                                {
                                    id = obj.Id,
                                    name = obj.Name ?? $"Wall {obj.Id}",
                                    position = placement != null ? new
                                    {
                                        x = placement.Origin.X,
                                        y = placement.Origin.Y,
                                        z = placement.Origin.Z
                                    } : null,
                                    height = height,
                                    thickness = thickness,
                                    baseline = baselineData,
                                    mesh = meshData
                                };
                                walls.Add(wallData);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        // Skip this object if there's an error
                        LogToFile($"Error processing wall object: {ex.Message}");
                    }
                }

                var responseData = new JObject
                {
                    ["walls"] = JArray.FromObject(walls)
                };

                return new ConnectionResponse
                {
                    Id = message.Id,
                    Success = true,
                    Data = responseData
                };
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"GetWallsHandler error: {ex.Message}\n{ex.StackTrace}");
                return new ConnectionResponse
                {
                    Id = message.Id,
                    Success = false,
                    Error = $"Error getting walls: {ex.Message}"
                };
            }
        }

        private object ExtractBaselineData(Renga.IModelObject wallObj, Renga.IPlacement3D placement)
        {
            try
            {
                LogToFile($"Wall {wallObj.Id}: Starting baseline extraction...");
                
                // CRITICAL FIX: Use IWallParams -> GetContour() -> IWallContour.GetBaseline()
                // This gives us baseline2D which is already in GLOBAL coordinates
                var wallParams = wallObj as Renga.IWallParams;
                if (wallParams == null)
                {
                    LogToFile($"Wall {wallObj.Id}: Failed to cast to IWallParams");
                    return null;
                }

                var wallContour = wallParams.GetContour();
                if (wallContour == null)
                {
                    LogToFile($"Wall {wallObj.Id}: GetContour() returned null");
                    return null;
                }

                var baseline2D = wallContour.GetBaseline();
                if (baseline2D == null)
                {
                    LogToFile($"Wall {wallObj.Id}: GetBaseline() returned null");
                    return null;
                }

                // Get Z coordinate from placement if available
                double zCoord = 0.0;
                if (placement != null)
                {
                    zCoord = placement.Origin.Z;
                    LogToFile($"Wall {wallObj.Id}: placement origin=({placement.Origin.X}, {placement.Origin.Y}, {placement.Origin.Z})");
                }

                // CRITICAL: baseline2D from IWallContour.GetBaseline() is already in GLOBAL coordinates
                // Do NOT use CreateCurve3D with placement, as it applies incorrect transformation
                // Use baseline2D coordinates directly (they are already global)

                var baselineData = new Dictionary<string, object>();
                
                // Get curve type
                var curveType = baseline2D.Curve2DType;
                string curveTypeStr = curveType.ToString();
                baselineData["type"] = curveTypeStr;
                LogToFile($"Wall {wallObj.Id}: baseline2D type={curveTypeStr}");

                // Get start and end points from baseline2D (already global coordinates)
                var startPoint2D = baseline2D.GetBeginPoint();
                var endPoint2D = baseline2D.GetEndPoint();
                
                LogToFile($"Wall {wallObj.Id}: baseline2D start=({startPoint2D.X}, {startPoint2D.Y}), end=({endPoint2D.X}, {endPoint2D.Y})");

                // Convert to 3D by adding Z coordinate
                var startPoint3D = new Renga.Point3D
                {
                    X = startPoint2D.X,
                    Y = startPoint2D.Y,
                    Z = zCoord
                };

                var endPoint3D = new Renga.Point3D
                {
                    X = endPoint2D.X,
                    Y = endPoint2D.Y,
                    Z = zCoord
                };

                baselineData["startPoint"] = new Dictionary<string, object>
                {
                    { "x", startPoint3D.X },
                    { "y", startPoint3D.Y },
                    { "z", startPoint3D.Z }
                };

                baselineData["endPoint"] = new Dictionary<string, object>
                {
                    { "x", endPoint3D.X },
                    { "y", endPoint3D.Y },
                    { "z", endPoint3D.Z }
                };

                // Handle different curve types
                List<Dictionary<string, object>> sampledPoints = null;

                // Check curve type using Curve2DType enum
                if (curveType == Renga.Curve2DType.Curve2DType_LineSegment)
                {
                    // For lines, just use start and end points
                    sampledPoints = new List<Dictionary<string, object>>
                    {
                        new Dictionary<string, object> { { "x", startPoint3D.X }, { "y", startPoint3D.Y }, { "z", startPoint3D.Z } },
                        new Dictionary<string, object> { { "x", endPoint3D.X }, { "y", endPoint3D.Y }, { "z", endPoint3D.Z } }
                    };
                }
                else if (curveType == Renga.Curve2DType.Curve2DType_Arc)
                {
                    // CRITICAL: For arcs, we MUST use the exact direction from baseline2D
                    // Do NOT choose shortest arc - use the actual wall baseline direction
                    try
                    {
                        var arc2D = baseline2D as Renga.IArc2D;
                        if (arc2D != null)
                        {
                            var center2D = arc2D.GetCenter();
                            double radius = arc2D.GetRadius();
                            
                            // CRITICAL: Use CreateCurve3D to get sampled points in the CORRECT direction
                            // This preserves the actual wall baseline direction, not the shortest arc
                            // Use null placement to avoid double transformation (baseline2D is already global)
                            var baseline3D = baseline2D.CreateCurve3D(null);
                            if (baseline3D == null)
                            {
                                // Fallback: try with placement if null doesn't work
                                baseline3D = baseline2D.CreateCurve3D(placement);
                            }
                            
                            if (baseline3D != null)
                            {
                                LogToFile($"Wall {wallObj.Id}: Using CreateCurve3D for arc sampling to preserve correct direction");
                                
                                // Sample points from the actual curve direction
                                sampledPoints = new List<Dictionary<string, object>>();
                                int arcSamples = 50;
                                
                                // Always add start point
                                try
                                {
                                    var startPt = baseline3D.GetBeginPoint();
                                    sampledPoints.Add(new Dictionary<string, object>
                                    {
                                        { "x", startPt.X },
                                        { "y", startPt.Y },
                                        { "z", startPt.Z }
                                    });
                                }
                                catch (Exception ex)
                                {
                                    LogToFile($"Wall {wallObj.Id}: Error getting start point: {ex.Message}");
                                }
                                
                                // Try to get intermediate points
                                for (int i = 1; i < arcSamples; i++)
                                {
                                    double t = (double)i / arcSamples;
                                    try
                                    {
                                        // Try to get point at parameter using reflection
                                        Renga.Point3D? point3D = null;
                                        
                                        var baselineType = baseline3D.GetType();
                                        var getPointAtMethod = baselineType.GetMethod("GetPointAt", new Type[] { typeof(double) });
                                        if (getPointAtMethod != null)
                                        {
                                            var result = getPointAtMethod.Invoke(baseline3D, new object[] { t });
                                            if (result is Renga.Point3D pt)
                                            {
                                                point3D = pt;
                                            }
                                        }
                                        
                                        if (point3D.HasValue)
                                        {
                                            var pt = point3D.Value;
                                            sampledPoints.Add(new Dictionary<string, object>
                                            {
                                                { "x", pt.X },
                                                { "y", pt.Y },
                                                { "z", pt.Z }
                                            });
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        LogToFile($"Wall {wallObj.Id}: Error sampling arc point at t={t}: {ex.Message}");
                                    }
                                }
                                
                                // Always add end point
                                try
                                {
                                    var endPt = baseline3D.GetEndPoint();
                                    sampledPoints.Add(new Dictionary<string, object>
                                    {
                                        { "x", endPt.X },
                                        { "y", endPt.Y },
                                        { "z", endPt.Z }
                                    });
                                }
                                catch (Exception ex)
                                {
                                    LogToFile($"Wall {wallObj.Id}: Error getting end point: {ex.Message}");
                                }
                                
                                // Also store center and radius for reference
                                baselineData["center"] = new Dictionary<string, object>
                                {
                                    { "x", center2D.X },
                                    { "y", center2D.Y },
                                    { "z", zCoord }
                                };
                                baselineData["radius"] = radius;
                                
                                LogToFile($"Wall {wallObj.Id}: Sampled {sampledPoints.Count} points from arc using CreateCurve3D (preserving wall direction)");
                                
                                // If we got only 2 points (start and end), fall back to trigonometric calculation
                                if (sampledPoints.Count <= 2)
                                {
                                    LogToFile($"Wall {wallObj.Id}: Only got {sampledPoints.Count} points from CreateCurve3D, falling back to trigonometric calculation");
                                    sampledPoints = null; // Will trigger fallback below
                                }
                            }
                            
                            // Fallback: if CreateCurve3D failed or returned only 2 points, use trigonometric calculation
                            if (baseline3D == null || sampledPoints == null || sampledPoints.Count <= 2)
                            {
                                if (baseline3D == null)
                                    LogToFile($"Wall {wallObj.Id}: CreateCurve3D returned null, using trigonometric calculation");
                                else
                                    LogToFile($"Wall {wallObj.Id}: CreateCurve3D returned insufficient points, using trigonometric calculation");
                                
                                // Calculate angles from actual start and end points
                                double startAngle = Math.Atan2(startPoint2D.Y - center2D.Y, startPoint2D.X - center2D.X);
                                double endAngle = Math.Atan2(endPoint2D.Y - center2D.Y, endPoint2D.X - center2D.X);
                                
                                // CRITICAL: Do NOT choose shortest arc - preserve direction
                                // If endAngle < startAngle, we need to go the long way (add 2π)
                                if (endAngle < startAngle)
                                    endAngle += 2 * Math.PI;
                                
                                LogToFile($"Wall {wallObj.Id}: Arc center=({center2D.X}, {center2D.Y}), radius={radius}, startAngle={startAngle} ({startAngle * 180 / Math.PI}°), endAngle={endAngle} ({endAngle * 180 / Math.PI}°)");

                                // Sample points following the actual direction (not shortest)
                                sampledPoints = new List<Dictionary<string, object>>();
                                int arcSamples = 50;
                                for (int i = 0; i <= arcSamples; i++)
                                {
                                    double t = (double)i / arcSamples;
                                    double angle = startAngle + (endAngle - startAngle) * t;
                                    var point2D = new Renga.Point2D
                                    {
                                        X = center2D.X + radius * Math.Cos(angle),
                                        Y = center2D.Y + radius * Math.Sin(angle)
                                    };
                                    // Use baseline2D directly as global coordinates (no placement transformation)
                                    var point3D = new Renga.Point3D
                                    {
                                        X = point2D.X,
                                        Y = point2D.Y,
                                        Z = zCoord
                                    };
                                    sampledPoints.Add(new Dictionary<string, object>
                                    {
                                        { "x", point3D.X },
                                        { "y", point3D.Y },
                                        { "z", point3D.Z }
                                    });
                                }
                                baselineData["center"] = new Dictionary<string, object>
                                {
                                    { "x", center2D.X },
                                    { "y", center2D.Y },
                                    { "z", zCoord }
                                };
                                baselineData["radius"] = radius;
                                LogToFile($"Wall {wallObj.Id}: Recalculated {sampledPoints.Count} sampled points for Arc (preserving direction)");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        LogToFile($"Wall {wallObj.Id}: Error processing arc: {ex.Message}");
                    }
                }
                else if (curveType == Renga.Curve2DType.Curve2DType_PolyCurve)
                {
                    var polyCurve2D = baseline2D as Renga.IPolyCurve2D;
                    if (polyCurve2D != null)
                    {
                        // For polycurves, sample each segment
                        sampledPoints = new List<Dictionary<string, object>>();
                        int segmentCount = polyCurve2D.GetSegmentCount();
                        LogToFile($"Wall {wallObj.Id}: PolyCurve with {segmentCount} segments");
                        
                        for (int segIdx = 0; segIdx < segmentCount; segIdx++)
                        {
                            try
                            {
                                var segment = polyCurve2D.GetSegment(segIdx);
                                if (segment != null)
                                {
                                    // Sample the segment using CreateCurve3D for complex curves
                                    // But only for sampling, not for the main coordinates
                                    var segment3D = segment.CreateCurve3D(placement);
                                    if (segment3D != null)
                                    {
                                        int samples = 10;
                                        for (int i = 0; i <= samples; i++)
                                        {
                                            double t = (double)i / samples;
                                        try
                                        {
                                            // Try different methods to get point at parameter
                                            Renga.Point3D? point3D = null;
                                            
                                            // Try GetPointAt method via reflection
                                            var segmentType = segment3D.GetType();
                                            var getPointAtMethod = segmentType.GetMethod("GetPointAt", new Type[] { typeof(double) });
                                            if (getPointAtMethod != null)
                                            {
                                                var result = getPointAtMethod.Invoke(segment3D, new object[] { t });
                                                if (result is Renga.Point3D pt)
                                                {
                                                    point3D = pt;
                                                }
                                            }
                                            
                                            // Fallback: use start/end points for first/last sample
                                            if (!point3D.HasValue)
                                            {
                                                if (i == 0)
                                                    point3D = segment3D.GetBeginPoint();
                                                else if (i == samples)
                                                    point3D = segment3D.GetEndPoint();
                                                else
                                                    continue; // Skip intermediate points if we can't get them
                                            }
                                            
                                            if (point3D.HasValue)
                                            {
                                                var pt = point3D.Value;
                                                sampledPoints.Add(new Dictionary<string, object>
                                                {
                                                    { "x", pt.X },
                                                    { "y", pt.Y },
                                                    { "z", pt.Z }
                                                });
                                            }
                                        }
                                            catch { }
                                        }
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                LogToFile($"Wall {wallObj.Id}: Error sampling segment {segIdx}: {ex.Message}");
                            }
                        }
                        
                        if (sampledPoints.Count == 0)
                        {
                            // Fallback: use start and end points
                            sampledPoints.Add(new Dictionary<string, object> { { "x", startPoint3D.X }, { "y", startPoint3D.Y }, { "z", startPoint3D.Z } });
                            sampledPoints.Add(new Dictionary<string, object> { { "x", endPoint3D.X }, { "y", endPoint3D.Y }, { "z", endPoint3D.Z } });
                        }
                    }
                }
                else
                {
                    // For other curve types, try to sample using CreateCurve3D
                    // But remember: baseline2D coordinates are already global
                    try
                    {
                        var baseline3D = baseline2D.CreateCurve3D(placement);
                        if (baseline3D != null)
                        {
                            LogToFile($"Wall {wallObj.Id}: Using CreateCurve3D for sampling (type: {curveTypeStr})");
                            sampledPoints = new List<Dictionary<string, object>>();
                            int samples = 50;
                            for (int i = 0; i <= samples; i++)
                            {
                                double t = (double)i / samples;
                                try
                                {
                                    // Try different methods to get point at parameter
                                    Renga.Point3D? point3D = null;
                                    
                                    // Try GetPointAt method via reflection
                                    var baselineType = baseline3D.GetType();
                                    var getPointAtMethod = baselineType.GetMethod("GetPointAt", new Type[] { typeof(double) });
                                    if (getPointAtMethod != null)
                                    {
                                        var result = getPointAtMethod.Invoke(baseline3D, new object[] { t });
                                        if (result is Renga.Point3D pt)
                                        {
                                            point3D = pt;
                                        }
                                    }
                                    
                                    // Fallback: use start/end points for first/last sample
                                    if (!point3D.HasValue)
                                    {
                                        if (i == 0)
                                            point3D = baseline3D.GetBeginPoint();
                                        else if (i == samples)
                                            point3D = baseline3D.GetEndPoint();
                                        else
                                            continue; // Skip intermediate points if we can't get them
                                    }
                                    
                                    if (point3D.HasValue)
                                    {
                                        var pt = point3D.Value;
                                        sampledPoints.Add(new Dictionary<string, object>
                                        {
                                            { "x", pt.X },
                                            { "y", pt.Y },
                                            { "z", pt.Z }
                                        });
                                    }
                                }
                                catch { }
                            }
                            LogToFile($"Wall {wallObj.Id}: Sampled {sampledPoints.Count} points using CreateCurve3D");
                        }
                    }
                    catch (Exception ex)
                    {
                        LogToFile($"Wall {wallObj.Id}: Error using CreateCurve3D: {ex.Message}");
                    }
                    
                    // Fallback: use start and end points
                    if (sampledPoints == null || sampledPoints.Count == 0)
                    {
                        sampledPoints = new List<Dictionary<string, object>>
                        {
                            new Dictionary<string, object> { { "x", startPoint3D.X }, { "y", startPoint3D.Y }, { "z", startPoint3D.Z } },
                            new Dictionary<string, object> { { "x", endPoint3D.X }, { "y", endPoint3D.Y }, { "z", endPoint3D.Z } }
                        };
                    }
                }

                if (sampledPoints != null && sampledPoints.Count > 0)
                {
                    baselineData["sampledPoints"] = sampledPoints;
                }

                LogToFile($"Wall {wallObj.Id}: Baseline extraction completed successfully");
                return baselineData;
            }
            catch (Exception ex)
            {
                LogToFile($"Error extracting baseline: {ex.Message}\n{ex.StackTrace}");
                return null;
            }
        }

        private object ExtractBaselineFromCurve3D(Renga.ICurve3D baseline)
        {
            try
            {
                if (baseline == null)
                    return null;

                var baselineObj = new Dictionary<string, object>();
                
                // Get curve type using ICurve3D interface
                var curveType = baseline.Curve3DType;
                string curveTypeStr = curveType.ToString();
                baselineObj["type"] = curveTypeStr;
                
                // Get start and end points
                try
                {
                    var startPoint = baseline.GetBeginPoint();
                    var endPoint = baseline.GetEndPoint();
                    
                    baselineObj["startPoint"] = new { 
                        x = startPoint.X, 
                        y = startPoint.Y, 
                        z = startPoint.Z 
                    };
                    baselineObj["endPoint"] = new { 
                        x = endPoint.X, 
                        y = endPoint.Y, 
                        z = endPoint.Z 
                    };
                }
                catch (Exception ex)
                {
                    LogToFile($"Error getting start/end points: {ex.Message}");
                }

                // For polycurves, extract segments
                if (baseline is Renga.IPolyCurve3D polyCurve)
                {
                    var segments = new List<object>();
                    try
                    {
                        int segmentCount = polyCurve.GetSegmentCount();
                        for (int i = 0; i < segmentCount; i++)
                        {
                            try
                            {
                                var segment = polyCurve.GetSegment(i);
                                var segObj = ExtractBaselineFromCurve3D(segment);
                                if (segObj != null)
                                    segments.Add(segObj);
                            }
                            catch (Exception ex)
                            {
                                LogToFile($"Error extracting segment {i}: {ex.Message}");
                            }
                        }
                        baselineObj["segments"] = segments;
                    }
                    catch (Exception ex)
                    {
                        LogToFile($"Error extracting segments: {ex.Message}");
                    }
                }
                // For arcs, get center and radius
                else if (baseline is Renga.IArc3D arc)
                {
                    try
                    {
                        var center = arc.GetCenter();
                        baselineObj["center"] = new { 
                            x = center.X, 
                            y = center.Y, 
                            z = center.Z 
                        };
                        baselineObj["radius"] = arc.GetRadius();
                    }
                    catch (Exception ex)
                    {
                        LogToFile($"Error getting arc properties: {ex.Message}");
                    }
                }

                return baselineObj;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error extracting baseline from curve: {ex.Message}");
                return null;
            }
        }

        private object ExtractMeshDataFromExported(Renga.IExportedObject3D exportedObj3D)
        {
            try
            {
                if (exportedObj3D == null)
                    return null;

                var meshArray = new List<object>();
                int meshCount = exportedObj3D.MeshCount;
                
                LogToFile($"Extracting mesh from exported object, mesh count: {meshCount}");
                
                for (int i = 0; i < meshCount; i++)
                {
                    try
                    {
                        var mesh = exportedObj3D.GetMesh(i);
                        if (mesh == null)
                            continue;

                        // Get mesh type
                        var meshType = mesh.MeshType;
                        string meshTypeStr = meshType.ToString();
                        
                        var gridsList = new List<object>();
                        
                        // Get grids from mesh
                        int gridCount = mesh.GridCount;
                        for (int j = 0; j < gridCount; j++)
                        {
                            try
                            {
                                var gridObj = mesh.GetGrid(j);
                                if (gridObj == null)
                                    continue;

                                var gridData = new Dictionary<string, object>();
                                var vertices = new List<object>();
                                var triangles = new List<object>();

                                // Get grid type (for walls: FrontSide, BackSide, Bottom, Top, etc.)
                                var gridType = gridObj.GridType;
                                gridData["gridType"] = gridType.ToString();
                                
                                // Get vertices from grid
                                int vertexCount = gridObj.VertexCount;
                                for (int v = 0; v < vertexCount; v++)
                                {
                                    var vertex = gridObj.GetVertex(v);
                                    vertices.Add(new { x = vertex.X, y = vertex.Y, z = vertex.Z });
                                }

                                // Get triangles from grid
                                int triangleCount = gridObj.TriangleCount;
                                for (int t = 0; t < triangleCount; t++)
                                {
                                    var triangle = gridObj.GetTriangle(t);
                                    // Triangle has V0, V1, V2 properties (indices as uint)
                                    triangles.Add(new int[] { (int)triangle.V0, (int)triangle.V1, (int)triangle.V2 });
                                }
                                
                                if (vertices.Count > 0)
                                {
                                    gridData["vertices"] = vertices;
                                    gridData["triangles"] = triangles;
                                    gridsList.Add(gridData);
                                    LogToFile($"Grid {j} (type: {gridType}): {vertices.Count} vertices, {triangles.Count} triangles");
                                }
                            }
                            catch (Exception ex)
                            {
                                LogToFile($"Error extracting grid {j}: {ex.Message}");
                            }
                        }

                        if (gridsList.Count > 0)
                        {
                            var meshData = new Dictionary<string, object>
                            {
                                ["meshType"] = meshTypeStr,
                                ["grids"] = gridsList
                            };
                            meshArray.Add(meshData);
                            LogToFile($"Mesh {i} (type: {meshTypeStr}) extracted: {gridsList.Count} grids");
                        }
                    }
                    catch (Exception ex)
                    {
                        LogToFile($"Error extracting mesh {i}: {ex.Message}");
                    }
                }

                return meshArray.Count > 0 ? meshArray : null;
            }
            catch (Exception ex)
            {
                LogToFile($"Error extracting mesh from exported object: {ex.Message}\n{ex.StackTrace}");
                return null;
            }
        }

        private object ExtractMeshData(Renga.IModelObject modelObject)
        {
            try
            {
                if (modelObject == null)
                    return null;

                var meshArray = new List<object>();
                
                // Try to get geometry using reflection with multiple approaches
                try
                {
                    var objType = modelObject.GetType();
                    
                    // Try GetGeometry method (case-insensitive)
                    var allMethods = objType.GetMethods(System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
                    var getGeometryMethod = Array.Find(allMethods, m => 
                        m.Name.Equals("GetGeometry", StringComparison.OrdinalIgnoreCase) && 
                        m.GetParameters().Length == 0);
                    
                    if (getGeometryMethod != null)
                    {
                        var geometry = getGeometryMethod.Invoke(modelObject, null);
                        System.Diagnostics.Debug.WriteLine($"GetGeometry result: {(geometry != null ? "not null" : "null")}");
                        
                        if (geometry != null)
                        {
                            var geometryType = geometry.GetType();
                            var geometryMethods = geometryType.GetMethods(System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
                            var getMeshMethod = Array.Find(geometryMethods, m => 
                                m.Name.Equals("GetMesh", StringComparison.OrdinalIgnoreCase) && 
                                m.GetParameters().Length == 0);
                            
                            if (getMeshMethod != null)
                            {
                                var mesh = getMeshMethod.Invoke(geometry, null);
                                System.Diagnostics.Debug.WriteLine($"GetMesh result: {(mesh != null ? "not null" : "null")}");
                                
                                if (mesh != null)
                                {
                                    var meshType = mesh.GetType();
                                    var grid = new Dictionary<string, object>();
                                    var vertices = new List<object>();
                                    var triangles = new List<object>();

                                    // Get vertices
                                    var meshMethods = meshType.GetMethods(System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
                                    var getVertexCountMethod = Array.Find(meshMethods, m => 
                                        m.Name.Equals("GetVertexCount", StringComparison.OrdinalIgnoreCase) && 
                                        m.GetParameters().Length == 0);
                                    var getVertexMethod = Array.Find(meshMethods, m => 
                                        m.Name.Equals("GetVertex", StringComparison.OrdinalIgnoreCase) && 
                                        m.GetParameters().Length == 1 && 
                                        m.GetParameters()[0].ParameterType == typeof(int));
                                        
                                    if (getVertexCountMethod != null && getVertexMethod != null)
                                    {
                                        int vertexCount = (int)getVertexCountMethod.Invoke(mesh, null);
                                        System.Diagnostics.Debug.WriteLine($"Vertex count: {vertexCount}");
                                        
                                        for (int i = 0; i < vertexCount; i++)
                                        {
                                            var vertex = getVertexMethod.Invoke(mesh, new object[] { i });
                                            if (vertex != null)
                                            {
                                                var vertexType = vertex.GetType();
                                                var xProp = vertexType.GetProperty("X");
                                                var yProp = vertexType.GetProperty("Y");
                                                var zProp = vertexType.GetProperty("Z");
                                                if (xProp != null && yProp != null && zProp != null)
                                                {
                                                    double x = (double)xProp.GetValue(vertex);
                                                    double y = (double)yProp.GetValue(vertex);
                                                    double z = (double)zProp.GetValue(vertex);
                                                    vertices.Add(new { x = x, y = y, z = z });
                                                }
                                            }
                                        }
                                    }

                                    // Get triangles
                                    var getTriangleCountMethod = Array.Find(meshMethods, m => 
                                        m.Name.Equals("GetTriangleCount", StringComparison.OrdinalIgnoreCase) && 
                                        m.GetParameters().Length == 0);
                                    var getTriangleMethod = Array.Find(meshMethods, m => 
                                        m.Name.Equals("GetTriangle", StringComparison.OrdinalIgnoreCase) && 
                                        m.GetParameters().Length == 1 && 
                                        m.GetParameters()[0].ParameterType == typeof(int));
                                        
                                    if (getTriangleCountMethod != null && getTriangleMethod != null)
                                    {
                                        int triangleCount = (int)getTriangleCountMethod.Invoke(mesh, null);
                                        System.Diagnostics.Debug.WriteLine($"Triangle count: {triangleCount}");
                                        
                                        for (int i = 0; i < triangleCount; i++)
                                        {
                                            var triangle = getTriangleMethod.Invoke(mesh, new object[] { i });
                                            if (triangle != null)
                                            {
                                                var triangleType = triangle.GetType();
                                                var v0Prop = triangleType.GetProperty("Vertex0");
                                                var v1Prop = triangleType.GetProperty("Vertex1");
                                                var v2Prop = triangleType.GetProperty("Vertex2");
                                                if (v0Prop != null && v1Prop != null && v2Prop != null)
                                                {
                                                    int v0 = (int)v0Prop.GetValue(triangle);
                                                    int v1 = (int)v1Prop.GetValue(triangle);
                                                    int v2 = (int)v2Prop.GetValue(triangle);
                                                    triangles.Add(new int[] { v0, v1, v2 });
                                                }
                                            }
                                        }
                                    }

                                    if (vertices.Count > 0)
                                    {
                                        grid["vertices"] = vertices;
                                        grid["triangles"] = triangles;
                                        meshArray.Add(new Dictionary<string, object> { ["grids"] = new object[] { grid } });
                                        System.Diagnostics.Debug.WriteLine($"Mesh extracted: {vertices.Count} vertices, {triangles.Count} triangles");
                                    }
                                }
                            }
                            else
                            {
                                System.Diagnostics.Debug.WriteLine($"GetMesh method not found in geometry. Available methods: {string.Join(", ", geometryMethods.Select(m => m.Name))}");
                            }
                        }
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine($"GetGeometry method not found. Available methods: {string.Join(", ", allMethods.Select(m => m.Name))}");
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Error extracting mesh: {ex.Message}\n{ex.StackTrace}");
                }

                return meshArray.Count > 0 ? meshArray : null;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error extracting mesh data: {ex.Message}\n{ex.StackTrace}");
                return null;
            }
        }
    }
}

