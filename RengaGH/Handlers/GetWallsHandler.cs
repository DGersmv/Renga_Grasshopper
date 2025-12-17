using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
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

        public GetWallsHandler(Renga.IApplication app)
        {
            m_app = app;
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
                                var placement = wall.GetPlacement();
                                var modelObject = obj as Renga.IModelObject;
                                
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
                                
                                // Get baseline curve
                                object baselineData = ExtractBaselineData(obj);
                                
                                // Get mesh geometry
                                object meshData = ExtractMeshData(modelObject);
                                
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
                        System.Diagnostics.Debug.WriteLine($"Error processing wall object: {ex.Message}");
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

        private object ExtractBaselineData(Renga.IModelObject wallObj)
        {
            try
            {
                // Cast to ILevelObject to access wall-specific methods
                var levelObj = wallObj as Renga.ILevelObject;
                if (levelObj == null)
                    return null;

                // Try to get baseline using reflection (since IWall interface may not be directly accessible)
                try
                {
                    var wallType = wallObj.GetType();
                    var getBaselineMethod = wallType.GetMethod("GetBaseline");
                    if (getBaselineMethod != null)
                    {
                        var baseline = getBaselineMethod.Invoke(wallObj, null);
                        if (baseline != null)
                        {
                            return ExtractBaselineFromCurve(baseline);
                        }
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Error getting baseline via reflection: {ex.Message}");
                }

                return null;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error extracting baseline: {ex.Message}");
                return null;
            }
        }

        private object ExtractBaselineFromCurve(dynamic baseline)
        {
            try
            {
                if (baseline == null)
                    return null;

                var baselineObj = new Dictionary<string, object>();
                
                // Get curve type
                try
                {
                    var curveType = baseline.GetCurveType();
                    baselineObj["type"] = curveType.ToString();
                }
                catch
                {
                    baselineObj["type"] = "LineSegment";
                }

                // Get start and end points
                try
                {
                    var startPoint = baseline.GetStartPoint();
                    var endPoint = baseline.GetEndPoint();
                    
                    baselineObj["startPoint"] = new { x = startPoint.X, y = startPoint.Y, z = startPoint.Z };
                    baselineObj["endPoint"] = new { x = endPoint.X, y = endPoint.Y, z = endPoint.Z };
                }
                catch { }

                // Sample points along the curve - THIS IS CRITICAL for accurate representation
                var sampledPoints = new List<object>();
                try
                {
                    int sampleCount = 50; // Sample more points for accuracy
                    for (int i = 0; i <= sampleCount; i++)
                    {
                        double t = (double)i / sampleCount;
                        var point = baseline.GetPointOn(t);
                        sampledPoints.Add(new { x = point.X, y = point.Y, z = point.Z });
                    }
                    baselineObj["sampledPoints"] = sampledPoints;
                }
                catch { }

                // For arcs, get center and radius
                try
                {
                    var curveType = baseline.GetCurveType();
                    if (curveType.ToString().Contains("Arc"))
                    {
                        var arc = baseline as dynamic;
                        if (arc != null)
                        {
                            var center = arc.GetCenter();
                            baselineObj["center"] = new { x = center.X, y = center.Y, z = center.Z };
                            baselineObj["radius"] = arc.GetRadius();
                        }
                    }
                }
                catch { }

                // For polycurves, get segments with exact 3D coordinates
                try
                {
                    var curveType = baseline.GetCurveType();
                    if (curveType.ToString().Contains("PolyCurve") || curveType.ToString().Contains("Polyline"))
                    {
                        var segments = new List<object>();
                        // Try to get segments if available
                        // For now, rely on sampledPoints which are more accurate
                        baselineObj["segments"] = segments;
                    }
                }
                catch { }

                return baselineObj;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error extracting baseline from curve: {ex.Message}");
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
                
                // Try to get geometry using reflection
                try
                {
                    var objType = modelObject.GetType();
                    var getGeometryMethod = objType.GetMethod("GetGeometry");
                    if (getGeometryMethod != null)
                    {
                        var geometry = getGeometryMethod.Invoke(modelObject, null);
                        if (geometry != null)
                        {
                            var geometryType = geometry.GetType();
                            var getMeshMethod = geometryType.GetMethod("GetMesh");
                            if (getMeshMethod != null)
                            {
                                var mesh = getMeshMethod.Invoke(geometry, null);
                                if (mesh != null)
                                {
                                    var meshType = mesh.GetType();
                                    var grid = new Dictionary<string, object>();
                                    var vertices = new List<object>();
                                    var triangles = new List<object>();

                                    // Get vertices
                                    var getVertexCountMethod = meshType.GetMethod("GetVertexCount");
                                    var getVertexMethod = meshType.GetMethod("GetVertex", new Type[] { typeof(int) });
                                    if (getVertexCountMethod != null && getVertexMethod != null)
                                    {
                                        int vertexCount = (int)getVertexCountMethod.Invoke(mesh, null);
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
                                    var getTriangleCountMethod = meshType.GetMethod("GetTriangleCount");
                                    var getTriangleMethod = meshType.GetMethod("GetTriangle", new Type[] { typeof(int) });
                                    if (getTriangleCountMethod != null && getTriangleMethod != null)
                                    {
                                        int triangleCount = (int)getTriangleCountMethod.Invoke(mesh, null);
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
                                    }
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Error extracting mesh: {ex.Message}");
                }

                return meshArray.Count > 0 ? meshArray : null;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error extracting mesh data: {ex.Message}");
                return null;
            }
        }
    }
}

