using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using Grasshopper.Kernel;
using Grasshopper.Kernel.Attributes;
using Grasshopper.Kernel.Types;
using Grasshopper.Kernel.Data;
using Grasshopper.Kernel.Parameters;
using Rhino.Geometry;
using GrasshopperRNG.Connection;
using GrasshopperRNG.Commands;
using Newtonsoft.Json.Linq;

namespace GrasshopperRNG.Components
{
    /// <summary>
    /// Component for getting walls from Renga and displaying them in Grasshopper
    /// </summary>
    public class RengaGetWallsComponent : GH_Component
    {
        private bool lastUpdateValue = false;

        public RengaGetWallsComponent()
            : base("Get Walls", "GetWalls",
                "Input: Get walls from Renga and their parameters",
                "Renga", "Input")
        {
        }


        protected override void RegisterInputParams(GH_InputParamManager pManager)
        {
            pManager.AddGenericParameter("RengaConnect", "RC", "Renga Connect component (main node)", GH_ParamAccess.item);
            pManager.AddBooleanParameter("Update", "Update", "Trigger update on False->True change", GH_ParamAccess.item);
        }

        public override void CreateAttributes()
        {
            m_attributes = new RengaGetWallsComponentAttributes(this);
            // Set Update parameter as optional for backward compatibility
            if (Params.Input.Count > 1)
            {
                Params.Input[1].Optional = true;
            }
        }

        protected override void BeforeSolveInstance()
        {
            // Reset update state before each solve
            // This helps with backward compatibility
            base.BeforeSolveInstance();
        }

        protected override void RegisterOutputParams(GH_OutputParamManager pManager)
        {
            pManager.AddBooleanParameter("Success", "S", "Success status", GH_ParamAccess.item);
            pManager.AddTextParameter("Message", "M", "Status message", GH_ParamAccess.item);
            pManager.AddIntegerParameter("WallIds", "ID", "Wall IDs", GH_ParamAccess.list);
            pManager.AddTextParameter("WallNames", "N", "Wall names", GH_ParamAccess.list);
            pManager.AddGeometryParameter("Baselines", "BL", "Wall baselines (curves)", GH_ParamAccess.list);
            pManager.AddGeometryParameter("Meshes", "M", "Wall geometry as Mesh", GH_ParamAccess.list);
            pManager.AddNumberParameter("Thickness", "T", "Wall thickness", GH_ParamAccess.list);
        }

        protected override void SolveInstance(IGH_DataAccess DA)
        {
            object rengaConnectObj = null;
            bool updateValue = false;

            if (!DA.GetData(0, ref rengaConnectObj))
            {
                DA.SetData(0, false);
                DA.SetData(1, "Renga Connect component not connected");
                DA.SetDataList(2, new List<int>());
                DA.SetDataList(3, new List<string>());
                DA.SetDataList(4, new List<IGH_GeometricGoo>());
                DA.SetDataList(5, new List<IGH_GeometricGoo>());
                DA.SetDataList(6, new List<double>());
                return;
            }

            // Get Update input (optional for backward compatibility)
            updateValue = false;
            if (Params.Input.Count > 1)
            {
                DA.GetData(1, ref updateValue);
            }

            // Check for False->True transition (trigger)
            bool shouldUpdate = updateValue && !lastUpdateValue;
            lastUpdateValue = updateValue;

            // Only process if Update trigger occurred
            if (!shouldUpdate)
            {
                DA.SetData(0, false);
                DA.SetData(1, "Set Update to True to get walls from Renga");
                DA.SetDataList(2, new List<int>());
                DA.SetDataList(3, new List<string>());
                DA.SetDataList(4, new List<IGH_GeometricGoo>());
                DA.SetDataList(5, new List<IGH_GeometricGoo>());
                DA.SetDataList(6, new List<double>());
                return;
            }

            // Extract RengaConnectionClient from rengaConnectObj
            RengaConnectionClient client = null;
            
            // Try as RengaGhClientGoo
            if (rengaConnectObj is RengaGhClientGoo goo)
            {
                client = goo.Value;
            }
            // Try direct cast
            else if (rengaConnectObj is RengaConnectionClient directClient)
            {
                client = directClient;
            }
            // Try as IGH_Goo and cast
            else if (rengaConnectObj is IGH_Goo ghGoo)
            {
                try
                {
                    var scriptVar = ghGoo.ScriptVariable();
                    if (scriptVar is RengaConnectionClient scriptClient)
                    {
                        client = scriptClient;
                    }
                    else if (ghGoo is RengaGhClientGoo clientGoo)
                    {
                        client = clientGoo.Value;
                    }
                }
                catch
                {
                    // ScriptVariable may throw, ignore
                }
            }

            if (client == null || !client.IsServerReachable())
            {
                DA.SetData(0, false);
                DA.SetData(1, "Renga Connect is not connected. Connect to Renga first.");
                DA.SetDataList(2, new List<int>());
                DA.SetDataList(3, new List<string>());
                DA.SetDataList(4, new List<IGH_GeometricGoo>());
                DA.SetDataList(5, new List<IGH_GeometricGoo>());
                DA.SetDataList(6, new List<double>());
                return;
            }

            // Prepare command to get walls
            var message = GetWallsCommand.CreateMessage();

            // Send command to server
            ConnectionResponse response = null;
            try
            {
                response = client.Send(message);
            }
            catch (Exception ex)
            {
                DA.SetData(0, false);
                DA.SetData(1, $"Error: {ex.Message}");
                DA.SetDataList(2, new List<int>());
                DA.SetDataList(3, new List<string>());
                DA.SetDataList(4, new List<IGH_GeometricGoo>());
                DA.SetDataList(5, new List<IGH_GeometricGoo>());
                DA.SetDataList(6, new List<double>());
                return;
            }
            
            var wallIds = new List<int>();
            var wallNames = new List<string>();
            var baselines = new List<IGH_GeometricGoo>();
            var meshes = new List<IGH_GeometricGoo>();
            var thicknesses = new List<double>();
            bool success = false;
            string messageStr = "";

            if (response == null || !response.Success)
            {
                success = false;
                messageStr = response?.Error ?? "Failed to get response from Renga";
            }
            else
            {
                // Parse response
                try
                {
                    success = response.Success;
                    messageStr = response.Error ?? "Success";
                    var walls = response.Data?["walls"] as JArray;

                    if (success && walls != null)
                    {
                        foreach (var wallObj in walls)
                        {
                            var wall = wallObj as JObject;
                            if (wall == null)
                                continue;

                            // Extract ID
                            var id = wall["id"]?.Value<int>() ?? 0;
                            wallIds.Add(id);

                            // Extract name
                            var name = wall["name"]?.ToString() ?? "";
                            wallNames.Add(name);

                            // Extract baseline
                            IGH_GeometricGoo baselineGeo = null;
                            try
                            {
                                var baselineObj = wall["baseline"] as JObject;
                                if (baselineObj == null)
                                {
                                    System.Diagnostics.Debug.WriteLine($"Warning: baseline is null for wall {id}");
                                    baselines.Add(null);
                                    continue;
                                }
                                
                                var startPoint = baselineObj["startPoint"] as JObject;
                                var endPoint = baselineObj["endPoint"] as JObject;
                                var baselineType = baselineObj["type"]?.ToString() ?? "LineSegment";

                                double startX = 0, startY = 0, startZ = 0;
                                double endX = 0, endY = 0, endZ = 0;
                                
                                if (startPoint != null && endPoint != null)
                                {
                                    startX = startPoint["x"]?.Value<double>() ?? 0;
                                    startY = startPoint["y"]?.Value<double>() ?? 0;
                                    startZ = startPoint["z"]?.Value<double>() ?? 0;
                                    endX = endPoint["x"]?.Value<double>() ?? 0;
                                    endY = endPoint["y"]?.Value<double>() ?? 0;
                                    endZ = endPoint["z"]?.Value<double>() ?? 0;

                                    // Check if we have sampled points (preferred for all curve types)
                                    var sampledPoints = baselineObj["sampledPoints"] as JArray;
                                    if (sampledPoints != null && sampledPoints.Count > 1)
                                    {
                                        // Use sampled points to create a polyline curve
                                        var polyline = new Polyline();
                                        foreach (var pt in sampledPoints)
                                        {
                                            double px = pt["x"]?.Value<double>() ?? 0;
                                            double py = pt["y"]?.Value<double>() ?? 0;
                                            double pz = pt["z"]?.Value<double>() ?? 0;
                                            polyline.Add(px, py, pz);
                                        }
                                        baselineGeo = new GH_Curve(new PolylineCurve(polyline));
                                    }
                                    // If no sampled points, try to reconstruct based on type
                                    else if (baselineType == "Arc")
                                    {
                                        // Fallback to Arc reconstruction if sampled points not available
                                        var center = baselineObj["center"] as JObject;
                                        var radius = baselineObj["radius"]?.Value<double>() ?? 0;
                                        if (center != null)
                                        {
                                            double centerX = center["x"]?.Value<double>() ?? 0;
                                            double centerY = center["y"]?.Value<double>() ?? 0;
                                            double centerZ = center["z"]?.Value<double>() ?? 0;
                                            
                                            var startPt = new Point3d(startX, startY, startZ);
                                            var endPt = new Point3d(endX, endY, endZ);
                                            var centerPt = new Point3d(centerX, centerY, centerZ);
                                            
                                            var plane = new Plane(centerPt, Vector3d.ZAxis);
                                            var startVec = startPt - centerPt;
                                            var endVec = endPt - centerPt;
                                            double startAngle = Math.Atan2(startVec.Y, startVec.X);
                                            double endAngle = Math.Atan2(endVec.Y, endVec.X);
                                            
                                            var arc = new Arc();
                                            arc.Plane = plane;
                                            arc.Radius = radius;
                                            arc.StartAngle = startAngle;
                                            arc.EndAngle = endAngle;
                                            baselineGeo = new GH_Curve(arc.ToNurbsCurve());
                                        }
                                        else
                                        {
                                            // If center is missing, fallback to line
                                            var line = new Line(new Point3d(startX, startY, startZ), new Point3d(endX, endY, endZ));
                                            baselineGeo = new GH_Line(line);
                                        }
                                    }
                                    else if (baselineType == "PolyCurve")
                                    {
                                        var segments = baselineObj["segments"] as JArray;
                                        if (segments != null && segments.Count > 0)
                                        {
                                            var curves = new List<Rhino.Geometry.Curve>();
                                            
                                            // Start with first point
                                            var currentPoint = new Point3d(startX, startY, startZ);
                                            
                                            int segmentIndex = 0;
                                            foreach (var seg in segments)
                                            {
                                                var segType = seg["type"]?.ToString() ?? "LineSegment";
                                                
                                                // Get 3D coordinates if available, otherwise use 2D with Z from start
                                                double segStartX = seg["start3DX"]?.Value<double>() ?? seg["startX"]?.Value<double>() ?? 0;
                                                double segStartY = seg["start3DY"]?.Value<double>() ?? seg["startY"]?.Value<double>() ?? 0;
                                                double segStartZ = seg["start3DZ"]?.Value<double>() ?? startZ;
                                                double segEndX = seg["end3DX"]?.Value<double>() ?? seg["endX"]?.Value<double>() ?? 0;
                                                double segEndY = seg["end3DY"]?.Value<double>() ?? seg["endY"]?.Value<double>() ?? 0;
                                                double segEndZ = seg["end3DZ"]?.Value<double>() ?? startZ;
                                                
                                                var segStartPt = new Point3d(segStartX, segStartY, segStartZ);
                                                var segEndPt = new Point3d(segEndX, segEndY, segEndZ);
                                                
                                                // Use segment start point if current point is not set or different
                                                if (segmentIndex == 0 || currentPoint.DistanceTo(segStartPt) > 0.001)
                                                {
                                                    currentPoint = segStartPt;
                                                }
                                                
                                                segmentIndex++;
                                                
                                                if (segType == "Arc")
                                                {
                                                    // Handle Arc segment
                                                    double center3DX = seg["center3DX"]?.Value<double>() ?? seg["centerX"]?.Value<double>() ?? 0;
                                                    double center3DY = seg["center3DY"]?.Value<double>() ?? seg["centerY"]?.Value<double>() ?? 0;
                                                    double center3DZ = seg["center3DZ"]?.Value<double>() ?? segStartZ;
                                                    double radius = seg["radius"]?.Value<double>() ?? 0;
                                                    
                                                    var centerPt = new Point3d(center3DX, center3DY, center3DZ);
                                                    
                                                    var plane = new Plane(centerPt, Vector3d.ZAxis);
                                                    var startVec = currentPoint - centerPt;
                                                    var endVec = segEndPt - centerPt;
                                                    double startAngle = Math.Atan2(startVec.Y, startVec.X);
                                                    double endAngle = Math.Atan2(endVec.Y, endVec.X);
                                                    
                                                    var arc = new Arc();
                                                    arc.Plane = plane;
                                                    arc.Radius = radius;
                                                    arc.StartAngle = startAngle;
                                                    arc.EndAngle = endAngle;
                                                    
                                                    curves.Add(arc.ToNurbsCurve());
                                                    currentPoint = segEndPt;
                                                }
                                                else
                                                {
                                                    // Handle LineSegment
                                                    var line = new Line(currentPoint, segEndPt);
                                                    curves.Add(new LineCurve(line));
                                                    currentPoint = segEndPt;
                                                }
                                            }
                                            
                                            // Add final point if needed
                                            var finalPoint = new Point3d(endX, endY, endZ);
                                            if (currentPoint.DistanceTo(finalPoint) > 0.001)
                                            {
                                                var line = new Line(currentPoint, finalPoint);
                                                curves.Add(new LineCurve(line));
                                            }
                                            
                                            // Join all curves into one
                                            if (curves.Count > 0)
                                            {
                                                var joinedCurves = Rhino.Geometry.Curve.JoinCurves(curves);
                                                if (joinedCurves != null && joinedCurves.Length > 0)
                                                {
                                                    baselineGeo = new GH_Curve(joinedCurves[0]);
                                                }
                                                else if (curves.Count == 1)
                                                {
                                                    baselineGeo = new GH_Curve(curves[0]);
                                                }
                                            }
                                        }
                                        else
                                        {
                                            var line = new Line(new Point3d(startX, startY, startZ), new Point3d(endX, endY, endZ));
                                            baselineGeo = new GH_Line(line);
                                        }
                                    }
                                    else
                                    {
                                        // Default: LineSegment
                                        var line = new Line(new Point3d(startX, startY, startZ), new Point3d(endX, endY, endZ));
                                        baselineGeo = new GH_Line(line);
                                    }
                                    
                                    // Fallback: if no baseline was created, create a simple line from start/end points
                                    if (baselineGeo == null)
                                    {
                                        var line = new Line(new Point3d(startX, startY, startZ), new Point3d(endX, endY, endZ));
                                        baselineGeo = new GH_Line(line);
                                    }
                                }
                                else
                                {
                                    // If startPoint or endPoint is null, create a null baseline
                                    System.Diagnostics.Debug.WriteLine($"Warning: startPoint or endPoint is null for wall {id}");
                                    baselineGeo = null;
                                }
                            }
                            catch (Exception ex)
                            {
                                System.Diagnostics.Debug.WriteLine($"Error parsing baseline: {ex.Message}");
                            }
                            // Always add something, even if null (Grasshopper will handle it)
                            baselines.Add(baselineGeo);

                            // Extract mesh
                            IGH_GeometricGoo meshGeo = null;
                            try
                            {
                                var meshArray = wall["mesh"] as JArray;
                                if (meshArray != null && meshArray.Count > 0)
                                {
                                    var allVertices = new List<Point3f>();
                                    var allFaces = new List<MeshFace>();

                                    foreach (var meshObj in meshArray)
                                    {
                                        var grids = meshObj["grids"] as JArray;
                                        if (grids != null)
                                        {
                                            int vertexOffset = allVertices.Count;
                                            foreach (var gridObj in grids)
                                            {
                                                var vertices = gridObj["vertices"] as JArray;
                                                var triangles = gridObj["triangles"] as JArray;

                                                if (vertices != null)
                                                {
                                                    foreach (var v in vertices)
                                                    {
                                                        double x = v["x"]?.Value<double>() ?? 0;
                                                        double y = v["y"]?.Value<double>() ?? 0;
                                                        double z = v["z"]?.Value<double>() ?? 0;
                                                        allVertices.Add(new Point3f((float)x, (float)y, (float)z));
                                                    }
                                                }

                                                if (triangles != null)
                                                {
                                                    foreach (var tri in triangles)
                                                    {
                                                        var triArray = tri as JArray;
                                                        if (triArray != null && triArray.Count >= 3)
                                                        {
                                                            int v0 = triArray[0].Value<int>() + vertexOffset;
                                                            int v1 = triArray[1].Value<int>() + vertexOffset;
                                                            int v2 = triArray[2].Value<int>() + vertexOffset;
                                                            allFaces.Add(new MeshFace(v0, v1, v2));
                                                        }
                                                    }
                                                }
                                                
                                                vertexOffset = allVertices.Count;
                                            }
                                        }
                                    }

                                    if (allVertices.Count > 0)
                                    {
                                        var mesh = new Mesh();
                                        mesh.Vertices.AddVertices(allVertices);
                                        mesh.Faces.AddFaces(allFaces);
                                        mesh.Normals.ComputeNormals();
                                        mesh.Compact();
                                        meshGeo = new GH_Mesh(mesh);
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                System.Diagnostics.Debug.WriteLine($"Error parsing mesh: {ex.Message}");
                            }
                            meshes.Add(meshGeo);

                            // Extract thickness
                            var thickness = wall["thickness"]?.Value<double>() ?? 0.0;
                            thicknesses.Add(thickness);

                        }
                    }
                }
                catch (Exception ex)
                {
                    success = false;
                    messageStr = $"Error parsing response: {ex.Message}";
                }
            }

            DA.SetData(0, success);
            DA.SetData(1, messageStr);
            DA.SetDataList(2, wallIds);
            DA.SetDataList(3, wallNames);
            DA.SetDataList(4, baselines);
            DA.SetDataList(5, meshes);
            DA.SetDataList(6, thicknesses);
        }


        protected override Bitmap Icon
        {
            get
            {
                // TODO: Add icon
                return null;
            }
        }

        public override Guid ComponentGuid => new Guid("06660b3d-4ea9-4eaf-b9c0-d884b1667356");
    }
}

