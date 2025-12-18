using System;
using System.Collections.Generic;
using System.Drawing;
using Grasshopper.Kernel;
using Grasshopper.Kernel.Attributes;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;
using GrasshopperRNG.Connection;
using GrasshopperRNG.Commands;
using Newtonsoft.Json.Linq;

namespace GrasshopperRNG.Components
{
    /// <summary>
    /// Component for getting walls from Renga
    /// </summary>
    public class RengaGetWallsComponent : GH_Component
    {
        private bool lastUpdateValue = false;

        public RengaGetWallsComponent()
            : base("Renga Get Walls", "RengaGetWalls",
                "Get walls from Renga with baseline curves and mesh geometry",
                "Renga", "Walls")
        {
        }

        public override Guid ComponentGuid => new Guid("8d4e5f6a-7b8c-9d0e-1f2a-3b4c5d6e7f8a");

        public override void CreateAttributes()
        {
            m_attributes = new RengaGetWallsComponentAttributes(this);
        }

        protected override void RegisterInputParams(GH_InputParamManager pManager)
        {
            pManager.AddGenericParameter("RengaConnect", "RC", "Renga Connect component (main node)", GH_ParamAccess.item);
            pManager.AddBooleanParameter("Update", "U", "Trigger update on False->True change", GH_ParamAccess.item, false);
        }

        protected override void RegisterOutputParams(GH_OutputParamManager pManager)
        {
            pManager.AddBooleanParameter("Success", "S", "Success status", GH_ParamAccess.item);
            pManager.AddTextParameter("Message", "M", "Status message", GH_ParamAccess.item);
            pManager.AddCurveParameter("Baselines", "B", "Wall baseline curves", GH_ParamAccess.list);
            pManager.AddMeshParameter("Meshes", "M", "Wall mesh geometry", GH_ParamAccess.list);
        }

        protected override void SolveInstance(IGH_DataAccess DA)
        {
            object rengaConnectObj = null;
            bool updateValue = false;

            // Get Update input
            DA.GetData(1, ref updateValue);

            // Check for False->True transition (trigger)
            bool shouldUpdate = updateValue && !lastUpdateValue;
            lastUpdateValue = updateValue;

            if (!DA.GetData(0, ref rengaConnectObj))
            {
                DA.SetData(0, false);
                DA.SetData(1, "Renga Connect component not provided. Connect to Renga first.");
                DA.SetDataList(2, new List<Curve>());
                DA.SetDataList(3, new List<Mesh>());
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

            if (client == null)
            {
                DA.SetData(0, false);
                DA.SetData(1, "Renga Connect component not provided. Connect to Renga first.");
                DA.SetDataList(2, new List<Curve>());
                DA.SetDataList(3, new List<Mesh>());
                return;
            }

            if (!client.IsServerReachable())
            {
                DA.SetData(0, false);
                DA.SetData(1, "Renga Connect is not connected. Connect to Renga first.");
                DA.SetDataList(2, new List<Curve>());
                DA.SetDataList(3, new List<Mesh>());
                return;
            }

            // Only process if Update trigger occurred (False->True)
            if (!shouldUpdate)
            {
                DA.SetData(0, false);
                DA.SetData(1, "Set Update to True to get walls from Renga");
                DA.SetDataList(2, new List<Curve>());
                DA.SetDataList(3, new List<Mesh>());
                return;
            }

            // Prepare command
            var commandMessage = GetWallsCommand.CreateMessage();

            // Send command to server
            ConnectionResponse response = null;
            try
            {
                response = client.Send(commandMessage);
            }
            catch (Exception ex)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, $"Error: {ex.Message}");
                DA.SetData(0, false);
                DA.SetData(1, $"Error: {ex.Message}");
                DA.SetDataList(2, new List<Curve>());
                DA.SetDataList(3, new List<Mesh>());
                return;
            }

            if (response == null || !response.Success)
            {
                var errorMsg = response?.Error ?? "Failed to get walls from Renga or no response";
                DA.SetData(0, false);
                DA.SetData(1, errorMsg);
                DA.SetDataList(2, new List<Curve>());
                DA.SetDataList(3, new List<Mesh>());
                return;
            }

            // Parse response
            try
            {
                var walls = response.Data?["walls"] as JArray;
                if (walls == null)
                {
                    DA.SetData(0, false);
                    DA.SetData(1, "No walls found in response");
                    DA.SetDataList(2, new List<Curve>());
                    DA.SetDataList(3, new List<Mesh>());
                    return;
                }

                var baselines = new List<Curve>();
                var meshes = new List<Mesh>();

                foreach (var wallToken in walls)
                {
                    var wall = wallToken as JObject;
                    if (wall == null)
                        continue;

                    var id = wall["id"]?.Value<int>() ?? 0;
                    
                    // Process baseline
                    var baselineObj = wall["baseline"] as JObject;
                    Curve baselineGeo = null;
                    if (baselineObj == null)
                    {
                        System.Diagnostics.Debug.WriteLine($"Warning: baseline is null for wall {id}");
                        baselineGeo = null;
                    }
                    else
                    {
                        baselineGeo = ParseBaseline(baselineObj);
                    }

                    if (baselineGeo != null)
                    {
                        baselines.Add(baselineGeo);
                    }

                    // Process mesh
                    var meshArray = wall["mesh"] as JArray;
                    if (meshArray != null)
                    {
                        foreach (var meshToken in meshArray)
                        {
                            var meshObj = meshToken as JObject;
                            if (meshObj == null)
                                continue;

                            var grids = meshObj["grids"] as JArray;
                            if (grids != null)
                            {
                                foreach (var gridToken in grids)
                                {
                                    var grid = gridToken as JObject;
                                    if (grid == null)
                                        continue;

                                    var mesh = ParseMesh(grid);
                                    if (mesh != null)
                                    {
                                        meshes.Add(mesh);
                                    }
                                }
                            }
                        }
                    }
                }

                DA.SetData(0, true);
                DA.SetData(1, $"Found {walls.Count} walls");
                DA.SetDataList(2, baselines);
                DA.SetDataList(3, meshes);
            }
            catch (Exception ex)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, $"Error parsing response: {ex.Message}");
                DA.SetData(0, false);
                DA.SetData(1, $"Error parsing response: {ex.Message}");
                DA.SetDataList(2, new List<Curve>());
                DA.SetDataList(3, new List<Mesh>());
            }
        }

        private Curve ParseBaseline(JObject baselineObj)
        {
            try
            {
                var type = baselineObj["type"]?.Value<string>() ?? "";
                var startPointObj = baselineObj["startPoint"] as JObject;
                var endPointObj = baselineObj["endPoint"] as JObject;

                if (startPointObj == null || endPointObj == null)
                    return null;

                var startX = startPointObj["x"]?.Value<double>() ?? 0;
                var startY = startPointObj["y"]?.Value<double>() ?? 0;
                var startZ = startPointObj["z"]?.Value<double>() ?? 0;
                var endX = endPointObj["x"]?.Value<double>() ?? 0;
                var endY = endPointObj["y"]?.Value<double>() ?? 0;
                var endZ = endPointObj["z"]?.Value<double>() ?? 0;

                var startPt = new Point3d(startX, startY, startZ);
                var endPt = new Point3d(endX, endY, endZ);

                // Handle different curve types
                // For arcs, use sampled points to create accurate curve
                if (type.Contains("Arc") || type.Contains("arc"))
                {
                    // Use sampled points if available - they are already calculated correctly
                    var sampledPoints = baselineObj["sampledPoints"] as JArray;
                    if (sampledPoints != null && sampledPoints.Count >= 2)
                    {
                        var points = new List<Point3d>();
                        foreach (var ptToken in sampledPoints)
                        {
                            var ptObj = ptToken as JObject;
                            if (ptObj != null)
                            {
                                var x = ptObj["x"]?.Value<double>() ?? 0;
                                var y = ptObj["y"]?.Value<double>() ?? 0;
                                var z = ptObj["z"]?.Value<double>() ?? 0;
                                points.Add(new Point3d(x, y, z));
                            }
                        }
                        
                        // CRITICAL: For arcs, ALWAYS use PolylineCurve from ALL sampled points
                        // This preserves the exact wall baseline direction and ensures it matches the mesh
                        // Do NOT try to create Arc from three points - this may choose wrong direction
                        if (points.Count >= 2)
                        {
                            return new PolylineCurve(points);
                        }
                    }
                    
                    // Fallback: simple line if no sampled points
                    return new Line(startPt, endPt).ToNurbsCurve();
                }

                // Use sampled points if available
                var sampledPoints2 = baselineObj["sampledPoints"] as JArray;
                if (sampledPoints2 != null && sampledPoints2.Count >= 2)
                {
                    var points = new List<Point3d>();
                    foreach (var ptToken in sampledPoints2)
                    {
                        var ptObj = ptToken as JObject;
                        if (ptObj != null)
                        {
                            var x = ptObj["x"]?.Value<double>() ?? 0;
                            var y = ptObj["y"]?.Value<double>() ?? 0;
                            var z = ptObj["z"]?.Value<double>() ?? 0;
                            points.Add(new Point3d(x, y, z));
                        }
                    }
                    if (points.Count >= 2)
                    {
                        if (points.Count == 2)
                        {
                            return new Line(points[0], points[1]).ToNurbsCurve();
                        }
                        else
                        {
                            return new PolylineCurve(points);
                        }
                    }
                }

                // Fallback: simple line
                return new Line(startPt, endPt).ToNurbsCurve();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error parsing baseline: {ex.Message}");
                return null;
            }
        }

        private Mesh ParseMesh(JObject gridObj)
        {
            try
            {
                var vertices = gridObj["vertices"] as JArray;
                var triangles = gridObj["triangles"] as JArray;

                if (vertices == null || triangles == null)
                    return null;

                var mesh = new Mesh();

                // Add vertices
                foreach (var vToken in vertices)
                {
                    var vObj = vToken as JObject;
                    if (vObj != null)
                    {
                        var x = vObj["x"]?.Value<double>() ?? 0;
                        var y = vObj["y"]?.Value<double>() ?? 0;
                        var z = vObj["z"]?.Value<double>() ?? 0;
                        mesh.Vertices.Add(x, y, z);
                    }
                }

                // Add faces
                foreach (var tToken in triangles)
                {
                    var tArray = tToken as JArray;
                    if (tArray != null && tArray.Count >= 3)
                    {
                        var i0 = tArray[0]?.Value<int>() ?? 0;
                        var i1 = tArray[1]?.Value<int>() ?? 0;
                        var i2 = tArray[2]?.Value<int>() ?? 0;
                        mesh.Faces.AddFace(i0, i1, i2);
                    }
                }

                if (mesh.Vertices.Count > 0 && mesh.Faces.Count > 0)
                {
                    mesh.Normals.ComputeNormals();
                    return mesh;
                }

                return null;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error parsing mesh: {ex.Message}");
                return null;
            }
        }

        protected override Bitmap Icon
        {
            get
            {
                // TODO: Add icon
                return null;
            }
        }
    }
}
