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
using GrasshopperRNG.Client;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace GrasshopperRNG.Components
{
    /// <summary>
    /// Component for creating columns in Renga from Grasshopper points
    /// Prepares commands and sends them to main Renga Connect node
    /// </summary>
    public class RengaCreateColumnsComponent : GH_Component
    {
        private bool lastUpdateValue = false;
        private static Dictionary<string, string> pointGuidToColumnGuidMap = new Dictionary<string, string>();

        public RengaCreateColumnsComponent()
            : base("Create Columns", "CreateColumns",
                "Output: Create columns in Renga from Grasshopper points",
                "Renga", "Output")
        {
        }


        protected override void RegisterInputParams(GH_InputParamManager pManager)
        {
            pManager.AddPointParameter("Points", "P", "Points for column placement", GH_ParamAccess.list);
            pManager.AddGenericParameter("RengaConnect", "RC", "Renga Connect component (main node)", GH_ParamAccess.item);
            pManager.AddBooleanParameter("Update", "Update", "Trigger update on False->True change", GH_ParamAccess.item);
        }

        public override void CreateAttributes()
        {
            m_attributes = new RengaCreateColumnsComponentAttributes(this);
            // Set Update parameter as optional for backward compatibility
            if (Params.Input.Count > 2)
            {
                Params.Input[2].Optional = true;
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
            pManager.AddBooleanParameter("Success", "S", "Success status for each column", GH_ParamAccess.list);
            pManager.AddTextParameter("Message", "M", "Messages for each column", GH_ParamAccess.list);
            pManager.AddTextParameter("ColumnGuids", "CG", "Column GUIDs in Renga", GH_ParamAccess.list);
            pManager.AddGeometryParameter("Mesh", "M", "Column geometry as Mesh", GH_ParamAccess.list);
        }

        protected override void SolveInstance(IGH_DataAccess DA)
        {
            List<Point3d> points = new List<Point3d>();
            object rengaConnectObj = null;
            bool updateValue = false;

            if (!DA.GetDataList(0, points) || points.Count == 0)
            {
                DA.SetDataList(0, new List<bool>());
                DA.SetDataList(1, new List<string> { "No points provided" });
                DA.SetDataList(2, new List<string>());
                DA.SetDataList(3, new List<Mesh>());
                return;
            }

            if (!DA.GetData(1, ref rengaConnectObj))
            {
                DA.SetDataList(0, new List<bool>());
                DA.SetDataList(1, new List<string> { "Renga Connect component not connected" });
                DA.SetDataList(2, new List<string>());
                DA.SetDataList(3, new List<Mesh>());
                return;
            }

            // Get Update input (optional for backward compatibility)
            updateValue = false;
            if (Params.Input.Count > 2)
            {
                DA.GetData(2, ref updateValue);
            }

            // Check for False->True transition (trigger)
            bool shouldUpdate = updateValue && !lastUpdateValue;
            lastUpdateValue = updateValue;

            // Only process if Update trigger occurred
            if (!shouldUpdate)
            {
                DA.SetDataList(0, new List<bool>());
                DA.SetDataList(1, new List<string> { "Set Update to True to send points to Renga" });
                DA.SetDataList(2, new List<string>());
                DA.SetDataList(3, new List<Mesh>());
                return;
            }

            // Extract RengaGhClient from rengaConnectObj
            RengaGhClient client = null;
            
            // Try as RengaGhClientGoo
            if (rengaConnectObj is RengaGhClientGoo goo)
            {
                client = goo.Value;
            }
            // Try direct cast
            else if (rengaConnectObj is RengaGhClient directClient)
            {
                client = directClient;
            }
            // Try as IGH_Goo and cast
            else if (rengaConnectObj is IGH_Goo ghGoo)
            {
                try
                {
                    var scriptVar = ghGoo.ScriptVariable();
                    if (scriptVar is RengaGhClient scriptClient)
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
                DA.SetDataList(0, new List<bool>());
                DA.SetDataList(1, new List<string> { "Renga Connect component not provided. Connect to Renga first." });
                DA.SetDataList(2, new List<string>());
                DA.SetDataList(3, new List<Mesh>());
                return;
            }

            // Prepare command with points and GUIDs
            var command = PrepareCommand(points);
            if (command == null)
            {
                DA.SetDataList(0, new List<bool>());
                DA.SetDataList(1, new List<string> { "Failed to prepare command" });
                DA.SetDataList(2, new List<string>());
                DA.SetDataList(3, new List<Mesh>());
                return;
            }

            // Send command to server (Send() will create a new connection if needed)
            var json = JsonConvert.SerializeObject(command);
            var responseJson = client.Send(json);
            
            var successes = new List<bool>();
            var messages = new List<string>();
            var columnGuids = new List<string>();
            var meshes = new List<Mesh>();

            if (string.IsNullOrEmpty(responseJson))
            {
                // No response - assume failure
                for (int i = 0; i < points.Count; i++)
                {
                    successes.Add(false);
                    messages.Add("Failed to send data to Renga or no response");
                    columnGuids.Add("");
                    meshes.Add(null);
                }
            }
            else
            {
                // Parse response
                try
                {
                    var response = JsonConvert.DeserializeObject<JObject>(responseJson);
                    var results = response?["results"] as JArray;

                    if (results != null && results.Count == points.Count)
                    {
                        for (int i = 0; i < points.Count; i++)
                        {
                            var result = results[i] as JObject;
                            var success = result?["success"]?.Value<bool>() ?? false;
                            var message = result?["message"]?.ToString() ?? "Unknown";
                            var columnId = result?["columnId"]?.ToString() ?? "";
                            var geometry = result?["geometry"] as JObject;

                            successes.Add(success);
                            messages.Add(message);
                            columnGuids.Add(columnId);

                            // Build mesh from geometry if available
                            Mesh mesh = null;
                            if (success && geometry != null)
                            {
                                mesh = BuildMeshFromGeometry(geometry);
                            }
                            meshes.Add(mesh);

                            // Update mapping
                            if (success && !string.IsNullOrEmpty(columnId))
                            {
                                var pointGuid = GetPointGuid(points[i]);
                                pointGuidToColumnGuidMap[pointGuid] = columnId;
                            }
                        }
                    }
                    else
                    {
                        // Response format doesn't match
                        for (int i = 0; i < points.Count; i++)
                        {
                            successes.Add(false);
                            messages.Add("Invalid response format from Renga");
                            columnGuids.Add("");
                            meshes.Add(null);
                        }
                    }
                }
                catch (Exception ex)
                {
                    // Error parsing response
                    for (int i = 0; i < points.Count; i++)
                    {
                        successes.Add(false);
                        messages.Add($"Error parsing response: {ex.Message}");
                        columnGuids.Add("");
                        meshes.Add(null);
                    }
                }
            }

            DA.SetDataList(0, successes);
            DA.SetDataList(1, messages);
            DA.SetDataList(2, columnGuids);
            DA.SetDataList(3, meshes);
        }

        protected override Bitmap Icon
        {
            get
            {
                // TODO: Add icon
                return null;
            }
        }

        private object PrepareCommand(List<Point3d> points)
        {
            var pointData = new List<object>();

            foreach (var point in points)
            {
                var pointGuid = GetPointGuid(point);
                var rengaColumnGuid = pointGuidToColumnGuidMap.ContainsKey(pointGuid) 
                    ? pointGuidToColumnGuidMap[pointGuid] 
                    : null;

                pointData.Add(new
                {
                    x = point.X,
                    y = point.Y,
                    z = point.Z,
                    grasshopperGuid = pointGuid,
                    rengaColumnGuid = rengaColumnGuid
                });
            }

            return new
            {
                command = "update_points",
                points = pointData,
                timestamp = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
            };
        }

        private Mesh BuildMeshFromGeometry(JObject geometry)
        {
            try
            {
                var mesh = new Mesh();
                
                // Parse vertices
                var verticesArray = geometry["vertices"] as JArray;
                if (verticesArray == null || verticesArray.Count == 0)
                    return null;

                // Add all vertices to mesh (coordinates in mm, one-to-one)
                foreach (var vertexObj in verticesArray)
                {
                    var vertex = vertexObj as JObject;
                    if (vertex == null)
                        continue;

                    var x = vertex["x"]?.Value<float>() ?? 0f;
                    var y = vertex["y"]?.Value<float>() ?? 0f;
                    var z = vertex["z"]?.Value<float>() ?? 0f;

                    // Add vertex (coordinates in mm)
                    mesh.Vertices.Add(x, y, z);
                }

                // Parse triangles
                var trianglesArray = geometry["triangles"] as JArray;
                if (trianglesArray == null || trianglesArray.Count == 0)
                    return null;

                // Add all faces (triangles) to mesh
                foreach (var triangleObj in trianglesArray)
                {
                    var triangle = triangleObj as JObject;
                    if (triangle == null)
                        continue;

                    var v0 = triangle["v0"]?.Value<uint>() ?? 0;
                    var v1 = triangle["v1"]?.Value<uint>() ?? 0;
                    var v2 = triangle["v2"]?.Value<uint>() ?? 0;

                    // Add face with three vertex indices
                    mesh.Faces.AddFace((int)v0, (int)v1, (int)v2);
                }

                // Rebuild normals for proper rendering
                mesh.Normals.ComputeNormals();
                mesh.Compact();

                return mesh;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error building mesh from geometry: {ex.Message}");
                return null;
            }
        }

        private string GetPointGuid(Point3d point)
        {
            // Try to get Rhino GUID if available (from GH_Point)
            // For now, generate a stable GUID based on point coordinates
            // Round coordinates to 1mm precision for stability
            var roundedX = Math.Round(point.X, 0);
            var roundedY = Math.Round(point.Y, 0);
            var roundedZ = Math.Round(point.Z, 0);
            
            // Create a stable key
            var key = $"Point_{roundedX}_{roundedY}_{roundedZ}";
            
            // Use MD5 hash to generate a GUID-like string
            using (var md5 = System.Security.Cryptography.MD5.Create())
            {
                var hash = md5.ComputeHash(System.Text.Encoding.UTF8.GetBytes(key));
                return new Guid(hash).ToString();
            }
        }

        public override Guid ComponentGuid => new Guid("18960f2d-3491-4936-8f4d-0b5d432797f6");
    }
}

