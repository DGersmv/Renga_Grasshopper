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
    /// Component for creating columns in Renga from Grasshopper points
    /// Prepares commands and sends them to main Renga Connect node
    /// </summary>
    public class RengaCreateColumnsComponent : GH_Component
    {
        private bool lastUpdateValue = false;

        public RengaCreateColumnsComponent()
            : base("Renga Create Columns", "RengaCreateColumns",
                "Create columns in Renga from Grasshopper points",
                "Renga", "Columns")
        {
        }

        public override Guid ComponentGuid => new Guid("7c3bb3ab-6ac6-479c-a563-bb90b14ebdaf");

        public override void CreateAttributes()
        {
            m_attributes = new RengaCreateColumnsComponentAttributes(this);
        }

        protected override void RegisterInputParams(GH_InputParamManager pManager)
        {
            pManager.AddPointParameter("Points", "P", "Points for column placement", GH_ParamAccess.list);
            pManager.AddGenericParameter("RengaConnect", "RC", "Renga Connect component (main node)", GH_ParamAccess.item);
            pManager.AddNumberParameter("Height", "H", "Column height in millimeters (one per point, or single value for all, default: 3000)", GH_ParamAccess.list, 3000.0);
            pManager.AddBooleanParameter("Update", "U", "Trigger update on False->True change", GH_ParamAccess.item, false);
        }

        protected override void RegisterOutputParams(GH_OutputParamManager pManager)
        {
            pManager.AddBooleanParameter("Success", "S", "Success status for each column", GH_ParamAccess.list);
            pManager.AddTextParameter("Message", "M", "Messages for each column", GH_ParamAccess.list);
            pManager.AddTextParameter("ColumnGuids", "CG", "Column GUIDs in Renga", GH_ParamAccess.list);
        }

        protected override void SolveInstance(IGH_DataAccess DA)
        {
            List<Point3d> points = new List<Point3d>();
            object rengaConnectObj = null;
            List<double> heights = new List<double>();
            bool updateValue = false;

            // Get Update input
            DA.GetData(3, ref updateValue);

            // Check for False->True transition (trigger)
            bool shouldUpdate = updateValue && !lastUpdateValue;
            lastUpdateValue = updateValue;

            // Validate inputs
            if (!DA.GetDataList(0, points) || points.Count == 0)
            {
                DA.SetDataList(0, new List<bool>());
                DA.SetDataList(1, new List<string> { "No points provided" });
                DA.SetDataList(2, new List<string>());
                return;
            }

            if (!DA.GetData(1, ref rengaConnectObj))
            {
                DA.SetDataList(0, new List<bool>());
                DA.SetDataList(1, new List<string> { "Renga Connect component not connected" });
                DA.SetDataList(2, new List<string>());
                return;
            }

            // Get height parameter (list)
            if (!DA.GetDataList(2, heights) || heights.Count == 0)
            {
                heights = new List<double> { 3000.0 }; // default
            }

            // Normalize heights list: if shorter than points, use last value for remaining
            while (heights.Count < points.Count)
            {
                if (heights.Count > 0)
                    heights.Add(heights[heights.Count - 1]);
                else
                    heights.Add(3000.0);
            }

            // Validate heights
            for (int i = 0; i < heights.Count; i++)
            {
                if (heights[i] <= 0)
                {
                    AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, $"Height at index {i} must be positive. Using default 3000mm");
                    heights[i] = 3000.0;
                }
            }

            // Validate points
            string validationError;
            if (!ValidateInputs(points, out validationError))
            {
                DA.SetDataList(0, new List<bool>());
                DA.SetDataList(1, new List<string> { validationError });
                DA.SetDataList(2, new List<string>());
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, validationError);
                return;
            }

            // Only process if Update trigger occurred (False->True)
            if (!shouldUpdate)
            {
                DA.SetDataList(0, new List<bool>());
                DA.SetDataList(1, new List<string> { "Set Update to True to send points to Renga" });
                DA.SetDataList(2, new List<string>());
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
                DA.SetDataList(0, new List<bool>());
                DA.SetDataList(1, new List<string> { "Renga Connect is not connected. Connect to Renga first." });
                DA.SetDataList(2, new List<string>());
                return;
            }

            // Prepare command with points, heights and GUIDs
            var commandMessage = CreateColumnsCommand.CreateMessage(points, heights);

            // Send command to server
            ConnectionResponse response = null;
            try
            {
                response = client.Send(commandMessage);
            }
            catch (Exception ex)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, $"Error: {ex.Message}");
                for (int i = 0; i < points.Count; i++)
                {
                    DA.SetDataList(0, new List<bool> { false });
                    DA.SetDataList(1, new List<string> { $"Error: {ex.Message}" });
                    DA.SetDataList(2, new List<string> { "" });
                }
                return;
            }
            
            var successes = new List<bool>();
            var messages = new List<string>();
            var columnGuids = new List<string>();

            if (response == null || !response.Success)
            {
                // No response or error
                var errorMsg = response?.Error ?? "Failed to send data to Renga or no response";
                for (int i = 0; i < points.Count; i++)
                {
                    successes.Add(false);
                    messages.Add(errorMsg);
                    columnGuids.Add("");
                }
            }
            else
            {
                // Parse response
                try
                {
                    var results = response.Data?["results"] as JArray;

                    if (results != null && results.Count == points.Count)
                    {
                        for (int i = 0; i < points.Count; i++)
                        {
                            var result = results[i] as JObject;
                            var success = result?["success"]?.Value<bool>() ?? false;
                            var resultMessage = result?["message"]?.ToString() ?? "Unknown";
                            var columnId = result?["columnId"]?.ToString() ?? "";

                            successes.Add(success);
                            messages.Add(resultMessage);
                            columnGuids.Add(columnId);

                            // Update mapping
                            if (success && !string.IsNullOrEmpty(columnId))
                            {
                                var pointGuid = result?["grasshopperGuid"]?.ToString();
                                if (!string.IsNullOrEmpty(pointGuid))
                                {
                                    CreateColumnsCommand.UpdateMapping(pointGuid, columnId);
                                }
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
                    }
                }
            }

            DA.SetDataList(0, successes);
            DA.SetDataList(1, messages);
            DA.SetDataList(2, columnGuids);
        }

        protected override Bitmap Icon
        {
            get
            {
                // TODO: Add icon
                return null;
            }
        }


        private bool ValidateInputs(List<Point3d> points, out string errorMessage)
        {
            errorMessage = "";

            if (points == null || points.Count == 0)
            {
                errorMessage = "No points provided";
                return false;
            }

            foreach (var point in points)
            {
                if (point == null)
                {
                    errorMessage = "One or more points are null";
                    return false;
                }

                if (double.IsNaN(point.X) || double.IsInfinity(point.X) ||
                    double.IsNaN(point.Y) || double.IsInfinity(point.Y) ||
                    double.IsNaN(point.Z) || double.IsInfinity(point.Z))
                {
                    errorMessage = $"Invalid point coordinates: ({point.X}, {point.Y}, {point.Z})";
                    return false;
                }

                // Check for reasonable coordinate values (within ±1000000)
                if (Math.Abs(point.X) > 1000000 || Math.Abs(point.Y) > 1000000 || Math.Abs(point.Z) > 1000000)
                {
                    errorMessage = $"Point coordinates out of reasonable range: ({point.X}, {point.Y}, {point.Z})";
                    return false;
                }
            }

            return true;
        }

    }
}
