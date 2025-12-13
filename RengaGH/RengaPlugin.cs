/*  Renga_Grasshopper Integration Plugin
 *
 *  This plugin creates a TCP server to receive data from Grasshopper
 *  and creates/updates columns in Renga based on point coordinates.
 *
 *  Copyright Renga Software LLC, 2025. All rights reserved.
 */

#nullable disable
using System;
using Renga;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Threading.Tasks;
using System.Text;
using System.Windows.Forms;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace RengaPlugin
{
    public class RengaPlugin : Renga.IPlugin
    {
        private Renga.IApplication m_app;
        private TcpListener tcpListener;
        private bool isServerRunning = false;
        private int serverPort = 50100; // Default port
        private Dictionary<string, int> guidToColumnIdMap = new Dictionary<string, int>(); // Grasshopper Point GUID -> Renga Column ID

        private List<Renga.ActionEventSource> m_eventSources = new List<Renga.ActionEventSource>();

        public bool Initialize(string pluginFolder)
        {
            m_app = new Renga.Application();
            var ui = m_app.UI;
            var panelExtension = ui.CreateUIPanelExtension();

            // Create menu - only connection management with Grasshopper
            var mainAction = ui.CreateAction();
            mainAction.DisplayName = "Renga";
            mainAction.ToolTip = "Connect to Grasshopper and manage TCP server";

            var mainEvents = new Renga.ActionEventSource(mainAction);
            mainEvents.Triggered += (s, e) =>
            {
                ShowServerSettings();
            };
            m_eventSources.Add(mainEvents);
            panelExtension.AddToolButton(mainAction);

            ui.AddExtensionToPrimaryPanel(panelExtension);

            // Don't start server automatically - user will start it from settings
            // StartTcpServer(serverPort);

            return true;
        }

        private void ShowServerSettings()
        {
            using (var form = new ServerSettingsForm(serverPort, isServerRunning))
            {
                form.StartServerRequested += (s, e) =>
                {
                    StartTcpServer(form.Port);
                    form.UpdateServerStatus(isServerRunning);
                };
                form.StopServerRequested += (s, e) =>
                {
                    StopTcpServer();
                    form.UpdateServerStatus(isServerRunning);
                };
                form.PortChanged += (s, port) =>
                {
                    serverPort = port;
                };

                form.ShowDialog();
            }
        }

        public void Stop()
        {
            StopTcpServer();
            foreach (var eventSource in m_eventSources)
                eventSource.Dispose();
            m_eventSources.Clear();
        }

        private void StartTcpServer(int port)
        {
            try
            {
                if (port < 1024 || port > 65535)
                {
                    m_app.UI.ShowMessageBox(
                        Renga.MessageIcon.MessageIcon_Error,
                        "Renga_Grasshopper Plugin",
                        $"Invalid port number: {port}. Port must be in range 1024-65535.");
                    return;
                }

                serverPort = port;
                tcpListener = new TcpListener(IPAddress.Any, port);
                tcpListener.Start();
                isServerRunning = true;

                // Start accepting connections asynchronously
                _ = Task.Run(async () => await AcceptConnectionsAsync());

                // Server started successfully - message will be shown in settings form
                System.Diagnostics.Debug.WriteLine($"TCP server started on port {port}");
            }
            catch (Exception ex)
            {
                m_app.UI.ShowMessageBox(
                    Renga.MessageIcon.MessageIcon_Error,
                    "Renga_Grasshopper Plugin",
                    $"Failed to start TCP server: {ex.Message}");
            }
        }

        private void StopTcpServer()
        {
            isServerRunning = false;
            tcpListener?.Stop();
        }

        private async Task AcceptConnectionsAsync()
        {
            while (isServerRunning)
            {
                try
                {
                    var client = await tcpListener!.AcceptTcpClientAsync();
                    _ = Task.Run(async () => await HandleClientAsync(client));
                }
                catch (ObjectDisposedException)
                {
                    // Server was stopped
                    break;
                }
                catch (Exception ex)
                {
                    // Log error but continue accepting connections
                    System.Diagnostics.Debug.WriteLine($"Error accepting connection: {ex.Message}");
                }
            }
        }

        private async Task HandleClientAsync(TcpClient client)
        {
            try
            {
                var stream = client.GetStream();
                stream.ReadTimeout = 5000; // 5 second timeout
                
                // Read all data - wait for complete message
                var buffer = new List<byte>();
                var readBuffer = new byte[4096];
                
                // Read until no more data available or timeout
                while (client.Connected)
                {
                    int bytesRead = 0;
                    try
                    {
                        bytesRead = await stream.ReadAsync(readBuffer, 0, readBuffer.Length);
                    }
                    catch (System.IO.IOException)
                    {
                        // Timeout or connection closed
                        break;
                    }
                    
                    if (bytesRead == 0) break;
                    
                    for (int i = 0; i < bytesRead; i++)
                    {
                        buffer.Add(readBuffer[i]);
                    }
                    
                    // Check if we have complete JSON (simple check - no more data available)
                    if (!stream.DataAvailable)
                    {
                        // Small delay to ensure all data is received
                        await Task.Delay(50);
                        if (!stream.DataAvailable) break;
                    }
                }

                if (buffer.Count > 0)
                {
                    var json = Encoding.UTF8.GetString(buffer.ToArray());
                    System.Diagnostics.Debug.WriteLine($"Received JSON ({buffer.Count} bytes): {json.Substring(0, Math.Min(500, json.Length))}...");
                    
                    var command = ParseAndProcessCommand(json);
                    
                    // Send response back to client
                    var response = CreateResponse(command);
                    var responseJson = JsonConvert.SerializeObject(response);
                    var responseData = Encoding.UTF8.GetBytes(responseJson);
                    
                    System.Diagnostics.Debug.WriteLine($"Sending response ({responseData.Length} bytes): {responseJson.Substring(0, Math.Min(500, responseJson.Length))}...");
                    
                    await stream.WriteAsync(responseData, 0, responseData.Length);
                    await stream.FlushAsync();
                    
                    System.Diagnostics.Debug.WriteLine("Response sent successfully");
                    
                    // Wait a bit to ensure response is fully sent before closing
                    await Task.Delay(100);
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine("No data received from client");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error handling client: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"Stack trace: {ex.StackTrace}");
            }
            finally
            {
                try
                {
                    // Give time for response to be fully sent
                    await Task.Delay(50);
                    client.Close();
                }
                catch { }
            }
        }

        private CommandResult ParseAndProcessCommand(string json)
        {
            try
            {
                var jsonObj = JObject.Parse(json);
                var command = jsonObj["command"]?.ToString();

                // Handle get_walls command
                if (command == "get_walls")
                {
                    return GetWallsCommand();
                }

                // Handle update_points command (existing)
                var points = jsonObj["points"] as JArray;

                if (points == null || points.Count == 0)
                {
                    return new CommandResult { Success = false, Message = "No points provided" };
                }

                var results = new List<PointResult>();

                foreach (var pointObj in points)
                {
                    var pointResult = ProcessPoint(pointObj as JObject);
                    results.Add(pointResult);
                }

                return new CommandResult
                {
                    Success = true,
                    Message = $"Processed {results.Count} points",
                    Results = results
                };
            }
            catch (Exception ex)
            {
                return new CommandResult
                {
                    Success = false,
                    Message = $"Error parsing command: {ex.Message}"
                };
            }
        }

        private PointResult ProcessPoint(JObject? pointObj)
        {
            if (pointObj == null)
            {
                return new PointResult { Success = false, Message = "Invalid point data" };
            }

            try
            {
                var x = pointObj["x"]?.Value<double>() ?? 0;
                var y = pointObj["y"]?.Value<double>() ?? 0;
                var z = pointObj["z"]?.Value<double>() ?? 0;
                var grasshopperGuid = pointObj["grasshopperGuid"]?.ToString();
                var rengaColumnGuid = pointObj["rengaColumnGuid"]?.ToString();

                if (string.IsNullOrEmpty(grasshopperGuid))
                {
                    return new PointResult { Success = false, Message = "Missing grasshopperGuid" };
                }

                // Check if column already exists
                int columnId = 0;
                bool columnExists = false;

                // First check by grasshopperGuid in our map
                if (guidToColumnIdMap.ContainsKey(grasshopperGuid))
                {
                    columnId = guidToColumnIdMap[grasshopperGuid];
                    columnExists = true;
                }
                // If not found, check by rengaColumnGuid (if provided)
                else if (!string.IsNullOrEmpty(rengaColumnGuid))
                {
                    if (int.TryParse(rengaColumnGuid, out int parsedId))
                    {
                        // Check if this column ID exists in our map
                        if (guidToColumnIdMap.ContainsValue(parsedId))
                        {
                            columnId = parsedId;
                            columnExists = true;
                            // Update mapping with grasshopperGuid
                            var existingKey = guidToColumnIdMap.FirstOrDefault(kvp => kvp.Value == parsedId).Key;
                            if (!string.IsNullOrEmpty(existingKey))
                            {
                                guidToColumnIdMap.Remove(existingKey);
                            }
                            guidToColumnIdMap[grasshopperGuid] = parsedId;
                        }
                    }
                }
                else
                {
                    columnId = 0;
                }

                if (columnExists)
                {
                    // Update column position
                    return UpdateColumn(columnId, x, y, z, grasshopperGuid);
                }
                else
                {
                    // Create new column
                    return CreateColumn(x, y, z, grasshopperGuid);
                }
            }
            catch (Exception ex)
            {
                return new PointResult { Success = false, Message = $"Error processing point: {ex.Message}" };
            }
        }

        private PointResult CreateColumn(double x, double y, double z, string grasshopperGuid)
        {
            try
            {
                var model = m_app.Project.Model;
                if (model == null)
                {
                    return new PointResult { Success = false, Message = "No active model" };
                }

                // Get active level or first level
                Renga.ILevel? level = GetActiveLevel();
                if (level == null)
                {
                    return new PointResult { Success = false, Message = "No active level found" };
                }

                var args = model.CreateNewEntityArgs();
                args.TypeId = Renga.ObjectTypes.Column;

                var op = m_app.Project.CreateOperationWithUndo(model.Id);
                op.Start();
                var column = model.CreateObject(args) as Renga.ILevelObject;
                
                if (column == null)
                {
                    op.Rollback();
                    return new PointResult { Success = false, Message = m_app.LastError };
                }

                // Set placement - get existing placement and modify origin
                var placement = column.GetPlacement();
                if (placement != null)
                {
                    // Create a copy and move it to the new position
                    var newPlacement = placement.GetCopy();
                    var moveVector = new Renga.Vector3D 
                    { 
                        X = x - placement.Origin.X, 
                        Y = y - placement.Origin.Y, 
                        Z = z - placement.Origin.Z 
                    };
                    newPlacement.Move(moveVector);
                    column.SetPlacement(newPlacement);
                }
                op.Apply();

                // Store mapping - get ID from column object (cast to IModelObject)
                int columnId = (column as Renga.IModelObject).Id;
                
                guidToColumnIdMap[grasshopperGuid] = columnId;

                // Get geometry for the created column
                var geometry = GetColumnGeometry(columnId);

                return new PointResult
                {
                    Success = true,
                    Message = "Column created",
                    ColumnId = columnId.ToString(),
                    GrasshopperGuid = grasshopperGuid,
                    Geometry = geometry
                };
            }
            catch (Exception ex)
            {
                return new PointResult { Success = false, Message = $"Error creating column: {ex.Message}" };
            }
        }

        private PointResult UpdateColumn(int columnId, double x, double y, double z, string grasshopperGuid)
        {
            try
            {
                var model = m_app.Project.Model;
                if (model == null)
                {
                    return new PointResult { Success = false, Message = "No active model" };
                }

                var column = model.GetObjects().GetById(columnId) as ILevelObject;
                if (column == null)
                {
                    return new PointResult { Success = false, Message = "Column not found" };
                }

                var op = m_app.Project.CreateOperationWithUndo(model.Id);
                op.Start();

                // Update placement - get existing placement and move it to new position
                var placement = column.GetPlacement();
                if (placement != null)
                {
                    // Create a copy and move it to the new position
                    var newPlacement = placement.GetCopy();
                    var moveVector = new Renga.Vector3D 
                    { 
                        X = x - placement.Origin.X, 
                        Y = y - placement.Origin.Y, 
                        Z = z - placement.Origin.Z 
                    };
                    newPlacement.Move(moveVector);
                    column.SetPlacement(newPlacement);
                }
                
                op.Apply();

                // Get geometry for the updated column
                var geometry = GetColumnGeometry(columnId);

                return new PointResult
                {
                    Success = true,
                    Message = "Column updated",
                    ColumnId = columnId.ToString(),
                    GrasshopperGuid = grasshopperGuid,
                    Geometry = geometry
                };
            }
            catch (Exception ex)
            {
                return new PointResult { Success = false, Message = $"Error updating column: {ex.Message}" };
            }
        }

        private CommandResult GetWallsCommand()
        {
            try
            {
                var project = m_app.Project;
                if (project == null)
                {
                    return new CommandResult { Success = false, Message = "No active project" };
                }

                var model = project.Model;
                if (model == null)
                {
                    return new CommandResult { Success = false, Message = "No active model" };
                }

                // Get selected objects
                var selection = m_app.Selection;
                if (selection == null)
                {
                    return new CommandResult { Success = false, Message = "Selection not available" };
                }

                var selectedObjectIds = selection.GetSelectedObjects();
                var selectedIds = new List<int>();
                
                // Try to convert selectedObjectIds to list of integers
                try
                {
                    // GetSelectedObjects might return an array or collection
                    if (selectedObjectIds is int[] intArray)
                    {
                        selectedIds.AddRange(intArray);
                    }
                    else if (selectedObjectIds is System.Collections.ICollection collection)
                    {
                        foreach (var item in collection)
                        {
                            if (item is int id)
                            {
                                selectedIds.Add(id);
                            }
                        }
                    }
                    else
                    {
                        // Try to cast directly
                        var ids = selectedObjectIds as System.Collections.IEnumerable;
                        if (ids != null)
                        {
                            foreach (var item in ids)
                            {
                                if (item is int id)
                                {
                                    selectedIds.Add(id);
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Warning: Could not parse selected objects: {ex.Message}");
                }
                
                if (selectedIds.Count == 0)
                {
                    return new CommandResult
                    {
                        Success = true,
                        Message = "Found 0 selected walls",
                        Walls = new List<object>()
                    };
                }

                var objects = model.GetObjects();
                var dataExporter = project.DataExporter;
                if (dataExporter == null)
                {
                    return new CommandResult { Success = false, Message = "DataExporter not available" };
                }

                var object3dCollection = dataExporter.GetObjects3D();
                if (object3dCollection == null)
                {
                    return new CommandResult { Success = false, Message = "Failed to get 3D objects" };
                }

                // Create map of object IDs to 3D objects for faster lookup
                var object3dMap = new Dictionary<int, object>();
                if (object3dCollection != null)
                {
                    for (int i = 0; i < object3dCollection.Count; i++)
                    {
                        var obj3d = object3dCollection.Get(i);
                        if (obj3d != null)
                        {
                            object3dMap[obj3d.ModelObjectId] = obj3d;
                        }
                    }
                }

                var walls = new List<object>();

                // Process selected walls
                foreach (int wallId in selectedIds)
                {
                    var modelObject = objects.GetById(wallId);
                    if (modelObject == null || modelObject.ObjectType != Renga.ObjectTypes.Wall)
                        continue;

                    var wallData = new Dictionary<string, object>
                    {
                        { "id", modelObject.Id },
                        { "name", modelObject.Name ?? "" }
                    };

                    // Get baseline (line of attachment) - convert from 2D to 3D
                    try
                    {
                        var baseline2DObject = modelObject as Renga.IBaseline2DObject;
                        if (baseline2DObject != null)
                        {
                            var baseline2D = baseline2DObject.GetBaseline();
                            if (baseline2D != null)
                            {
                                var baselineData = new Dictionary<string, object>();
                                
                                // Get placement to convert 2D to 3D
                                var levelObject = modelObject as Renga.ILevelObject;
                                var placement = levelObject?.GetPlacement();
                                
                                // Create 3D curve from 2D baseline for all types
                                Renga.ICurve3D baseline3D = null;
                                if (placement != null)
                                {
                                    baseline3D = baseline2D.CreateCurve3D(placement);
                                    if (baseline3D != null)
                                    {
                                        // Get 3D start and end points
                                        var startPoint3D = baseline3D.GetBeginPoint();
                                        var endPoint3D = baseline3D.GetEndPoint();
                                        
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
                                        
                                        // Sample points from 3D curve for better representation
                                        var sampledPoints = new List<Dictionary<string, object>>();
                                        sampledPoints.Add(new Dictionary<string, object>
                                        {
                                            { "x", startPoint3D.X },
                                            { "y", startPoint3D.Y },
                                            { "z", startPoint3D.Z }
                                        });
                                        
                                        // Sample intermediate points (more samples for complex curves)
                                        int samples = 50;
                                        for (int i = 1; i < samples; i++)
                                        {
                                            double t = (double)i / samples;
                                            // Try to get point at parameter t
                                            try
                                            {
                                                // For now, interpolate linearly - will be improved with actual curve evaluation
                                                var interpolatedPoint = new Renga.Point3D
                                                {
                                                    X = startPoint3D.X + (endPoint3D.X - startPoint3D.X) * t,
                                                    Y = startPoint3D.Y + (endPoint3D.Y - startPoint3D.Y) * t,
                                                    Z = startPoint3D.Z + (endPoint3D.Z - startPoint3D.Z) * t
                                                };
                                                sampledPoints.Add(new Dictionary<string, object>
                                                {
                                                    { "x", interpolatedPoint.X },
                                                    { "y", interpolatedPoint.Y },
                                                    { "z", interpolatedPoint.Z }
                                                });
                                            }
                                            catch { }
                                        }
                                        
                                        sampledPoints.Add(new Dictionary<string, object>
                                        {
                                            { "x", endPoint3D.X },
                                            { "y", endPoint3D.Y },
                                            { "z", endPoint3D.Z }
                                        });
                                        
                                        baselineData["sampledPoints"] = sampledPoints;
                                    }
                                }
                                else
                                {
                                    // Fallback: use 2D coordinates if placement not available
                                    var startPoint2D = baseline2D.GetBeginPoint();
                                    var endPoint2D = baseline2D.GetEndPoint();
                                    
                                    baselineData["startPoint"] = new Dictionary<string, object>
                                    {
                                        { "x", startPoint2D.X },
                                        { "y", startPoint2D.Y },
                                        { "z", 0.0 }
                                    };
                                    
                                    baselineData["endPoint"] = new Dictionary<string, object>
                                    {
                                        { "x", endPoint2D.X },
                                        { "y", endPoint2D.Y },
                                        { "z", 0.0 }
                                    };
                                }

                                // Check curve type - try to QueryInterface for COM objects
                                // Try Arc first using QueryInterface pattern
                                Renga.IArc2D arc = null;
                                try
                                {
                                    // Try direct cast first
                                    arc = baseline2D as Renga.IArc2D;
                                    
                                    // If that fails, try QueryInterface (for COM objects)
                                    if (arc == null && baseline2D is System.Runtime.InteropServices.ComTypes.IConnectionPointContainer)
                                    {
                                        try
                                        {
                                            var comObj = baseline2D as System.Runtime.InteropServices.ComTypes.IConnectionPointContainer;
                                            // Try to get IArc2D interface
                                            var guid = typeof(Renga.IArc2D).GUID;
                                            System.IntPtr ppv;
                                            if (System.Runtime.InteropServices.Marshal.QueryInterface(
                                                System.Runtime.InteropServices.Marshal.GetIUnknownForObject(baseline2D),
                                                ref guid, out ppv) == 0)
                                            {
                                                arc = System.Runtime.InteropServices.Marshal.GetObjectForIUnknown(ppv) as Renga.IArc2D;
                                                System.Runtime.InteropServices.Marshal.Release(ppv);
                                            }
                                        }
                                        catch { }
                                    }
                                    
                                    if (arc != null)
                                    {
                                        baselineData["type"] = "Arc";
                                        var center2D = arc.GetCenter();
                                        baselineData["radius"] = arc.GetRadius();
                                        
                                        System.Diagnostics.Debug.WriteLine($"Wall {modelObject.Id}: Baseline is Arc, radius={arc.GetRadius()}, center=({center2D.X}, {center2D.Y})");
                                        
                                        // Convert center to 3D if placement available
                                        if (placement != null)
                                        {
                                            var origin = placement.Origin;
                                            var axisX = placement.AxisX;
                                            var axisY = placement.AxisY;
                                            
                                            var center3DTransformed = new Renga.Point3D
                                            {
                                                X = origin.X + center2D.X * axisX.X + center2D.Y * axisY.X,
                                                Y = origin.Y + center2D.X * axisX.Y + center2D.Y * axisY.Y,
                                                Z = origin.Z + center2D.X * axisX.Z + center2D.Y * axisY.Z
                                            };
                                            
                                            baselineData["center"] = new Dictionary<string, object>
                                            {
                                                { "x", center3DTransformed.X },
                                                { "y", center3DTransformed.Y },
                                                { "z", center3DTransformed.Z }
                                            };
                                        }
                                        else
                                        {
                                            baselineData["center2D"] = new Dictionary<string, object>
                                            {
                                                { "x", center2D.X },
                                                { "y", center2D.Y }
                                            };
                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    System.Diagnostics.Debug.WriteLine($"Error checking Arc type for wall {modelObject.Id}: {ex.Message}");
                                }
                                
                                // Try PolyCurve if not Arc
                                if (arc == null)
                                {
                                    try
                                    {
                                        var polyCurve = baseline2D as Renga.IPolyCurve2D;
                                        if (polyCurve != null)
                                        {
                                            baselineData["type"] = "PolyCurve";
                                            var segments = new List<Dictionary<string, object>>();
                                            
                                            // Convert each segment to 3D separately
                                            for (int i = 0; i < polyCurve.GetSegmentCount(); i++)
                                            {
                                                var segment = polyCurve.GetSegment(i);
                                                var segStart2D = segment.GetBeginPoint();
                                                var segEnd2D = segment.GetEndPoint();
                                                
                                                var segmentData = new Dictionary<string, object>
                                                {
                                                    { "startX", segStart2D.X },
                                                    { "startY", segStart2D.Y },
                                                    { "endX", segEnd2D.X },
                                                    { "endY", segEnd2D.Y }
                                                };
                                                
                                                // Convert segment to 3D if placement available
                                                if (placement != null)
                                                {
                                                    var segment3D = segment.CreateCurve3D(placement);
                                                    if (segment3D != null)
                                                    {
                                                        var segStart3D = segment3D.GetBeginPoint();
                                                        var segEnd3D = segment3D.GetEndPoint();
                                                        
                                                        segmentData["start3DX"] = segStart3D.X;
                                                        segmentData["start3DY"] = segStart3D.Y;
                                                        segmentData["start3DZ"] = segStart3D.Z;
                                                        segmentData["end3DX"] = segEnd3D.X;
                                                        segmentData["end3DY"] = segEnd3D.Y;
                                                        segmentData["end3DZ"] = segEnd3D.Z;
                                                    }
                                                }
                                                
                                                // Check if segment is an Arc
                                                if (segment is Renga.IArc2D arcSegment)
                                                {
                                                    segmentData["type"] = "Arc";
                                                    var center2D = arcSegment.GetCenter();
                                                    segmentData["centerX"] = center2D.X;
                                                    segmentData["centerY"] = center2D.Y;
                                                    segmentData["radius"] = arcSegment.GetRadius();
                                                    
                                                    // Convert center to 3D if placement available
                                                    if (placement != null)
                                                    {
                                                        var origin = placement.Origin;
                                                        var axisX = placement.AxisX;
                                                        var axisY = placement.AxisY;
                                                        
                                                        var center3D = new Renga.Point3D
                                                        {
                                                            X = origin.X + center2D.X * axisX.X + center2D.Y * axisY.X,
                                                            Y = origin.Y + center2D.X * axisX.Y + center2D.Y * axisY.Y,
                                                            Z = origin.Z + center2D.X * axisX.Z + center2D.Y * axisY.Z
                                                        };
                                                        
                                                        segmentData["center3DX"] = center3D.X;
                                                        segmentData["center3DY"] = center3D.Y;
                                                        segmentData["center3DZ"] = center3D.Z;
                                                    }
                                                }
                                                else
                                                {
                                                    segmentData["type"] = "LineSegment";
                                                }
                                                
                                                segments.Add(segmentData);
                                            }
                                            
                                            baselineData["segments"] = segments;
                                        }
                                        else
                                        {
                                            baselineData["type"] = "LineSegment";
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        System.Diagnostics.Debug.WriteLine($"Error checking PolyCurve type for wall {modelObject.Id}: {ex.Message}");
                                        baselineData["type"] = "LineSegment";
                                    }
                                }
                                
                                // Default to LineSegment if type not set
                                if (!baselineData.ContainsKey("type"))
                                {
                                    baselineData["type"] = "LineSegment";
                                }

                                wallData["baseline"] = baselineData;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Warning: Could not get baseline for wall {modelObject.Id}: {ex.Message}");
                    }

                    // Get contour (3D points from wall contour)
                    try
                    {
                        var wallContour = modelObject as Renga.IWallContour;
                        if (wallContour != null)
                        {
                            var contour2D = wallContour.GetContour();
                            if (contour2D != null)
                            {
                                var levelObject = modelObject as Renga.ILevelObject;
                                var placement = levelObject?.GetPlacement();
                                
                                if (placement != null)
                                {
                                    var contour3D = contour2D.CreateCurve3D(placement);
                                    if (contour3D != null)
                                    {
                                        // Sample points from 3D curve
                                        var contourPoints = new List<Dictionary<string, object>>();
                                        var startPoint = contour3D.GetBeginPoint();
                                        var endPoint = contour3D.GetEndPoint();
                                        
                                        contourPoints.Add(new Dictionary<string, object>
                                        {
                                            { "x", startPoint.X },
                                            { "y", startPoint.Y },
                                            { "z", startPoint.Z }
                                        });
                                        
                                        // Sample intermediate points (you can adjust the number of samples)
                                        // Note: ICurve3D may not have GetPointAt, so we'll interpolate between start and end points
                                        int samples = 10;
                                        for (int i = 1; i < samples; i++)
                                        {
                                            // Interpolate between start and end points
                                            double t = (double)i / samples;
                                            var interpolatedPoint = new Renga.Point3D
                                            {
                                                X = startPoint.X + (endPoint.X - startPoint.X) * t,
                                                Y = startPoint.Y + (endPoint.Y - startPoint.Y) * t,
                                                Z = startPoint.Z + (endPoint.Z - startPoint.Z) * t
                                            };
                                            contourPoints.Add(new Dictionary<string, object>
                                            {
                                                { "x", interpolatedPoint.X },
                                                { "y", interpolatedPoint.Y },
                                                { "z", interpolatedPoint.Z }
                                            });
                                        }
                                        
                                        contourPoints.Add(new Dictionary<string, object>
                                        {
                                            { "x", endPoint.X },
                                            { "y", endPoint.Y },
                                            { "z", endPoint.Z }
                                        });
                                        
                                        wallData["contour"] = contourPoints;
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Warning: Could not get contour for wall {modelObject.Id}: {ex.Message}");
                    }

                    // Get all parameters
                    try
                    {
                        var parameterContainer = modelObject as Renga.IParameterContainer;
                        var parametersDict = new Dictionary<string, object>();
                        
                        // Add wall GUID first
                        parametersDict["WallGuid"] = modelObject.Id.ToString();
                        
                        if (parameterContainer != null)
                        {
                            // Get all parameters
                            var parameterIds = parameterContainer.GetIds();
                            if (parameterIds != null)
                            {
                                int paramCount = parameterIds.Count;
                                for (int i = 0; i < paramCount; i++)
                                {
                                    try
                                    {
                                        var paramId = parameterIds.Get(i);
                                        var param = parameterContainer.Get(paramId);
                                        if (param != null)
                                        {
                                            // IParameter doesn't have Name property, use GUID as key
                                            var paramKey = paramId.ToString();
                                            var paramValue = GetParameterValue(param);
                                            parametersDict[paramKey] = paramValue;
                                        }
                                    }
                                    catch
                                    {
                                        // Skip parameters that can't be read
                                    }
                                }
                            }
                        }
                        
                        wallData["parameters"] = parametersDict;
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Warning: Could not get parameters for wall {modelObject.Id}: {ex.Message}");
                        wallData["parameters"] = new Dictionary<string, object>
                        {
                            { "WallGuid", modelObject.Id.ToString() }
                        };
                    }

                    // Get thickness
                    try
                    {
                        double? thickness = null;
                        
                        // Try to get from quantities
                        var quantityContainer = modelObject as Renga.IQuantityContainer;
                        if (quantityContainer != null)
                        {
                            try
                            {
                                var thicknessQuantity = quantityContainer.Get(Renga.QuantityIds.NominalThickness);
                                if (thicknessQuantity != null)
                                {
                                    // Try different ways to get the value
                                    try
                                    {
                                        // Method 1: Try GetValue() method
                                        var getValueMethod = thicknessQuantity.GetType().GetMethod("GetValue");
                                        if (getValueMethod != null)
                                        {
                                            var value = getValueMethod.Invoke(thicknessQuantity, null);
                                            if (value is double d)
                                            {
                                                thickness = d;
                                            }
                                        }
                                    }
                                    catch { }
                                    
                                    // Method 2: Try Value property
                                    if (!thickness.HasValue)
                                    {
                                        try
                                        {
                                            var valueProp = thicknessQuantity.GetType().GetProperty("Value");
                                            if (valueProp != null)
                                            {
                                                var value = valueProp.GetValue(thicknessQuantity);
                                                if (value is double d)
                                                {
                                                    thickness = d;
                                                }
                                                else if (value != null)
                                                {
                                                    // Try to convert
                                                    if (double.TryParse(value.ToString(), out double parsedValue))
                                                    {
                                                        thickness = parsedValue;
                                                    }
                                                }
                                            }
                                        }
                                        catch { }
                                    }
                                    
                                    // Method 3: Try to get from IQuantity interface directly
                                    if (!thickness.HasValue)
                                    {
                                        try
                                        {
                                            // IQuantity might have a direct property access
                                            var quantityType = thicknessQuantity.GetType();
                                            var fields = quantityType.GetFields(System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance);
                                            foreach (var field in fields)
                                            {
                                                if (field.Name.Contains("Value") || field.Name.Contains("Thickness"))
                                                {
                                                    var value = field.GetValue(thicknessQuantity);
                                                    if (value is double d)
                                                    {
                                                        thickness = d;
                                                        break;
                                                    }
                                                }
                                            }
                                        }
                                        catch { }
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                System.Diagnostics.Debug.WriteLine($"Error getting thickness quantity: {ex.Message}");
                            }
                        }
                        
                        // If still no thickness, try to get from parameters
                        if (!thickness.HasValue)
                        {
                            var parameterContainer = modelObject as Renga.IParameterContainer;
                            if (parameterContainer != null)
                            {
                                try
                                {
                                    // Try common thickness parameter GUIDs or names
                                    var paramIds = parameterContainer.GetIds();
                                    if (paramIds != null)
                                    {
                                        int paramCount = paramIds.Count;
                                        for (int i = 0; i < paramCount; i++)
                                        {
                                            try
                                            {
                                                var paramId = paramIds.Get(i);
                                                var param = parameterContainer.Get(paramId);
                                                if (param != null)
                                                {
                                                    // Try to get double value
                                                    try
                                                    {
                                                        var paramValue = param.GetDoubleValue();
                                                        // Check if this might be thickness (heuristic: positive value between 50-1000mm)
                                                        if (paramValue > 0 && paramValue < 10000)
                                                        {
                                                            thickness = paramValue;
                                                            break;
                                                        }
                                                    }
                                                    catch { }
                                                }
                                            }
                                            catch { }
                                        }
                                    }
                                }
                                catch { }
                            }
                        }
                        
                        if (thickness.HasValue)
                        {
                            wallData["thickness"] = thickness.Value;
                        }
                        else
                        {
                            wallData["thickness"] = 0.0;
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Warning: Could not get thickness for wall {modelObject.Id}: {ex.Message}");
                        wallData["thickness"] = 0.0;
                    }

                    // Get mesh geometry
                    try
                    {
                        if (object3dMap.ContainsKey(modelObject.Id))
                        {
                            var exportedObject3D = object3dMap[modelObject.Id] as Renga.IExportedObject3D;
                            if (exportedObject3D != null)
                            {
                                var meshArray = new List<Dictionary<string, object>>();
                                
                                for (int meshIndex = 0; meshIndex < exportedObject3D.MeshCount; meshIndex++)
                                {
                                    var mesh = exportedObject3D.GetMesh(meshIndex);
                                    if (mesh == null)
                                        continue;
                                    
                                    var meshData = new Dictionary<string, object>
                                    {
                                        { "meshType", mesh.GetType().GUID.ToString() },
                                        { "meshTypeGuid", mesh.GetType().GUID.ToString().ToLower() }
                                    };
                                    
                                    var gridsArray = new List<Dictionary<string, object>>();
                                    
                                    for (int gridIndex = 0; gridIndex < mesh.GridCount; gridIndex++)
                                    {
                                        var grid = mesh.GetGrid(gridIndex);
                                        if (grid == null)
                                            continue;
                                        
                                        var gridData = new Dictionary<string, object>
                                        {
                                            { "gridType", grid.GridType },
                                            { "gridTypeName", GetWallGridTypeName(grid.GridType) }
                                        };
                                        
                                        var vertices = new List<Dictionary<string, object>>();
                                        int vertexCount = grid.VertexCount;
                                        for (int vIndex = 0; vIndex < vertexCount; vIndex++)
                                        {
                                            var vertex = grid.GetVertex(vIndex);
                                            vertices.Add(new Dictionary<string, object>
                                            {
                                                { "x", vertex.X },
                                                { "y", vertex.Y },
                                                { "z", vertex.Z }
                                            });
                                        }
                                        gridData["vertices"] = vertices;
                                        
                                        var triangles = new List<int[]>();
                                        int triangleCount = grid.TriangleCount;
                                        for (int tIndex = 0; tIndex < triangleCount; tIndex++)
                                        {
                                            var triangle = grid.GetTriangle(tIndex);
                                            triangles.Add(new int[] { (int)triangle.V0, (int)triangle.V1, (int)triangle.V2 });
                                        }
                                        gridData["triangles"] = triangles;
                                        
                                        gridsArray.Add(gridData);
                                    }
                                    
                                    meshData["grids"] = gridsArray;
                                    meshArray.Add(meshData);
                                }
                                
                                if (meshArray.Count > 0)
                                {
                                    wallData["mesh"] = meshArray;
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Warning: Could not get mesh for wall {modelObject.Id}: {ex.Message}");
                    }

                    walls.Add(wallData);
                }

                return new CommandResult
                {
                    Success = true,
                    Message = $"Found {walls.Count} selected walls",
                    Walls = walls
                };
            }
            catch (Exception ex)
            {
                return new CommandResult
                {
                    Success = false,
                    Message = $"Error getting walls: {ex.Message}"
                };
            }
        }

        private object GetParameterValue(Renga.IParameter param)
        {
            try
            {
                // Try GetDoubleValue first (most common)
                try
                {
                    return param.GetDoubleValue();
                }
                catch { }
                
                // Try GetStringValue
                try
                {
                    return param.GetStringValue() ?? "";
                }
                catch { }
                
                // Fallback: return empty string
                return "";
            }
            catch
            {
                return "";
            }
        }

        private string GetWallGridTypeName(int gridType)
        {
            // Wall grid types from Renga API
            switch (gridType)
            {
                case 0: return "FrontSide";
                case 1: return "BackSide";
                case 2: return "LeftSide";
                case 3: return "RightSide";
                case 4: return "TopSide";
                case 5: return "BottomSide";
                default: return $"Unknown_{gridType}";
            }
        }

        private ColumnGeometry? GetObjectGeometry(int objectId, Guid objectType)
        {
            try
            {
                var project = m_app.Project;
                if (project == null)
                    return null;

                var dataExporter = project.DataExporter;
                if (dataExporter == null)
                    return null;

                var objectCollection = dataExporter.GetObjects3D();
                if (objectCollection == null)
                    return null;

                // Find the object by ID and type
                for (int objectIndex = 0; objectIndex < objectCollection.Count; objectIndex++)
                {
                    var exportedObject = objectCollection.Get(objectIndex);
                    if (exportedObject == null)
                        continue;

                    // Check if this is our object
                    if (exportedObject.ModelObjectId == objectId && 
                        exportedObject.ModelObjectType == objectType)
                    {
                        var geometry = new ColumnGeometry();
                        uint vertexOffset = 0; // Track vertex offset for triangle indices

                        // Iterate through all meshes
                        for (int meshIndex = 0; meshIndex < exportedObject.MeshCount; meshIndex++)
                        {
                            var mesh = exportedObject.GetMesh(meshIndex);
                            if (mesh == null)
                                continue;

                            // Iterate through all grids in the mesh
                            for (int gridIndex = 0; gridIndex < mesh.GridCount; gridIndex++)
                            {
                                var grid = mesh.GetGrid(gridIndex);
                                if (grid == null)
                                    continue;

                                // Get all vertices and add to global list
                                int vertexCount = grid.VertexCount;
                                for (int vertexIndex = 0; vertexIndex < vertexCount; vertexIndex++)
                                {
                                    var vertex = grid.GetVertex(vertexIndex);
                                    geometry.Vertices.Add(new VertexData
                                    {
                                        X = vertex.X,
                                        Y = vertex.Y,
                                        Z = vertex.Z
                                    });
                                }

                                // Get all triangles and adjust indices with vertex offset
                                int triangleCount = grid.TriangleCount;
                                for (int triangleIndex = 0; triangleIndex < triangleCount; triangleIndex++)
                                {
                                    var triangle = grid.GetTriangle(triangleIndex);
                                    geometry.Triangles.Add(new TriangleData
                                    {
                                        V0 = triangle.V0 + vertexOffset,
                                        V1 = triangle.V1 + vertexOffset,
                                        V2 = triangle.V2 + vertexOffset
                                    });
                                }

                                // Update vertex offset for next grid
                                vertexOffset += (uint)vertexCount;
                            }
                        }

                        return geometry;
                    }
                }

                return null;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error getting object geometry: {ex.Message}");
                return null;
            }
        }

        private ColumnGeometry? GetColumnGeometry(int columnId)
        {
            try
            {
                var project = m_app.Project;
                if (project == null)
                    return null;

                var dataExporter = project.DataExporter;
                if (dataExporter == null)
                    return null;

                var objectCollection = dataExporter.GetObjects3D();
                if (objectCollection == null)
                    return null;

                // Find the column object by ID and type
                for (int objectIndex = 0; objectIndex < objectCollection.Count; objectIndex++)
                {
                    var exportedObject = objectCollection.Get(objectIndex);
                    if (exportedObject == null)
                        continue;

                    // Check if this is our column
                    if (exportedObject.ModelObjectId == columnId && 
                        exportedObject.ModelObjectType == Renga.ObjectTypes.Column)
                    {
                        var geometry = new ColumnGeometry();
                        uint vertexOffset = 0; // Track vertex offset for triangle indices

                        // Iterate through all meshes
                        for (int meshIndex = 0; meshIndex < exportedObject.MeshCount; meshIndex++)
                        {
                            var mesh = exportedObject.GetMesh(meshIndex);
                            if (mesh == null)
                                continue;

                            // Iterate through all grids in the mesh
                            for (int gridIndex = 0; gridIndex < mesh.GridCount; gridIndex++)
                            {
                                var grid = mesh.GetGrid(gridIndex);
                                if (grid == null)
                                    continue;

                                // Get all vertices and add to global list
                                int vertexCount = grid.VertexCount;
                                for (int vertexIndex = 0; vertexIndex < vertexCount; vertexIndex++)
                                {
                                    var vertex = grid.GetVertex(vertexIndex);
                                    geometry.Vertices.Add(new VertexData
                                    {
                                        X = vertex.X,
                                        Y = vertex.Y,
                                        Z = vertex.Z
                                    });
                                }

                                // Get all triangles and adjust indices with vertex offset
                                int triangleCount = grid.TriangleCount;
                                for (int triangleIndex = 0; triangleIndex < triangleCount; triangleIndex++)
                                {
                                    var triangle = grid.GetTriangle(triangleIndex);
                                    geometry.Triangles.Add(new TriangleData
                                    {
                                        V0 = triangle.V0 + vertexOffset,
                                        V1 = triangle.V1 + vertexOffset,
                                        V2 = triangle.V2 + vertexOffset
                                    });
                                }

                                // Update vertex offset for next grid
                                vertexOffset += (uint)vertexCount;
                            }
                        }

                        return geometry;
                    }
                }

                return null;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error getting column geometry: {ex.Message}");
                return null;
            }
        }

        private Renga.ILevel? GetActiveLevel()
        {
            try
            {
                var view = m_app.ActiveView;
                if (view?.Type == Renga.ViewType.ViewType_View3D || view?.Type == Renga.ViewType.ViewType_Level)
                {
                    var model = m_app.Project.Model;
                    if (model != null)
                    {
                        var objects = model.GetObjects();
                        int count = objects.Count;
                        for (int i = 0; i < count; i++)
                        {
                            var obj = objects.GetByIndex(i);
                            if (obj.ObjectType == Renga.ObjectTypes.Level)
                            {
                                return obj as Renga.ILevel;
                            }
                        }
                    }
                }
            }
            catch
            {
                // Fall through
            }
            return null;
        }

        private object CreateResponse(CommandResult result)
        {
            // Handle get_walls response
            if (result.Walls != null)
            {
                return new
                {
                    success = result.Success,
                    message = result.Message,
                    walls = result.Walls
                };
            }

            // Handle update_points response
            return new
            {
                success = result.Success,
                message = result.Message,
                results = result.Results?.Select(r => new
                {
                    success = r.Success,
                    message = r.Message,
                    columnId = r.ColumnId,
                    grasshopperGuid = r.GrasshopperGuid,
                    geometry = r.Geometry != null ? new
                    {
                        vertices = r.Geometry.Vertices.Select(v => new { x = v.X, y = v.Y, z = v.Z }),
                        triangles = r.Geometry.Triangles.Select(t => new { v0 = t.V0, v1 = t.V1, v2 = t.V2 })
                    } : null
                })
            };
        }
    }

    // Helper classes
    internal class CommandResult
    {
        public bool Success { get; set; }
        public string Message { get; set; } = "";
        public List<PointResult>? Results { get; set; }
        public List<object>? Walls { get; set; }
    }

    internal class PointResult
    {
        public bool Success { get; set; }
        public string Message { get; set; } = "";
        public string? ColumnId { get; set; }
        public string? GrasshopperGuid { get; set; }
        public ColumnGeometry? Geometry { get; set; }
    }

    internal class ColumnGeometry
    {
        public List<VertexData> Vertices { get; set; } = new List<VertexData>();
        public List<TriangleData> Triangles { get; set; } = new List<TriangleData>();
    }

    internal class VertexData
    {
        public float X { get; set; }
        public float Y { get; set; }
        public float Z { get; set; }
    }

    internal class TriangleData
    {
        public uint V0 { get; set; }
        public uint V1 { get; set; }
        public uint V2 { get; set; }
    }
}

