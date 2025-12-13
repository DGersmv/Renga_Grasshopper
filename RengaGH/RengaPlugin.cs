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
using System.IO;
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
        
        // Log file path
        private static readonly string LogFilePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "Renga",
            "RengaGH_Server.log");

        /// <summary>
        /// Log message to file with timestamp
        /// </summary>
        private static void Log(string message)
        {
            try
            {
                var logDir = Path.GetDirectoryName(LogFilePath);
                if (!Directory.Exists(logDir))
                    Directory.CreateDirectory(logDir);
                
                var logMessage = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] {message}";
                File.AppendAllText(LogFilePath, logMessage + Environment.NewLine);
                System.Diagnostics.Debug.WriteLine(logMessage);
            }
            catch { }
        }

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
                Log($"TCP server started on port {port}");
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
            Log("Waiting for client connection...");
            while (isServerRunning)
            {
                try
                {
                    var client = await tcpListener!.AcceptTcpClientAsync();
                    var clientEndpoint = client.Client.RemoteEndPoint?.ToString() ?? "unknown";
                    Log($"Client connection accepted from {clientEndpoint}");
                    Log($"Starting HandleClientAsync task for {clientEndpoint}");
                    _ = Task.Run(async () => await HandleClientAsync(client));
                    Log("Waiting for client connection...");
                }
                catch (ObjectDisposedException)
                {
                    // Server was stopped
                    break;
                }
                catch (Exception ex)
                {
                    // Log error but continue accepting connections
                    Log($"Error accepting connection: {ex.Message}");
                    Log($"Stack trace: {ex.StackTrace}");
                }
            }
        }

        private async Task HandleClientAsync(TcpClient client)
        {
            var clientEndpoint = client.Client.RemoteEndPoint?.ToString() ?? "unknown";
            try
            {
                Log($"=== New client connected from {clientEndpoint} ===");
                Log($"Client connected: {client.Connected}");
                var stream = client.GetStream();
                stream.ReadTimeout = 10000; // 10 second timeout
                Log($"Stream obtained, setting read timeout to 10000ms");
                Log("Starting to read data from client...");
                Log("Attempting to read data from stream...");
                
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
                    
                    Log($"Read {bytesRead} bytes from client");
                    for (int i = 0; i < bytesRead; i++)
                    {
                        buffer.Add(readBuffer[i]);
                    }
                    Log($"Total bytes read: {buffer.Count}");
                    
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
                    Log($"Received JSON ({buffer.Count} bytes): {json.Substring(0, Math.Min(500, json.Length))}...");
                    
                    var command = ParseAndProcessCommand(json);
                    
                    // Send response back to client
                    var response = CreateResponse(command);
                    var responseJson = JsonConvert.SerializeObject(response);
                    var responseData = Encoding.UTF8.GetBytes(responseJson);
                    
                    Log($"Sending response ({responseData.Length} bytes): {responseJson.Substring(0, Math.Min(500, responseJson.Length))}...");
                    
                    await stream.WriteAsync(responseData, 0, responseData.Length);
                    await stream.FlushAsync();
                    
                    Log("Response sent successfully");
                    
                    // Wait a bit to ensure response is fully sent before closing
                    await Task.Delay(100);
                }
                else
                {
                    Log("No data received from client");
                }
            }
            catch (Exception ex)
            {
                Log($"Error handling client: {ex.Message}");
                Log($"Stack trace: {ex.StackTrace}");
            }
            finally
            {
                try
                {
                    // Give time for response to be fully sent
                    await Task.Delay(50);
                    client.Close();
                    Log($"Client connection closed");
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
                    Log($"Warning: Could not parse selected objects: {ex.Message}");
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

                    // Get baseline using IWallParams -> IWallContour -> GetBaseline() for accurate curve representation
                    Log($"Wall {modelObject.Id}: Starting baseline extraction...");
                    try
                    {
                        // First get IWallParams from IModelObject
                        var wallParams = modelObject as Renga.IWallParams;
                        Log($"Wall {modelObject.Id}: wallParams = {(wallParams != null ? "NOT NULL" : "NULL")}");
                        if (wallParams != null)
                        {
                            // Get IWallContour from IWallParams
                            var wallContour = wallParams.GetContour();
                            Log($"Wall {modelObject.Id}: wallContour = {(wallContour != null ? "NOT NULL" : "NULL")}");
                            if (wallContour != null)
                            {
                                Log($"Wall {modelObject.Id}: Calling GetBaseline()...");
                                var baseline2D = wallContour.GetBaseline();
                                Log($"Wall {modelObject.Id}: GetBaseline() returned {(baseline2D != null ? "NOT NULL" : "NULL")}");
                            if (baseline2D != null)
                            {
                                var baselineData = new Dictionary<string, object>();
                                
                                // Get baseline 2D coordinates first to check if they're already in global coordinates
                                var startPoint2D = baseline2D.GetBeginPoint();
                                var endPoint2D = baseline2D.GetEndPoint();
                                Log($"Wall {modelObject.Id}: baseline2D start=({startPoint2D.X}, {startPoint2D.Y}), end=({endPoint2D.X}, {endPoint2D.Y})");
                                
                                // Get placement to convert 2D to 3D
                                var levelObject = modelObject as Renga.ILevelObject;
                                var placement = levelObject?.GetPlacement();
                                if (placement != null)
                                {
                                    Log($"Wall {modelObject.Id}: placement origin=({placement.Origin.X}, {placement.Origin.Y}, {placement.Origin.Z})");
                                    Log($"Wall {modelObject.Id}: placement axisX=({placement.AxisX.X}, {placement.AxisX.Y}, {placement.AxisX.Z})");
                                    Log($"Wall {modelObject.Id}: placement axisY=({placement.AxisY.X}, {placement.AxisY.Y}, {placement.AxisY.Z})");
                                }
                                else
                                {
                                    Log($"Wall {modelObject.Id}: placement = NULL");
                                }
                                
                                // CRITICAL FIX: baseline2D from IWallContour.GetBaseline() is already in GLOBAL coordinates
                                // Do NOT use CreateCurve3D with placement, as it applies incorrect transformation
                                // Use baseline2D coordinates directly (they are already global)
                                
                                // Get Z coordinate from placement if available, otherwise use 0
                                double zCoord = 0.0;
                                if (placement != null)
                                {
                                    zCoord = placement.Origin.Z;
                                }
                                
                                // Use baseline2D directly as global coordinates (X, Y from baseline2D, Z from placement)
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
                                
                                Log($"Wall {modelObject.Id}: Using baseline2D directly as global coordinates (no placement transformation), Z={zCoord}");
                                Log($"Wall {modelObject.Id}: baseline3D start=({startPoint3D.X}, {startPoint3D.Y}, {startPoint3D.Z}), end=({endPoint3D.X}, {endPoint3D.Y}, {endPoint3D.Z})");
                                
                                // Set start and end points
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
                                
                                // Initialize sampledPoints - will be updated for arcs below
                                var sampledPoints = new List<Dictionary<string, object>>();
                                sampledPoints.Add(new Dictionary<string, object>
                                {
                                    { "x", startPoint3D.X },
                                    { "y", startPoint3D.Y },
                                    { "z", startPoint3D.Z }
                                });
                                
                                // Add intermediate points for line segments (will be overridden for arcs)
                                int samples = 50;
                                for (int i = 1; i < samples; i++)
                                {
                                    double t = (double)i / samples;
                                    sampledPoints.Add(new Dictionary<string, object>
                                    {
                                        { "x", startPoint3D.X + (endPoint3D.X - startPoint3D.X) * t },
                                        { "y", startPoint3D.Y + (endPoint3D.Y - startPoint3D.Y) * t },
                                        { "z", startPoint3D.Z }
                                    });
                                }
                                
                                sampledPoints.Add(new Dictionary<string, object>
                                {
                                    { "x", endPoint3D.X },
                                    { "y", endPoint3D.Y },
                                    { "z", endPoint3D.Z }
                                });
                                
                                baselineData["sampledPoints"] = sampledPoints;

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
                                        var radius = arc.GetRadius();
                                        baselineData["radius"] = radius;
                                        
                                        Log($"Wall {modelObject.Id}: Baseline is Arc, radius={radius}, center=({center2D.X}, {center2D.Y})");
                                        
                                        // Calculate angles
                                        var startVec = new Renga.Vector2D { X = startPoint2D.X - center2D.X, Y = startPoint2D.Y - center2D.Y };
                                        var endVec = new Renga.Vector2D { X = endPoint2D.X - center2D.X, Y = endPoint2D.Y - center2D.Y };
                                        double startAngle = Math.Atan2(startVec.Y, startVec.X);
                                        double endAngle = Math.Atan2(endVec.Y, endVec.X);
                                        
                                        // Normalize angles to ensure proper arc direction
                                        if (endAngle < startAngle)
                                            endAngle += 2 * Math.PI;
                                        
                                        // Convert center to 3D - use baseline2D directly as global coordinates
                                        if (placement != null)
                                        {
                                            var center3D = new Renga.Point3D
                                            {
                                                X = center2D.X,
                                                Y = center2D.Y,
                                                Z = zCoord
                                            };
                                            
                                            baselineData["center"] = new Dictionary<string, object>
                                            {
                                                { "x", center3D.X },
                                                { "y", center3D.Y },
                                                { "z", center3D.Z }
                                            };
                                            
                                            // Recompute sampledPoints for arc using proper arc calculation
                                            sampledPoints = new List<Dictionary<string, object>>();
                                            int arcSamples = 50; // More samples for smooth arc
                                            
                                            for (int i = 0; i <= arcSamples; i++)
                                            {
                                                double t = (double)i / arcSamples;
                                                double angle = startAngle + (endAngle - startAngle) * t;
                                                
                                                // Calculate 2D point on arc
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
                                            
                                            baselineData["sampledPoints"] = sampledPoints;
                                            Log($"Wall {modelObject.Id}: Recalculated {sampledPoints.Count} sampled points for Arc");
                                        }
                                        else
                                        {
                                            baselineData["center2D"] = new Dictionary<string, object>
                                            {
                                                { "x", center2D.X },
                                                { "y", center2D.Y }
                                            };
                                            
                                            // For 2D arc, compute sampled points in 2D
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
                                                
                                                sampledPoints.Add(new Dictionary<string, object>
                                                {
                                                    { "x", point2D.X },
                                                    { "y", point2D.Y },
                                                    { "z", 0.0 }
                                                });
                                            }
                                            
                                            baselineData["sampledPoints"] = sampledPoints;
                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    Log($"Error checking Arc type for wall {modelObject.Id}: {ex.Message}");
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
                                        Log($"Error checking PolyCurve type for wall {modelObject.Id}: {ex.Message}");
                                        baselineData["type"] = "LineSegment";
                                    }
                                }
                                
                                // Default to LineSegment if type not set
                                if (!baselineData.ContainsKey("type"))
                                {
                                    baselineData["type"] = "LineSegment";
                                }

                                Log($"Wall {modelObject.Id}: About to add baseline to wallData. baselineData keys: {string.Join(", ", baselineData.Keys)}");
                                wallData["baseline"] = baselineData;
                                Log($"Wall {modelObject.Id}: ✓ Baseline added to wallData! has startPoint: {baselineData.ContainsKey("startPoint")}, has endPoint: {baselineData.ContainsKey("endPoint")}, has sampledPoints: {baselineData.ContainsKey("sampledPoints")}");
                            }
                            else
                            {
                                Log($"Wall {modelObject.Id}: ERROR - GetBaseline() returned null! This should not happen - every wall must have a baseline!");
                                // Even if baseline2D is null, try to create a minimal baseline from wall placement
                                try
                                {
                                    var levelObject = modelObject as Renga.ILevelObject;
                                    var placement = levelObject?.GetPlacement();
                                    if (placement != null)
                                    {
                                        var origin = placement.Origin;
                                        // Create a minimal baseline (point at origin) - this is a fallback
                                        var baselineData = new Dictionary<string, object>();
                                        baselineData["startPoint"] = new Dictionary<string, object>
                                        {
                                            { "x", origin.X },
                                            { "y", origin.Y },
                                            { "z", origin.Z }
                                        };
                                        baselineData["endPoint"] = new Dictionary<string, object>
                                        {
                                            { "x", origin.X },
                                            { "y", origin.Y },
                                            { "z", origin.Z }
                                        };
                                        baselineData["type"] = "LineSegment";
                                        var sampledPoints = new List<Dictionary<string, object>>();
                                        sampledPoints.Add(new Dictionary<string, object>
                                        {
                                            { "x", origin.X },
                                            { "y", origin.Y },
                                            { "z", origin.Z }
                                        });
                                        baselineData["sampledPoints"] = sampledPoints;
                                        wallData["baseline"] = baselineData;
                                        Log($"Wall {modelObject.Id}: Created fallback baseline from placement origin");
                                    }
                                }
                                catch (Exception fallbackEx)
                                {
                                    Log($"Wall {modelObject.Id}: Failed to create fallback baseline: {fallbackEx.Message}");
                                }
                            }
                            }
                            else
                            {
                                Log($"Wall {modelObject.Id}: ERROR - GetContour() returned null! Cannot get IWallContour from IWallParams.");
                            }
                        }
                        else
                        {
                            Log($"Wall {modelObject.Id}: ERROR - wallParams is null or modelObject is not IWallParams! ObjectType: {modelObject?.ObjectType}");
                        }
                    }
                    catch (Exception ex)
                    {
                        Log($"ERROR: Could not get baseline for wall {modelObject.Id}: {ex.Message}");
                        Log($"  Stack trace: {ex.StackTrace}");
                        Log($"  Exception type: {ex.GetType().FullName}");
                        if (ex.InnerException != null)
                        {
                            Log($"  Inner exception: {ex.InnerException.Message}");
                        }
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
                        Log($"Warning: Could not get contour for wall {modelObject.Id}: {ex.Message}");
                    }

                    // Get all parameters
                    try
                    {
                        var parameterContainer = modelObject as Renga.IParameterContainer;
                        var parametersDict = new Dictionary<string, object>();
                        
                        // Add wall GUID first - use unique wall ID
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
                        Log($"Warning: Could not get parameters for wall {modelObject.Id}: {ex.Message}");
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
                                Log($"Error getting thickness quantity: {ex.Message}");
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
                        Log($"Warning: Could not get thickness for wall {modelObject.Id}: {ex.Message}");
                        wallData["thickness"] = 0.0;
                    }

                    // Get mesh geometry
                    try
                    {
                        if (object3dMap.ContainsKey(modelObject.Id))
                        {
                            Log($"Found object3d in map for wall {modelObject.Id}, type: {object3dMap[modelObject.Id]?.GetType().Name ?? "null"}");
                            var exportedObject3D = object3dMap[modelObject.Id] as Renga.IExportedObject3D;
                            if (exportedObject3D != null)
                            {
                                Log($"Wall {modelObject.Id} has {exportedObject3D.MeshCount} meshes");
                                var meshArray = new List<Dictionary<string, object>>();
                                
                                for (int meshIndex = 0; meshIndex < exportedObject3D.MeshCount; meshIndex++)
                                {
                                    var mesh = exportedObject3D.GetMesh(meshIndex);
                                    if (mesh == null)
                                        continue;
                                    
                                    Log($"Mesh {meshIndex} has {mesh.GridCount} grids");
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
                                        if (vertexCount > 0)
                                        {
                                            var firstVertex = grid.GetVertex(0);
                                            Log($"Wall {modelObject.Id}: Mesh {meshIndex} Grid {gridIndex} first vertex=({firstVertex.X}, {firstVertex.Y}, {firstVertex.Z})");
                                        }
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
                                    Log($"Successfully added {meshArray.Count} meshes for wall {modelObject.Id}");
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Log($"Warning: Could not get mesh for wall {modelObject.Id}: {ex.Message}");
                    }

                    // Log wallData contents before adding to list
                    Log($"Wall {modelObject.Id}: Final wallData keys: {string.Join(", ", wallData.Keys)}");
                    Log($"Wall {modelObject.Id}: Has baseline in wallData: {wallData.ContainsKey("baseline")}");
                    if (wallData.ContainsKey("baseline"))
                    {
                        var baseline = wallData["baseline"] as Dictionary<string, object>;
                        if (baseline != null)
                        {
                            Log($"Wall {modelObject.Id}: Baseline keys: {string.Join(", ", baseline.Keys)}");
                        }
                        else
                        {
                            Log($"Wall {modelObject.Id}: WARNING - baseline value is not Dictionary!");
                        }
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
                Log($"Error getting object geometry: {ex.Message}");
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
                Log($"Error getting column geometry: {ex.Message}");
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

