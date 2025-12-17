using System;
using System.Collections.Generic;
using System.Linq;
using Renga;
using RengaPlugin.Commands;
using RengaPlugin.Connection;
using Newtonsoft.Json.Linq;

namespace RengaPlugin.Handlers
{
    /// <summary>
    /// Handler for update_points command (create/update columns)
    /// </summary>
    public class CreateColumnsHandler : ICommandHandler
    {
        private Renga.IApplication m_app;
        private static Dictionary<string, int> guidToColumnIdMap = new Dictionary<string, int>();

        public CreateColumnsHandler(Renga.IApplication app)
        {
            m_app = app;
        }

        public ConnectionResponse Handle(ConnectionMessage message)
        {
            try
            {
                var points = message.Data?["points"] as JArray;
                if (points == null || points.Count == 0)
                {
                    return new ConnectionResponse
                    {
                        Id = message.Id,
                        Success = false,
                        Error = "No points provided"
                    };
                }

                var results = new List<object>();

                foreach (var pointObj in points)
                {
                    var pointResult = ProcessPoint(pointObj as JObject);
                    results.Add(new
                    {
                        success = pointResult.Success,
                        message = pointResult.Message,
                        columnId = pointResult.ColumnId,
                        grasshopperGuid = pointResult.GrasshopperGuid
                    });
                }

                var responseData = new JObject
                {
                    ["results"] = JArray.FromObject(results)
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
                return new ConnectionResponse
                {
                    Id = message.Id,
                    Success = false,
                    Error = $"Error processing points: {ex.Message}"
                };
            }
        }

        private PointResult ProcessPoint(JObject pointObj)
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
                var height = pointObj["height"]?.Value<double>() ?? 3000.0;
                var grasshopperGuid = pointObj["grasshopperGuid"]?.ToString();
                var rengaColumnGuid = pointObj["rengaColumnGuid"]?.ToString();

                if (string.IsNullOrEmpty(grasshopperGuid))
                {
                    return new PointResult { Success = false, Message = "Missing grasshopperGuid" };
                }

                // Check if column already exists
                int columnId = 0;
                bool columnExists = false;

                if (guidToColumnIdMap.ContainsKey(grasshopperGuid))
                {
                    columnId = guidToColumnIdMap[grasshopperGuid];
                    // Verify column actually exists in Renga
                    if (ColumnExistsInRenga(columnId))
                    {
                        columnExists = true;
                    }
                    else
                    {
                        // Column was deleted in Renga, remove from map
                        guidToColumnIdMap.Remove(grasshopperGuid);
                    }
                }
                else if (!string.IsNullOrEmpty(rengaColumnGuid))
                {
                    if (int.TryParse(rengaColumnGuid, out int parsedId))
                    {
                        // Verify column actually exists in Renga
                        if (ColumnExistsInRenga(parsedId))
                        {
                            columnId = parsedId;
                            columnExists = true;
                            var existingKey = guidToColumnIdMap.FirstOrDefault(kvp => kvp.Value == parsedId).Key;
                            if (!string.IsNullOrEmpty(existingKey))
                            {
                                guidToColumnIdMap.Remove(existingKey);
                            }
                            guidToColumnIdMap[grasshopperGuid] = parsedId;
                        }
                    }
                }

                if (columnExists)
                {
                    return UpdateColumn(columnId, x, y, z, height, grasshopperGuid);
                }
                else
                {
                    return CreateColumn(x, y, z, height, grasshopperGuid);
                }
            }
            catch (Exception ex)
            {
                return new PointResult { Success = false, Message = $"Error processing point: {ex.Message}" };
            }
        }

        private PointResult CreateColumn(double x, double y, double z, double height, string grasshopperGuid)
        {
            try
            {
                var model = m_app.Project.Model;
                if (model == null)
                {
                    return new PointResult { Success = false, Message = "No active model" };
                }

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

                // Set placement
                var placement = column.GetPlacement();
                if (placement != null)
                {
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

                // Set column height
                try
                {
                    var modelObject = column as Renga.IModelObject;
                    if (modelObject != null)
                    {
                        var parameters = modelObject.GetParameters();
                        if (parameters != null)
                        {
                            try
                            {
                                var heightParameter = parameters.Get(Renga.ParameterIds.ColumnHeight);
                                if (heightParameter != null)
                                {
                                    heightParameter.SetDoubleValue(height);
                                }
                            }
                            catch { }
                        }
                    }
                }
                catch { }

                op.Apply();

                int columnId = (column as Renga.IModelObject).Id;
                guidToColumnIdMap[grasshopperGuid] = columnId;

                return new PointResult
                {
                    Success = true,
                    Message = "Column created",
                    ColumnId = columnId.ToString(),
                    GrasshopperGuid = grasshopperGuid
                };
            }
            catch (Exception ex)
            {
                return new PointResult { Success = false, Message = $"Error creating column: {ex.Message}" };
            }
        }

        private PointResult UpdateColumn(int columnId, double x, double y, double z, double height, string grasshopperGuid)
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

                // Update placement
                var placement = column.GetPlacement();
                if (placement != null)
                {
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

                // Update height
                try
                {
                    var modelObject = column as Renga.IModelObject;
                    if (modelObject != null)
                    {
                        var parameters = modelObject.GetParameters();
                        if (parameters != null)
                        {
                            try
                            {
                                var heightParameter = parameters.Get(Renga.ParameterIds.ColumnHeight);
                                if (heightParameter != null)
                                {
                                    heightParameter.SetDoubleValue(height);
                                }
                            }
                            catch { }
                        }
                    }
                }
                catch { }
                
                op.Apply();

                return new PointResult
                {
                    Success = true,
                    Message = "Column updated",
                    ColumnId = columnId.ToString(),
                    GrasshopperGuid = grasshopperGuid
                };
            }
            catch (Exception ex)
            {
                return new PointResult { Success = false, Message = $"Error updating column: {ex.Message}" };
            }
        }

        private bool ColumnExistsInRenga(int columnId)
        {
            try
            {
                var model = m_app.Project.Model;
                if (model == null)
                    return false;

                var objects = model.GetObjects();
                var column = objects.GetById(columnId);
                return column != null && column.ObjectType == Renga.ObjectTypes.Column;
            }
            catch
            {
                return false;
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
            catch { }
            return null;
        }
    }

    internal class PointResult
    {
        public bool Success { get; set; }
        public string Message { get; set; } = "";
        public string? ColumnId { get; set; }
        public string? GrasshopperGuid { get; set; }
    }
}

