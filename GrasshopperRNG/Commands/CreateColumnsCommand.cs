using System;
using System.Collections.Generic;
using Rhino.Geometry;
using GrasshopperRNG.Connection;
using Newtonsoft.Json.Linq;

namespace GrasshopperRNG.Commands
{
    /// <summary>
    /// Command for creating/updating columns in Renga
    /// </summary>
    public class CreateColumnsCommand
    {
        private static Dictionary<string, string> pointGuidToColumnGuidMap = new Dictionary<string, string>();
        private static Dictionary<Point3d, string> pointToGuidMap = new Dictionary<Point3d, string>();
        private static int guidCounter = 0;

        private class Point3dEqualityComparer : IEqualityComparer<Point3d>
        {
            private const double tolerance = 0.001;

            public bool Equals(Point3d x, Point3d y)
            {
                return Math.Abs(x.X - y.X) < tolerance &&
                       Math.Abs(x.Y - y.Y) < tolerance &&
                       Math.Abs(x.Z - y.Z) < tolerance;
            }

            public int GetHashCode(Point3d obj)
            {
                return obj.X.GetHashCode() ^ obj.Y.GetHashCode() ^ obj.Z.GetHashCode();
            }
        }

        public static ConnectionMessage CreateMessage(List<Point3d> points, List<double> heights)
        {
            var pointData = new JArray();

            for (int i = 0; i < points.Count; i++)
            {
                var point = points[i];
                var height = i < heights.Count ? heights[i] : (heights.Count > 0 ? heights[heights.Count - 1] : 3000.0);
                
                var pointGuid = GetPointGuid(point);
                var rengaColumnGuid = pointGuidToColumnGuidMap.ContainsKey(pointGuid) 
                    ? pointGuidToColumnGuidMap[pointGuid] 
                    : null;

                pointData.Add(new JObject
                {
                    ["x"] = point.X,
                    ["y"] = point.Y,
                    ["z"] = point.Z,
                    ["height"] = height,
                    ["grasshopperGuid"] = pointGuid,
                    ["rengaColumnGuid"] = rengaColumnGuid != null ? JToken.FromObject(rengaColumnGuid) : JValue.CreateNull()
                });
            }

            var data = new JObject
            {
                ["points"] = pointData
            };

            return new ConnectionMessage
            {
                Command = "update_points",
                Data = data
            };
        }

        public static void UpdateMapping(string pointGuid, string columnId)
        {
            if (!string.IsNullOrEmpty(columnId))
            {
                pointGuidToColumnGuidMap[pointGuid] = columnId;
            }
        }

        private static string GetPointGuid(Point3d point)
        {
            const double tolerance = 0.001;
            foreach (var kvp in pointToGuidMap)
            {
                var existingPoint = kvp.Key;
                if (Math.Abs(existingPoint.X - point.X) < tolerance &&
                    Math.Abs(existingPoint.Y - point.Y) < tolerance &&
                    Math.Abs(existingPoint.Z - point.Z) < tolerance)
                {
                    return kvp.Value;
                }
            }

            var newGuid = $"GH_Point_{++guidCounter}_{Guid.NewGuid():N}";
            pointToGuidMap[point] = newGuid;
            return newGuid;
        }
    }
}

