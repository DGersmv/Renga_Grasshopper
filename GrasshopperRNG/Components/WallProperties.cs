using System;
using System.Collections.Generic;

namespace GrasshopperRNG.Components
{
    /// <summary>
    /// Class to hold wall properties extracted from Renga
    /// </summary>
    public class WallProperties
    {
        public int WallId { get; set; }
        public string WallName { get; set; }
        
        // Многослойная структура
        public bool HasMultilayerStructure { get; set; }
        public List<LayerInfo> Layers { get; set; }  // Если многослойная
        public double TotalThickness { get; set; }  // Общая толщина
        
        // Материалы
        public List<string> Materials { get; set; }
        
        // Линия привязки
        public string AlignmentLinePosition { get; set; }  // "Left", "Center", "Right", etc.
        public double AlignmentOffset { get; set; }
        
        // Уровень
        public int LevelId { get; set; }
        public string LevelName { get; set; }
        public double LevelOffset { get; set; }
        
        // Высота
        public double Height { get; set; }
        
        public WallProperties()
        {
            Layers = new List<LayerInfo>();
            Materials = new List<string>();
        }
    }
    
    /// <summary>
    /// Information about a single layer in a multilayer wall
    /// </summary>
    public class LayerInfo
    {
        public int Index { get; set; }
        public double Thickness { get; set; }
        public string Material { get; set; }
    }
}

