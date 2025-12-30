using System;
using System.Collections.Generic;
using System.Drawing;
using Grasshopper.Kernel;
using Grasshopper.Kernel.Data;
using Grasshopper.Kernel.Types;
using Grasshopper.Kernel.Parameters;

namespace GrasshopperRNG.Components
{
    /// <summary>
    /// Component for displaying wall properties from Renga
    /// </summary>
    public class RengaWallPropertiesComponent : GH_Component
    {
        public RengaWallPropertiesComponent()
            : base("Renga Wall Properties", "RengaWallProps",
                "Display wall properties from Renga Get Walls component",
                "Renga", "Walls")
        {
        }

        public override Guid ComponentGuid => new Guid("f9e8d7c6-b5a4-3210-9876-543210fedcba");

        public override void CreateAttributes()
        {
            m_attributes = new RengaWallPropertiesComponentAttributes(this);
        }

        protected override void RegisterInputParams(GH_InputParamManager pManager)
        {
            pManager.AddGenericParameter("Properties", "P", "Wall properties from Renga Get Walls", GH_ParamAccess.list);
        }

        protected override void RegisterOutputParams(GH_OutputParamManager pManager)
        {
            pManager.AddIntegerParameter("Wall IDs", "ID", "Wall IDs", GH_ParamAccess.list);
            pManager.AddTextParameter("Wall Names", "Name", "Wall names", GH_ParamAccess.list);
            pManager.AddBooleanParameter("Has Multilayer", "ML", "Has multilayer structure", GH_ParamAccess.list);
            pManager.AddIntegerParameter("Layer Count", "LC", "Number of layers (0 if no multilayer)", GH_ParamAccess.list);
            pManager.AddNumberParameter("Layer Thicknesses", "LT", "Layer thicknesses (list of lists)", GH_ParamAccess.tree);
            pManager.AddTextParameter("Layer Materials", "LM", "Layer materials (list of lists)", GH_ParamAccess.tree);
            pManager.AddNumberParameter("Total Thickness", "TT", "Total wall thickness", GH_ParamAccess.list);
            pManager.AddTextParameter("Materials", "M", "All materials in wall", GH_ParamAccess.tree);
            pManager.AddTextParameter("Alignment Position", "AP", "Alignment line position (Left/Center/Right)", GH_ParamAccess.list);
            pManager.AddNumberParameter("Alignment Offset", "AO", "Offset from alignment line", GH_ParamAccess.list);
            pManager.AddIntegerParameter("Level ID", "LID", "Level (floor) ID", GH_ParamAccess.list);
            pManager.AddTextParameter("Level Name", "LN", "Level (floor) name", GH_ParamAccess.list);
            pManager.AddNumberParameter("Level Offset", "LO", "Offset from level", GH_ParamAccess.list);
            pManager.AddNumberParameter("Height", "H", "Wall height", GH_ParamAccess.list);
        }

        protected override void SolveInstance(IGH_DataAccess DA)
        {
            var propertiesList = new List<WallProperties>();
            
            if (!DA.GetDataList(0, propertiesList))
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, "No properties provided");
                return;
            }

            if (propertiesList.Count == 0)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, "Properties list is empty");
                return;
            }

            var wallIds = new List<int>();
            var wallNames = new List<string>();
            var hasMultilayer = new List<bool>();
            var layerCounts = new List<int>();
            var layerThicknessesTree = new GH_Structure<GH_Number>();
            var layerMaterialsTree = new GH_Structure<GH_String>();
            var totalThicknesses = new List<double>();
            var materialsTree = new GH_Structure<GH_String>();
            var alignmentPositions = new List<string>();
            var alignmentOffsets = new List<double>();
            var levelIds = new List<int>();
            var levelNames = new List<string>();
            var levelOffsets = new List<double>();
            var heights = new List<double>();

            for (int i = 0; i < propertiesList.Count; i++)
            {
                var props = propertiesList[i];
                
                wallIds.Add(props.WallId);
                wallNames.Add(props.WallName ?? "Unknown");
                hasMultilayer.Add(props.HasMultilayerStructure);
                
                int layerCount = props.Layers != null ? props.Layers.Count : 0;
                layerCounts.Add(layerCount);
                
                // Layer thicknesses and materials (as tree structure)
                var thicknessPath = new GH_Path(i);
                var materialPath = new GH_Path(i);
                
                if (props.Layers != null && props.Layers.Count > 0)
                {
                    foreach (var layer in props.Layers)
                    {
                        layerThicknessesTree.Append(new GH_Number(layer.Thickness), thicknessPath);
                        layerMaterialsTree.Append(new GH_String(layer.Material ?? "Unknown"), materialPath);
                    }
                }
                
                totalThicknesses.Add(props.TotalThickness);
                
                // Materials (as tree - one branch per wall)
                var materialsPath = new GH_Path(i);
                if (props.Materials != null)
                {
                    foreach (var material in props.Materials)
                    {
                        materialsTree.Append(new GH_String(material), materialsPath);
                    }
                }
                
                alignmentPositions.Add(props.AlignmentLinePosition ?? "Unknown");
                alignmentOffsets.Add(props.AlignmentOffset);
                levelIds.Add(props.LevelId);
                levelNames.Add(props.LevelName ?? "Unknown");
                levelOffsets.Add(props.LevelOffset);
                heights.Add(props.Height);
            }

            DA.SetDataList(0, wallIds);
            DA.SetDataList(1, wallNames);
            DA.SetDataList(2, hasMultilayer);
            DA.SetDataList(3, layerCounts);
            DA.SetDataTree(4, layerThicknessesTree);
            DA.SetDataTree(5, layerMaterialsTree);
            DA.SetDataList(6, totalThicknesses);
            DA.SetDataTree(7, materialsTree);
            DA.SetDataList(8, alignmentPositions);
            DA.SetDataList(9, alignmentOffsets);
            DA.SetDataList(10, levelIds);
            DA.SetDataList(11, levelNames);
            DA.SetDataList(12, levelOffsets);
            DA.SetDataList(13, heights);
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

