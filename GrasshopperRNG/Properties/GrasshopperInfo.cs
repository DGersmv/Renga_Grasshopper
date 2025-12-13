using Grasshopper.Kernel;
using System;
using System.Drawing;

namespace GrasshopperRNG.Properties
{
    /// <summary>
    /// Assembly info for Grasshopper plugin registration
    /// This class is automatically discovered by Grasshopper to register the plugin
    /// </summary>
    public class GrasshopperInfo : GH_AssemblyInfo
    {
        public override string Name => "Renga_Grasshopper";
        public override Bitmap Icon => null; // Icon will be set by category
        public override string Description => "Renga_Grasshopper Integration - Create columns in Renga from Grasshopper points";
        public override Guid Id => new Guid("3b5803f7-4a55-49b6-b251-c0acf25f42b8");
        public override string AuthorName => "Renga Software LLC";
        public override string AuthorContact => "";
        public override string Version => "1.0.0.0";
    }
}

