using System;
using System.Drawing;
using Grasshopper.GUI.Canvas;
using Grasshopper.Kernel.Attributes;

namespace GrasshopperRNG.Components
{
    /// <summary>
    /// Custom attributes for RengaWallPropertiesComponent
    /// </summary>
    public class RengaWallPropertiesComponentAttributes : GH_ComponentAttributes
    {
        public RengaWallPropertiesComponentAttributes(RengaWallPropertiesComponent owner) : base(owner)
        {
        }

        protected override void Layout()
        {
            base.Layout();
        }

        protected override void Render(GH_Canvas canvas, Graphics graphics, GH_CanvasChannel channel)
        {
            // Render base component
            base.Render(canvas, graphics, channel);
        }
    }
}

