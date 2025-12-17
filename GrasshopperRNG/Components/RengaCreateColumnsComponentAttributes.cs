using System;
using System.Drawing;
using Grasshopper.GUI.Canvas;
using Grasshopper.Kernel.Attributes;

namespace GrasshopperRNG.Components
{
    /// <summary>
    /// Custom attributes for RengaCreateColumnsComponent
    /// </summary>
    public class RengaCreateColumnsComponentAttributes : GH_ComponentAttributes
    {
        public RengaCreateColumnsComponentAttributes(RengaCreateColumnsComponent owner) : base(owner)
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
