using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using Grasshopper.GUI.Canvas;
using Grasshopper.Kernel.Attributes;

namespace GrasshopperRNG.Components
{
    /// <summary>
    /// Custom attributes for RengaConnectComponent with Update button
    /// </summary>
    public class RengaConnectComponentAttributes : GH_ComponentAttributes
    {
        public RengaConnectComponentAttributes(RengaConnectComponent owner) : base(owner)
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

