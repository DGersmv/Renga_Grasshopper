using System;
using System.Drawing;
using Grasshopper.Kernel;
using Grasshopper.Kernel.Attributes;
using Grasshopper.Kernel.Parameters;
using GrasshopperRNG.Connection;

namespace GrasshopperRNG.Components
{
    /// <summary>
    /// Component for connecting to Renga TCP server
    /// Main node that manages connection and data transmission
    /// </summary>
    public class RengaConnectComponent : GH_Component
    {
        private RengaConnectionClient client;

        public RengaConnectComponent()
            : base("Renga Connect", "RengaConnect",
                "Main: Connect to Renga and manage data transmission",
                "Renga", "Main")
        {
        }

        public override void CreateAttributes()
        {
            m_attributes = new RengaConnectComponentAttributes(this);
        }

        protected override void RegisterInputParams(GH_InputParamManager pManager)
        {
            pManager.AddIntegerParameter("Port", "P", "TCP server port number (default: 50100)", GH_ParamAccess.item, 50100);
            pManager.AddBooleanParameter("Connect", "C", "Enable/disable connection", GH_ParamAccess.item, false);
        }

        protected override void RegisterOutputParams(GH_OutputParamManager pManager)
        {
            pManager.AddBooleanParameter("Connected", "C", "Connection status", GH_ParamAccess.item);
            pManager.AddTextParameter("Message", "M", "Status message", GH_ParamAccess.item);
            pManager.AddGenericParameter("Client", "Client", "Client object for other components", GH_ParamAccess.item);
        }

        protected override void SolveInstance(IGH_DataAccess DA)
        {
            int port = 50100;
            bool connect = false;

            DA.GetData(0, ref port);
            DA.GetData(1, ref connect);


            if (client == null)
            {
                client = new RengaConnectionClient { Port = port };
            }
            else if (client.Port != port)
            {
                client.Port = port;
            }

            // Check server reachability (new client doesn't maintain persistent connection)
            bool isReachable = client.IsServerReachable();

            // Handle connection status
            if (connect)
            {
                if (isReachable)
                {
                    DA.SetData(0, true);
                    DA.SetData(1, $"Server reachable on port {port}");
                    DA.SetData(2, new RengaGhClientGoo(client));
                }
                else
                {
                    DA.SetData(0, false);
                    DA.SetData(1, $"Server not reachable on port {port}. Make sure Renga plugin is running and server is started.");
                    DA.SetData(2, null);
                }
            }
            else
            {
                DA.SetData(0, false);
                DA.SetData(1, "Not connected");
                DA.SetData(2, null);
            }
        }

        protected override Bitmap Icon
        {
            get
            {
                // TODO: Add icon
                return null;
            }
        }

        public override Guid ComponentGuid => new Guid("6569e153-5300-47c4-a44e-418b4ebed893");
    }
}

