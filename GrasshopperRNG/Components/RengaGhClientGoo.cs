using System;
using Grasshopper.Kernel;
using Grasshopper.Kernel.Types;
using GrasshopperRNG.Connection;

namespace GrasshopperRNG.Components
{
    /// <summary>
    /// GH_Goo wrapper for RengaConnectionClient to pass between components
    /// </summary>
    public class RengaGhClientGoo : GH_Goo<RengaConnectionClient>
    {
        public RengaGhClientGoo()
        {
        }

        public RengaGhClientGoo(RengaConnectionClient client)
        {
            Value = client;
        }

        public override bool IsValid => Value != null;

        public override string TypeName => "RengaConnectionClient";

        public override string TypeDescription => "Renga TCP Client for Grasshopper";

        public override IGH_Goo Duplicate()
        {
            return new RengaGhClientGoo(Value);
        }

        public override string ToString()
        {
            if (Value == null)
                return "Null RengaConnectionClient";
            
            return Value.IsServerReachable() 
                ? $"RengaConnectionClient (Server reachable on port {Value.Port})" 
                : $"RengaConnectionClient (Server not reachable)";
        }

        public override bool CastFrom(object source)
        {
            if (source is RengaConnectionClient client)
            {
                Value = client;
                return true;
            }
            return false;
        }

        public bool CastTo<T>(out T target)
        {
            if (typeof(T).IsAssignableFrom(typeof(RengaConnectionClient)))
            {
                target = (T)(object)Value;
                return true;
            }
            target = default(T);
            return false;
        }
    }
}

