using GrasshopperRNG.Connection;
using Newtonsoft.Json.Linq;

namespace GrasshopperRNG.Commands
{
    /// <summary>
    /// Command for getting walls from Renga
    /// </summary>
    public class GetWallsCommand
    {
        public static ConnectionMessage CreateMessage()
        {
            var data = new JObject();

            return new ConnectionMessage
            {
                Command = "get_walls",
                Data = data
            };
        }
    }
}



