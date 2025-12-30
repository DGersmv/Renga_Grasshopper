using RengaPlugin.Connection;

namespace RengaPlugin.Commands
{
    /// <summary>
    /// Interface for command handlers
    /// </summary>
    public interface ICommandHandler
    {
        /// <summary>
        /// Handle a command and return response
        /// </summary>
        ConnectionResponse Handle(ConnectionMessage message);
    }
}





