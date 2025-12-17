using System;
using System.Collections.Generic;
using Renga;
using RengaPlugin.Connection;
using RengaPlugin.Handlers;

namespace RengaPlugin.Commands
{
    /// <summary>
    /// Routes commands to appropriate handlers
    /// </summary>
    public class CommandRouter
    {
        private Dictionary<string, ICommandHandler> handlers;
        private Renga.IApplication m_app;

        public CommandRouter(Renga.IApplication app)
        {
            m_app = app;
            handlers = new Dictionary<string, ICommandHandler>
            {
                { "get_walls", new GetWallsHandler(app) },
                { "update_points", new CreateColumnsHandler(app) }
            };
        }

        /// <summary>
        /// Route a message to appropriate handler
        /// </summary>
        public ConnectionResponse Route(ConnectionMessage message)
        {
            if (message == null || string.IsNullOrEmpty(message.Command))
            {
                return new ConnectionResponse
                {
                    Id = message?.Id ?? "",
                    Success = false,
                    Error = "Invalid message: missing command"
                };
            }

            if (handlers.TryGetValue(message.Command, out var handler))
            {
                try
                {
                    return handler.Handle(message);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Error handling command {message.Command}: {ex.Message}\n{ex.StackTrace}");
                    return new ConnectionResponse
                    {
                        Id = message.Id,
                        Success = false,
                        Error = $"Error handling command: {ex.Message}"
                    };
                }
            }
            else
            {
                return new ConnectionResponse
                {
                    Id = message.Id,
                    Success = false,
                    Error = $"Unknown command: {message.Command}"
                };
            }
        }
    }
}

