/*  Renga_Grasshopper Integration Plugin
 *
 *  This plugin creates a TCP server to receive data from Grasshopper
 *  and creates/updates columns in Renga based on point coordinates.
 *
 *  Copyright Renga Software LLC, 2025. All rights reserved.
 */

#nullable disable
using System;
using Renga;
using System.Collections.Generic;
using System.Net;
using System.Net.Sockets;
using System.Threading.Tasks;
using System.Windows.Forms;
using RengaPlugin.Connection;
using RengaPlugin.Commands;

namespace RengaPlugin
{
    public partial class RengaPlugin : Renga.IPlugin
    {
        private Renga.IApplication m_app;
        private TcpListener tcpListener;
        private bool isServerRunning = false;
        private int serverPort = 50100; // Default port
        private CommandRouter commandRouter;

        private List<Renga.ActionEventSource> m_eventSources = new List<Renga.ActionEventSource>();

        public bool Initialize(string pluginFolder)
        {
            m_app = new Renga.Application();
            
            // Initialize command router
            commandRouter = new CommandRouter(m_app);
            
            var ui = m_app.UI;
            var panelExtension = ui.CreateUIPanelExtension();

            // Add button for server settings
            var settingsAction = ui.CreateAction();
            settingsAction.DisplayName = "Server Settings";
            settingsAction.ToolTip = "Configure TCP server port and manage server";

            var settingsEvents = new Renga.ActionEventSource(settingsAction);
            settingsEvents.Triggered += (s, e) =>
            {
                ShowServerSettings();
            };
            m_eventSources.Add(settingsEvents);
            panelExtension.AddToolButton(settingsAction);

            ui.AddExtensionToPrimaryPanel(panelExtension);

            // Don't start server automatically - user will start it from settings
            // StartTcpServer(serverPort);

            return true;
        }

        private void ShowServerSettings()
        {
            using (var form = new ServerSettingsForm(serverPort, isServerRunning))
            {
                form.StartServerRequested += (s, e) =>
                {
                    StartTcpServer(form.Port);
                    form.UpdateServerStatus(isServerRunning);
                };
                form.StopServerRequested += (s, e) =>
                {
                    StopTcpServer();
                    form.UpdateServerStatus(isServerRunning);
                };
                form.PortChanged += (s, port) =>
                {
                    serverPort = port;
                };

                form.ShowDialog();
            }
        }

        public void Stop()
        {
            StopTcpServer();
            foreach (var eventSource in m_eventSources)
                eventSource.Dispose();
            m_eventSources.Clear();
        }

        private void StartTcpServer(int port)
        {
            try
            {
                if (port < 1024 || port > 65535)
                {
                    m_app.UI.ShowMessageBox(
                        Renga.MessageIcon.MessageIcon_Error,
                        "Renga_Grasshopper Plugin",
                        $"Invalid port number: {port}. Port must be in range 1024-65535.");
                    return;
                }

                serverPort = port;
                tcpListener = new TcpListener(IPAddress.Any, port);
                tcpListener.Start();
                isServerRunning = true;

                // Start accepting connections asynchronously
                _ = Task.Run(async () => await AcceptConnectionsAsync());

                // Server started successfully - message will be shown in settings form
                System.Diagnostics.Debug.WriteLine($"TCP server started on port {port}");
            }
            catch (Exception ex)
            {
                m_app.UI.ShowMessageBox(
                    Renga.MessageIcon.MessageIcon_Error,
                    "Renga_Grasshopper Plugin",
                    $"Failed to start TCP server: {ex.Message}");
            }
        }

        private void StopTcpServer()
        {
            isServerRunning = false;
            tcpListener?.Stop();
        }

        private async Task AcceptConnectionsAsync()
        {
            while (isServerRunning)
            {
                try
                {
                    var client = await tcpListener!.AcceptTcpClientAsync();
                    _ = Task.Run(async () => await HandleClientAsync(client));
                }
                catch (ObjectDisposedException)
                {
                    // Server was stopped
                    break;
                }
                catch (Exception ex)
                {
                    // Log error but continue accepting connections
                    System.Diagnostics.Debug.WriteLine($"Error accepting connection: {ex.Message}");
                }
            }
        }

        private async Task HandleClientAsync(TcpClient client)
        {
            try
            {
                var stream = client.GetStream();
                stream.ReadTimeout = 10000; // 10 seconds timeout
                
                // Receive message using new protocol
                var json = await Connection.ConnectionProtocol.ReceiveMessageAsync(stream, 10000);
                System.Diagnostics.Debug.WriteLine($"Received JSON ({json.Length} bytes): {json.Substring(0, Math.Min(200, json.Length))}...");
                
                // Parse message
                var message = ConnectionMessage.FromJson(json);
                
                // Route to appropriate handler
                var response = commandRouter.Route(message);
                
                // Send response back to client
                var responseJson = response.ToJson();
                System.Diagnostics.Debug.WriteLine($"Sending response ({responseJson.Length} bytes)");
                
                await Connection.ConnectionProtocol.SendMessageAsync(stream, responseJson);
                System.Diagnostics.Debug.WriteLine($"Response sent successfully");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error handling client: {ex.Message}\n{ex.StackTrace}");
                
                // Try to send error response if possible
                try
                {
                    var errorResponse = new ConnectionResponse
                    {
                        Id = "",
                        Success = false,
                        Error = $"Server error: {ex.Message}"
                    };
                    var stream = client.GetStream();
                    await Connection.ConnectionProtocol.SendMessageAsync(stream, errorResponse.ToJson());
                }
                catch { }
            }
            finally
            {
                try
                {
                    client.Close();
                }
                catch { }
            }
        }

    }
}

