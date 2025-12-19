using System;
using System.IO;
using System.Net.Sockets;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace GrasshopperRNG.Connection
{
    /// <summary>
    /// Reliable TCP client for communicating with Renga plugin
    /// Uses length-prefixed protocol for reliable message delivery
    /// </summary>
    public class RengaConnectionClient
    {
        private static readonly string LogFilePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), 
            "Grasshopper", 
            "RengaGH_Client.log");

        public string Host { get; set; } = "127.0.0.1";
        public int Port { get; set; } = 50100;
        public TimeSpan Timeout { get; set; } = TimeSpan.FromSeconds(10);

        private static void Log(string message)
        {
            try
            {
                var logDir = Path.GetDirectoryName(LogFilePath);
                if (!Directory.Exists(logDir))
                    Directory.CreateDirectory(logDir);

                var logMessage = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] {message}";
                File.AppendAllText(LogFilePath, logMessage + Environment.NewLine);
                System.Diagnostics.Debug.WriteLine(logMessage);
            }
            catch { }
        }

        /// <summary>
        /// Send a message and receive response
        /// Creates a new connection for each request
        /// </summary>
        public ConnectionResponse Send(ConnectionMessage message)
        {
            TcpClient client = null;
            NetworkStream stream = null;

            try
            {
                Log($"=== Creating new connection to {Host}:{Port} ===");
                
                // Create connection
                client = new TcpClient();
                var connectTask = client.ConnectAsync(Host, Port);
                
                if (!connectTask.Wait(Timeout))
                {
                    Log("❌ Connection timeout");
                    return new ConnectionResponse
                    {
                        Id = message.Id,
                        Success = false,
                        Error = "Connection timeout"
                    };
                }

                if (!client.Connected)
                {
                    Log("❌ Failed to connect to server");
                    return new ConnectionResponse
                    {
                        Id = message.Id,
                        Success = false,
                        Error = "Failed to connect to server"
                    };
                }

                Log($"✓ Connected to server {Host}:{Port}");

                stream = client.GetStream();
                stream.ReadTimeout = (int)Timeout.TotalMilliseconds;
                stream.WriteTimeout = (int)Timeout.TotalMilliseconds;

                // Send message
                var json = message.ToJson();
                Log($"→ Sending {json.Length} bytes to server");
                Log($"  Data preview: {json.Substring(0, Math.Min(300, json.Length))}...");
                
                ConnectionProtocol.SendMessage(stream, json);
                Log("✓ Data sent, waiting for response...");

                // Receive response
                var responseJson = ConnectionProtocol.ReceiveMessage(stream, (int)Timeout.TotalMilliseconds);
                Log($"✓ Received {responseJson.Length} bytes from server");
                Log($"  Response preview: {responseJson.Substring(0, Math.Min(300, responseJson.Length))}...");

                var response = ConnectionResponse.FromJson(responseJson);
                return response;
            }
            catch (IOException ioEx)
            {
                Log($"❌ IO Error: {ioEx.Message}");
                return new ConnectionResponse
                {
                    Id = message.Id,
                    Success = false,
                    Error = $"IO Error: {ioEx.Message}"
                };
            }
            catch (SocketException socketEx)
            {
                Log($"❌ Socket Error: {socketEx.Message}");
                return new ConnectionResponse
                {
                    Id = message.Id,
                    Success = false,
                    Error = $"Socket Error: {socketEx.Message}"
                };
            }
            catch (Exception ex)
            {
                Log($"❌ Error: {ex.Message}");
                Log($"Stack trace: {ex.StackTrace}");
                return new ConnectionResponse
                {
                    Id = message.Id,
                    Success = false,
                    Error = $"Error: {ex.Message}"
                };
            }
            finally
            {
                try
                {
                    stream?.Close();
                    client?.Close();
                }
                catch { }
            }
        }

        /// <summary>
        /// Check if server is reachable (quick connection test)
        /// </summary>
        public bool IsServerReachable()
        {
            try
            {
                using (var client = new TcpClient())
                {
                    var connectTask = client.ConnectAsync(Host, Port);
                    if (connectTask.Wait(TimeSpan.FromSeconds(2)))
                    {
                        return client.Connected;
                    }
                }
            }
            catch
            {
                return false;
            }
            return false;
        }
    }
}



