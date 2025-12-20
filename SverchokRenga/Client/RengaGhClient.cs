using System;
using System.Collections.Generic;
using System.Net.Sockets;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace SverchokRenga.Client
{
    /// <summary>
    /// TCP client for communicating with Renga plugin
    /// Updated: Improved response reading and reconnection logic
    /// </summary>
    public class RengaGhClient
    {
        private TcpClient tcpClient;
        private NetworkStream stream;
        private static readonly string LogFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Grasshopper", "RengaGH_Client.log");
        
        public string Host { get; set; } = "127.0.0.1";
        public int Port { get; set; } = 50100;
        public TimeSpan Timeout { get; set; } = TimeSpan.FromSeconds(10);
        public bool IsConnected => tcpClient != null && tcpClient.Connected;
        
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
        /// Connect to Renga TCP server
        /// </summary>
        public bool Connect()
        {
            try
            {
                tcpClient = new TcpClient();
                var connectTask = tcpClient.ConnectAsync(Host, Port);
                if (connectTask.Wait(Timeout))
                {
                    if (tcpClient.Connected)
                    {
                        stream = tcpClient.GetStream();
                        return true;
                    }
                }
                return false;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Disconnect from server
        /// </summary>
        public void Disconnect()
        {
            stream?.Close();
            tcpClient?.Close();
            stream = null;
            tcpClient = null;
        }

        /// <summary>
        /// Send JSON data to server and get response
        /// Creates a new connection for each request to ensure clean communication
        /// </summary>
        public string Send(string jsonData)
        {
            TcpClient tempClient = null;
            NetworkStream tempStream = null;
            
            try
            {
                // Create a new connection for this request
                // Server closes connection after response, so we need fresh connection each time
                Log($"=== Creating new connection to {Host}:{Port} ===");
                tempClient = new TcpClient();
                var connectTask = tempClient.ConnectAsync(Host, Port);
                if (!connectTask.Wait(Timeout))
                {
                    Log("❌ Connection timeout");
                    return null;
                }
                
                if (!tempClient.Connected)
                {
                    Log("❌ Failed to connect to server");
                    return null;
                }
                
                Log($"✓ Connected to server {Host}:{Port}");
                
                tempStream = tempClient.GetStream();
                tempStream.ReadTimeout = 10000; // 10 second timeout
                
                // Send data
                var data = Encoding.UTF8.GetBytes(jsonData);
                Log($"→ Sending {data.Length} bytes to server");
                Log($"  Data preview: {jsonData.Substring(0, Math.Min(300, jsonData.Length))}...");
                tempStream.Write(data, 0, data.Length);
                tempStream.Flush();
                Log("✓ Data sent, waiting for response...");

                // Read response - server will send response and close connection
                var buffer = new List<byte>();
                var readBuffer = new byte[8192];
                bool dataReceived = false;
                int attempts = 0;
                const int maxAttempts = 100; // 10 seconds (100 * 100ms)
                
                // Wait a bit for server to process and send response
                System.Threading.Thread.Sleep(200);
                
                // Try to read response - use simple blocking read with timeout
                try
                {
                    // Read response in a loop until we get all data or connection closes
                    while (true)
                    {
                        // Check if connection is still alive
                        if (!tempClient.Connected && buffer.Count == 0)
                        {
                            System.Diagnostics.Debug.WriteLine("  Connection closed before any data received");
                            break;
                        }
                        
                        // Try to read data
                        if (tempStream.DataAvailable)
                        {
                            int bytesRead = tempStream.Read(readBuffer, 0, readBuffer.Length);
                            if (bytesRead > 0)
                            {
                                dataReceived = true;
                                Log($"  Read {bytesRead} bytes from stream");
                                
                                for (int i = 0; i < bytesRead; i++)
                                {
                                    buffer.Add(readBuffer[i]);
                                }
                                
                                // Continue reading if more data available
                                if (tempStream.DataAvailable)
                                    continue;
                            }
                        }
                        
                        // If we received data and no more is available, check if connection is closed
                        if (dataReceived)
                        {
                            // Wait a bit to see if more data arrives
                            System.Threading.Thread.Sleep(100);
                            
                            // If connection closed or no more data, we're done
                            if (!tempClient.Connected || !tempStream.DataAvailable)
                            {
                                Log("  Message complete (connection closed or no more data)");
                                break;
                            }
                        }
                        else
                        {
                            // No data yet, wait and retry
                            System.Threading.Thread.Sleep(100);
                            attempts++;
                            
                            if (attempts >= maxAttempts)
                            {
                                Log($"  Timeout waiting for response (attempts: {attempts})");
                                break;
                            }
                            
                            // Check connection status
                            if (!tempClient.Connected)
                            {
                                Log("  Connection closed before receiving response");
                                break;
                            }
                        }
                    }
                }
                catch (System.IO.IOException ioEx)
                {
                    // Connection closed or timeout - this is expected after server sends response
                    Log($"  Read IOException: {ioEx.Message}");
                    // If we received data, it's complete
                    if (!dataReceived)
                    {
                        Log("  Connection closed before receiving response");
                    }
                }
                catch (Exception ex)
                {
                    Log($"  Unexpected read error: {ex.Message}");
                }
                
                if (buffer.Count > 0)
                {
                    var response = Encoding.UTF8.GetString(buffer.ToArray());
                    Log($"✓ Received {buffer.Count} bytes from server");
                    Log($"  Response preview: {response.Substring(0, Math.Min(300, response.Length))}...");
                    return response;
                }
                else
                {
                    Log("❌ No response received from server");
                    Log($"  Connection status: {tempClient?.Connected ?? false}");
                    Log($"  Data received flag: {dataReceived}");
                    Log($"  Attempts: {attempts}");
                    return null;
                }
            }
            catch (Exception ex)
            {
                Log($"Error sending data: {ex.Message}");
                Log($"Stack trace: {ex.StackTrace}");
                return null;
            }
            finally
            {
                // Clean up temporary connection
                try
                {
                    tempStream?.Close();
                    tempClient?.Close();
                }
                catch { }
            }
        }

        public void Dispose()
        {
            Disconnect();
        }
    }
}


