using System;
using System.IO;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading.Tasks;

namespace GrasshopperRNG.Connection
{
    /// <summary>
    /// Protocol for reliable TCP communication with length-prefixed messages
    /// </summary>
    public static class ConnectionProtocol
    {
        /// <summary>
        /// Send a message with length prefix (4 bytes big-endian + JSON data)
        /// </summary>
        public static async Task SendMessageAsync(NetworkStream stream, string json)
        {
            if (stream == null || !stream.CanWrite)
                throw new InvalidOperationException("Stream is not writable");

            var data = Encoding.UTF8.GetBytes(json);
            var length = BitConverter.GetBytes(IPAddress.HostToNetworkOrder(data.Length));
            
            // Send length (4 bytes)
            await stream.WriteAsync(length, 0, 4);
            
            // Send data
            await stream.WriteAsync(data, 0, data.Length);
            
            // Flush to ensure data is sent
            await stream.FlushAsync();
        }

        /// <summary>
        /// Receive a message with length prefix
        /// </summary>
        public static async Task<string> ReceiveMessageAsync(NetworkStream stream, int timeoutMs = 10000)
        {
            if (stream == null || !stream.CanRead)
                throw new InvalidOperationException("Stream is not readable");

            // Set read timeout
            stream.ReadTimeout = timeoutMs;

            // Read length (4 bytes)
            var lengthBytes = new byte[4];
            int totalRead = 0;
            
            while (totalRead < 4)
            {
                var read = await stream.ReadAsync(lengthBytes, totalRead, 4 - totalRead);
                if (read == 0)
                    throw new IOException("Connection closed while reading message length");
                totalRead += read;
            }
            
            int length = IPAddress.NetworkToHostOrder(BitConverter.ToInt32(lengthBytes, 0));
            
            if (length < 0 || length > 10 * 1024 * 1024) // Max 10MB
                throw new IOException($"Invalid message length: {length}");

            // Read JSON data
            var buffer = new byte[length];
            totalRead = 0;
            
            while (totalRead < length)
            {
                var read = await stream.ReadAsync(buffer, totalRead, length - totalRead);
                if (read == 0)
                    throw new IOException("Connection closed while reading message data");
                totalRead += read;
            }
            
            return Encoding.UTF8.GetString(buffer, 0, length);
        }

        /// <summary>
        /// Synchronous version for compatibility
        /// </summary>
        public static void SendMessage(NetworkStream stream, string json)
        {
            if (stream == null || !stream.CanWrite)
                throw new InvalidOperationException("Stream is not writable");

            var data = Encoding.UTF8.GetBytes(json);
            var length = BitConverter.GetBytes(IPAddress.HostToNetworkOrder(data.Length));
            
            // Send length (4 bytes)
            stream.Write(length, 0, 4);
            
            // Send data
            stream.Write(data, 0, data.Length);
            
            // Flush to ensure data is sent
            stream.Flush();
        }

        /// <summary>
        /// Synchronous version for compatibility
        /// </summary>
        public static string ReceiveMessage(NetworkStream stream, int timeoutMs = 10000)
        {
            if (stream == null || !stream.CanRead)
                throw new InvalidOperationException("Stream is not readable");

            // Set read timeout
            stream.ReadTimeout = timeoutMs;

            // Read length (4 bytes)
            var lengthBytes = new byte[4];
            int totalRead = 0;
            
            while (totalRead < 4)
            {
                var read = stream.Read(lengthBytes, totalRead, 4 - totalRead);
                if (read == 0)
                    throw new IOException("Connection closed while reading message length");
                totalRead += read;
            }
            
            int length = IPAddress.NetworkToHostOrder(BitConverter.ToInt32(lengthBytes, 0));
            
            if (length < 0 || length > 10 * 1024 * 1024) // Max 10MB
                throw new IOException($"Invalid message length: {length}");

            // Read JSON data
            var buffer = new byte[length];
            totalRead = 0;
            
            while (totalRead < length)
            {
                var read = stream.Read(buffer, totalRead, length - totalRead);
                if (read == 0)
                    throw new IOException("Connection closed while reading message data");
                totalRead += read;
            }
            
            return Encoding.UTF8.GetString(buffer, 0, length);
        }
    }
}



