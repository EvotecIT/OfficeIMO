using System;
using System.IO;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading;

/// <summary>A test-owned loopback server with synchronous port reservation and bounded shutdown.</summary>
public sealed class OfficeIMOTestHttpFileServer : IDisposable {
    private readonly TcpListener listener;
    private readonly Thread worker;
    private readonly string filePath;
    private readonly string contentType;
    private readonly int requestCount;
    private volatile bool stopped;
    private TcpClient activeClient;

    public OfficeIMOTestHttpFileServer(string path, string type, int count) {
        filePath = path;
        contentType = type;
        requestCount = count;
        listener = new TcpListener(IPAddress.Loopback, 0);
        listener.Start();
        Port = ((IPEndPoint)listener.LocalEndpoint).Port;
        worker = new Thread(Serve);
        worker.IsBackground = true;
        worker.Start();
    }

    public int Port { get; private set; }
    public Exception Failure { get; private set; }

    private void Serve() {
        try {
            for (int index = 0; index < requestCount && !stopped; index++) {
                using (TcpClient client = listener.AcceptTcpClient()) {
                    activeClient = client;
                    if (stopped) return;
                    using (NetworkStream stream = client.GetStream()) {
                        stream.ReadTimeout = 5000;
                        stream.WriteTimeout = 5000;
                        byte[] request = new byte[4096];
                        int length = 0;
                        while (length < request.Length) {
                            int read = stream.Read(request, length, request.Length - length);
                            if (read <= 0) break;
                            length += read;
                            if (Encoding.ASCII.GetString(request, 0, length).Contains("\r\n\r\n")) break;
                        }
                        byte[] body = File.ReadAllBytes(filePath);
                        byte[] header = Encoding.ASCII.GetBytes(
                            "HTTP/1.1 200 OK\r\nContent-Type: " + contentType +
                            "\r\nContent-Length: " + body.Length + "\r\nConnection: close\r\n\r\n");
                        stream.Write(header, 0, header.Length);
                        stream.Write(body, 0, body.Length);
                        stream.Flush();
                    }
                    activeClient = null;
                }
            }
        } catch (Exception error) {
            if (!stopped) Failure = error;
        } finally {
            listener.Stop();
        }
    }

    public void Dispose() {
        stopped = true;
        listener.Stop();
        TcpClient client = activeClient;
        if (client != null) client.Close();
        if (!worker.Join(6000)) throw new TimeoutException("The test HTTP server did not stop.");
    }
}
