using System;
using System.IO;
using System.IO.Pipes;
using System.Threading;
using System.Threading.Tasks;

namespace MsgToPdfConverter.Utils
{
    public class SingleInstanceManager : IDisposable
    {
        private static readonly string UserSid = System.Security.Principal.WindowsIdentity.GetCurrent().User?.Value ?? "default";
        private static readonly string MutexName = $"Local\\MsgToPdfConverter_SingleInstance_Mutex_{UserSid}";
        private static readonly string PipeName = $"Local\\MsgToPdfConverter_SingleInstance_Pipe_{UserSid}";
        private Mutex _mutex;
        private bool _isFirstInstance;
        private CancellationTokenSource _cts;

        private static void Log(string msg)
        {
            try
            {
               // System.IO.File.AppendAllText("SingleInstanceManager.log", $"{System.DateTime.Now:yyyy-MM-dd HH:mm:ss.fff} {msg}\r\n");
            }
            catch { }
        }

        public bool IsFirstInstance => _isFirstInstance;

        public event Action<string> FileReceived;

        public SingleInstanceManager()
        {
            Log($"[CTOR] PID: {System.Diagnostics.Process.GetCurrentProcess().Id}, User: {Environment.UserName}, MutexName: {MutexName}, PipeName: {PipeName}");
            _mutex = new Mutex(true, MutexName, out _isFirstInstance);
            Log($"[CTOR] IsFirstInstance: {_isFirstInstance}");
            if (_isFirstInstance)
            {
                _cts = new CancellationTokenSource();
                Task.Run(() => ListenForFiles(_cts.Token));
                Log("[CTOR] Started pipe listener");
            }
        }

        public void SendFileToFirstInstance(string filePath)
        {
            Log($"[SendFile] PID: {System.Diagnostics.Process.GetCurrentProcess().Id}, User: {Environment.UserName}, PipeName: {PipeName}, File: {filePath}");
            try
            {
                using (var client = new NamedPipeClientStream(".", PipeName, PipeDirection.Out))
                {
                    client.Connect(1000);
                    using (var writer = new StreamWriter(client))

                    {
                        writer.WriteLine(filePath);
                        writer.Flush();
                    }
                }
                Log("[SendFile] Success");
            }
            catch (Exception ex)
            {
                Log($"[SendFile] Exception: {ex}");
            }
        }

        private void ListenForFiles(CancellationToken token)
        {
            while (!token.IsCancellationRequested)
            {
                try
                {
                    using (var server = new NamedPipeServerStream(PipeName, PipeDirection.In))
                    {
                        server.WaitForConnection();
                        using (var reader = new StreamReader(server))
                        {
                            string file = reader.ReadLine();
                            if (!string.IsNullOrEmpty(file))
                                FileReceived?.Invoke(file);
                        }
                    }
                }
                catch { }
            }
        }

        public void Dispose()
        {
            _cts?.Cancel();
            _mutex?.Dispose();
        }
    }
}
