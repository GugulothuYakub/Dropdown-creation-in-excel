namespace entitycore.MiddleWare
{
    public class FileLogger : ILogger
    {
        private readonly string _path;
        private readonly object _lock = new object();

        public FileLogger(string path)
        {
            _path = path;
        }

        public IDisposable BeginScope<TState>(TState state)
        {
            return null;
        }

        public bool IsEnabled(LogLevel logLevel)
        {
            return true;
        }
        public void Log<TState>(LogLevel logLevel, EventId eventId, TState state, Exception exception, Func<TState, Exception, string> formatter)
        {
            if (logLevel == LogLevel.None)
            {
                return;
            }

            var message = formatter(state, exception);
            #region
            /* This code is commented by yakub.gugulothu@cesltd.com*/

            // Skip logging the specific message
            if (message == "An unhandled exception occurred.")
            {
                return;
            }
            #endregion

            var logMessage = $"{DateTime.UtcNow} [{logLevel}] : {message}{Environment.NewLine}";

            if (exception != null)
            {
                logMessage += $"Exception: {exception.Message}{Environment.NewLine}";
                logMessage += $"Stack Trace: {exception.StackTrace}{Environment.NewLine}";
            }

            lock (_lock)
            {
                File.AppendAllText(_path, logMessage);
            }
        }
    }
}
