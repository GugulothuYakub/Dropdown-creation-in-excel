using Microsoft.Extensions.Logging.Abstractions;

namespace entitycore.MiddleWare
{
    public class FileLoggerProvider : ILoggerProvider
    {
        private readonly string _path;

        public FileLoggerProvider(string path)
        {
            _path = path;
        }

        public ILogger CreateLogger(string categoryName)
        {
            if (categoryName.StartsWith("Microsoft.Hosting.Lifetime"))
            {
                return NullLogger.Instance;
            }

            if (categoryName == "Microsoft.AspNetCore.HttpsPolicy.HttpsRedirectionMiddleware")
            {
                return NullLogger.Instance;
            }

            var className = categoryName.Split(".").LastOrDefault();
            var filePath = Path.Combine(_path, $"{className}_log_{DateTime.Now.ToString("MMddyyyy")}.log");
            return new FileLogger(filePath);
        }

        public void Dispose()
        {
        }
    }
}
