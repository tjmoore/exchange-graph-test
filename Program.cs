using Serilog;
using System;
using System.Threading.Tasks;

namespace ExchangeGraphTool
{
    class Program
    {
        public static Version? AppVersion = null;

        static async Task<int> Main(string[] args)
        {
            AppVersion = typeof(Program).Assembly.GetName().Version;

            Log.Logger = new LoggerConfiguration()
                .WriteTo.Console()
                .CreateLogger();

            return await new CommandLineHandler().Process(args);
        }
    }
}
