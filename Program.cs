using System;
using System.CommandLine;
using System.CommandLine.Invocation;
using System.Linq;
using System.Threading.Tasks;

namespace ExchangeGraphTool
{
    class Program
    {
        static async Task<int> Main(string[] args)
        {
            var rootCommand = new RootCommand("Exchange Graph API test tool for creating bulk events");

            rootCommand.AddGlobalOption(new Option<string>("--clientId", "Graph API Client ID", ArgumentArity.ExactlyOne));
            rootCommand.AddGlobalOption(new Option<string>("--tenantId", "Graph API Tenant ID", ArgumentArity.ZeroOrOne));
            rootCommand.AddGlobalOption(new Option<string>("--clientSecret", "Graph API Client Secret", ArgumentArity.ExactlyOne));
            rootCommand.AddGlobalOption(new Option<string>("--mailboxTemplate", "Mailbox address template (format <name>{0}@<domain>)", ArgumentArity.ExactlyOne));
            rootCommand.AddGlobalOption(new Option<int>("--numMailbox", "Number of mailboxes to use in template", ArgumentArity.ExactlyOne));
            rootCommand.AddGlobalOption(new Option<int?>("--startMailbox", "Start number of mailboxes to use in template, default zero", ArgumentArity.ZeroOrOne));

            var getCommand = new Command("get", "Fetches events matching specified transaction ID, or all events if not specified")
            {
                new Option<string>("--transactionId", "Use specified ID for transaction ID on events or return all events otherwise", ArgumentArity.ZeroOrOne),
                new Option<bool>("--dumpEvents", "Dump event detail", ArgumentArity.ZeroOrOne)
            };

            var createCommmand = new Command("create", "Creates sample events")
            {
                new Option<int?>("--maxEvents", "Max number of events per mailbox, default 1", ArgumentArity.ZeroOrOne),
                new Option<string>("--transactionId", "Use specified ID for transaction ID on events, otherwise generates a new GUID", ArgumentArity.ZeroOrOne)                
            };

            var deleteCommand = new Command("delete", "Deletes events matching specified transaction ID")
            {
                new Option<string>("--transactionId", "Use specified ID for transaction ID to match events to delete", ArgumentArity.ExactlyOne)
            };

            getCommand.Handler = CommandHandler.Create<string, string, string, string, int, int?, string, bool>(HandleGet);
            createCommmand.Handler = CommandHandler.Create<string, string, string, string, int, int?, int?, string>(HandleCreate);
            deleteCommand.Handler = CommandHandler.Create<string, string, string, string, int, int?, string>(HandleDelete);

            rootCommand.AddCommand(getCommand);
            rootCommand.AddCommand(createCommmand);
            rootCommand.AddCommand(deleteCommand);

            return await rootCommand.InvokeAsync(args);
        }

        static async Task HandleGet(string clientId, string tenantId, string clientSecret, string mailboxTemplate, int numMailbox, int? startMailbox, string transactionId, bool dumpEvents)
        {
            if (string.IsNullOrEmpty(tenantId))
                tenantId = "common";

            var factory = new GraphApiFactory(clientId, tenantId, clientSecret);
            var calendar = new Calendar(factory.Client);
            var mailboxes = Enumerable.Range(startMailbox ?? 1, numMailbox).Select(i => string.Format(mailboxTemplate, i));

            Console.WriteLine("Find events...");
            var eventLists = await calendar.FindEvents(mailboxes, transactionId);

            int total = 0;

            foreach(var list in eventLists)
            {
                int count = list.Value.Count;
                total += count;

                Console.WriteLine($"Found {count} events for {list.Key}");

                if (dumpEvents)
                {
                    foreach (var e in list.Value)
                    {
                        Console.WriteLine($"{e.Start.DateTime} - {e.End.DateTime}: {e.Subject}");
                    }
                }
            }

            Console.WriteLine($"Total: {total} events");
        }

        static async Task HandleCreate(string clientId, string tenantId, string clientSecret, string mailboxTemplate, int numMailbox, int? startMailbox, int? maxEvents, string transactionId)
        {
            if (string.IsNullOrEmpty(tenantId))
                tenantId = "common";

            var factory = new GraphApiFactory(clientId, tenantId, clientSecret);
            var calendar = new Calendar(factory.Client);
            var mailboxes = Enumerable.Range(startMailbox ?? 1, numMailbox).Select(i => string.Format(mailboxTemplate, i));

            if (string.IsNullOrEmpty(transactionId))
                transactionId = Guid.NewGuid().ToString();

            Console.WriteLine($"Transaction ID = {transactionId}");

            Console.WriteLine("Creating events...");
            await calendar.CreateSampleEvents(mailboxes, maxEvents ?? 1, transactionId);
        }

        static async Task HandleDelete(string clientId, string tenantId, string clientSecret, string mailboxTemplate, int numMailbox, int? startMailbox, string transactionId)
        {
            if (string.IsNullOrEmpty(tenantId))
                tenantId = "common";

            var factory = new GraphApiFactory(clientId, tenantId, clientSecret);
            var calendar = new Calendar(factory.Client);
            var mailboxes = Enumerable.Range(startMailbox ?? 1, numMailbox).Select(i => string.Format(mailboxTemplate, i));

            var events = await calendar.FindEvents(mailboxes, transactionId);

            Console.WriteLine($"Deleting {events.Values.SelectMany(e => e).Count()} events...");

            await calendar.DeleteEvents(events);
        }
    }
}
