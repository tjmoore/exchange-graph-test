using Serilog;
using System;
using System.Collections.Generic;
using System.CommandLine;
using System.Linq;
using System.Threading;
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

            var clientIdOption = new Option<string>("--client-id", [ "-cid" ])
            { Description = "Graph API Client ID", Arity = ArgumentArity.ExactlyOne };

            var tenantIdOption = new Option<string>("--tenant-id", [ "-tid" ])
            { Description = "Graph API Tenant ID", Arity = ArgumentArity.ZeroOrOne };

            var clientSecretOption = new Option<string>("--client-secret", [ "-cs" ])
            { Description = "Graph API Client Secret", Arity = ArgumentArity.ExactlyOne };

            var mailBoxTemplateOption = new Option<string>("--mailbox-template", [ "-mt" ])
            { Description = "Mailbox address template (format <name>{0}@<domain>)", Arity = ArgumentArity.ExactlyOne };

            var numMailboxOption = new Option<int>("--num-mailbox", [ "-nm" ])
            { Description = "Number of mailboxes to use in template", Arity = ArgumentArity.ExactlyOne };

            var startMailboxOption = new Option<int?>("--start-mailbox", [ "-sm" ])
            { Description = "Start number of mailboxes to use in template, default one", Arity = ArgumentArity.ZeroOrOne };

            var transactionIdOption = new Option<string>("--transaction-id", [ "-trid" ])
            { Description = "Use specified ID as prefix for transaction ID on events", Arity = ArgumentArity.ZeroOrOne };

            var dumpEventsOption = new Option<bool>("--dump-events", [ "-dump" ])
            { Description = "Dump event detail", Arity = ArgumentArity.ZeroOrOne };

            var maxEventsOption = new Option<int?>("--max-events", [ "-me" ])
            { Description = "Max number of events per mailbox, default 1", Arity = ArgumentArity.ZeroOrOne };

            // There used to be global options in previous version of CommandLine but can't find that now
            // so have to add these for each command
            Option[] authOptions = [
                clientIdOption,
                tenantIdOption,
                clientSecretOption
            ];

            Option[] eventOptions = [
                mailBoxTemplateOption,
                numMailboxOption,
                startMailboxOption
            ];

            var rootCommand = new RootCommand($"Exchange Graph API test tool v{AppVersion?.Major}.{AppVersion?.Minor}.{AppVersion?.Build}");
            rootCommand.AddOptions([clientIdOption, tenantIdOption, clientSecretOption]);

            var getCommand = new Command("get-events", "Fetches events matching specified transaction ID, or all events if not specified");
            getCommand.AddOptions(
            [
                ..authOptions,
                ..eventOptions,
                transactionIdOption,
                dumpEventsOption
            ]);

            var createCommmand = new Command("create-events", "Creates sample events");
            createCommmand.AddOptions(
            [
                ..authOptions,
                ..eventOptions,
                maxEventsOption,
                transactionIdOption
            ]);

            var deleteCommand = new Command("delete-events", "Deletes events matching specified transaction ID");
            deleteCommand.AddOptions(
            [
                ..authOptions,
                ..eventOptions,
                transactionIdOption
            ]);

            getCommand.SetAction((parseResult, token) =>
            {
                string? clientId = parseResult.GetValue(clientIdOption);
                string? tenantId = parseResult.GetValue(tenantIdOption);
                string? clientSecret = parseResult.GetValue(clientSecretOption);
                string? mailboxTemplate = parseResult.GetValue(mailBoxTemplateOption);
                int numMailbox = parseResult.GetValue(numMailboxOption);
                int? startMailbox = parseResult.GetValue(startMailboxOption);
                string? transactionId = parseResult.GetValue(transactionIdOption);
                bool dumpEvents = parseResult.GetValue(dumpEventsOption);

                return HandleGet(clientId!, tenantId, clientSecret!, mailboxTemplate!, numMailbox, startMailbox, transactionId, dumpEvents, token);
            });

            createCommmand.SetAction((parseResult, token) =>
            {
                string? clientId = parseResult.GetValue(clientIdOption);
                string? tenantId = parseResult.GetValue(tenantIdOption);
                string? clientSecret = parseResult. GetValue(clientSecretOption);
                string? mailboxTemplate = parseResult.GetValue(mailBoxTemplateOption);
                int numMailbox = parseResult.GetValue(numMailboxOption);
                int? startMailbox = parseResult.GetValue(startMailboxOption);
                int? maxEvents = parseResult.GetValue(maxEventsOption);
                string? transactionId = parseResult.GetValue(transactionIdOption);

                return HandleCreate(clientId!, tenantId, clientSecret!, mailboxTemplate!, numMailbox, startMailbox, maxEvents, transactionId, token);
            });

            deleteCommand.SetAction((parseResult, token) =>
            {
                string? clientId = parseResult.GetValue(clientIdOption);
                string? tenantId = parseResult.GetValue(tenantIdOption);
                string? clientSecret = parseResult.GetValue(clientSecretOption);
                string? mailboxTemplate = parseResult.GetValue(mailBoxTemplateOption);
                int numMailbox = parseResult.GetValue(numMailboxOption);
                int? startMailbox = parseResult.GetValue(startMailboxOption);
                string? transactionId = parseResult.GetValue(transactionIdOption);

                return HandleDelete(clientId!, tenantId, clientSecret!, mailboxTemplate!, numMailbox, startMailbox, transactionId, token);
            });

            rootCommand.Add(getCommand);
            rootCommand.Add(createCommmand);
            rootCommand.Add(deleteCommand);

            var parseResult = rootCommand.Parse(args);
            return await parseResult.InvokeAsync();
        }

        static async Task HandleGet(string clientId, string? tenantId, string clientSecret, string mailboxTemplate, int numMailbox, int? startMailbox, string? transactionId, bool dumpEvents, CancellationToken token)
        {
            if (string.IsNullOrEmpty(tenantId))
                tenantId = "common";

            var factory = new GraphApiFactory(clientId, tenantId, clientSecret);
            var calendar = new Calendar(factory.Client);
            var mailboxes = Enumerable.Range(startMailbox ?? 1, numMailbox).Select(i => string.Format(mailboxTemplate, i));

            Log.Information("Find events...");
            var eventLists = await calendar.FindEvents(mailboxes, transactionId, token);

            int total = 0;

            foreach(var list in eventLists)
            {
                int count = list.Value.Count;
                total += count;

                Log.Information("Found {count} events for {list}", count, list.Key);

                if (dumpEvents)
                {
                    foreach (var e in list.Value)
                    {
                        Log.Information("{start} - {end}: {subject}", e.Start?.DateTime, e.End?.DateTime, e.Subject);
                    }
                }
            }

            Log.Information("Total: {total} events", total);
        }

        static async Task HandleCreate(string clientId, string? tenantId, string clientSecret, string mailboxTemplate, int numMailbox, int? startMailbox, int? maxEvents, string? transactionId, CancellationToken token)
        {
            if (string.IsNullOrEmpty(tenantId))
                tenantId = "common";

            var factory = new GraphApiFactory(clientId, tenantId, clientSecret);
            var calendar = new Calendar(factory.Client);
            var mailboxes = Enumerable.Range(startMailbox ?? 1, numMailbox).Select(i => string.Format(mailboxTemplate, i));

            if (string.IsNullOrEmpty(transactionId))
                transactionId = Guid.NewGuid().ToString();

            Log.Information("Transaction ID = {transactionId}", transactionId);

            Log.Information("Creating events...");
            await calendar.CreateSampleEvents(mailboxes, maxEvents ?? 1, transactionId, token);
        }

        static async Task HandleDelete(string clientId, string? tenantId, string clientSecret, string mailboxTemplate, int numMailbox, int? startMailbox, string? transactionId, CancellationToken token)
        {
            if (string.IsNullOrEmpty(tenantId))
                tenantId = "common";

            var factory = new GraphApiFactory(clientId, tenantId, clientSecret);
            var calendar = new Calendar(factory.Client);
            var mailboxes = Enumerable.Range(startMailbox ?? 1, numMailbox).Select(i => string.Format(mailboxTemplate, i));

            var events = await calendar.FindEvents(mailboxes, transactionId, token);

            Log.Information("Deleting {numEvents} events...", events.Values.SelectMany(e => e).Count());

            await calendar.DeleteEvents(events, token);
        }
    }

    static class CommandOptionsHelper
    {
        public static void AddOptions(this Command command, IEnumerable<Option> options)
        {
            foreach (var option in options)
            {
                command.Add(option);
            }
        }
    }
}
