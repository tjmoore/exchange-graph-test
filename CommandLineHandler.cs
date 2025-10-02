using Serilog;
using System;
using System.Collections.Generic;
using System.CommandLine;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace ExchangeGraphTool
{
    internal class CommandLineHandler
    {
        private readonly Option<string> _clientIdOption = new("--client-id", ["-cid"])
        { Description = "Graph API Client ID", Arity = ArgumentArity.ExactlyOne };

        private readonly Option<string> _tenantIdOption = new("--tenant-id", ["-tid"])
        { Description = "Graph API Tenant ID", Arity = ArgumentArity.ZeroOrOne };

        private readonly Option<string> _clientSecretOption = new("--client-secret", ["-cs"])
        { Description = "Graph API Client Secret", Arity = ArgumentArity.ExactlyOne };

        private readonly Option<string> _mailBoxTemplateOption = new("--mailbox-template", ["-mt"])
        { Description = "Mailbox address template (format <name>{0}@<domain>)", Arity = ArgumentArity.ExactlyOne };

        private readonly Option<int> _numMailboxOption = new("--num-mailbox", ["-nm"])
        { Description = "Number of mailboxes to use in template", Arity = ArgumentArity.ExactlyOne };

        private readonly Option<int?> _startMailboxOption = new("--start-mailbox", ["-sm"])
        { Description = "Start number of mailboxes to use in template, default one", Arity = ArgumentArity.ZeroOrOne };

        private readonly Option<string> _transactionIdOption = new("--transaction-id", ["-trid"])
        { Description = "Use specified ID as prefix for transaction ID on events", Arity = ArgumentArity.ZeroOrOne };

        private readonly Option<bool> _dumpEventsOption = new("--dump-events", ["-dump"])
        { Description = "Dump event detail", Arity = ArgumentArity.ZeroOrOne };

        private readonly Option<int?> _maxEventsOption = new("--max-events", ["-me"])
        { Description = "Max number of events per mailbox per run, default 1, max 4", Arity = ArgumentArity.ZeroOrOne };

        private readonly Option<int?> _batchSize = new("--batch-size", ["-bs"])
        { Description = "Max batch size for Graph API calls, default 20", Arity = ArgumentArity.ZeroOrOne };

        private struct CommandParams
        {
            public string ClientId;
            public string? TenantId;
            public string ClientSecret;
            public string MailboxTemplate;
            public int NumMailbox;
            public int? StartMailbox;
            public string? TransactionId;
            public int? MaxEvents;
            public bool DumpEvents;
            public int? BatchSize;
        }

        public async Task<int> Process(string[] args)
        {
            // There used to be global options in previous version of CommandLine but can't find that now
            // so have to add these for each command
            Option[] authOptions = [
                _clientIdOption,
                _tenantIdOption,
                _clientSecretOption
            ];

            Option[] eventOptions = [
                _mailBoxTemplateOption,
                _numMailboxOption,
                _startMailboxOption,
                _batchSize,
                _transactionIdOption
            ];

            var rootCommand = new RootCommand($"Exchange Graph API test tool v{Program.AppVersion?.Major}.{Program.AppVersion?.Minor}.{Program.AppVersion?.Build}");

            var getCommand = new Command("get-events", "Fetches events matching specified transaction ID, or all events if not specified");
            getCommand.AddOptions(
            [
                ..authOptions,
                ..eventOptions,
                _dumpEventsOption
            ]);

            var createCommmand = new Command("create-events", "Creates sample events");
            createCommmand.AddOptions(
            [
                ..authOptions,
                ..eventOptions,
                _maxEventsOption
            ]);

            var deleteCommand = new Command("delete-events", "Deletes events matching specified transaction ID");
            deleteCommand.AddOptions(
            [
                ..authOptions,
                ..eventOptions
            ]);

            getCommand.SetAction((parseResult, token) =>
            {
                var commandParams = ParseCommandParams(parseResult);

                using var graphApiFactory = GetGraphFactory(commandParams);

                return HandleGet(
                    graphApiFactory,
                    commandParams.MailboxTemplate,
                    commandParams.NumMailbox,
                    commandParams.StartMailbox,
                    commandParams.TransactionId,
                    commandParams.DumpEvents,
                    commandParams.BatchSize,
                    token);
            });

            createCommmand.SetAction((parseResult, token) =>
            {
                var commandParams = ParseCommandParams(parseResult);

                using var graphApiFactory = GetGraphFactory(commandParams);

                return HandleCreate(
                    graphApiFactory,
                    commandParams.MailboxTemplate,
                    commandParams.NumMailbox,
                    commandParams.StartMailbox,
                    commandParams.MaxEvents,
                    commandParams.TransactionId,
                    commandParams.BatchSize,
                    token);
            });

            deleteCommand.SetAction((parseResult, token) =>
            {
                var commandParams = ParseCommandParams(parseResult);

                using var graphApiFactory = GetGraphFactory(commandParams);

                return HandleDelete(
                    graphApiFactory,
                    commandParams.MailboxTemplate,
                    commandParams.NumMailbox,
                    commandParams.StartMailbox,
                    commandParams.TransactionId,
                    commandParams.BatchSize,
                    token);
            });

            rootCommand.Add(getCommand);
            rootCommand.Add(createCommmand);
            rootCommand.Add(deleteCommand);

            var parseResult = rootCommand.Parse(args);
            return await parseResult.InvokeAsync();
        }

        private CommandParams ParseCommandParams(ParseResult parseResult)
        {
            return new CommandParams
            {
                ClientId = parseResult.GetValue(_clientIdOption)!,
                TenantId = parseResult.GetValue(_tenantIdOption),
                ClientSecret = parseResult.GetValue(_clientSecretOption)!,
                MailboxTemplate = parseResult.GetValue(_mailBoxTemplateOption)!,
                NumMailbox = parseResult.GetValue(_numMailboxOption),
                StartMailbox = parseResult.GetValue(_startMailboxOption),
                TransactionId = parseResult.GetValue(_transactionIdOption),
                MaxEvents = parseResult.GetValue(_maxEventsOption),
                DumpEvents = parseResult.GetValue(_dumpEventsOption),
                BatchSize = parseResult.GetValue(_batchSize)
            };
        }

        private static GraphApiFactory GetGraphFactory(CommandParams commandParams)
        {
            string? tenantId = commandParams.TenantId;
            if (string.IsNullOrEmpty(tenantId))
                tenantId = "common";

            return new GraphApiFactory(commandParams.ClientId, tenantId, commandParams.ClientSecret);
        }

        private static async Task HandleGet(
            GraphApiFactory graphApiFactory,
            string mailboxTemplate,
            int numMailbox,
            int? startMailbox,
            string? transactionId,
            bool dumpEvents,
            int? maxBatchSize,
            CancellationToken token)
        {
            var calendar = new Calendar(graphApiFactory.Client) { BatchSize = maxBatchSize ?? 20 };
            var mailboxes = Enumerable.Range(startMailbox ?? 1, numMailbox).Select(i => string.Format(mailboxTemplate, i));

            Log.Information("Find events...");
            var eventLists = await calendar.FindEvents(mailboxes, transactionId, token);

            int total = 0;

            foreach (var list in eventLists)
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

        private static async Task HandleCreate(
            GraphApiFactory graphApiFactory,
            string mailboxTemplate,
            int numMailbox,
            int? startMailbox,
            int? maxEvents,
            string? transactionId,
            int? maxBatchSize,
            CancellationToken token)
        {
            var calendar = new Calendar(graphApiFactory.Client) { BatchSize = maxBatchSize ?? 20 };
            var mailboxes = Enumerable.Range(startMailbox ?? 1, numMailbox).Select(i => string.Format(mailboxTemplate, i));

            if (string.IsNullOrEmpty(transactionId))
                transactionId = Guid.NewGuid().ToString();

            Log.Information("Transaction ID = {transactionId}", transactionId);

            Log.Information("Creating events...");
            await calendar.CreateSampleEvents(mailboxes, maxEvents ?? 1, transactionId, token);
        }

        private static async Task HandleDelete(
            GraphApiFactory graphApiFactory,
            string mailboxTemplate,
            int numMailbox,
            int? startMailbox,
            string? transactionId,
            int? maxBatchSize,
            CancellationToken token)
        {
            var calendar = new Calendar(graphApiFactory.Client) { BatchSize = maxBatchSize ?? 20 };
            var mailboxes = Enumerable.Range(startMailbox ?? 1, numMailbox).Select(i => string.Format(mailboxTemplate, i));

            var events = await calendar.FindEvents(mailboxes, transactionId, token);

            Log.Information("Deleting {numEvents} events...", events.Values.SelectMany(e => e).Count());

            await calendar.DeleteEvents(events, token);
        }
    }

    internal static class CommandOptionsHelper
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
