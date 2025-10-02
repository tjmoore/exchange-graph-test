using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Kiota.Abstractions;
using Serilog;
using System;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.Linq;
using System.Net;
using System.Threading;
using System.Threading.Tasks;

namespace ExchangeGraphTool
{
    public class Calendar(GraphServiceClient client)
    {
        private const int DefaultBatchSize = 20;

        private readonly GraphServiceClient _client = client;

        private const string TimeZoneUtcHeader = "outlook.timezone=\"UTC\"";

        public int BatchSize { get; set; } = DefaultBatchSize;

        /// <summary>
        /// Create sample events in each of the specified mailbox calendars, with random number of events up to maxEventsPerMailbox
        /// </summary>
        /// <param name="mailboxes"></param>
        /// <param name="maxEventsPerMailbox">Max number of events to create per mailbox</param>
        /// <returns></returns>
        public async Task CreateSampleEvents(IEnumerable<string> mailboxes, int maxEventsPerMailbox, string transactionId, CancellationToken token = default)
        {
            if (mailboxes == null || !mailboxes.Any())
                return;

            // TODO: batches work in parallel on the server but per mailbox there's a 4 request concurrent limit.
            // Group into 4s, or spread across batches ensuring no more than 4 of same mailbox in a batch?
            maxEventsPerMailbox = Math.Min(maxEventsPerMailbox, 4);

            int eventNum = 1;

            DateTime start = DateTime.UtcNow;

            for (int numRuns = 0; numRuns < 10; numRuns++)
            {
                Log.Information("Create events in {mailboxes} mailboxes starting from {start}", mailboxes.Count(), start);

                eventNum = await CreateEvents(mailboxes, start, maxEventsPerMailbox, transactionId, eventNum, token);
                start = start.AddHours(2);
            }
        }

        private async Task<int> CreateEvents(IEnumerable<string> mailboxes, DateTime start, int maxEventsPerMailbox, string transactionId, int eventNum, CancellationToken token)
        {
            var rnd = new Random();

            var requests = new Dictionary<string, RequestInformation>();

            foreach (var mailbox in mailboxes)
            {
                DateTime nextStart = start;

                int numEvents = rnd.Next(0, maxEventsPerMailbox + 1);

                Log.Debug("{mailbox} - Creating {numEvents} events", mailbox, numEvents);

                for (int n = 1; n <= numEvents; n++)
                {
                    string stepId = Guid.NewGuid().ToString();
                    requests.Add(stepId, BuildCreateEventRequest(mailbox, $"Event {eventNum}", nextStart, 15, $"{transactionId}_{stepId}"));

                    nextStart = nextStart.AddMinutes(30);

                    eventNum++;
                }
            }

            var batches = requests.Batch(BatchSize).ToList();

            Log.Information("Sending {requests} requests in {batches} batches", requests.Count, batches.Count);

            int requestNum = 1;
            foreach (var batch in batches)
            {
                var batchRequest = new BatchRequestContentCollection(_client);
                foreach (var step in batch)
                {
                    await batchRequest.AddBatchRequestStepAsync(step.Value, step.Key);
                }

                Log.Information("Batch - {requestNum}", requestNum);

                var batchResponse = await _client
                    .Batch
                    .PostAsync(batchRequest, token);

                var statusCodes = await batchResponse.GetResponsesStatusCodesAsync();
                foreach (var statusCode in statusCodes)
                {
                    if (!IsSuccessStatusCode(statusCode.Value))
                    {
                        Log.Error("Request failed: {id} - {statusCode}", statusCode.Key, statusCode.Value);
                    }
                }
                requestNum++;
            }

            return eventNum;
        }

        /// <summary>
        /// Find sample events created based on transaction ID or all events if transaction ID is null or empty
        /// </summary>
        /// <param name="mailboxes"></param>
        /// <param name="transactionId"></param>
        /// <param name="token"></param>
        /// <returns></returns>
        public async Task<IDictionary<string, IList<Event>>> FindEvents(IEnumerable<string> mailboxes, string? transactionId, CancellationToken token = default)
        {
            if (mailboxes == null || !mailboxes.Any())
                return ImmutableDictionary<string, IList<Event>>.Empty;

            var batches = mailboxes.Batch(BatchSize);

            var events = new Dictionary<string, IList<Event>>();

            Log.Information("Sending {requests} requests in {batches} batches", mailboxes.Count(), batches.Count());

            foreach (var batch in batches)
            {
                var batchSteps = from mailbox in batch
                                 select new KeyValuePair<string, RequestInformation>
                                 (mailbox, BuildFindEventsRequest(mailbox));

                var batchRequest = new BatchRequestContentCollection(_client);
                foreach (var step in batchSteps)
                {
                    await batchRequest.AddBatchRequestStepAsync(step.Value, step.Key);
                }

                Log.Debug("Batch request: {@batchRequest}", batchRequest);

                var batchResponse = await _client
                    .Batch
                    .PostAsync(batchRequest, token);

                var statusCodes = await batchResponse.GetResponsesStatusCodesAsync();

                foreach (var statusCode in statusCodes)
                {
                    if (IsSuccessStatusCode(statusCode.Value))
                    {
                        try
                        {
                            var collectionResponse =
                                await batchResponse.GetResponseByIdAsync<EventCollectionResponse>(statusCode.Key);

                            if (collectionResponse.Value != null)
                            {
                                var responseEvents = collectionResponse.Value.Where(e =>
                                    string.IsNullOrEmpty(transactionId) || (e.TransactionId != null && e.TransactionId.StartsWith(transactionId)))
                                    .ToList();

                                // statusCode.Key should be the mailbox
                                if (responseEvents.Count != 0)
                                    events.Add(statusCode.Key, responseEvents);
                            }
                        }
                        catch (Exception ex)
                        {
                            Log.Error(ex, "Error with mailbox: {mailbox}: {message}", statusCode.Key, ex.Message);
                        }
                    }
                    else
                    {
                        Log.Error("Request failed: {id} - {statusCode}", statusCode.Key, statusCode.Value);
                    }
                }
            }

            return events;
        }

        /// <summary>
        /// Delete the specified events
        /// </summary>
        /// <param name="events"></param>
        /// <param name="token"></param>
        /// <returns></returns>
        public async Task DeleteEvents(IDictionary<string, IList<Event>> eventLists, CancellationToken token = default)
        {
            if (eventLists == null || !eventLists.Any())
                return;

            var requests = (from l in eventLists
                            from e in l.Value
                            where !string.IsNullOrEmpty(e.ICalUId) && !string.IsNullOrEmpty(e.Id)
                            select new KeyValuePair<string, RequestInformation>(e.ICalUId!, BuildDeleteEventRequest(l.Key, e.Id!)))
                            .ToList();

            // TODO: this may fail if there are more than 4 events in a mailbox due to concurrent limits per mailbox (4) as the batch will run
            // concurrently on the server.

            var batches = requests.Batch(BatchSize).ToList();

            Log.Information("Sending {requests} requests in {batches} batches", requests.Count, batches.Count);

            foreach (var batch in batches)
            {
                var batchRequest = new BatchRequestContentCollection(_client);
                foreach (var step in batch)
                {
                    await batchRequest.AddBatchRequestStepAsync(step.Value, step.Key);
                }

                var batchResponse = await _client
                    .Batch
                    .PostAsync(batchRequest, token);

                var statusCodes = await batchResponse.GetResponsesStatusCodesAsync();
                foreach (var statusCode in statusCodes)
                {
                    if (!IsSuccessStatusCode(statusCode.Value))
                    {
                        Log.Error("Request failed: {id} - {statusCode}", statusCode.Key, statusCode.Value);
                    }
                }
            }
        }


        private RequestInformation BuildCreateEventRequest(string mailbox, string subject, DateTime startUtc, int duration, string transactionId)
        {
            var newEvent = new Event
            {
                TransactionId = transactionId,
                Subject = subject,
                Start = new DateTimeTimeZone
                {
                    DateTime = startUtc.ToString("yyyy-MM-dd HH:mm:ss"),
                    TimeZone = "UTC"
                },
                End = new DateTimeTimeZone
                {
                    DateTime = startUtc.AddMinutes(duration).ToString("yyyy-MM-dd HH:mm:ss"),
                    TimeZone = "UTC"
                }
            };

            return _client
                .Users[mailbox]
                .Events
                .ToPostRequestInformation(newEvent, (config) =>
                {
                    config.Headers.Add("Prefer", TimeZoneUtcHeader);
                });
        }

        private RequestInformation BuildFindEventsRequest(string mailbox)
        {
            return _client
                .Users[mailbox]
                .Events
                .ToGetRequestInformation((config) =>
                {
                    config.Headers.Add("Prefer", TimeZoneUtcHeader);
                    config.QueryParameters.Top = 99999;
                    config.QueryParameters.Orderby = ["start/dateTime"];
                });
        }

        private RequestInformation BuildDeleteEventRequest(string mailbox, string eventId)
        {
            return _client
                .Users[mailbox]
                .Events[eventId]
                .ToDeleteRequestInformation();
        }

        public static bool IsSuccessStatusCode(HttpStatusCode statusCode) =>
            ((int)statusCode >= 200) && ((int)statusCode <= 299);
    }
}
