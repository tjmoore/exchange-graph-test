using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.Linq;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;

namespace ExchangeGraphTool
{
    public class Calendar
    {
        private const int MaxBatchSize = 20; // Batch max size = 20 at present with Graph API

        private int _batchSize = MaxBatchSize;

        private readonly GraphServiceClient _client;

        public int BatchSize
        {
            get => _batchSize;
            set => Math.Min(value, MaxBatchSize);
        }

        public Calendar(GraphServiceClient client)
        {
            _client = client;
        }

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
                Console.WriteLine($"Create events in {mailboxes.Count()} mailboxes starting from {start}");

                eventNum = await CreateEvents(mailboxes, start, maxEventsPerMailbox, transactionId, eventNum, token);
                start = start.AddHours(2);
            }
        }

        private async Task<int> CreateEvents(IEnumerable<string> mailboxes, DateTime start, int maxEventsPerMailbox, string transactionId, int eventNum, CancellationToken token)
        {
            var rnd = new Random();

            var requests = new List<BatchRequestStep>();

            foreach (var mailbox in mailboxes)
            {
                DateTime nextStart = start;

                int numEvents = rnd.Next(0, maxEventsPerMailbox + 1);

                //Console.WriteLine($"{mailbox} - Creating {numEvents} events");

                for (int n = 1; n <= numEvents; n++)
                {
                    string stepId = Guid.NewGuid().ToString();
                    requests.Add(new BatchRequestStep(stepId, BuildCreateEventRequest(mailbox, $"Event {eventNum}", nextStart, 15, $"{transactionId}_{stepId}")));

                    nextStart = nextStart.AddMinutes(30);

                    eventNum++;
                }
            }

            var batches = requests.Batch(BatchSize).ToList();

            Console.WriteLine($"Sending {requests.Count} requests in {batches.Count} batches");

            int requestNum = 1;
            foreach (var batch in batches)
            {
                var batchRequest = new BatchRequestContent(batch.ToArray());
                Console.WriteLine($"Batch - {requestNum}");

                var batchResponse = await _client
                    .Batch
                    .Request()
                    .PostAsync(batchRequest, token);

                var responses = await batchResponse.GetResponsesAsync();
                foreach (var response in responses)
                {
                    if (!response.Value.IsSuccessStatusCode)
                    {
                        var request = batchRequest.BatchRequestSteps[response.Key].Request;

                        string requestUri = request.RequestUri.ToString();
                        string requestContent = await request.Content.ReadAsStringAsync();
                        string content = await response.Value.Content.ReadAsStringAsync();
                        Console.WriteLine($"Request failed: {requestUri} - {requestContent}{Environment.NewLine}{(int)response.Value.StatusCode} : {response.Value.ReasonPhrase} - {content}");
                    }
                }
                requestNum++;
            }

            return eventNum;
        }

        /// <summary>
        /// Find sample events created based on transaction ID
        /// </summary>
        /// <param name="mailboxes"></param>
        /// <param name="token"></param>
        /// <returns></returns>
        public async Task<IDictionary<string, IList<Event>>> FindEvents(IEnumerable<string> mailboxes, string transactionId, CancellationToken token = default)
        {
            if (mailboxes == null || !mailboxes.Any())
                return ImmutableDictionary<string, IList<Event>>.Empty;

            var batches = mailboxes.Batch(BatchSize);

            var events = new Dictionary<string, IList<Event>>();

            foreach (var batch in batches)
            {
                var batchSteps = from mailbox in batch
                                 select new BatchRequestStep(mailbox, BuildFindEventsRequest(mailbox).GetHttpRequestMessage());

                var batchRequest = new BatchRequestContent(batchSteps.ToArray());

                var batchResponse = await _client
                    .Batch
                    .Request()
                    .PostAsync(batchRequest, token);

                foreach (var mailbox in batch)
                {
                    try
                    {
                        var collectionPage = await batchResponse.GetResponseByIdAsync<UserEventsCollectionResponse>(mailbox);

                        List<Event> responseEvents = collectionPage.Value.Where(e => string.IsNullOrEmpty(transactionId) || (e.TransactionId != null && e.TransactionId.StartsWith(transactionId))).ToList();

                        if (responseEvents.Any())
                            events.Add(mailbox, responseEvents);
                    }
                    catch (ServiceException ex)
                    {
                        Console.WriteLine($"Error with mailbox: {mailbox}: {ex.Message}");
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
                            select new BatchRequestStep(e.ICalUId, BuildDeleteEventRequest(l.Key, e.Id))).ToList();

            // TODO: this may fail if there are more than 4 events in a mailbox due to concurrent limits per mailbox (4) as the batch will run
            // concurrently on the server.

            var batches = requests.Batch(BatchSize).ToList();

            Console.WriteLine($"Sending {requests.Count} requests in {batches.Count} batches");

            foreach (var batch in batches)
            {
                var batchRequest = new BatchRequestContent(batch.ToArray());

                await _client
                    .Batch
                    .Request()
                    .PostAsync(batchRequest, token);
            }
        }


        private HttpRequestMessage BuildCreateEventRequest(string mailbox, string subject, DateTime startUtc, int duration, string transactionId)
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

            var jsonEvent = _client.HttpProvider.Serializer.SerializeAsJsonContent(newEvent);

            var addEventRequest = _client
                .Users[mailbox]
                .Events
                .Request()
                .Header("Prefer", $"outlook.timezone=\"UTC\"")
                .GetHttpRequestMessage();

            addEventRequest.Method = HttpMethod.Post;
            addEventRequest.Content = jsonEvent;

            return addEventRequest;
        }

        private IUserEventsCollectionRequest BuildFindEventsRequest(string mailbox)
        {
            return _client
                .Users[mailbox]
                .Events
                .Request()
                .Top(99999)
                .OrderBy("start/dateTime");
        }

        private HttpRequestMessage BuildDeleteEventRequest(string mailbox, string eventId)
        {
            var deleteEventRequest = _client
                .Users[mailbox]
                .Events[eventId]
                .Request()
                .GetHttpRequestMessage();

            deleteEventRequest.Method = HttpMethod.Delete;

            return deleteEventRequest;
        }
    }
}
