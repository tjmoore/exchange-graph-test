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

            int eventNum = 1;

            var rnd = new Random();

            var requests = new List<BatchRequestStep>();

            foreach (var mailbox in mailboxes)
            {
                DateTime start = DateTime.UtcNow;

                int numEvents = rnd.Next(0, maxEventsPerMailbox + 1);

                Console.WriteLine($"{mailbox} - Creating {numEvents} events");

                for (int n = 1; n <= numEvents; n++)
                {
                    requests.Add(new BatchRequestStep(mailbox, BuildCreateEventRequest(mailbox, $"Event {eventNum}", start, 15, transactionId)));
                    start = start.AddMinutes(30);
                    eventNum++;
                }
            }

            var batches = requests.Batch(BatchSize);

            foreach (var batch in batches)
            {
                var batchRequest = new BatchRequestContent(batch.ToArray());

                await _client
                    .Batch
                    .Request()
                    .PostAsync(batchRequest, token);
            }
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
                    var collectionPage = await batchResponse.GetResponseByIdAsync<UserEventsCollectionResponse>(mailbox);

                    List<Event> responseEvents = collectionPage.Value.Where(e => string.IsNullOrEmpty(transactionId) || e.TransactionId == transactionId).ToList();

                    if (responseEvents.Any())
                        events.Add(mailbox, responseEvents);
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

            // Deletes in a batch doesn't appear to work. It just deletes the first each time, so using a loop which will be slow.

            foreach (var l in eventLists)
            {
                foreach (var e in l.Value)
                {
                    await _client
                        .Users[l.Key]
                        .Events[e.Id]
                        .Request()
                        .DeleteAsync(token);
                }
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
    }
}
