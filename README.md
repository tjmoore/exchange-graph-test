# exchange-graph-test

Simple example and test tool for Microsoft Graph API for various operations

At present this is focused on a specific use case of bulk creating calendar events in multiple mailboxes.
It illustrates fetching, creating and deleting events and use of batching in Graph API also.

Components used:

* .NET 8
* Microsoft Graph SDK
* System.CommandLine

## Detail

This has specific use cases and may not be of use to others directly but feel free to fork and customise or just use as an example.

This is based on a test set up with test mailboxes with the same name pattern differing by number.

e.g. testroom1@mydomain, testroom2@mydomain, etc

`--mailbox-template` must be supplied as a format string, e.g. `testroom{0}@mydomain` for it to populate the mailbox addresses to use.

`--num-mailbox` specifies the number of mailboxes to iterrate and `--start-mailbox` the initial number (default 1)

A transactionId is used per event to track for testing and to allow the `delete` command to find the events. `--transaction-id` allows specifying of an ID. This is used as a prefix, each event will have a unique ID appended. On delete it searches for the prefix in each event.

`create-events` does 10 runs through all mailboxes incrementing date every 2 hours and creates a random number of events per mailbox. Due to concurrency limits per mailbox (4 requests) it will only create up to 4 events per mailbox in a run. `--max-events` controls how many events to create per mailbox (a value greater than 4 will result in a max of 4).

`get-events` fetches events for the mailboxes with a summary count of events, or with `--dump-events` it will trace the event responses. If `--transaction-id` is specified it filters to only those with that prefix in TransactionId.

`delete-events` deletes the events based on `--transaction-id` used as prefix for the TransactionId.

For example:

Create events in first 10 mailboxes with max 10 events per mailbox and transaction ID "test123":
```powershell
create-events --mailbox-template "testroom{0}@mydomain" --num-mailbox 10 --max-events 10 --transaction-id "test123" --client-id "<client-id>" --tenant-id "<tenant-id>" --client-secret "<client-secret>"
```

Fetch events in first 10 mailboxes with transaction ID "test123" and dump event details:
```powershell
get-events --mailbox-template "testroom{0}@mydomain" --num-mailbox 10 --transaction-id "test123" --dump-events --client-id "<client-id>" --tenant-id "<tenant-id>" --client-secret "<client-secret>"
```

Delete events in first 10 mailboxes with transaction ID "test123":
```powershell
delete-events --mailbox-template "testroom{0}@mydomain" --num-mailbox 10 --transaction-id "test123" --client-id "<client-id>" --tenant-id "<tenant-id>" --client-secret "<client-secret>"
```



## Azure AD Application requirements

This requires an application created and configured in Azure AD with application permissions for Graph API for calendar access. See https://docs.microsoft.com/en-us/graph/auth-v2-service

`Calendars.ReadWrite` is the main permissions required and admin consent will be required for the application permission. It is also possible to scope the permissions to specific mailboxes using ApplicationAccessPolicy. https://docs.microsoft.com/en-us/graph/auth-limit-mailbox-access


## Usage

```
ExchangeGraphTool
  Exchange Graph API test tool

Usage:
  ExchangeGraphTool [options] [command]

Options:
  --client-id <clientId>                Graph API Client ID
  --tenant-id <tenantId>                Graph API Tenant ID
  --client-secret <clientSecret>        Graph API Client Secret
  --version                             Show version information
  -?, -h, --help                        Show help and usage information

Commands:
  get-events     Fetches events matching specified transaction ID, or all events if not specified
  Options:
    --mailbox-template <mailboxTemplate>  Mailbox address template (format <name>{0}@<domain>)
    --num-mailbox <numMailbox>            Number of mailboxes to use in template
    --start-mailbox <startMailbox>        Start number of mailboxes to use in template, default one
    --transaction-id                      Use specified ID as prefix for transaction ID on events or return all events otherwise
    --dump-events                         Dump event detail

  create-events  Creates sample events
  Options:
    --mailbox-template <mailboxTemplate>  Mailbox address template (format <name>{0}@<domain>)
    --num-mailbox <numMailbox>            Number of mailboxes to use in template
    --start-mailbox <startMailbox>        Start number of mailboxes to use in template, default one
    --max-events                          Max number of events per mailbox, default 1
    --transaction-id                      Use specified ID as prefix for transaction ID on events, otherwise generates a new GUID
  delete-events  Deletes events matching specified transaction ID

  Options:
    --mailbox-template <mailboxTemplate>  Mailbox address template (format <name>{0}@<domain>)
    --num-mailbox <numMailbox>            Number of mailboxes to use in template
    --start-mailbox <startMailbox>        Start number of mailboxes to use in template, default one
    --transaction-id                      Use specified ID as prefix for transaction ID to match events to delete
 ```
 
 
