# exchange-graph-test

Simple test tool for Microsoft Graph API for populating Exchange mailboxes (e.g. room mailboxes) with calendar events. It illustrates fetching, creating and deleting events and use of batching in Graph API also.

Components used:

* .NET 5
* Microsoft Graph SDK
* System.CommandLine

## Detail

This is a specific use case and may not be of use to others directly but feel free to fork and customise or just use as an example.

This is based on a test set up with test mailboxes with the same name pattern differing by number.

e.g. testroom1@mydomain, testroom2@mydomain, etc

`--mailboxTemplate` must be supplied as a format string, e.g. `testroom{0}@mydomain` for it to populate the mailbox addresses to use.

`--numMailbox` specifies the number of mailboxes to iterrate and `--startMailbox` the initial number (default 1)

A transactionId is used per event to track for testing and to allow the `delete` command to find the events. `--transactionId` allows specifying of an ID. This is used as a prefix, each event will have a unique ID appended. On delete it searches for the prefix in each event.

`create` does 10 runs through all mailboxes incrementing date every 2 hours and creates a random number of events per mailbox. Due to concurrency limits per mailbox (4 requests) it will only create up to 4 events per mailbox in a run. `--maxEvents` controls how many events to create per mailbox (a value greater than 4 will result in a max of 4).

`get` fetches events for the mailboxes with a summary count of events, or with `--dumpEvents` it will trace the event responses. If `--transactionId` is specified it filters to only those with that prefix in TransactionId.

`delete` deletes the events based on `--transactionId` used as prefix for the TransactionId.

## Azure AD Application requirements

This requires an application created and configured in Azure AD with application permissions for Graph API for calendar access. See https://docs.microsoft.com/en-us/graph/auth-v2-service

`Calendars.ReadWrite` is the main permissions required and admin consent will be required for the application permission. It is also possible to scope the permissions to specific mailboxes using ApplicationAccessPolicy. https://docs.microsoft.com/en-us/graph/auth-limit-mailbox-access


## Usage

```
ExchangeGraphTool
  Exchange Graph API test tool for creating bulk events

Usage:
  ExchangeGraphTool [options] [command]

Options:
  --clientId <clientId>                Graph API Client ID
  --tenantId <tenantId>                Graph API Tenant ID
  --clientSecret <clientSecret>        Graph API Client Secret
  --mailboxTemplate <mailboxTemplate>  Mailbox address template (format <name>{0}@<domain>)
  --numMailbox <numMailbox>            Number of mailboxes to use in template
  --startMailbox <startMailbox>        Start number of mailboxes to use in template, default one
  --version                            Show version information
  -?, -h, --help                       Show help and usage information

Commands:
  get     Fetches events matching specified transaction ID, or all events if not specified
  Options:
    --transactionId                    Use specified ID as prefix for transaction ID on events or return all events otherwise
    --dumpEvents                       Dump event detail
  create  Creates sample events
  Options:
    --maxEvents                        Max number of events per mailbox, default 1
    --transactionId                    Use specified ID as prefix for transaction ID on events, otherwise generates a new GUID
  delete  Deletes events matching specified transaction ID
  Options:
    --transactionId                    Use specified ID as prefix for transaction ID to match events to delete
 ```
 
 
