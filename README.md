# Google Calendar to Notion Connector

This Google Workspace Add-on allows you to easily send Google Calendar events to a Notion database. It checks if an event already exists in your Notion database and provides a direct link, or allows you to create a new page in Notion with a single click.

## Overview

The script `onDefaultHomePageOpen.js` implements the core logic for the add-on. It interacts with the Google Calendar API to fetch event details and the Notion API to query and create pages.

## Features

- **Event Details View**: Displays the event title, start time, and end time in the add-on sidebar.
- **Duplicate Check**: Automatically queries the configured Notion database to see if the current calendar event has already been exported (based on the Event ID).
- **Send to Notion**: If the event is not in Notion, a "SEND TO NOTION" button is displayed. Clicking it creates a new page in the Notion database with the event details.
- **Open in Notion**: If the event is already linked, an "OPEN IN NOTION" button is displayed to directly open the corresponding Notion page.
- **Seamless UI Update**: After sending an event to Notion, the card automatically updates to show the "OPEN IN NOTION" button without requiring a refresh.

## Specifications

### Entry Points
- **`onHomepageOpen(e)`**: Default homepage handler. Shows a welcome message.
- **`onEventOpen(e)`**: Event trigger handler. Delegates to `onDefaultHomePageOpen`.

### Notion Integration
The script maps Google Calendar fields to Notion properties as follows:

| Google Calendar Field | Notion Property Name | Notion Property Type |
|-----------------------|----------------------|----------------------|
| Event Title           | `Name`               | Title                |
| Start & End Time      | `Date`               | Date                 |
| Creators              | `Organizer`          | Rich Text            |
| Location              | `Location`           | Rich Text            |
| Event ID              | `ID`                 | Rich Text            |

### Required Configuration (Script Properties)
The following script properties must be set in the Google Apps Script project:

- **`databaseId`**: The ID of the Notion database where events will be stored.
- **`notionAPIToken`**: The Notion Internal Integration Token (Secret).

### Dependencies
- **CalendarApp**: Used to fetch detailed event information.
- **UrlFetchApp**: Used to communicate with the Notion API (v1).
- **PropertiesService**: Used to securely retrieve API credentials.
