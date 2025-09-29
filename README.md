# Outlook to Google Calendar Sync

A Google Apps Script that automatically syncs events from an Outlook calendar ICS feed to a Google Calendar every 30 minutes.

## Features

- ✅ Syncs events every 30 minutes automatically
- ✅ Handles recurring events (daily, weekly, monthly, yearly)
- ✅ Proper timezone conversion with DST support
- ✅ Avoids duplicate events
- ✅ Syncs events up to 8 weeks ahead
- ✅ Comprehensive timezone mapping for Outlook timezones

## Setup Instructions

### 1. Get Your Outlook Calendar ICS Feed URL

1. Open Outlook Calendar (web version)
2. Go to Settings → View all Outlook settings → Calendar → Shared calendars
3. Publish your calendar and copy the ICS link

### 2. Create Google Apps Script Project

1. Go to [script.google.com](https://script.google.com)
2. Click "New Project"
3. Copy the contents of `Code.gs` into the script editor
4. Save the project (give it a name like "Outlook Calendar Sync")

### 3. Configure the Script

At the top of `Code.gs`, update the configuration:

\`\`\`javascript
const CONFIG = {
  ICS_FEED_URL: 'YOUR_OUTLOOK_ICS_URL_HERE',
  TARGET_CALENDAR_NAME: 'Stryker',  // Change to your target calendar name
  SYNC_WEEKS_AHEAD: 8
};
\`\`\`

### 4. Create Target Calendar (if needed)

1. Open Google Calendar
2. Click the "+" next to "Other calendars"
3. Select "Create new calendar"
4. Name it (e.g., "Stryker" or whatever you set in CONFIG)

### 5. Authorize and Test

1. In the Apps Script editor, select `syncCalendarEvents` from the function dropdown
2. Click "Run" (▶️)
3. Authorize the script when prompted:
   - Click "Review Permissions"
   - Select your Google account
   - Click "Advanced" → "Go to [Project Name] (unsafe)"
   - Click "Allow"
4. Check the logs (View → Logs) to see the sync results

### 6. Set Up Automatic Sync

1. In the Apps Script editor, select `setupTrigger` from the function dropdown
2. Click "Run" (▶️)
3. This creates a trigger to run every 30 minutes automatically

## Viewing Logs

To see what's happening during syncs:

1. In Apps Script editor, go to "Executions" (left sidebar)
2. Click on any execution to see detailed logs
3. Look for:
   - Total events parsed
   - Events added
   - Recurring events found
   - Any errors or warnings

## Troubleshooting

### Events showing at wrong times

- Check that your timezone is correctly set in Google Calendar settings
- Verify the Outlook calendar is publishing with proper timezone info

### Missing events

- Check the logs for "Skipped" or "ERROR" messages
- Recurring events older than 6 months are skipped by default
- Events beyond 8 weeks in the future are skipped

### Duplicate events

- The script checks for duplicates based on title + start time + end time
- If you see duplicates, they might have slightly different times or titles

## Customization

### Change sync frequency

In `setupTrigger()`, modify:
\`\`\`javascript
.everyMinutes(30)  // Change to 15, 30, or 60
\`\`\`

### Change sync window

In the CONFIG object:
\`\`\`javascript
SYNC_WEEKS_AHEAD: 8  // Change to desired number of weeks
\`\`\`

### Add more timezones

If you encounter "Unknown timezone" warnings in logs, add mappings to the `TIMEZONE_MAP` object in the format:
\`\`\`javascript
'Outlook Timezone Name': 'IANA/Timezone/Identifier'
\`\`\`

## License

MIT License - Feel free to use and modify as needed.
