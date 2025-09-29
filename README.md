# Outlook to Google Calendar Sync

A Google Apps Script that automatically syncs events from an Outlook calendar ICS feed to a Google Calendar every 30 minutes.

## Features

- ✅ Syncs events every 30 minutes automatically
- ✅ Handles recurring events (daily, weekly, monthly, yearly)
- ✅ Proper timezone conversion with DST support
- ✅ Avoids duplicate events
- ✅ One-way sync (removes deleted events)
- ✅ Syncs events up to 8 weeks ahead
- ✅ Comprehensive timezone mapping for Outlook timezones
- ✅ Optional event color customization
- ✅ Use default calendar or create a separate one

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
  // REQUIRED: Your Outlook calendar ICS feed URL
  ICS_FEED_URL: 'YOUR_OUTLOOK_ICS_URL_HERE',
  
  // OPTIONAL: Target calendar name
  // Set to null to use your default calendar, or specify a name like 'Work'
  TARGET_CALENDAR_NAME: 'Work',
  
  // OPTIONAL: Event color
  // Set to null for default color, or use one of:
  // CalendarApp.EventColor.PALE_BLUE, CalendarApp.EventColor.PALE_GREEN,
  // CalendarApp.EventColor.MAUVE, CalendarApp.EventColor.PALE_RED,
  // CalendarApp.EventColor.YELLOW, CalendarApp.EventColor.ORANGE,
  // CalendarApp.EventColor.CYAN, CalendarApp.EventColor.GRAY,
  // CalendarApp.EventColor.BLUE, CalendarApp.EventColor.GREEN,
  // CalendarApp.EventColor.RED
  EVENT_COLOR: null,
  
  // How many weeks ahead to sync
  SYNC_WEEKS_AHEAD: 8
};
\`\`\`

**Configuration Examples:**

\`\`\`javascript
// Use default calendar with blue events
TARGET_CALENDAR_NAME: null,
EVENT_COLOR: CalendarApp.EventColor.BLUE,

// Use separate 'Work' calendar with default color
TARGET_CALENDAR_NAME: 'Work',
EVENT_COLOR: null,

// Use 'Personal' calendar with green events
TARGET_CALENDAR_NAME: 'Personal',
EVENT_COLOR: CalendarApp.EventColor.GREEN,
\`\`\`

### 4. Create Target Calendar (Optional)

If you specified a `TARGET_CALENDAR_NAME` (not null), the script will automatically create it if it doesn't exist. You can also create it manually:

1. Open Google Calendar
2. Click the "+" next to "Other calendars"
3. Select "Create new calendar"
4. Name it to match your `TARGET_CALENDAR_NAME`

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
   - Events added/deleted
   - Recurring events expanded
   - Any errors or warnings

## How It Works

### One-Way Sync
- Events are synced from Outlook → Google Calendar
- Events deleted from Outlook are automatically removed from Google Calendar
- All synced events are tagged with `[Synced from Outlook]` in their description

### Recurring Events
- Recurring events are expanded into individual occurrences
- Only future occurrences within the sync window are created
- Past occurrences are not created

### Duplicate Prevention
- Events are tracked by their unique UID from Outlook
- Duplicate events are automatically detected and removed

## Troubleshooting

### Events showing at wrong times

- Check that your timezone is correctly set in Google Calendar settings
- Verify the Outlook calendar is publishing with proper timezone info

### Missing events

- Check the logs for "Skipped" or "ERROR" messages
- Recurring events that have ended are skipped
- Events beyond the sync window (8 weeks) are skipped

### Duplicate events

- The script automatically detects and removes duplicates
- If you see duplicates, run the sync again - it will clean them up

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
