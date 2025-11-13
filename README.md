# Outlook Calendar Sync for Google Calendar

A free, open-source Google Apps Script that solves the frustrating limitations of Outlook calendar feeds in Google Calendar.

## The Problem This Solves

If you've tried syncing your Outlook calendar to Google Calendar using the built-in ICS feed subscription, you've likely encountered these issues:

- **24-hour sync delays** - Google Calendar only updates subscribed calendars once every 24 hours, making them nearly useless for real-time scheduling
- **Wrong timezones** - Events appear hours off due to Outlook's non-standard timezone identifiers that Google Calendar can't interpret correctly
- **Missing recurring events** - Many recurring events simply don't show up, especially if they started in the past
- **No deletion sync** - Events deleted from Outlook remain in Google Calendar indefinitely
- **Incomplete syncs** - Only some events appear with no clear pattern
- **No control** - Can't force a refresh or customize sync behavior

This script provides a **true one-way sync** that runs automatically every 30 minutes, handles timezones correctly, expands recurring events properly, and removes deleted events.

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

At the top of Code.gs, update the configuration:

    const CONFIG = {
      ICS_FEED_URL: 'YOUR_OUTLOOK_ICS_URL_HERE',
      TARGET_CALENDAR_NAME: 'Work',
      TARGET_TIMEZONE: null,
      EVENT_COLOR: null,
      SYNC_WEEKS_AHEAD: 8,
      SYNC_MARKER: '[Synced from Outlook]'
    };

**Configuration options:**

- **ICS_FEED_URL** (required): Your Outlook calendar ICS feed URL
- **TARGET_CALENDAR_NAME** (optional): Set to null to use your default calendar, or specify a name like 'Work'
- **TARGET_TIMEZONE** (optional): Set to null to use the calendar's default timezone, or use a valid IANA timezone identifier like "America/Los_Angeles", "America/New_York", "Europe/London", "Asia/Tokyo". Full list: [List of tz database time zones](https://en.wikipedia.org/wiki/List_of_tz_database_time_zones)
- **EVENT_COLOR** (optional): Set to null for default color, or use a color name like "BLUE", "GREEN", "RED", "PALE_BLUE", "PALE_GREEN", "MAUVE", "PALE_RED", "YELLOW", "ORANGE", "CYAN", "GRAY"
- **SYNC_WEEKS_AHEAD**: How many weeks ahead to sync (default: 8)
- **SYNC_MARKER**: Marker to identify synced events in the description (default: '[Synced from Outlook]')

**Configuration examples:**

Use default calendar with blue events:

    TARGET_CALENDAR_NAME: null,
    EVENT_COLOR: "BLUE",

Use separate 'Work' calendar with default color:

    TARGET_CALENDAR_NAME: 'Work',
    EVENT_COLOR: null,

Use 'Personal' calendar with green events:

    TARGET_CALENDAR_NAME: 'Personal',
    EVENT_COLOR: "GREEN",

### 4. Create Target Calendar (Optional)

If you specified a TARGET_CALENDAR_NAME (not null), the script will automatically create it if it doesn't exist. You can also create it manually:

1. Open Google Calendar
2. Click the "+" next to "Other calendars"
3. Select "Create new calendar"
4. Name it to match your TARGET_CALENDAR_NAME

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

In setupTrigger(), modify the line:

    .everyMinutes(30)

Change to 15, 30, or 60 minutes

### Change sync window

In the CONFIG object, change:

    SYNC_WEEKS_AHEAD: 8

Set to desired number of weeks

### Add more timezones

If you encounter "Unknown timezone" warnings in logs, add mappings to the TIMEZONE_MAP object in the format:

    'Outlook Timezone Name': 'IANA/Timezone/Identifier'

## License

MIT License - Feel free to use and modify as needed.

## Copying the Code

For easy access, here's the full `Code.gs` script you can copy directly:

    // ============================================================================
    // CONFIGURATION
    // ============================================================================
    
    const CONFIG = {
      // IMPORTANT: Replace this with your Outlook calendar ICS feed URL
      ICS_FEED_URL: 'YOUR_OUTLOOK_ICS_URL_HERE',
      
      // Calendar Configuration
      // Set to null to use your default/primary calendar, or specify a calendar name
      // Examples: 'Work', 'Personal', null (for default calendar)
      TARGET_CALENDAR_NAME: 'Work',
      
      // Target timezone for the Google Calendar
      // Set to null to use the calendar's default timezone
      // Examples: 'America/Los_Angeles', 'America/New_York', 'Europe/London', 'Asia/Tokyo'
      // Full list: https://en.wikipedia.org/wiki/List_of_tz_database_time_zones
      TARGET_TIMEZONE: null,
      
      // Event Color Configuration
      // Set to null to keep default calendar color, or use one of these color names:
      // "PALE_BLUE", "PALE_GREEN", "MAUVE", "PALE_RED", "YELLOW", "ORANGE",
      // "CYAN", "GRAY", "BLUE", "GREEN", "RED"
      // Example: "BLUE" or null (default)
      EVENT_COLOR: null,
      
      // How many weeks ahead to sync events
      SYNC_WEEKS_AHEAD: 8,
      
      SYNC_MARKER: '[Synced from Outlook]'
    };
    
    // ============================================================================
    // MAIN SYNC FUNCTION
    // ============================================================================
    
    function syncCalendarEvents() {
      try {
        // Validate configuration
        if (CONFIG.ICS_FEED_URL === 'YOUR_OUTLOOK_ICS_URL_HERE') {
          throw new Error('Please configure your ICS_FEED_URL in the CONFIG object at the top of the script.');
        }
        
        // Fetch the ICS feed
        Logger.log('Fetching ICS feed...');
        const response = UrlFetchApp.fetch(CONFIG.ICS_FEED_URL);
        const icsData = response.getContentText();
        Logger.log(`Fetched ${icsData.length} characters from ICS feed`);
        
        // Parse ICS data
        const events = parseICS(icsData);
        Logger.log(`Successfully parsed ${events.length} total events from ICS feed`);
        
        let calendar;
        if (CONFIG.TARGET_CALENDAR_NAME === null) {
          // Use default/primary calendar
          calendar = CalendarApp.getDefaultCalendar();
          Logger.log('Using default calendar');
        } else {
          // Use named calendar, create if doesn't exist
          const calendars = CalendarApp.getCalendarsByName(CONFIG.TARGET_CALENDAR_NAME);
          if (calendars.length === 0) {
            Logger.log(`Calendar "${CONFIG.TARGET_CALENDAR_NAME}" not found. Creating it...`);
            calendar = CalendarApp.createCalendar(CONFIG.TARGET_CALENDAR_NAME);
            Logger.log(`Created calendar "${CONFIG.TARGET_CALENDAR_NAME}"`);
          } else {
            calendar = calendars[0];
          }
        }
        
        const now = new Date();
        const oneWeekAgo = new Date(now.getTime() - (7 * 24 * 60 * 60 * 1000));
        const futureDate = new Date(now.getTime() + (CONFIG.SYNC_WEEKS_AHEAD * 7 * 24 * 60 * 60 * 1000));
        Logger.log(`Syncing events from ${now.toISOString()} to ${futureDate.toISOString()}`);
        
        const existingEvents = calendar.getEvents(oneWeekAgo, futureDate);
        const syncedEvents = existingEvents.filter(event => {
          const desc = event.getDescription();
          return desc && desc.includes(CONFIG.SYNC_MARKER);
        });
        Logger.log(`Found ${syncedEvents.length} existing synced events in target calendar`);
        
        const expandedEvents = [];
        events.forEach(event => {
          if (event.recurrence && event.rruleString) {
            // Expand recurring event into individual occurrences
            const occurrences = expandRecurringEvent(event, now, futureDate);
            expandedEvents.push(...occurrences);
            Logger.log(`Expanded recurring event "${event.title}" into ${occurrences.length} occurrences`);
          } else {
            // Regular one-time event
            if (event.endTime >= now && event.startTime <= futureDate) {
              expandedEvents.push(event);
            }
          }
        });
        
        Logger.log(`Total events after expansion: ${expandedEvents.length}`);
        
        const feedEventKeys = new Set(); // Events from the feed
        const existingEventMap = {}; // Existing events by key
        
        syncedEvents.forEach(event => {
          const desc = event.getDescription();
          const uidMatch = desc ? desc.match(/UID:([^\n]+)/) : null;
          
          let key;
          if (uidMatch) {
            // Trim whitespace from UID
            const uid = uidMatch[1].trim();
            key = `${uid}_${event.getStartTime().getTime()}`;
          } else {
            // Fallback to title+time key
            key = generateEventKey(event.getTitle(), event.getStartTime(), event.getEndTime());
          }
          
          if (existingEventMap[key]) {
            Logger.log(`WARNING: Found duplicate existing event with key ${key}: "${event.getTitle()}" at ${event.getStartTime().toISOString()}`);
            // Keep track of duplicate for potential cleanup
            if (!existingEventMap[key].duplicates) {
              existingEventMap[key].duplicates = [];
            }
            existingEventMap[key].duplicates.push(event);
          } else {
            existingEventMap[key] = { event: event, duplicates: [] };
          }
        });
        
        const calendarTimezone = calendar.getTimeZone();
        Logger.log(`Target calendar timezone: ${calendarTimezone}`);
        if (CONFIG.TARGET_TIMEZONE) {
          Logger.log(`Config TARGET_TIMEZONE: ${CONFIG.TARGET_TIMEZONE}`);
          if (CONFIG.TARGET_TIMEZONE !== calendarTimezone) {
            Logger.log(`WARNING: TARGET_TIMEZONE (${CONFIG.TARGET_TIMEZONE}) differs from calendar timezone (${calendarTimezone})`);
          }
        }
        
        // Add events from feed
        let addedCount = 0;
        let skippedDuplicate = 0;
        let errorCount = 0;
        
        expandedEvents.forEach((feedEvent, index) => {
          // Generate key for this occurrence
          let key;
          if (feedEvent.uid) {
            const uid = feedEvent.uid.trim();
            key = `${uid}_${feedEvent.startTime.getTime()}`;
          } else {
            key = generateEventKey(feedEvent.title, feedEvent.startTime, feedEvent.endTime);
          }
          
          feedEventKeys.add(key);
          
          // Skip if event already exists
          if (existingEventMap[key]) {
            skippedDuplicate++;
            return;
          }
          
          let description = '';
          if (feedEvent.description) {
            description = feedEvent.description;
          }
          
          if (feedEvent.uid) {
            const uid = feedEvent.uid.trim();
            description = description ? `${description}\n\nUID:${uid}\n${CONFIG.SYNC_MARKER}` : `UID:${uid}\n${CONFIG.SYNC_MARKER}`;
          } else {
            description = description ? `${description}\n\n${CONFIG.SYNC_MARKER}` : CONFIG.SYNC_MARKER;
          }
          
          Logger.log(`\n--- Creating Event ---`);
          Logger.log(`Title: "${feedEvent.title}"`);
          Logger.log(`Start (UTC): ${feedEvent.startTime.toISOString()}`);
          Logger.log(`Start (Local): ${feedEvent.startTime.toString()}`);
          Logger.log(`Start (Timestamp): ${feedEvent.startTime.getTime()}`);
          if (!feedEvent.isAllDay) {
            Logger.log(`End (UTC): ${feedEvent.endTime.toISOString()}`);
            Logger.log(`End (Local): ${feedEvent.endTime.toString()}`);
          }
          Logger.log(`Is All Day: ${feedEvent.isAllDay}`);
          Logger.log(`UID: ${feedEvent.uid || 'none'}`);
          Logger.log(`Key: ${key}`);
          
          // Create the event
          try {
            let newEvent;
            if (feedEvent.isAllDay) {
              newEvent = calendar.createAllDayEvent(feedEvent.title, feedEvent.startTime, {
                description: description,
                location: feedEvent.location
              });
            } else {
              newEvent = calendar.createEvent(feedEvent.title, feedEvent.startTime, feedEvent.endTime, {
                description: description,
                location: feedEvent.location
              });
            }
            
            if (CONFIG.EVENT_COLOR !== null) {
              const colorKey = CONFIG.EVENT_COLOR.toUpperCase();
              const eventColor = COLOR_MAP[colorKey];
              if (eventColor) {
                newEvent.setColor(eventColor);
              } else {
                Logger.log(`WARNING: Unknown color name "${CONFIG.EVENT_COLOR}". Valid options: ${Object.keys(COLOR_MAP).join(', ')}`);
              }
            }
            
            Logger.log(`✓ Created event at: ${newEvent.getStartTime().toISOString()} (${newEvent.getStartTime().toString()})`);
            addedCount++;
          } catch (e) {
            Logger.log(`ERROR creating event "${feedEvent.title}": ${e.toString()}`);
            errorCount++;
          }
        });
        
        let deletedCount = 0;
        Object.keys(existingEventMap).forEach(key => {
          if (!feedEventKeys.has(key)) {
            const eventData = existingEventMap[key];
            
            // Delete the main event
            try {
              eventData.event.deleteEvent();
              Logger.log(`Deleted event no longer in feed: "${eventData.event.getTitle()}" at ${eventData.event.getStartTime().toISOString()} (key: ${key})`);
              deletedCount++;
            } catch (e) {
              Logger.log(`ERROR deleting event "${eventData.event.getTitle()}": ${e.toString()}`);
            }
            
            // Delete any duplicates
            if (eventData.duplicates && eventData.duplicates.length > 0) {
              eventData.duplicates.forEach(dupEvent => {
                try {
                  dupEvent.deleteEvent();
                  Logger.log(`Deleted duplicate event: "${dupEvent.getTitle()}" at ${dupEvent.getStartTime().toISOString()}`);
                  deletedCount++;
                } catch (e) {
                  Logger.log(`ERROR deleting duplicate event: ${dupEvent.getTitle()}": ${e.toString()}`);
                }
              });
            }
          } else {
            // Event is in feed, but delete any duplicates
            const eventData = existingEventMap[key];
            if (eventData.duplicates && eventData.duplicates.length > 0) {
              Logger.log(`Found ${eventData.duplicates.length} duplicate(s) for event "${eventData.event.getTitle()}" - cleaning up`);
              eventData.duplicates.forEach(dupEvent => {
                try {
                  dupEvent.deleteEvent();
                  Logger.log(`Deleted duplicate event: "${dupEvent.getTitle()}" at ${dupEvent.getStartTime().toISOString()}`);
                  deletedCount++;
                } catch (e) {
                  Logger.log(`ERROR deleting duplicate event: ${dupEvent.getTitle()}": ${e.toString()}`);
                }
              });
            }
          }
        });
        
        Logger.log('\n=== SECOND PASS: Cleaning up same-day DST duplicates ===');
        const dupDeleteCount = cleanupSameDayDuplicates(calendar, CONFIG.SYNC_MARKER);
        deletedCount += dupDeleteCount;
    
        Logger.log('\n=== SYNC SUMMARY ===');
        Logger.log(`Total events in feed: ${events.length}`);
        Logger.log(`Total occurrences after expansion: ${expandedEvents.length}`);
        Logger.log(`Added: ${addedCount}`);
        Logger.log(`Deleted (no longer in feed or duplicates): ${deletedCount}`);
        Logger.log(`Skipped (duplicates): ${skippedDuplicate}`);
        Logger.log(`Errors: ${errorCount}`);
        
      } catch (error) {
        Logger.log(`FATAL ERROR syncing calendar: ${error.toString()}`);
      }
    }
    
    // ... Rest of Code.gs continues - see Code.gs file for complete implementation

For the full, detailed code including all helper functions, timezone mappings, and the complete implementation, please copy directly from the `Code.gs` file in this repository.
