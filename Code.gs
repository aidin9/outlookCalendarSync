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
// COLOR MAPPING
// ============================================================================

const COLOR_MAP = {
  'PALE_BLUE': CalendarApp.EventColor.PALE_BLUE,
  'PALE_GREEN': CalendarApp.EventColor.PALE_GREEN,
  'MAUVE': CalendarApp.EventColor.MAUVE,
  'PALE_RED': CalendarApp.EventColor.PALE_RED,
  'YELLOW': CalendarApp.EventColor.YELLOW,
  'ORANGE': CalendarApp.EventColor.ORANGE,
  'CYAN': CalendarApp.EventColor.CYAN,
  'GRAY': CalendarApp.EventColor.GRAY,
  'BLUE': CalendarApp.EventColor.BLUE,
  'GREEN': CalendarApp.EventColor.GREEN,
  'RED': CalendarApp.EventColor.RED
};

// ============================================================================
// TIMEZONE MAPPINGS
// ============================================================================

// Comprehensive map of Outlook timezone names to IANA timezone identifiers
const TIMEZONE_MAP = {
  'Pacific Standard Time': 'America/Los_Angeles',
  'Mountain Standard Time': 'America/Denver',
  'Central Standard Time': 'America/Chicago',
  'Eastern Standard Time': 'America/New_York',
  'US Mountain Standard Time': 'America/Phoenix',
  'Alaskan Standard Time': 'America/Anchorage',
  'Hawaiian Standard Time': 'Pacific/Honolulu',
  'Atlantic Standard Time': 'America/Halifax',
  'Newfoundland Standard Time': 'America/St_Johns',
  'SA Pacific Standard Time': 'America/Bogota',
  'SA Western Standard Time': 'America/La_Paz',
  'SA Eastern Standard Time': 'America/Cayenne',
  'Argentina Standard Time': 'America/Buenos_Aires',
  'E. South America Standard Time': 'America/Sao_Paulo',
  'Greenland Standard Time': 'America/Godthab',
  'Montevideo Standard Time': 'America/Montevideo',
  'UTC': 'UTC',
  'GMT Standard Time': 'Europe/London',
  'Greenwich Standard Time': 'Atlantic/Reykjavik',
  'W. Europe Standard Time': 'Europe/Berlin',
  'Central Europe Standard Time': 'Europe/Budapest',
  'Romance Standard Time': 'Europe/Paris',
  'Central European Standard Time': 'Europe/Warsaw',
  'W. Central Africa Standard Time': 'Africa/Lagos',
  'Namibia Standard Time': 'Africa/Windhoek',
  'Jordan Standard Time': 'Asia/Amman',
  'GTB Standard Time': 'Europe/Bucharest',
  'Middle East Standard Time': 'Asia/Beirut',
  'Egypt Standard Time': 'Africa/Cairo',
  'Syria Standard Time': 'Asia/Damascus',
  'E. Europe Standard Time': 'Europe/Chisinau',
  'South Africa Standard Time': 'Africa/Johannesburg',
  'FLE Standard Time': 'Europe/Kiev',
  'Turkey Standard Time': 'Europe/Istanbul',
  'Israel Standard Time': 'Asia/Jerusalem',
  'Libya Standard Time': 'Africa/Tripoli',
  'Arabic Standard Time': 'Asia/Baghdad',
  'Arab Standard Time': 'Asia/Riyadh',
  'Arabian Standard Time': 'Asia/Dubai',
  'Belarus Standard Time': 'Europe/Minsk',
  'Russian Standard Time': 'Europe/Moscow',
  'E. Africa Standard Time': 'Africa/Nairobi',
  'Iran Standard Time': 'Asia/Tehran',
  'Caucasus Standard Time': 'Asia/Yerevan',
  'Azerbaijan Standard Time': 'Asia/Baku',
  'Mauritius Standard Time': 'Indian/Mauritius',
  'Georgian Standard Time': 'Asia/Tbilisi',
  'Afghanistan Standard Time': 'Asia/Kabul',
  'West Asia Standard Time': 'Asia/Tashkent',
  'Pakistan Standard Time': 'Asia/Karachi',
  'India Standard Time': 'Asia/Kolkata',
  'Sri Lanka Standard Time': 'Asia/Colombo',
  'Nepal Standard Time': 'Asia/Kathmandu',
  'Central Asia Standard Time': 'Asia/Almaty',
  'Bangladesh Standard Time': 'Asia/Dhaka',
  'Ekaterinburg Standard Time': 'Asia/Yekaterinburg',
  'Myanmar Standard Time': 'Asia/Rangoon',
  'SE Asia Standard Time': 'Asia/Bangkok',
  'N. Central Asia Standard Time': 'Asia/Novosibirsk',
  'China Standard Time': 'Asia/Shanghai',
  'North Asia Standard Time': 'Asia/Krasnoyarsk',
  'Singapore Standard Time': 'Asia/Singapore',
  'W. Australia Standard Time': 'Australia/Perth',
  'Taipei Standard Time': 'Asia/Taipei',
  'Ulaanbaatar Standard Time': 'Asia/Ulaanbaatar',
  'North Asia East Standard Time': 'Asia/Irkutsk',
  'Japan Standard Time': 'Asia/Tokyo',
  'Korea Standard Time': 'Asia/Seoul',
  'Cen. Australia Standard Time': 'Australia/Adelaide',
  'AUS Central Standard Time': 'Australia/Darwin',
  'E. Australia Standard Time': 'Australia/Brisbane',
  'AUS Eastern Standard Time': 'Australia/Sydney',
  'West Pacific Standard Time': 'Pacific/Port_Moresby',
  'Tasmania Standard Time': 'Australia/Hobart',
  'Yakutsk Standard Time': 'Asia/Yakutsk',
  'Central Pacific Standard Time': 'Pacific/Guadalcanal',
  'Vladivostok Standard Time': 'Asia/Vladivostok',
  'New Zealand Standard Time': 'Pacific/Auckland',
  'Fiji Standard Time': 'Pacific/Fiji',
  'Kamchatka Standard Time': 'Asia/Kamchatka',
  'Tonga Standard Time': 'Pacific/Tongatapu',
  'Samoa Standard Time': 'Pacific/Apia',
  'Azores Standard Time': 'Atlantic/Azores',
  'Cape Verde Standard Time': 'Atlantic/Cape_Verde',
  'Morocco Standard Time': 'Africa/Casablanca',
  'Dateline Standard Time': 'Etc/GMT+12',
  'UTC-11': 'Etc/GMT+11',
  'UTC-02': 'Etc/GMT+2',
  'UTC+12': 'Etc/GMT-12'
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
          const eventColor = COLOR_MAP[CONFIG.EVENT_COLOR];
          if (eventColor) {
            newEvent.setColor(eventColor);
          } else {
            Logger.log(`WARNING: Unknown color name "${CONFIG.EVENT_COLOR}". Valid options: ${Object.keys(COLOR_MAP).join(', ')}`);
          }
        }
        
        Logger.log(`Added event: "${feedEvent.title}" at ${feedEvent.startTime.toISOString()} (key: ${key})`);
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
              Logger.log(`ERROR deleting duplicate event: ${e.toString()}`);
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
              Logger.log(`ERROR deleting duplicate event: ${e.toString()}`);
            }
          });
        }
      }
    });
    
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

// ============================================================================
// ICS PARSING FUNCTIONS
// ============================================================================

function parseICS(icsData) {
  const events = [];
  const lines = icsData.split(/\r?\n/);
  
  // Extract timezone definitions from the ICS file
  const timezones = extractTimezones(icsData);
  Logger.log(`Found ${Object.keys(timezones).length} timezone definitions in ICS`);
  
  let currentEvent = null;
  let eventCount = 0;
  
  for (let i = 0; i < lines.length; i++) {
    let line = lines[i].trim();
    
    // Handle line continuation (lines starting with space or tab)
    while (i + 1 < lines.length && (lines[i + 1].startsWith(' ') || lines[i + 1].startsWith('\t'))) {
      i++;
      line += lines[i].trim();
    }
    
    if (line === 'BEGIN:VEVENT') {
      currentEvent = {
        title: '(No title)',
        startTime: null,
        endTime: null,
        description: '',
        location: '',
        isAllDay: false,
        recurrence: null,
        rruleString: null,
        uid: null,  // Added UID field
        recurrenceEndDate: null  // Track UNTIL date
      };
      eventCount++;
    } else if (line === 'END:VEVENT' && currentEvent) {
      if (currentEvent.startTime && currentEvent.endTime) {
        if (currentEvent.rruleString) {
          const result = parseRRule(currentEvent.rruleString);
          currentEvent.recurrence = result.recurrence;
          currentEvent.recurrenceEndDate = result.endDate;
        }
        events.push(currentEvent);
      } else {
        Logger.log(`WARNING: Skipped incomplete event #${eventCount} - Start: ${currentEvent.startTime}, End: ${currentEvent.endTime}`);
      }
      currentEvent = null;
    } else if (currentEvent) {
      const colonIndex = line.indexOf(':');
      
      if (colonIndex === -1) continue;
      
      const fieldPart = line.substring(0, colonIndex);
      const fieldValue = line.substring(colonIndex + 1);
      
      // Parse field name and parameters
      const fieldParts = fieldPart.split(';');
      const fieldName = fieldParts[0];
      
      // Extract parameters (like TZID)
      const params = {};
      for (let j = 1; j < fieldParts.length; j++) {
        const paramParts = fieldParts[j].split('=');
        if (paramParts.length === 2) {
          params[paramParts[0]] = paramParts[1];
        }
      }
      
      switch (fieldName) {
        case 'UID':  // Parse UID field
          currentEvent.uid = fieldValue.trim();
          break;
        case 'SUMMARY':
          const summary = decodeICSText(fieldValue);
          if (summary && summary.trim()) {
            currentEvent.title = summary;
          }
          break;
        case 'DTSTART':
          currentEvent.startTime = parseICSDate(fieldValue, params, timezones);
          if (params['VALUE'] === 'DATE') {
            currentEvent.isAllDay = true;
          }
          break;
        case 'DTEND':
          currentEvent.endTime = parseICSDate(fieldValue, params, timezones);
          break;
        case 'DESCRIPTION':
          currentEvent.description = decodeICSText(fieldValue);
          break;
        case 'LOCATION':
          currentEvent.location = decodeICSText(fieldValue);
          break;
        case 'RRULE':
          currentEvent.rruleString = fieldValue;
          break;
      }
    }
  }
  
  return events;
}

function parseRRule(rruleString) {
  try {
    // Parse RRULE format: FREQ=WEEKLY;BYDAY=MO,WE,FR;UNTIL=20251231T235959Z
    const parts = rruleString.split(';');
    const rules = {};
    
    parts.forEach(part => {
      const [key, value] = part.split('=');
      if (key && value) {
        rules[key] = value;
      }
    });
    
    if (!rules['FREQ']) {
      return { recurrence: null, endDate: null };
    }
    
    let recurrence;
    let endDate = null;
    
    switch (rules['FREQ']) {
      case 'DAILY':
        const dailyInterval = rules['INTERVAL'] ? parseInt(rules['INTERVAL']) : 1;
        recurrence = CalendarApp.newRecurrence().addDailyRule().interval(dailyInterval);
        break;
        
      case 'WEEKLY':
        const weeklyInterval = rules['INTERVAL'] ? parseInt(rules['INTERVAL']) : 1;
        recurrence = CalendarApp.newRecurrence().addWeeklyRule().interval(weeklyInterval);
        
        if (rules['BYDAY']) {
          const days = rules['BYDAY'].split(',');
          days.forEach(day => {
            switch (day) {
              case 'SU': recurrence.onlyOnWeekday(CalendarApp.Weekday.SUNDAY); break;
              case 'MO': recurrence.onlyOnWeekday(CalendarApp.Weekday.MONDAY); break;
              case 'TU': recurrence.onlyOnWeekday(CalendarApp.Weekday.TUESDAY); break;
              case 'WE': recurrence.onlyOnWeekday(CalendarApp.Weekday.WEDNESDAY); break;
              case 'TH': recurrence.onlyOnWeekday(CalendarApp.Weekday.THURSDAY); break;
              case 'FR': recurrence.onlyOnWeekday(CalendarApp.Weekday.FRIDAY); break;
              case 'SA': recurrence.onlyOnWeekday(CalendarApp.Weekday.SATURDAY); break;
            }
          });
        }
        break;
        
      case 'MONTHLY':
        const monthlyInterval = rules['INTERVAL'] ? parseInt(rules['INTERVAL']) : 1;
        recurrence = CalendarApp.newRecurrence().addMonthlyRule().interval(monthlyInterval);
        
        if (rules['BYMONTHDAY']) {
          recurrence.onlyOnMonthDay(parseInt(rules['BYMONTHDAY']));
        }
        break;
        
      case 'YEARLY':
        const yearlyInterval = rules['INTERVAL'] ? parseInt(rules['INTERVAL']) : 1;
        recurrence = CalendarApp.newRecurrence().addYearlyRule().interval(yearlyInterval);
        break;
        
      default:
        Logger.log(`Unknown FREQ type: ${rules['FREQ']}`);
        return { recurrence: null, endDate: null };
    }
    
    if (rules['UNTIL']) {
      endDate = parseICSDate(rules['UNTIL'], {}, {});
      recurrence.until(endDate);
    }
    
    // Handle COUNT (number of occurrences)
    if (rules['COUNT']) {
      recurrence.times(parseInt(rules['COUNT']));
    }
    
    return { recurrence, endDate };
    
  } catch (e) {
    Logger.log(`Error parsing RRULE "${rruleString}": ${e.toString()}`);
    return { recurrence: null, endDate: null };
  }
}

// ============================================================================
// TIMEZONE HELPER FUNCTIONS
// ============================================================================

function isDST(date) {
  const year = date.getFullYear();
  
  // DST starts: Second Sunday in March at 2:00 AM
  const marchFirst = new Date(year, 2, 1);
  const dstStart = new Date(year, 2, 1 + (7 - marchFirst.getDay()) + 7, 2, 0, 0);
  
  // DST ends: First Sunday in November at 2:00 AM
  const novemberFirst = new Date(year, 10, 1);
  const dstEnd = new Date(year, 10, 1 + (7 - novemberFirst.getDay()), 2, 0, 0);
  
  return date >= dstStart && date < dstEnd;
}

function getTimezoneOffset(ianaTimezone, date) {
  const isDSTActive = isDST(date);
  
  const offsets = {
    'America/Los_Angeles': isDSTActive ? -7 * 60 * 60 * 1000 : -8 * 60 * 60 * 1000,
    'America/Denver': isDSTActive ? -6 * 60 * 60 * 1000 : -7 * 60 * 60 * 1000,
    'America/Chicago': isDSTActive ? -5 * 60 * 60 * 1000 : -6 * 60 * 60 * 1000,
    'America/New_York': isDSTActive ? -4 * 60 * 60 * 1000 : -5 * 60 * 60 * 1000,
    'America/Anchorage': isDSTActive ? -8 * 60 * 60 * 1000 : -9 * 60 * 60 * 1000,
    'America/Halifax': isDSTActive ? -3 * 60 * 60 * 1000 : -4 * 60 * 60 * 1000,
    'America/St_Johns': isDSTActive ? -2.5 * 60 * 60 * 1000 : -3.5 * 60 * 60 * 1000,
    'America/Phoenix': -7 * 60 * 60 * 1000,
    'Pacific/Honolulu': -10 * 60 * 60 * 1000,
    'Asia/Kolkata': 5.5 * 60 * 60 * 1000,
    'Asia/Dubai': 4 * 60 * 60 * 1000,
    'Asia/Shanghai': 8 * 60 * 60 * 1000,
    'Asia/Tokyo': 9 * 60 * 60 * 1000,
    'Asia/Seoul': 9 * 60 * 60 * 1000,
    'Asia/Singapore': 8 * 60 * 60 * 1000,
    'Asia/Bangkok': 7 * 60 * 60 * 1000,
    'Asia/Riyadh': 3 * 60 * 60 * 1000,
    'Europe/London': isDSTActive ? 1 * 60 * 60 * 1000 : 0,
    'Europe/Berlin': isDSTActive ? 2 * 60 * 60 * 1000 : 1 * 60 * 60 * 1000,
    'Europe/Paris': isDSTActive ? 2 * 60 * 60 * 1000 : 1 * 60 * 60 * 1000,
    'UTC': 0
  };
  
  return offsets[ianaTimezone] || 0;
}

// ============================================================================
// TRIGGER SETUP
// ============================================================================

function setupTrigger() {
  // Delete existing triggers to avoid duplicates
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'syncCalendarEvents') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // Create new trigger to run every 30 minutes
  ScriptApp.newTrigger('syncCalendarEvents')
    .timeBased()
    .everyMinutes(30)
    .create();
  
  Logger.log('Trigger set up successfully! Script will run every 30 minutes.');
}

function generateEventKey(title, startTime, endTime) {
  return `${title}_${startTime.getTime()}_${endTime.getTime()}`;
}

function extractTimezones(icsData) {
  const timezones = {};
  const lines = icsData.split(/\r?\n/);
  
  let inTimezone = false;
  let currentTzid = null;
  
  for (let i = 0; i < lines.length; i++) {
    let line = lines[i].trim();
    
    if (line.startsWith('BEGIN:VTIMEZONE')) {
      inTimezone = true;
    } else if (line.startsWith('END:VTIMEZONE')) {
      inTimezone = false;
      currentTzid = null;
    } else if (inTimezone) {
      if (line.startsWith('TZID:')) {
        currentTzid = line.substring(5);
        if (TIMEZONE_MAP[currentTzid]) {
          timezones[currentTzid] = TIMEZONE_MAP[currentTzid];
          Logger.log(`Mapped timezone: ${currentTzid} -> ${TIMEZONE_MAP[currentTzid]}`);
        } else {
          Logger.log(`WARNING: Unknown timezone: ${currentTzid}`);
        }
      }
    }
  }
  
  return timezones;
}

function parseICSDate(dateString, params = {}, timezones = {}) {
  dateString = dateString.trim();
  
  // Check if it's a DATE only (all-day event)
  if (dateString.length === 8 || params['VALUE'] === 'DATE') {
    const year = parseInt(dateString.substring(0, 4));
    const month = parseInt(dateString.substring(4, 6)) - 1;
    const day = parseInt(dateString.substring(6, 8));
    return new Date(year, month, day);
  }
  
  // Parse datetime: YYYYMMDDTHHMMSS or YYYYMMDDTHHMMSSZ
  const year = parseInt(dateString.substring(0, 4));
  const month = parseInt(dateString.substring(4, 6)) - 1;
  const day = parseInt(dateString.substring(6, 8));
  const hour = parseInt(dateString.substring(9, 11));
  const minute = parseInt(dateString.substring(11, 13));
  const second = parseInt(dateString.substring(13, 15)) || 0;
  
  // If it ends with Z, it's explicitly UTC time
  if (dateString.endsWith('Z')) {
    return new Date(Date.UTC(year, month, day, hour, minute, second));
  }
  
  // If there's a TZID parameter, try to handle it
  if (params['TZID']) {
    const tzid = params['TZID'];
    const ianaTimezone = timezones[tzid] || TIMEZONE_MAP[tzid];
    
    if (ianaTimezone) {
      const utcDate = new Date(Date.UTC(year, month, day, hour, minute, second));
      const offset = getTimezoneOffset(ianaTimezone, utcDate);
      return new Date(utcDate.getTime() - offset);
    }
  }
  
  // Default: treat as local time
  return new Date(year, month, day, hour, minute, second);
}

function decodeICSText(text) {
  return text
    .replace(/\\n/g, '\n')
    .replace(/\\,/g, ',')
    .replace(/\\;/g, ';')
    .replace(/\\\\/g, '\\');
}

// ============================================================================
// NEW FUNCTION TO EXPAND RECURRING EVENTS
// ============================================================================

function expandRecurringEvent(event, startDate, endDate) {
  const occurrences = [];
  
  try {
    const rules = {};
    const parts = event.rruleString.split(';');
    parts.forEach(part => {
      const [key, value] = part.split('=');
      if (key && value) {
        rules[key] = value;
      }
    });
    
    if (!rules['FREQ']) {
      return occurrences;
    }
    
    // Determine recurrence end date
    let recurrenceEnd = endDate;
    if (rules['UNTIL']) {
      const untilDate = parseICSDate(rules['UNTIL'], {}, {});
      if (untilDate < recurrenceEnd) {
        recurrenceEnd = untilDate;
      }
    }
    
    // Skip if recurrence already ended
    if (recurrenceEnd < startDate) {
      Logger.log(`Skipping recurring event "${event.title}" - ended on ${recurrenceEnd.toISOString()}`);
      return occurrences;
    }
    
    const duration = event.endTime.getTime() - event.startTime.getTime();
    let currentDate = new Date(event.startTime);
    const interval = rules['INTERVAL'] ? parseInt(rules['INTERVAL']) : 1;
    const maxOccurrences = rules['COUNT'] ? parseInt(rules['COUNT']) : 1000; // Safety limit
    let count = 0;
    
    // Generate occurrences based on frequency
    while (currentDate <= recurrenceEnd && count < maxOccurrences) {
      // Check if this occurrence is within our sync window
      if (currentDate >= startDate && currentDate <= endDate) {
        // Check if this day matches the BYDAY rule (for weekly events)
        let includeOccurrence = true;
        
        if (rules['FREQ'] === 'WEEKLY' && rules['BYDAY']) {
          const dayOfWeek = currentDate.getDay();
          const dayMap = { 'SU': 0, 'MO': 1, 'TU': 2, 'WE': 3, 'TH': 4, 'FR': 5, 'SA': 6 };
          const allowedDays = rules['BYDAY'].split(',').map(d => dayMap[d]);
          includeOccurrence = allowedDays.includes(dayOfWeek);
        }
        
        if (includeOccurrence) {
          occurrences.push({
            title: event.title,
            startTime: new Date(currentDate),
            endTime: new Date(currentDate.getTime() + duration),
            description: event.description,
            location: event.location,
            isAllDay: event.isAllDay,
            uid: event.uid,
            recurrence: null,
            rruleString: null
          });
        }
      }
      
      // Move to next occurrence
      switch (rules['FREQ']) {
        case 'DAILY':
          currentDate.setDate(currentDate.getDate() + interval);
          break;
        case 'WEEKLY':
          currentDate.setDate(currentDate.getDate() + (7 * interval));
          break;
        case 'MONTHLY':
          currentDate.setMonth(currentDate.getMonth() + interval);
          break;
        case 'YEARLY':
          currentDate.setFullYear(currentDate.getFullYear() + interval);
          break;
        default:
          // Unknown frequency, stop
          break;
      }
      
      count++;
    }
    
  } catch (e) {
    Logger.log(`Error expanding recurring event "${event.title}": ${e.toString()}`);
  }
  
  return occurrences;
}
