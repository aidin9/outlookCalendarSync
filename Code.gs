// ============================================================================
// CONFIGURATION
// ============================================================================

const CONFIG = {
  // IMPORTANT: Replace this with your Outlook calendar ICS feed URL
  ICS_FEED_URL: 'YOUR_OUTLOOK_ICS_URL_HERE',
  
  // Name of the Google Calendar to sync events to (must already exist)
  TARGET_CALENDAR_NAME: 'Work',
  
  // How many weeks ahead to sync events
  SYNC_WEEKS_AHEAD: 8
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
    
    // Get the target calendar by name
    const calendars = CalendarApp.getCalendarsByName(CONFIG.TARGET_CALENDAR_NAME);
    if (calendars.length === 0) {
      throw new Error(`Calendar "${CONFIG.TARGET_CALENDAR_NAME}" not found. Please create it first.`);
    }
    const calendar = calendars[0];
    
    // Get existing events to avoid duplicates
    const now = new Date();
    const futureDate = new Date(now.getTime() + (CONFIG.SYNC_WEEKS_AHEAD * 7 * 24 * 60 * 60 * 1000));
    Logger.log(`Syncing events from ${now.toISOString()} to ${futureDate.toISOString()}`);
    
    const existingEvents = calendar.getEvents(now, futureDate);
    Logger.log(`Found ${existingEvents.length} existing events in target calendar`);
    
    // Create a map of existing events for quick lookup
    const existingEventMap = {};
    existingEvents.forEach(event => {
      const key = `${event.getTitle()}_${event.getStartTime().getTime()}_${event.getEndTime().getTime()}`;
      existingEventMap[key] = true;
    });
    
    // Add new events
    let addedCount = 0;
    let skippedPast = 0;
    let skippedDuplicate = 0;
    let skippedFuture = 0;
    let skippedOldRecurring = 0;
    let errorCount = 0;
    let recurringCount = 0;
    
    events.forEach((event, index) => {
      Logger.log(`\n--- Processing event #${index + 1}: "${event.title}" ---`);
      Logger.log(`  Start: ${event.startTime ? event.startTime.toISOString() : 'NULL'}`);
      Logger.log(`  End: ${event.endTime ? event.endTime.toISOString() : 'NULL'}`);
      Logger.log(`  Has recurrence: ${event.recurrence ? 'YES' : 'NO'}`);
      if (event.rruleString) {
        Logger.log(`  RRULE string: ${event.rruleString}`);
      }
      
      // Handle recurring events differently - don't skip if they have recurrence
      if (event.recurrence) {
        recurringCount++;
        Logger.log(`  → This is a recurring event`);
        
        // Skip recurring events that started more than 6 months ago to avoid very old series
        const sixMonthsAgo = new Date(now.getTime() - (180 * 24 * 60 * 60 * 1000));
        if (event.startTime < sixMonthsAgo) {
          Logger.log(`  → SKIPPED: Recurring event started too long ago (${event.startTime.toISOString()})`);
          skippedOldRecurring++;
          return;
        }
      } else {
        Logger.log(`  → This is a one-time event`);
        
        // For non-recurring events, skip if in the past
        if (event.endTime < now) {
          Logger.log(`  → SKIPPED: Event ended in the past (${event.endTime.toISOString()})`);
          skippedPast++;
          return;
        }
        
        // Skip events too far in the future
        if (event.startTime > futureDate) {
          Logger.log(`  → SKIPPED: Event is beyond ${CONFIG.SYNC_WEEKS_AHEAD} weeks (${event.startTime.toISOString()})`);
          skippedFuture++;
          return;
        }
      }
      
      const key = `${event.title}_${event.startTime.getTime()}_${event.endTime.getTime()}`;
      
      // Skip if event already exists
      if (existingEventMap[key]) {
        Logger.log(`  → SKIPPED: Duplicate event already exists`);
        skippedDuplicate++;
        return;
      }
      
      // Create the event
      try {
        // Create recurring event series if recurrence exists
        if (event.recurrence) {
          if (event.isAllDay) {
            calendar.createAllDayEventSeries(
              event.title,
              event.startTime,
              event.recurrence,
              {
                description: event.description,
                location: event.location
              }
            );
          } else {
            calendar.createEventSeries(
              event.title,
              event.startTime,
              event.endTime,
              event.recurrence,
              {
                description: event.description,
                location: event.location
              }
            );
          }
          Logger.log(`  → ADDED: Recurring event created successfully`);
        } else {
          // Regular one-time event
          if (event.isAllDay) {
            calendar.createAllDayEvent(event.title, event.startTime, {
              description: event.description,
              location: event.location
            });
          } else {
            calendar.createEvent(event.title, event.startTime, event.endTime, {
              description: event.description,
              location: event.location
            });
          }
          Logger.log(`  → ADDED: One-time event created successfully`);
        }
        
        addedCount++;
      } catch (e) {
        Logger.log(`  → ERROR: Failed to create event - ${e.toString()}`);
        errorCount++;
      }
    });
    
    Logger.log('\n=== SYNC SUMMARY ===');
    Logger.log(`Total events in feed: ${events.length}`);
    Logger.log(`Recurring events found: ${recurringCount}`);
    Logger.log(`Added: ${addedCount}`);
    Logger.log(`Skipped (duplicates): ${skippedDuplicate}`);
    Logger.log(`Skipped (past events): ${skippedPast}`);
    Logger.log(`Skipped (old recurring): ${skippedOldRecurring}`);
    Logger.log(`Skipped (beyond ${CONFIG.SYNC_WEEKS_AHEAD} weeks): ${skippedFuture}`);
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
        rruleString: null
      };
      eventCount++;
    } else if (line === 'END:VEVENT' && currentEvent) {
      if (currentEvent.startTime && currentEvent.endTime) {
        if (currentEvent.rruleString) {
          currentEvent.recurrence = parseRRule(currentEvent.rruleString);
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
      return null;
    }
    
    let recurrence;
    
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
        return null;
    }
    
    // Handle UNTIL (end date)
    if (rules['UNTIL']) {
      const untilDate = parseICSDate(rules['UNTIL'], {}, {});
      recurrence.until(untilDate);
    }
    
    // Handle COUNT (number of occurrences)
    if (rules['COUNT']) {
      recurrence.times(parseInt(rules['COUNT']));
    }
    
    return recurrence;
    
  } catch (e) {
    Logger.log(`Error parsing RRULE "${rruleString}": ${e.toString()}`);
    return null;
  }
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
  
  if (dateString.length === 8 || params['VALUE'] === 'DATE') {
    const year = parseInt(dateString.substring(0, 4));
    const month = parseInt(dateString.substring(4, 6)) - 1;
    const day = parseInt(dateString.substring(6, 8));
    return new Date(year, month, day);
  }
  
  const year = parseInt(dateString.substring(0, 4));
  const month = parseInt(dateString.substring(4, 6)) - 1;
  const day = parseInt(dateString.substring(6, 8));
  const hour = parseInt(dateString.substring(9, 11));
  const minute = parseInt(dateString.substring(11, 13));
  const second = parseInt(dateString.substring(13, 15)) || 0;
  
  if (dateString.endsWith('Z')) {
    return new Date(Date.UTC(year, month, day, hour, minute, second));
  }
  
  if (params['TZID']) {
    const tzid = params['TZID'];
    const ianaTimezone = timezones[tzid] || TIMEZONE_MAP[tzid];
    
    if (ianaTimezone) {
      const utcDate = new Date(Date.UTC(year, month, day, hour, minute, second));
      const offset = getTimezoneOffset(ianaTimezone, utcDate);
      return new Date(utcDate.getTime() - offset);
    }
  }
  
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
