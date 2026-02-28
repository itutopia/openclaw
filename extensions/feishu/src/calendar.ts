import type * as Lark from "@larksuiteoapi/node-sdk";
import { Type } from "@sinclair/typebox";
import type { OpenClawPluginApi } from "openclaw/plugin-sdk";
import { listEnabledFeishuAccounts } from "./accounts.js";
import { createFeishuClient } from "./client.js";
import { resolveToolsConfig } from "./tools-config.js";

// ============ Helpers ============

function json(data: unknown) {
  return {
    content: [{ type: "text" as const, text: JSON.stringify(data, null, 2) }],
    details: data,
  };
}

/**
 * Convert a date string or timestamp to Unix timestamp (seconds)
 * Supports: "2024-03-03", "2024-03-03 10:00", timestamp in ms
 */
function toTimestamp(date: string | number, defaultHour = 9): number {
  if (typeof date === "number") {
    // Timestamp in milliseconds, convert to seconds
    return Math.floor(date / 1000);
  }

  // Try to parse date string
  // Format: "2024-03-03" or "2024-03-03 10:00" or "2024-03-03T10:00:00"
  const parsed = new Date(date);
  if (!isNaN(parsed.getTime())) {
    // If no time specified, use default hour
    if (/^\d{4}-\d{2}-\d{2}$/.test(date)) {
      parsed.setHours(defaultHour, 0, 0, 0);
    }
    return Math.floor(parsed.getTime() / 1000);
  }

  throw new Error(`Invalid date format: ${date}`);
}

/**
 * Calculate end time based on start time and duration
 */
function calculateEndTime(timestamp: number, durationMinutes: number): number {
  return timestamp + durationMinutes * 60;
}

// ============ Schema ============

const FeishuCalendarSchema = Type.Object({
  action: Type.Union([
    Type.Literal("list_calendars"),
    Type.Literal("get_calendar"),
    Type.Literal("get_primary_calendar"),
    Type.Literal("create_calendar"),
    Type.Literal("list_events"),
    Type.Literal("get_event"),
    Type.Literal("create_event"),
    Type.Literal("update_event"),
    Type.Literal("delete_event"),
  ]),

  // Calendar operations
  calendar_id: Type.Optional(Type.String({ description: "Calendar ID" })),
  calendar_summary: Type.Optional(Type.String({ description: "Calendar name/summary" })),
  calendar_description: Type.Optional(Type.String({ description: "Calendar description" })),

  // Event operations
  event_id: Type.Optional(Type.String({ description: "Event ID" })),
  summary: Type.Optional(Type.String({ description: "Event title/summary" })),
  description: Type.Optional(Type.String({ description: "Event description" })),
  start_time: Type.Optional(
    Type.String({
      description:
        'Event start time. Formats: "2024-03-03", "2024-03-03 10:00", ISO string, or timestamp in ms',
    }),
  ),
  end_time: Type.Optional(
    Type.String({
      description: "Event end time. Same formats as start_time",
    }),
  ),
  duration_minutes: Type.Optional(
    Type.Number({
      description: "Event duration in minutes (used if end_time not specified, default: 60)",
    }),
  ),
  location: Type.Optional(Type.String({ description: "Event location" })),
  attendee_ids: Type.Optional(
    Type.Array(Type.String({ description: "Attendee user ID (open_id or user_id)" })),
  ),

  // List filters
  from_time: Type.Optional(Type.String({ description: "List events from this time" })),
  to_time: Type.Optional(Type.String({ description: "List events until this time" })),
  page_size: Type.Optional(Type.Number({ description: "Page size for listing (default: 50)" })),
  page_token: Type.Optional(Type.String({ description: "Pagination token" })),
});

type FeishuCalendarParams = Static<typeof FeishuCalendarSchema>;

// ============ Calendar Actions ============

async function getPrimaryCalendar(client: Lark.Client) {
  const res = await client.calendar.calendar.primary({
    data: {},
  });
  if (res.code !== 0) {
    throw new Error(res.msg);
  }

  return {
    calendar_id: res.data?.calendar?.calendar_id,
    summary: res.data?.calendar?.summary,
    description: res.data?.calendar?.description,
    type: res.data?.calendar?.type,
    permission: res.data?.calendar?.permission,
  };
}

async function listCalendars(client: Lark.Client) {
  const res = await client.calendar.calendar.list({
    data: {},
  });
  if (res.code !== 0) {
    throw new Error(res.msg);
  }

  return {
    calendars: res.data?.calendar_list?.map((cal) => ({
      calendar_id: cal.calendar_id,
      summary: cal.summary,
      description: cal.description,
      type: cal.type,
      permission: cal.permission,
    })),
    page_token: res.data?.page_token,
    has_more: res.data?.has_more,
  };
}

async function getCalendar(client: Lark.Client, calendarId: string) {
  const res = await client.calendar.calendar.get({
    path: { calendar_id: calendarId },
  });
  if (res.code !== 0) {
    throw new Error(res.msg);
  }

  return {
    calendar: res.data?.calendar,
  };
}

async function createCalendar(client: Lark.Client, summary: string, description?: string) {
  const res = await client.calendar.calendar.create({
    data: {
      summary,
      description,
    },
  });
  if (res.code !== 0) {
    throw new Error(res.msg);
  }

  return {
    calendar_id: res.data?.calendar?.calendar_id,
    summary: res.data?.calendar?.summary,
    description: res.data?.calendar?.description,
  };
}

// ============ Event Actions ============

async function listEvents(
  client: Lark.Client,
  calendarId: string,
  fromTime?: string,
  toTime?: string,
  pageSize?: number,
  pageToken?: string,
) {
  const res = await client.calendar.calendarEvent.list({
    path: { calendar_id: calendarId },
    params: {
      start_time: fromTime,
      end_time: toTime,
      page_size: pageSize ?? 50,
      page_token: pageToken,
    },
  });
  if (res.code !== 0) {
    throw new Error(res.msg);
  }

  return {
    events: res.data?.events?.map((event) => ({
      event_id: event.event_id,
      summary: event.summary,
      description: event.description,
      location: event.location?.name,
      start_time: event.start_time,
      end_time: event.end_time,
      status: event.status,
      visibility: event.visibility,
    })),
    page_token: res.data?.page_token,
    has_more: res.data?.has_more,
  };
}

async function getEvent(client: Lark.Client, calendarId: string, eventId: string) {
  const res = await client.calendar.calendarEvent.get({
    path: { calendar_id: calendarId, event_id: eventId },
  });
  if (res.code !== 0) {
    throw new Error(res.msg);
  }

  return {
    event: res.data?.event,
  };
}

interface CreateEventParams {
  client: Lark.Client;
  calendarId: string;
  summary: string;
  startTime: string;
  endTime?: string;
  durationMinutes?: number;
  description?: string;
  location?: string;
  attendeeIds?: string[];
}

async function createEvent(params: CreateEventParams) {
  const {
    client,
    calendarId,
    summary,
    startTime,
    endTime,
    durationMinutes = 60,
    description,
    location,
    attendeeIds,
  } = params;

  const start = toTimestamp(startTime);
  const end = endTime ? toTimestamp(endTime) : calculateEndTime(start, durationMinutes);

  const eventData: Record<string, unknown> = {
    summary,
    start_time: { timestamp: start },
    end_time: { timestamp: end },
  };

  if (description) {
    eventData.description = description;
  }

  if (location) {
    eventData.location = { name: location };
  }

  if (attendeeIds && attendeeIds.length > 0) {
    eventData.attendees = attendeeIds.map((id) => ({
      type: "user" as const,
      user_id: id,
    }));
  }

  const res = await client.calendar.calendarEvent.create({
    path: { calendar_id: calendarId },
    data: eventData,
  });
  if (res.code !== 0) {
    throw new Error(res.msg);
  }

  return {
    event_id: res.data?.event?.event_id,
    summary: res.data?.event?.summary,
    start_time: res.data?.event?.start_time,
    end_time: res.data?.event?.end_time,
    location: res.data?.event?.location?.name,
  };
}

interface UpdateEventParams {
  client: Lark.Client;
  calendarId: string;
  eventId: string;
  summary?: string;
  startTime?: string;
  endTime?: string;
  durationMinutes?: number;
  description?: string;
  location?: string;
  attendeeIds?: string[];
}

async function updateEvent(params: UpdateEventParams) {
  const {
    client,
    calendarId,
    eventId,
    summary,
    startTime,
    endTime,
    durationMinutes,
    description,
    location,
    attendeeIds,
  } = params;

  const eventData: Record<string, unknown> = {};

  if (summary) {
    eventData.summary = summary;
  }

  if (startTime) {
    const start = toTimestamp(startTime);
    eventData.start_time = { timestamp: start };

    if (endTime) {
      eventData.end_time = { timestamp: toTimestamp(endTime) };
    } else if (durationMinutes) {
      eventData.end_time = { timestamp: calculateEndTime(start, durationMinutes) };
    }
  }

  if (description !== undefined) {
    eventData.description = description;
  }

  if (location !== undefined) {
    eventData.location = location ? { name: location } : null;
  }

  if (attendeeIds !== undefined) {
    eventData.attendees = attendeeIds.map((id) => ({
      type: "user" as const,
      user_id: id,
    }));
  }

  const res = await client.calendar.calendarEvent.patch({
    path: { calendar_id: calendarId, event_id: eventId },
    data: eventData,
  });
  if (res.code !== 0) {
    throw new Error(res.msg);
  }

  return {
    success: true,
    event_id: res.data?.event?.event_id,
    summary: res.data?.event?.summary,
  };
}

async function deleteEvent(client: Lark.Client, calendarId: string, eventId: string) {
  const res = await client.calendar.calendarEvent.delete({
    path: { calendar_id: calendarId, event_id: eventId },
  });
  if (res.code !== 0) {
    throw new Error(res.msg);
  }

  return { success: true, deleted_event_id: eventId };
}

// ============ Tool Registration ============

export function registerFeishuCalendarTools(api: OpenClawPluginApi) {
  if (!api.config) {
    api.logger.debug?.("feishu_calendar: No config available, skipping calendar tools");
    return;
  }

  const accounts = listEnabledFeishuAccounts(api.config);
  if (accounts.length === 0) {
    api.logger.debug?.("feishu_calendar: No Feishu accounts configured, skipping calendar tools");
    return;
  }

  const firstAccount = accounts[0];
  const toolsCfg = resolveToolsConfig(firstAccount.config.tools);

  if (!toolsCfg.calendar) {
    api.logger.debug?.("feishu_calendar: Calendar tools disabled in config");
    return;
  }

  const getClient = () => createFeishuClient(firstAccount);

  // Default to primary calendar (user's main calendar)
  const PRIMARY_CALENDAR_ID = "primary";

  api.registerTool(
    {
      name: "feishu_calendar",
      label: "Feishu Calendar",
      description:
        "Feishu calendar operations. Actions: list_calendars, get_calendar, get_primary_calendar, create_calendar, list_events, get_event, create_event, update_event, delete_event",
      parameters: FeishuCalendarSchema,
      async execute(_toolCallId, params) {
        const p = params as FeishuCalendarParams;
        try {
          const client = getClient();
          const calendarId = p.calendar_id ?? PRIMARY_CALENDAR_ID;

          switch (p.action) {
            case "list_calendars":
              return json(await listCalendars(client));

            case "get_calendar":
              return json(await getCalendar(client, calendarId));

            case "get_primary_calendar":
              return json(await getPrimaryCalendar(client));

            case "create_calendar":
              if (!p.calendar_summary) {
                return json({ error: "calendar_summary is required for create_calendar" });
              }
              return json(await createCalendar(client, p.calendar_summary, p.calendar_description));

            case "list_events":
              // If using default "primary", get the actual primary calendar ID first
              let listCalendarId = calendarId;
              if (calendarId === PRIMARY_CALENDAR_ID) {
                const primaryCal = await getPrimaryCalendar(client);
                listCalendarId = primaryCal.calendar_id;
              }
              return json(
                await listEvents(
                  client,
                  listCalendarId,
                  p.from_time,
                  p.to_time,
                  p.page_size,
                  p.page_token,
                ),
              );

            case "get_event":
              if (!p.event_id) {
                return json({ error: "event_id is required for get_event" });
              }
              return json(await getEvent(client, calendarId, p.event_id));

            case "create_event":
              if (!p.summary || !p.start_time) {
                return json({ error: "summary and start_time are required for create_event" });
              }
              // If using default "primary", get the actual primary calendar ID first
              let targetCalendarId = calendarId;
              if (calendarId === PRIMARY_CALENDAR_ID) {
                const primaryCal = await getPrimaryCalendar(client);
                targetCalendarId = primaryCal.calendar_id;
              }
              return json(
                await createEvent({
                  client,
                  calendarId: targetCalendarId,
                  summary: p.summary,
                  startTime: p.start_time,
                  endTime: p.end_time,
                  durationMinutes: p.duration_minutes,
                  description: p.description,
                  location: p.location,
                  attendeeIds: p.attendee_ids,
                }),
              );

            case "update_event":
              if (!p.event_id) {
                return json({ error: "event_id is required for update_event" });
              }
              return json(
                await updateEvent({
                  client,
                  calendarId,
                  eventId: p.event_id,
                  summary: p.summary,
                  startTime: p.start_time,
                  endTime: p.end_time,
                  durationMinutes: p.duration_minutes,
                  description: p.description,
                  location: p.location,
                  attendeeIds: p.attendee_ids,
                }),
              );

            case "delete_event":
              if (!p.event_id) {
                return json({ error: "event_id is required for delete_event" });
              }
              return json(await deleteEvent(client, calendarId, p.event_id));

            default:
              return json({ error: `Unknown action: ${(p as { action: string }).action}` });
          }
        } catch (err) {
          return json({ error: err instanceof Error ? err.message : String(err) });
        }
      },
    },
    { name: "feishu_calendar" },
  );

  api.logger.info?.("feishu_calendar: Registered feishu_calendar tool");
}
