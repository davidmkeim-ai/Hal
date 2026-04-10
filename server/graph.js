const GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0";

export async function createTeamsCalendarEvent(accessToken, meeting) {
  const response = await fetch(`${GRAPH_BASE_URL}/me/events`, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      subject: meeting.title,
      body: {
        contentType: "HTML",
        content: `
          <p>Created by HAL.</p>
          <p><strong>Source note category:</strong> ${escapeHtml(meeting.category || "general")}</p>
          <p>${escapeHtml(meeting.notes).replace(/\n/g, "<br>")}</p>
        `,
      },
      start: {
        dateTime: meeting.start,
        timeZone: meeting.timeZone,
      },
      end: {
        dateTime: meeting.end,
        timeZone: meeting.timeZone,
      },
      attendees: (meeting.attendees || []).map((address) => ({
        emailAddress: {
          address,
        },
        type: "required",
      })),
      isOnlineMeeting: true,
    }),
  });

  const payload = await response.json();

  if (!response.ok) {
    throw new Error(payload.error?.message || "Graph request failed");
  }

  return {
    id: payload.id,
    subject: payload.subject,
    webLink: payload.webLink,
    joinUrl: payload.onlineMeeting?.joinUrl || payload.onlineMeetingUrl || "",
  };
}

export async function getOutlookCalendarEventsForDay(accessToken, date, timeZone = "America/Chicago") {
  const nextDate = addDays(date, 1);
  const url = new URL(`${GRAPH_BASE_URL}/me/calendarView`);
  url.searchParams.set("startDateTime", `${date}T00:00:00`);
  url.searchParams.set("endDateTime", `${nextDate}T00:00:00`);
  url.searchParams.set("$select", "id,subject,start,end,webLink,isAllDay,onlineMeeting");
  url.searchParams.set("$orderby", "start/dateTime");

  const response = await fetch(url, {
    headers: {
      Authorization: `Bearer ${accessToken}`,
      Prefer: `outlook.timezone="${timeZone}"`,
    },
  });

  const payload = await response.json();

  if (!response.ok) {
    throw new Error(payload.error?.message || "Graph calendar request failed");
  }

  return (payload.value || []).map((event) => ({
    id: event.id,
    title: event.subject || "Untitled event",
    date,
    time: event.isAllDay ? "All day" : formatGraphEventTime(event.start?.dateTime),
    source: "outlook",
    webLink: event.webLink || "",
  }));
}

function addDays(dateString, count) {
  const [year, month, day] = dateString.split("-").map(Number);
  const value = new Date(year, month - 1, day + count);
  const nextYear = value.getFullYear();
  const nextMonth = String(value.getMonth() + 1).padStart(2, "0");
  const nextDay = String(value.getDate()).padStart(2, "0");
  return `${nextYear}-${nextMonth}-${nextDay}`;
}

function formatGraphEventTime(dateTime) {
  if (!dateTime) {
    return "";
  }

  const value = new Date(dateTime);
  if (Number.isNaN(value.getTime())) {
    const match = String(dateTime).match(/T(\d{2}):(\d{2})/);
    if (!match) {
      return "";
    }
    const [_, hourText, minuteText] = match;
    const hours = Number(hourText);
    const suffix = hours >= 12 ? "PM" : "AM";
    const normalizedHour = hours % 12 || 12;
    return `${normalizedHour}:${minuteText} ${suffix}`;
  }

  return value.toLocaleTimeString([], {
    hour: "numeric",
    minute: "2-digit",
  });
}

function escapeHtml(value = "") {
  return value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}
