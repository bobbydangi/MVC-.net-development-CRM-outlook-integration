using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Graph;
using Microsoft.Identity.Web;
using System;
using System.Linq;
using System.Threading.Tasks;

namespace YourNamespace.Controllers
{
    [Authorize]
    public class CalendarController : Controller
    {
        private readonly GraphServiceClient _graphServiceClient;

        public CalendarController(GraphServiceClient graphServiceClient)
        {
            _graphServiceClient = graphServiceClient;
        }

        // GET: Calendar/Events
        public async Task<IActionResult> Events()
        {
            var events = await _graphServiceClient.Me.Events.Request()
                .Select(e => new {
                    e.Subject,
                    e.Organizer,
                    e.Start,
                    e.End
                })
                .OrderBy("createdDateTime DESC")
                .GetAsync();

            return View(events.CurrentPage);
        }

        // POST: Calendar/CreateEvent
        [HttpPost]
        public async Task<IActionResult> CreateEvent(string subject, DateTime start, DateTime end)
        {
            var @event = new Event
            {
                Subject = subject,
                Start = new DateTimeTimeZone { DateTime = start.ToString("o"), TimeZone = "UTC" },
                End = new DateTimeTimeZone { DateTime = end.ToString("o"), TimeZone = "UTC" }
            };

            await _graphServiceClient.Me.Events.Request().AddAsync(@event);

            return RedirectToAction(nameof(Events));
        }

        // POST: Calendar/UpdateEvent
        [HttpPost]
        public async Task<IActionResult> UpdateEvent(string id, string subject, DateTime start, DateTime end)
        {
            var @event = new Event
            {
                Subject = subject,
                Start = new DateTimeTimeZone { DateTime = start.ToString("o"), TimeZone = "UTC" },
                End = new DateTimeTimeZone { DateTime = end.ToString("o"), TimeZone = "UTC" }
            };

            await _graphServiceClient.Me.Events[id].Request().UpdateAsync(@event);

            return RedirectToAction(nameof(Events));
        }

        // POST: Calendar/DeleteEvent
        [HttpPost]
        public async Task<IActionResult> DeleteEvent(string id)
        {
            await _graphServiceClient.Me.Events[id].Request().DeleteAsync();
            return RedirectToAction(nameof(Events));
        }

        // GET: Calendar/SyncEvents
        public async Task<IActionResult> SyncEvents()
        {
            var events = await _graphServiceClient.Me.Events.Delta().Request().GetAsync();
            return View(events.CurrentPage);
        }

        // GET: Calendar/FindMeetingTimes
        public async Task<IActionResult> FindMeetingTimes()
        {
            var meetingTimes = await _graphServiceClient.Me.FindMeetingTimes(new FindMeetingTimesRequestBuilder())
                .Request()
                .PostAsync();
            return View(meetingTimes.MeetingTimeSuggestions);
        }

        // GET: Calendar/GetAttachments
        public async Task<IActionResult> GetAttachments(string eventId)
        {
            var attachments = await _graphServiceClient.Me.Events[eventId].Attachments.Request().GetAsync();
            return View(attachments.CurrentPage);
        }

        // POST: Calendar/CreateAttachment
        [HttpPost]
        public async Task<IActionResult> CreateAttachment(string eventId, string name, string content)
        {
            var attachment = new FileAttachment
            {
                Name = name,
                ContentBytes = System.Text.Encoding.UTF8.GetBytes(content),
                ContentType = "text/plain"
            };

            await _graphServiceClient.Me.Events[eventId].Attachments.Request().AddAsync(attachment);
            return RedirectToAction(nameof(GetAttachments), new { eventId });
        }

        // POST: Calendar/DeleteAttachment
        [HttpPost]
        public async Task<IActionResult> DeleteAttachment(string eventId, string attachmentId)
        {
            await _graphServiceClient.Me.Events[eventId].Attachments[attachmentId].Request().DeleteAsync();
            return RedirectToAction(nameof(GetAttachments), new { eventId });
        }

        // GET: Calendar/GetReminders
        public async Task<IActionResult> GetReminders()
        {
            var reminders = await _graphServiceClient.Me.ReminderView("startDateTime", "endDateTime").Request().GetAsync();
            return View(reminders.CurrentPage);
        }

        // POST: Calendar/SnoozeReminder
        [HttpPost]
        public async Task<IActionResult> SnoozeReminder(string eventId, DateTime newTime)
        {
            var snoozeReminder = new SnoozeReminder
            {
                NewReminderTime = new DateTimeTimeZone
                {
                    DateTime = newTime.ToString("o"),
                    TimeZone = "UTC"
                }
            };

            await _graphServiceClient.Me.Events[eventId].SnoozeReminder(snoozeReminder).Request().PostAsync();
            return RedirectToAction(nameof(GetReminders));
        }

        // POST: Calendar/DismissReminder
        [HttpPost]
        public async Task<IActionResult> DismissReminder(string eventId)
        {
            await _graphServiceClient.Me.Events[eventId].DismissReminder().Request().PostAsync();
            return RedirectToAction(nameof(GetReminders));
        }

        // GET: Calendar/GetCalendars
        public async Task<IActionResult> GetCalendars()
        {
            var calendars = await _graphServiceClient.Me.Calendars.Request().GetAsync();
            return View(calendars.CurrentPage);
        }

        // POST: Calendar/CreateCalendar
        [HttpPost]
        public async Task<IActionResult> CreateCalendar(string name)
        {
            var calendar = new Calendar
            {
                Name = name
            };

            await _graphServiceClient.Me.Calendars.Request().AddAsync(calendar);
            return RedirectToAction(nameof(GetCalendars));
        }

        // POST: Calendar/UpdateCalendar
        [HttpPost]
        public async Task<IActionResult> UpdateCalendar(string id, string name)
        {
            var calendar = new Calendar
            {
                Name = name
            };

            await _graphServiceClient.Me.Calendars[id].Request().UpdateAsync(calendar);
            return RedirectToAction(nameof(GetCalendars));
        }

        // POST: Calendar/DeleteCalendar
        [HttpPost]
        public async Task<IActionResult> DeleteCalendar(string id)
        {
            await _graphServiceClient.Me.Calendars[id].Request().DeleteAsync();
            return RedirectToAction(nameof(GetCalendars));
        }
    }
}
