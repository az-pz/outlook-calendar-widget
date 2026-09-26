using System.Net;
using System.Text;
using OutlookCalendarWidget.Core;
using static OutlookCalendarWidget.Core.Tests.TestData;

namespace OutlookCalendarWidget.Core.Tests;

public class GraphCalendarClientTests
{
    private const string PageOne = """
    {
      "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#users('me')/calendarView",
      "@odata.nextLink": "https://graph.microsoft.com/v1.0/me/calendarView?$skip=2",
      "value": [
        {
          "id": "AAMk-1",
          "subject": "  Sprint planning ",
          "isAllDay": false,
          "isCancelled": false,
          "showAs": "busy",
          "webLink": "https://outlook.office365.com/owa/?itemid=AAMk-1",
          "start": { "dateTime": "2026-09-26T13:00:00.0000000", "timeZone": "Eastern Standard Time" },
          "end": { "dateTime": "2026-09-26T14:00:00.0000000", "timeZone": "Eastern Standard Time" },
          "location": { "displayName": "Building 7 / Room 1201" },
          "responseStatus": { "response": "accepted", "time": "0001-01-01T00:00:00Z" },
          "onlineMeeting": { "joinUrl": "https://teams.microsoft.com/l/meetup-join/abc" },
          "onlineMeetingUrl": null
        },
        {
          "id": "AAMk-2",
          "subject": "Company holiday",
          "isAllDay": true,
          "isCancelled": false,
          "showAs": "oof",
          "start": { "dateTime": "2026-09-28T00:00:00.0000000", "timeZone": "Eastern Standard Time" },
          "end": { "dateTime": "2026-09-29T00:00:00.0000000", "timeZone": "Eastern Standard Time" },
          "location": { "displayName": "" },
          "responseStatus": { "response": "organizer" },
          "onlineMeeting": null
        }
      ]
    }
    """;

    private const string PageTwo = """
    {
      "value": [
        {
          "id": "AAMk-3",
          "subject": "",
          "isAllDay": false,
          "isCancelled": true,
          "showAs": "tentative",
          "start": { "dateTime": "2026-09-26T18:30:00.0000000", "timeZone": "UTC" },
          "end": { "dateTime": "2026-09-26T19:00:00.0000000", "timeZone": "UTC" },
          "responseStatus": { "response": "tentativelyAccepted" },
          "onlineMeetingUrl": "https://legacy.example/join"
        }
      ]
    }
    """;

    [Fact]
    public void Parser_maps_graph_events_to_local_time()
    {
        var (events, nextLink) = GraphCalendarParser.ParsePage(PageOne, Eastern);

        Assert.Equal("https://graph.microsoft.com/v1.0/me/calendarView?$skip=2", nextLink);
        Assert.Equal(2, events.Count);

        var planning = events[0];
        Assert.Equal("AAMk-1", planning.Id);
        Assert.Equal("Sprint planning", planning.Subject);
        Assert.Equal(At(26, 13), planning.Start);
        Assert.Equal(TimeSpan.FromHours(-4), planning.Start.Offset);
        Assert.Equal(At(26, 14), planning.End);
        Assert.Equal("Building 7 / Room 1201", planning.Location);
        Assert.Equal(FreeBusyStatus.Busy, planning.ShowAs);
        Assert.Equal(MeetingResponse.Accepted, planning.Response);
        Assert.Equal("https://teams.microsoft.com/l/meetup-join/abc", planning.JoinUrl);
        Assert.Equal("https://outlook.office365.com/owa/?itemid=AAMk-1", planning.WebLink);

        var holiday = events[1];
        Assert.True(holiday.IsAllDay);
        Assert.Null(holiday.Location);
        Assert.Null(holiday.JoinUrl);
        Assert.Null(holiday.WebLink);
        Assert.Equal(FreeBusyStatus.Oof, holiday.ShowAs);
        Assert.Equal(MeetingResponse.Organizer, holiday.Response);
        Assert.Equal(new DateTimeOffset(2026, 9, 28, 0, 0, 0, TimeSpan.FromHours(-4)), holiday.Start);
    }

    [Fact]
    public void Parser_converts_other_time_zones_and_uses_legacy_join_url()
    {
        var (events, nextLink) = GraphCalendarParser.ParsePage(PageTwo, Eastern);

        Assert.Null(nextLink);
        var evt = Assert.Single(events);
        Assert.Equal(At(26, 14, 30), evt.Start);
        Assert.Equal(TimeSpan.FromHours(-4), evt.Start.Offset);
        Assert.True(evt.IsCancelled);
        Assert.Equal(string.Empty, evt.Subject);
        Assert.Equal(FreeBusyStatus.Tentative, evt.ShowAs);
        Assert.Equal(MeetingResponse.TentativelyAccepted, evt.Response);
        Assert.Equal("https://legacy.example/join", evt.JoinUrl);
    }

    [Fact]
    public async Task Client_sends_token_and_timezone_and_follows_next_links()
    {
        var handler = new StubHandler(request => request.RequestUri!.Query.Contains("$skip=2", StringComparison.Ordinal)
            ? Json(HttpStatusCode.OK, PageTwo)
            : Json(HttpStatusCode.OK, PageOne));
        using var http = new HttpClient(handler);
        var client = new GraphCalendarClient(http, _ => Task.FromResult("token-123"));

        var events = await client.GetEventsAsync(At(26, 0), At(27, 0), Eastern, TestContext.Current.CancellationToken);

        Assert.Equal(3, events.Count);
        Assert.Equal(2, handler.Requests.Count);

        var first = handler.Requests[0];
        Assert.Equal("Bearer", first.Headers.Authorization!.Scheme);
        Assert.Equal("token-123", first.Headers.Authorization.Parameter);
        Assert.Equal("outlook.timezone=\"Eastern Standard Time\"", first.Headers.GetValues("Prefer").Single());
        var url = first.RequestUri!.ToString();
        Assert.StartsWith("https://graph.microsoft.com/v1.0/me/calendarView?", url, StringComparison.Ordinal);
        Assert.Contains("startDateTime=2026-09-26T04:00:00Z", Uri.UnescapeDataString(url), StringComparison.Ordinal);
        Assert.Contains("endDateTime=2026-09-27T04:00:00Z", Uri.UnescapeDataString(url), StringComparison.Ordinal);
        Assert.Contains("$orderby=start/dateTime", url, StringComparison.Ordinal);
        Assert.Contains("onlineMeeting", url, StringComparison.Ordinal);
    }

    [Fact]
    public async Task Client_maps_unauthorized_to_authentication_exception()
    {
        using var http = new HttpClient(new StubHandler(_ => Json(HttpStatusCode.Unauthorized, "{}")));
        var client = new GraphCalendarClient(http, _ => Task.FromResult("expired"));

        await Assert.ThrowsAsync<GraphAuthenticationException>(
            () => client.GetEventsAsync(At(26, 0), At(27, 0), Eastern, TestContext.Current.CancellationToken));
    }

    [Fact]
    public async Task Client_surfaces_graph_error_messages()
    {
        const string error = """{ "error": { "code": "ErrorAccessDenied", "message": "Access is denied. Check credentials and try again." } }""";
        using var http = new HttpClient(new StubHandler(_ => Json(HttpStatusCode.Forbidden, error)));
        var client = new GraphCalendarClient(http, _ => Task.FromResult("token"));

        var ex = await Assert.ThrowsAsync<GraphRequestException>(
            () => client.GetEventsAsync(At(26, 0), At(27, 0), Eastern, TestContext.Current.CancellationToken));

        Assert.Equal(HttpStatusCode.Forbidden, ex.StatusCode);
        Assert.Equal("Access is denied. Check credentials and try again.", ex.Message);
    }

    [Fact]
    public async Task Client_reports_status_code_when_error_body_is_not_json()
    {
        using var http = new HttpClient(new StubHandler(_ => Json(HttpStatusCode.ServiceUnavailable, "<html>oops</html>")));
        var client = new GraphCalendarClient(http, _ => Task.FromResult("token"));

        var ex = await Assert.ThrowsAsync<GraphRequestException>(
            () => client.GetEventsAsync(At(26, 0), At(27, 0), Eastern, TestContext.Current.CancellationToken));

        Assert.Equal("Microsoft Graph returned 503 (ServiceUnavailable).", ex.Message);
    }

    private static HttpResponseMessage Json(HttpStatusCode status, string body) =>
        new(status) { Content = new StringContent(body, Encoding.UTF8, "application/json") };

    private sealed class StubHandler(Func<HttpRequestMessage, HttpResponseMessage> respond) : HttpMessageHandler
    {
        public List<HttpRequestMessage> Requests { get; } = [];

        protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
        {
            Requests.Add(request);
            return Task.FromResult(respond(request));
        }
    }
}
