using System.Net.Http.Headers;
using System.Text.Json;

namespace OneNoteWatcher.Core.Graph;

/// <summary>
/// Thin Microsoft Graph OneNote reader (delegated Notes.Read). Auth is injected as a token factory so
/// this stays platform-neutral and testable; the tray supplies MSAL. Endpoints per
/// docs/reference-data-sources.md §2. Handles paging and 429 back-off.
/// </summary>
public sealed class GraphClient
{
    private readonly HttpClient _http;
    private readonly Func<CancellationToken, Task<string>> _token;
    public const string BaseUrl = "https://graph.microsoft.com/v1.0";

    public GraphClient(Func<CancellationToken, Task<string>> tokenFactory, HttpClient? http = null)
    {
        _token = tokenFactory;
        _http = http ?? new HttpClient { Timeout = TimeSpan.FromSeconds(60) };
    }

    public async Task<GraphSnapshot> FetchAsync(CancellationToken ct)
    {
        var notebooks = new List<GraphNotebook>();
        await foreach (var el in PagedAsync($"{BaseUrl}/me/onenote/notebooks?$select=id,displayName,lastModifiedDateTime&$top=100", ct))
            notebooks.Add(new GraphNotebook(Str(el, "id")!, Str(el, "displayName") ?? "", Time(el, "lastModifiedDateTime")));

        var sections = new List<GraphSection>();
        await foreach (var el in PagedAsync(
            $"{BaseUrl}/me/onenote/sections?$select=id,displayName,lastModifiedDateTime,links,parentNotebook&$expand=parentNotebook($select=id,displayName)&$top=100", ct))
        {
            string? nbId = null, nbName = null, clientUrl = null, webUrl = null;
            if (el.TryGetProperty("parentNotebook", out var pn) && pn.ValueKind == JsonValueKind.Object)
            { nbId = Str(pn, "id"); nbName = Str(pn, "displayName"); }
            if (el.TryGetProperty("links", out var links) && links.ValueKind == JsonValueKind.Object)
            {
                if (links.TryGetProperty("oneNoteClientUrl", out var c) && c.ValueKind == JsonValueKind.Object) clientUrl = Str(c, "href");
                if (links.TryGetProperty("oneNoteWebUrl", out var w) && w.ValueKind == JsonValueKind.Object) webUrl = Str(w, "href");
            }
            sections.Add(new GraphSection(Str(el, "id")!, Str(el, "displayName") ?? "", Time(el, "lastModifiedDateTime"),
                nbId, nbName, clientUrl, webUrl));
        }
        return new GraphSnapshot(DateTimeOffset.UtcNow, notebooks, sections);
    }

    private async IAsyncEnumerable<JsonElement> PagedAsync(string url,
        [System.Runtime.CompilerServices.EnumeratorCancellation] CancellationToken ct)
    {
        string? next = url;
        while (next is not null)
        {
            using var doc = await GetJsonAsync(next, ct);
            if (doc.RootElement.TryGetProperty("value", out var arr) && arr.ValueKind == JsonValueKind.Array)
                foreach (var el in arr.EnumerateArray()) yield return el.Clone();
            next = doc.RootElement.TryGetProperty("@odata.nextLink", out var nl) ? nl.GetString() : null;
        }
    }

    private async Task<JsonDocument> GetJsonAsync(string url, CancellationToken ct)
    {
        for (var attempt = 0; ; attempt++)
        {
            using var req = new HttpRequestMessage(HttpMethod.Get, url);
            req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", await _token(ct));
            using var resp = await _http.SendAsync(req, ct);
            if ((int)resp.StatusCode == 429 || (int)resp.StatusCode >= 500)
            {
                if (attempt >= 3) resp.EnsureSuccessStatusCode();
                var wait = resp.Headers.RetryAfter?.Delta ?? TimeSpan.FromSeconds(Math.Pow(2, attempt + 1));
                await Task.Delay(wait, ct);
                continue;
            }
            resp.EnsureSuccessStatusCode();
            return JsonDocument.Parse(await resp.Content.ReadAsStringAsync(ct));
        }
    }

    private static string? Str(JsonElement e, string n) =>
        e.TryGetProperty(n, out var v) && v.ValueKind == JsonValueKind.String ? v.GetString() : null;

    private static DateTimeOffset? Time(JsonElement e, string n) =>
        DateTimeOffset.TryParse(Str(e, n), null, System.Globalization.DateTimeStyles.AssumeUniversal, out var t) ? t : null;
}
