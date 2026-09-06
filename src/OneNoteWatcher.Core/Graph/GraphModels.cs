namespace OneNoteWatcher.Core.Graph;

public sealed record GraphNotebook(string Id, string DisplayName, DateTimeOffset? LastModified);

public sealed record GraphSection(
    string Id, string DisplayName, DateTimeOffset? LastModified,
    string? NotebookId, string? NotebookName, string? ClientUrl, string? WebUrl);

/// <summary>Server-side truth fetched from Microsoft Graph in one poll.</summary>
public sealed record GraphSnapshot(
    DateTimeOffset FetchedUtc,
    IReadOnlyList<GraphNotebook> Notebooks,
    IReadOnlyList<GraphSection> Sections);
