using System.Security.Cryptography;
using Microsoft.Identity.Client;

namespace OneNoteWatcher;

/// <summary>Why the Microsoft Graph check can or cannot run.</summary>
public enum GraphAuthState { Unknown, SignedIn, NeverSignedIn, Expired }

/// <summary>
/// MSAL public-client auth for Microsoft Graph (delegated Notes.Read, personal account, device-code
/// flow). Token cache is DPAPI-protected per user. See docs/setup-graph.md.
/// </summary>
public sealed class GraphAuth
{
    private static readonly string[] Scopes = ["Notes.Read"];
    private readonly IPublicClientApplication _app;
    private readonly string _cachePath;

    public GraphAuth(string clientId, string tenant)
    {
        _cachePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "OneNoteWatcher", "msal.cache");
        Directory.CreateDirectory(Path.GetDirectoryName(_cachePath)!);

        _app = PublicClientApplicationBuilder.Create(clientId)
            .WithAuthority($"https://login.microsoftonline.com/{tenant}")
            .WithDefaultRedirectUri()
            .Build();

        _app.UserTokenCache.SetBeforeAccess(a =>
        {
            try
            {
                if (File.Exists(_cachePath))
                    a.TokenCache.DeserializeMsalV3(ProtectedData.Unprotect(File.ReadAllBytes(_cachePath), null, DataProtectionScope.CurrentUser));
            }
            catch (Exception) { /* corrupt cache → start fresh */ }
        });
        _app.UserTokenCache.SetAfterAccess(a =>
        {
            if (!a.HasStateChanged) return;
            try { File.WriteAllBytes(_cachePath, ProtectedData.Protect(a.TokenCache.SerializeMsalV3(), null, DataProtectionScope.CurrentUser)); }
            catch (Exception) { }
        });
    }

    public async Task<bool> IsSignedInAsync()
    {
        var accounts = await _app.GetAccountsAsync();
        return accounts.Any();
    }

    /// <summary>
    /// Why the cloud check is unavailable. An account that exists but can no longer get a token is an
    /// EXPIRED session — the watcher was working and went blind, which must alarm, not whisper.
    /// </summary>
    public GraphAuthState State { get; private set; } = GraphAuthState.Unknown;

    /// <summary>Silent token or null (never prompts). Used by the background poll.</summary>
    public async Task<string?> TryGetTokenSilentAsync(CancellationToken ct)
    {
        var acct = (await _app.GetAccountsAsync()).FirstOrDefault();
        if (acct is null) { State = GraphAuthState.NeverSignedIn; return null; }
        try
        {
            var token = (await _app.AcquireTokenSilent(Scopes, acct).ExecuteAsync(ct)).AccessToken;
            State = GraphAuthState.SignedIn;
            return token;
        }
        catch (MsalUiRequiredException)
        {
            // the refresh token has expired or consent was revoked: we HAD access and lost it
            State = GraphAuthState.Expired;
            return null;
        }
        catch (MsalServiceException)
        {
            State = GraphAuthState.Expired;
            return null;
        }
    }

    /// <summary>Interactive device-code sign-in; <paramref name="showCode"/> presents URL + code to the user.</summary>
    public async Task<string> SignInAsync(Action<string, string> showCode, CancellationToken ct)
    {
        var result = await _app.AcquireTokenWithDeviceCode(Scopes, dc =>
        {
            showCode(dc.VerificationUrl, dc.UserCode);
            return Task.CompletedTask;
        }).ExecuteAsync(ct);
        return result.AccessToken;
    }

    public async Task SignOutAsync()
    {
        foreach (var a in await _app.GetAccountsAsync()) await _app.RemoveAsync(a);
        try { File.Delete(_cachePath); } catch (IOException) { }
    }
}
