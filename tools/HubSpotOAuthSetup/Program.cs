using System.Diagnostics;
using System.Net;
using System.Net.Http.Json;
using System.Text.Json;
using System.Text.Json.Serialization;
using Microsoft.Extensions.Configuration;
using ProjectSync.Options;

// -----------------------------------------------------------------------------
// HubSpot OAuth setup — ONE-TIME. Captures a long-lived refresh token.
//
// Runs the authorization-code flow: opens the HubSpot consent page, catches the
// redirect on a local listener, exchanges the code, and prints the refresh token
// to paste into local.settings.json ("HubSpot:RefreshToken").
//
// Prerequisites (in a HubSpot developer account → your app → Auth tab):
//   • Client ID + Client Secret  → HubSpot:ClientId / HubSpot:ClientSecret
//   • Redirect URL must include   http://localhost:5127/callback  (or your HubSpot:RedirectUri)
//   • Scopes must include the ones in HubSpot:Scopes
// -----------------------------------------------------------------------------

var localSettingsPath = Path.GetFullPath(Path.Combine(
    AppContext.BaseDirectory, "..", "..", "..", "..", "..",
    "src", "ProjectSync.Functions", "local.settings.json"));

var configuration = new ConfigurationBuilder()
    .AddInMemoryCollection(LoadFunctionsValues(localSettingsPath))
    .AddEnvironmentVariables()
    .Build();

var options = new HubSpotOptions();
configuration.GetSection(HubSpotOptions.SectionName).Bind(options);

if (string.IsNullOrWhiteSpace(options.ClientId) || string.IsNullOrWhiteSpace(options.ClientSecret))
{
    Console.Error.WriteLine("❌ HubSpot:ClientId / HubSpot:ClientSecret not set.");
    Console.Error.WriteLine($"   Add them to the 'Values' in {localSettingsPath} (from your HubSpot app's Auth tab).");
    return 1;
}

var redirectUri = options.RedirectUri;
var scope = string.Join(' ', options.Scopes);
var authorizeUrl =
    "https://app.hubspot.com/oauth/authorize" +
    $"?client_id={Uri.EscapeDataString(options.ClientId)}" +
    $"&redirect_uri={Uri.EscapeDataString(redirectUri)}" +
    $"&scope={Uri.EscapeDataString(scope)}";

// Listen on the redirect origin (root, so the exact callback path matches regardless of trailing slash).
var origin = new Uri(redirectUri).GetLeftPart(UriPartial.Authority) + "/";
using var listener = new HttpListener();
listener.Prefixes.Add(origin);
try
{
    listener.Start();
}
catch (Exception ex)
{
    Console.Error.WriteLine($"❌ Could not start local listener on {origin}: {ex.Message}");
    Console.Error.WriteLine("   Ensure HubSpot:RedirectUri uses http://localhost:<port>/... and the port is free.");
    return 2;
}

Console.WriteLine("=== HubSpot OAuth setup ===");
Console.WriteLine($"Scopes  : {scope}");
Console.WriteLine($"Redirect: {redirectUri}");
Console.WriteLine();
Console.WriteLine("Opening the HubSpot consent page in your browser…");
Console.WriteLine("If it doesn't open, paste this URL manually:");
Console.WriteLine();
Console.WriteLine("  " + authorizeUrl);
Console.WriteLine();
try
{
    Process.Start(new ProcessStartInfo(authorizeUrl) { UseShellExecute = true });
}
catch { /* headless: user pastes manually */ }

Console.WriteLine("Waiting for the redirect (approve access, and select your account)…");

var context = await listener.GetContextAsync();
var code = context.Request.QueryString["code"];
var error = context.Request.QueryString["error"];

// Respond in the browser so the user knows they can close the tab.
var html = error is null
    ? "<h2>✅ Authorized. You can close this tab and return to the terminal.</h2>"
    : $"<h2>❌ Authorization failed: {WebUtility.HtmlEncode(error)}</h2>";
var buffer = System.Text.Encoding.UTF8.GetBytes($"<html><body style='font-family:sans-serif'>{html}</body></html>");
context.Response.ContentType = "text/html";
context.Response.OutputStream.Write(buffer);
context.Response.OutputStream.Close();
listener.Stop();

if (!string.IsNullOrEmpty(error))
{
    Console.Error.WriteLine($"❌ HubSpot returned an error: {error}");
    return 3;
}
if (string.IsNullOrEmpty(code))
{
    Console.Error.WriteLine("❌ No authorization code in the redirect.");
    return 3;
}

Console.WriteLine("→ Exchanging the authorization code for tokens…");
using var http = new HttpClient { Timeout = TimeSpan.FromSeconds(options.TimeoutSeconds) };
var form = new Dictionary<string, string>
{
    ["grant_type"] = "authorization_code",
    ["client_id"] = options.ClientId,
    ["client_secret"] = options.ClientSecret,
    ["redirect_uri"] = redirectUri,
    ["code"] = code!,
};
using var response = await http.PostAsync(
    $"{options.BaseUrl.TrimEnd('/')}/oauth/v1/token", new FormUrlEncodedContent(form));
var respBody = await response.Content.ReadAsStringAsync();
if (!response.IsSuccessStatusCode)
{
    Console.Error.WriteLine($"❌ Token exchange failed ({(int)response.StatusCode}): {respBody}");
    return 4;
}

var token = JsonSerializer.Deserialize<TokenResponse>(respBody);
if (token is null || string.IsNullOrEmpty(token.RefreshToken))
{
    Console.Error.WriteLine("❌ Token exchange returned no refresh token.");
    return 4;
}

Console.WriteLine();
Console.WriteLine("✅ Success. Add this to local.settings.json \"Values\" (git-ignored — never commit):");
Console.WriteLine();
Console.WriteLine($"    \"HubSpot:RefreshToken\": \"{token.RefreshToken}\"");
Console.WriteLine();
Console.WriteLine($"(access token valid ~{token.ExpiresIn}s; the app refreshes it automatically from now on.)");
Console.WriteLine("Then run:  dotnet run --project tools/HubSpotConnectivityTest");
return 0;

static Dictionary<string, string?> LoadFunctionsValues(string path)
{
    var result = new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase);
    if (!File.Exists(path)) return result;
    using var doc = JsonDocument.Parse(File.ReadAllText(path));
    if (doc.RootElement.TryGetProperty("Values", out var values))
    {
        foreach (var prop in values.EnumerateObject())
        {
            result[prop.Name] = prop.Value.GetString();
        }
    }
    return result;
}

sealed record TokenResponse
{
    [JsonPropertyName("refresh_token")] public string RefreshToken { get; init; } = string.Empty;
    [JsonPropertyName("access_token")] public string AccessToken { get; init; } = string.Empty;
    [JsonPropertyName("expires_in")] public int ExpiresIn { get; init; }
}
