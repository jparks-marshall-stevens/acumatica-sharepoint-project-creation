using System.Security.Cryptography.X509Certificates;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using Microsoft.SharePoint.Client;
using PnP.Framework;
using ProjectSync.Options;

namespace ProjectSync.SharePoint;

/// <summary>
/// Builds authenticated CSOM <see cref="ClientContext"/> instances using Azure AD
/// app-only certificate authentication (PnP.Framework).
/// </summary>
public sealed class SharePointContextFactory
{
    private readonly SharePointOptions _options;
    private readonly ILogger<SharePointContextFactory> _logger;
    private readonly Lazy<X509Certificate2> _certificate;

    public SharePointContextFactory(IOptions<SharePointOptions> options, ILogger<SharePointContextFactory> logger)
    {
        _options = options.Value;
        _logger = logger;
        _certificate = new Lazy<X509Certificate2>(LoadCertificate);
    }

    public Task<ClientContext> CreateContextAsync(string siteUrl)
    {
        var authManager = new AuthenticationManager(_options.ClientId, _certificate.Value, _options.AzureAdTenant);
        return authManager.GetContextAsync(siteUrl);
    }

    private X509Certificate2 LoadCertificate()
    {
        // Option 1: certificate supplied inline as a base64 PFX.
        if (!string.IsNullOrWhiteSpace(_options.CertificateBase64))
        {
            _logger.LogDebug("Loading SharePoint certificate from inline base64 PFX.");
            var raw = Convert.FromBase64String(_options.CertificateBase64);
            return X509CertificateLoader.LoadPkcs12(
                raw,
                _options.CertificatePassword,
                X509KeyStorageFlags.EphemeralKeySet | X509KeyStorageFlags.Exportable);
        }

        // Option 2: certificate resolved from a store by thumbprint.
        if (!string.IsNullOrWhiteSpace(_options.CertificateThumbprint))
        {
            var thumbprint = _options.CertificateThumbprint!.Replace(" ", string.Empty).ToUpperInvariant();
            foreach (var location in new[] { StoreLocation.CurrentUser, StoreLocation.LocalMachine })
            {
                using var store = new X509Store(StoreName.My, location);
                store.Open(OpenFlags.ReadOnly);
                var found = store.Certificates.Find(X509FindType.FindByThumbprint, thumbprint, validOnly: false);
                if (found.Count > 0)
                {
                    _logger.LogDebug("Loaded SharePoint certificate {Thumbprint} from {Location}.", thumbprint, location);
                    return found[0];
                }
            }

            throw new InvalidOperationException(
                $"Certificate with thumbprint '{thumbprint}' was not found in CurrentUser or LocalMachine 'My' stores.");
        }

        throw new InvalidOperationException(
            "No SharePoint certificate configured. Set SharePoint:CertificateBase64 (+Password) or SharePoint:CertificateThumbprint.");
    }
}
