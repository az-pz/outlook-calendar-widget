using Microsoft.Identity.Client;
using Microsoft.Identity.Client.Broker;
using Microsoft.Identity.Client.Extensions.Msal;
using OutlookCalendarWidget.Core;

namespace OutlookCalendarWidget.Services;

/// <summary>
/// Microsoft identity sign-in through MSAL and the Windows Web Account Manager (WAM) broker.
/// WAM gives single sign-on with the account already signed in to Windows and keeps refresh tokens in the OS.
/// </summary>
[System.Diagnostics.CodeAnalysis.SuppressMessage("Design", "CA1001:Types that own disposable fields should be disposable", Justification = "Process-lifetime singleton.")]
internal sealed class AuthService
{
    private static readonly string[] Scopes = [GraphCalendarClient.GraphScope];

    private readonly SemaphoreSlim gate = new(1, 1);
    private IPublicClientApplication? app;
    private string? appKey;

    public static AuthService Shared { get; } = new();

    /// <summary>Returns an access token without any UI, or <c>null</c> when the user needs to sign in interactively.</summary>
    public async Task<AuthenticationResult?> TryAcquireTokenSilentAsync(AppSettings settings, CancellationToken cancellationToken)
    {
        var pca = await GetApplicationAsync(settings);
        var account = (await pca.GetAccountsAsync()).FirstOrDefault();
        if (account is null)
        {
            if (settings.SignedOut)
            {
                return null;
            }

            // No account used by this app yet: try SSO with the account signed in to Windows.
            account = PublicClientApplication.OperatingSystemAccount;
        }

        try
        {
            return await pca.AcquireTokenSilent(Scopes, account).ExecuteAsync(cancellationToken);
        }
        catch (MsalUiRequiredException)
        {
            return null;
        }
    }

    public async Task<AuthenticationResult> SignInAsync(AppSettings settings, nint parentWindow, CancellationToken cancellationToken)
    {
        var pca = await GetApplicationAsync(settings);
        var result = await pca.AcquireTokenInteractive(Scopes)
            .WithParentActivityOrWindow(parentWindow)
            .WithPrompt(Prompt.SelectAccount)
            .ExecuteAsync(cancellationToken);

        if (settings.SignedOut)
        {
            (settings with { SignedOut = false }).Save();
        }

        return result;
    }

    public async Task SignOutAsync(AppSettings settings)
    {
        if (settings.IsConfigured)
        {
            var pca = await GetApplicationAsync(settings);
            foreach (var account in await pca.GetAccountsAsync())
            {
                await pca.RemoveAsync(account);
            }
        }

        (settings with { SignedOut = true }).Save();
    }

    public async Task<IAccount?> GetSignedInAccountAsync(AppSettings settings)
    {
        if (!settings.IsConfigured)
        {
            return null;
        }

        var pca = await GetApplicationAsync(settings);
        return (await pca.GetAccountsAsync()).FirstOrDefault();
    }

    private async Task<IPublicClientApplication> GetApplicationAsync(AppSettings settings)
    {
        if (!settings.IsConfigured)
        {
            throw new InvalidOperationException("An Entra ID application (client) ID has not been configured.");
        }

        var key = $"{settings.ClientId}|{settings.TenantId}";
        await gate.WaitAsync();
        try
        {
            if (app is not null && appKey == key)
            {
                return app;
            }

            var pca = PublicClientApplicationBuilder.Create(settings.ClientId)
                .WithAuthority(AzureCloudInstance.AzurePublic, settings.TenantId)
                .WithRedirectUri($"ms-appx-web://microsoft.aad.brokerplugin/{settings.ClientId}")
                .WithBroker(new BrokerOptions(BrokerOptions.OperatingSystems.Windows)
                {
                    Title = "Outlook Calendar Widget",
                    ListOperatingSystemAccounts = true,
                })
                .Build();

            // Persist account metadata (encrypted with DPAPI) so both the app and the widget provider process see the
            // same signed-in account. With WAM, the refresh tokens themselves stay in the OS broker.
            var storage = new StorageCreationPropertiesBuilder($"msal-{settings.ClientId}.cache", AppSettings.DataFolder).Build();
            var cacheHelper = await MsalCacheHelper.CreateAsync(storage);
            cacheHelper.RegisterCache(pca.UserTokenCache);

            app = pca;
            appKey = key;
            return pca;
        }
        finally
        {
            gate.Release();
        }
    }
}
