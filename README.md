# <picture><source media="(prefers-color-scheme: dark)" srcset="docs/images/icon-dark.svg"><img src="docs/images/icon.svg" alt="icon" width="32" height="32"></picture> My Enterprise Apps

## Summary

A SharePoint Framework webpart that displays enterprise applications from Microsoft Entra ID (formerly Azure AD) that the current user has access to. This solution leverages the Microsoft Graph API to dynamically retrieve and display Entra enterprise apps with customizable sorting, icon sizes, and the ability to show or hide hidden applications.

## Screenshot

![My Enterprise Apps webpart in SharePoint](docs/images/My-apps-screenshot.png "My Enterprise Apps webpart in SharePoint")

## Video

https://github.com/user-attachments/assets/6cd1f83f-80c3-4986-b193-1923b633b056

## Features

The Entra Enterprise Apps webpart provides the following functionality:

- Integration with Microsoft Graph API to retrieve enterprise applications
- Configurable icon sizes (small, normal, large, huge)
- Option to show or hide hidden applications
- Display of Entra ID applications with custom sorting capabilities
- Localization support (English and Dutch)
- Custom property pane configuration for administrators

## Configuration

The webpart can be configured through the property pane with the following options:

- **Title**: Custom title for the webpart
- **Sort Order**: Enter app-name keywords (one per line) to prioritize apps matching those terms
- **Show Hidden Apps**: Toggle to display or hide hidden enterprise applications
- **Icon Size**: Choose from small, normal, large, or huge icon sizes

## Installation and Upgrades

### Download or compile
[Download the latest release](https://github.com/DwayneSelsig/spfx-my-enterprise-apps-webpart/releases) or compile the solution (`npm run build`). The `.sppkg` file will be in `sharepoint/solution/`.

### Installation
Go to the [SharePoint admin center → **More features**](https://go.microsoft.com/fwlink/?linkid=2185077) → **Apps** → **Open** → **Upload** the `.sppkg` file. Approve Microsoft Graph permissions (`User.Read` and `Application.Read.All`) when prompted.

### Upgrades
Upload the new `.sppkg` file and overwrite the existing one when prompted.

> **Note:** SharePoint add-ins are being retired, but SharePoint Framework (SPFx) solutions like this one are not affected and remain fully supported.

For more information, see the SharePoint App Catalog documentation:
https://learn.microsoft.com/sharepoint/use-app-catalog

## Contributing

We welcome contributions from the community! Here are some ways you can help:

- **Translations**: Help translate the webpart into additional languages. The current supported languages are English and Dutch. If you'd like to contribute translations, please submit a pull request with the updated localization files in the `loc` folder.
- **Feature Suggestions**: Have an idea for a new feature or improvement? Please open an issue to share your suggestion. We'd love to hear about features you'd like to see in the Entra Enterprise Apps webpart.

## Solution

| Solution    | Author(s)                                               |
| ----------- | ------------------------------------------------------- |
| spfx-my-enterprise-apps-webpart | Dwayne Selsig |

## Version history

| Version | Date             | Comments        |
| ------- | ---------------- | --------------- |
| 0.5.0.0 | 2025-12-25          | Initial release |
| 0.6.0.0 | 2026-01-01          | Added Microsoft Apps |
| 0.6.1.0 | 2026-05-25       | Upgraded to SPFx 1.23.0 |
| 0.7.0.0 | 2026-08-27       | Added end-user filtering and custom app sizing; updated SPFx packages and Copilot logo |
| 0.8.0.0 | 2026-08-30       | Added app detail view, configurable default apps and property-pane toggles; improved default icons, Teams link and localization |
| 0.8.1.0 | 2026-08-30       | Added Entra ID and PowerShell setup instructions; improved enterprise application selection and privacy handling |

## Used SharePoint Framework Version

![version](https://img.shields.io/badge/version-1.23.0-green.svg)


## Adding an app from Microsoft Entra ID

This webpart displays the enterprise applications for which the current user has an app-role assignment. Creating an **App registration** by itself is therefore not enough: the app must have a corresponding **Enterprise application** (service principal) and the user, or a group of which the user is a direct member, must be assigned to it.

### Microsoft Entra admin center

Sign in to the [Microsoft Entra admin center](https://entra.microsoft.com) as a Cloud Application Administrator or Application Administrator, then:

1. Go to **Entra ID** > **Enterprise applications** > **All applications** > **New application**.
2. Select an application from the gallery, or select **Create your own application** for a custom application. Configure SSO when the application requires it.
3. Open the application and go to **Properties**:
   - Set **Enabled for users to sign in?** to **Yes**.
   - Set **Visible to users?** to **Yes**.
   - Upload a logo. Use a PNG of exactly **215 x 215 pixels**, no larger than **100 KB**. This is the icon shown by Entra and used by the webpart when available.
4. Go to **Users and groups** > **Add user/group**, select an existing user or—preferably—an existing security group, select the appropriate role (or **Default Access**) and choose **Assign**.
5. Next, go to **Entra ID** > **App registrations** > **All applications** and open the app registration that has the same **Application (client) ID** as the enterprise application. Then:
   - Under **Branding & properties**, enter the **Home page URL**.
   - Under **Authentication**, choose the appropriate platform (for example, **Web**) and add the required **Redirect URI**.
6. Ensure the webpart's Microsoft Graph requests are approved by a tenant admin in the SharePoint admin center: `User.Read` and `Application.Read.All`.

The **App registrations** step applies to custom or tenant-owned applications. Gallery applications usually have an app registration owned by the software vendor, so configure their sign-in URL and SSO settings in the enterprise application instead.

If **Show Hidden Apps** is disabled in the webpart properties, applications tagged `HideApp` are deliberately excluded. It can take a few minutes for a new Entra assignment to become visible.

### PowerShell

The examples below use the current `Microsoft.Entra` PowerShell module, rather than the retired `AzureAD` module. Replace values between angle brackets before running them. The signed-in account needs a suitable Entra administrator role, such as Cloud Application Administrator. These examples only use existing users and groups; they don't create either.

Install the module once, then connect with the required delegated permissions:

```powershell
Install-Module Microsoft.Entra -Scope CurrentUser
Connect-Entra -Scopes 'Application.ReadWrite.All','Application.Read.All','AppRoleAssignment.ReadWrite.All','User.Read.All','Group.Read.All'
```

Create the app registration and its enterprise application first. Then configure the homepage and redirect URI, upload the icon, and assign an existing user:

```powershell
$app = New-EntraApplication -DisplayName '<App display name>'

$servicePrincipal = New-EntraServicePrincipal `
  -AppId $app.AppId `
  -DisplayName $app.DisplayName `
  -Tags @('WindowsAzureActiveDirectoryIntegratedApp')

$web = @{
  homePageUrl = 'https://app.contoso.com'
  redirectUris = @('https://app.contoso.com/signin-oidc')
}

Set-EntraApplication -ApplicationId $app.Id -Web $web

Set-EntraServicePrincipal `
  -ServicePrincipalId $servicePrincipal.Id `
  -AccountEnabled $true `
  -AppRoleAssignmentRequired $true `
  -Homepage $web.homePageUrl `
  -ReplyUrls $web.redirectUris

Set-EntraApplicationLogo `
  -ApplicationId $app.Id `
  -FilePath 'C:\path\to\app-logo.png'

$user = Get-EntraUser -UserId '<user@contoso.com>'
New-EntraUserAppRoleAssignment `
  -UserId $user.Id `
  -PrincipalId $user.Id `
  -ResourceId $servicePrincipal.Id `
  -AppRoleId ([Guid]::Empty)
```

To assign the newly created enterprise application to an existing group instead of a user, replace the final user-assignment block with the following. Group assignment requires the appropriate Microsoft Entra ID licence, and nested group membership is not supported.

```powershell
$group = Get-EntraGroup -SearchString '<Group display name>'

New-EntraGroupAppRoleAssignment `
  -GroupId $group.Id `
  -PrincipalId $group.Id `
  -ResourceId $servicePrincipal.Id `
  -AppRoleId ([Guid]::Empty)
```

For an app with defined app roles, replace `[Guid]::Empty` with the ID of the role to assign. See the [enterprise-application properties](https://learn.microsoft.com/en-us/entra/identity/enterprise-apps/application-properties), [user/group assignment](https://learn.microsoft.com/en-us/entra/identity/enterprise-apps/assign-user-or-group-access-portal), and [Microsoft Entra PowerShell](https://learn.microsoft.com/en-us/powershell/entra-powershell/overview) documentation for details.

## Applies to

- [SharePoint Framework](https://aka.ms/spfx)
- [Microsoft 365 tenant](https://docs.microsoft.com/sharepoint/dev/spfx/set-up-your-developer-tenant)

> Get your own free development tenant by subscribing to [Microsoft 365 developer program](http://aka.ms/o365devprogram)

## Prerequisites

Before getting started, ensure your development environment is properly set up by following the [SharePoint Framework development environment setup guide](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-development-environment).

Additional requirements:
- Node.js version 22.14.0 or higher (and lower than 23.0.0)
- Appropriate Microsoft Graph permissions configured in the tenant
- Access to a SharePoint site where the webpart can be deployed

## Disclaimer

**THIS CODE IS PROVIDED _AS IS_ WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

---

## Minimal Path to Awesome

- Clone this repository
- Ensure that you are at the solution folder
- In the command-line run:
  - `npm install @rushstack/heft --global`
  - `npm install`
  - `npm run start`

Other build commands can be listed using `heft --help`.

To build the solution for production:
- `npm run build`

## Microsoft Graph Permissions

This solution requires the following Microsoft Graph permissions:

- `User.Read` - To read the current user's profile
- `Application.Read.All` - To read all enterprise applications (requires admin consent)
  
`User.Read` is covered by basic sign-in/profile consent; `Application.Read.All` must be approved by a tenant admin via the API access page or enterprise app consent.

## Enterprise application selection and privacy

`/me/appRoleAssignments` remains the source of truth for applications available to the current user, but it cannot identify Enterprise Applications that are present in the tenant and assigned to nobody. The webpart therefore also reads paged `servicePrincipals` results, restricted to both `servicePrincipalType eq 'Application'` and the `WindowsAzureActiveDirectoryIntegratedApp` tag. Both conditions are required: the type excludes `ServiceIdentity` (including Entra Agent Identities), while the tag excludes Microsoft infrastructure and backend service principals without relying on app names, publishers, or maintained blacklists.

For an Integrated App that is not in the current user's assignments, the webpart checks `appRoleAssignedTo`. Only a successful response with an empty `value` array means that the app is unassigned and can be shown. An app with any assignment (user, group, or service principal) is not shown to users who do not have it themselves.

These checks use Microsoft Graph `/$batch`, with at most 20 `appRoleAssignedTo` requests per batch. The first page requests only one assignment because the check only needs to distinguish empty from non-empty results. Batch subrequest failures and other unknown assignment states are handled fail-closed: the affected app is not shown. Existing `HideApp` and **Show Hidden Apps** behavior is applied after this selection.

There is no webpart-specific automated test suite yet. When validating a tenant manually, check that an app assigned to the current user and an unassigned Integrated App are shown; apps assigned only to another user or group are not shown; `HideApp` follows **Show Hidden Apps**; `ServiceIdentity` and non-tagged `Application` service principals are excluded; a candidate set of 20 and 21 apps produces one and two batches respectively; paged service-principal results are complete; and a failed batch subrequest does not expose its app.

## References

- [Getting started with SharePoint Framework](https://docs.microsoft.com/sharepoint/dev/spfx/set-up-your-developer-tenant)
- [Building for Microsoft Teams](https://docs.microsoft.com/sharepoint/dev/spfx/build-for-teams-overview)
- [Use Microsoft Graph in your solution](https://docs.microsoft.com/sharepoint/dev/spfx/web-parts/get-started/using-microsoft-graph-apis)
- [Publish SharePoint Framework applications to the Marketplace](https://docs.microsoft.com/sharepoint/dev/spfx/publish-to-marketplace-overview)
- [Microsoft 365 Patterns and Practices](https://aka.ms/m365pnp) - Guidance, tooling, samples and open-source controls for your Microsoft 365 development
- [Heft Documentation](https://heft.rushstack.io/)
- [Microsoft Entra ID Documentation](https://learn.microsoft.com/en-us/entra/)

Icon from [Microsoft Fluent UI System Icons](https://github.com/microsoft/fluentui-system-icons) (MIT License)
