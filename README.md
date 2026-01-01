# <svg fill="currentColor" width="32" height="32" viewBox="0 0 20 20" xmlns="http://www.w3.org/2000/svg"><path d="M4.5 17A1.5 1.5 0 0 1 3 15.65V4.5c0-.78.6-1.42 1.36-1.5H9c.78 0 1.42.6 1.5 1.36v.4l2.19-2.27a1.5 1.5 0 0 1 2-.14l.12.1 2.76 2.72c.55.55.6 1.42.11 2.02l-.1.11-2.31 2.2h.23c.78 0 1.42.6 1.5 1.36v4.64c0 .78-.6 1.42-1.36 1.5H4.5Zm5-6.5H4v5c0 .22.14.4.33.47l.08.02.09.01h5v-5.5Zm6 0h-5V16h5a.5.5 0 0 0 .5-.4V11a.5.5 0 0 0-.41-.5h-.09Zm-5-2.79V9.5h1.79L10.5 7.71ZM9 4.01H4.5a.5.5 0 0 0-.5.4v5.1h5.5v-5a.5.5 0 0 0-.33-.48l-.08-.02H9Z" fill="currentColor"/></svg> My Enterprise Apps

## Summary

A SharePoint Framework webpart that displays enterprise applications from Microsoft Entra ID (formerly Azure AD) that the current user has access to. This solution leverages the Microsoft Graph API to dynamically retrieve and display Entra enterprise apps with customizable sorting, icon sizes, and the ability to show or hide hidden applications.

## Screenshot

![My Enterprise Apps webpart in SharePoint](docs/images/My-apps-screenshot.png "My Enterprise Apps webpart in SharePoint")

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
| 0.5.0.0 | Current          | Initial release |

## Used SharePoint Framework Version

![version](https://img.shields.io/badge/version-1.22.0-green.svg)

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
  - `heft start`

Other build commands can be listed using `heft --help`.

To build the solution for production:
- `npm run build`

## Microsoft Graph Permissions

This solution requires the following Microsoft Graph permissions:

- `User.Read` - To read the current user's profile
- `Application.Read.All` - To read all enterprise applications (requires admin consent)
  
`User.Read` is covered by basic sign-in/profile consent; `Application.Read.All` must be approved by a tenant admin via the API access page or enterprise app consent.

## References

- [Getting started with SharePoint Framework](https://docs.microsoft.com/sharepoint/dev/spfx/set-up-your-developer-tenant)
- [Building for Microsoft Teams](https://docs.microsoft.com/sharepoint/dev/spfx/build-for-teams-overview)
- [Use Microsoft Graph in your solution](https://docs.microsoft.com/sharepoint/dev/spfx/web-parts/get-started/using-microsoft-graph-apis)
- [Publish SharePoint Framework applications to the Marketplace](https://docs.microsoft.com/sharepoint/dev/spfx/publish-to-marketplace-overview)
- [Microsoft 365 Patterns and Practices](https://aka.ms/m365pnp) - Guidance, tooling, samples and open-source controls for your Microsoft 365 development
- [Heft Documentation](https://heft.rushstack.io/)
- [Microsoft Entra ID Documentation](https://learn.microsoft.com/en-us/entra/)

Icon from [Microsoft Fluent UI System Icons](https://github.com/microsoft/fluentui-system-icons) (MIT License)