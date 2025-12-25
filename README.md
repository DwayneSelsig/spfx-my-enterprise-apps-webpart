# spfx-entra-enterprise-apps

## Summary

A SharePoint Framework webpart that displays enterprise applications from Microsoft Entra ID (formerly Azure AD) that the current user has access to. This solution leverages the Microsoft Graph API to dynamically retrieve and display Entra enterprise apps with customizable sorting, icon sizes, and the ability to show or hide hidden applications.

## Screenshot

![Entra Enterprise Apps webpart in SharePoint (My Apps)](docs/images/My-apps-screenshot.png "Entra Enterprise Apps webpart in SharePoint (My Apps)")

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

## Solution

| Solution    | Author(s)                                               |
| ----------- | ------------------------------------------------------- |
| spfx-my-enterprise-apps-webpart | Dwayne Selsig |

## Version history

| Version | Date             | Comments        |
| ------- | ---------------- | --------------- |
| 0.5.0.0 | Current          | Initial release |

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

## Features

The Entra Enterprise Apps webpart provides the following functionality:

This webpart illustrates the following concepts:

- Integration with Microsoft Graph API to retrieve enterprise applications
- Configurable icon sizes (small, normal, large, huge)
- Option to show or hide hidden applications
- Display of Entra ID applications with custom sorting capabilities
- Localization support (English and Dutch)
- Custom property pane configuration for administrators

## Configuration

The webpart can be configured through the property pane with the following options:

- **Title**: Custom title for the webpart
- **Sort Order**: Define a custom sort order for applications (supports comma-separated list)
- **Show Hidden Apps**: Toggle to display or hide hidden enterprise applications
- **Icon Size**: Choose from small, normal, large, or huge icon sizes

## Microsoft Graph Permissions

This solution requires the following Microsoft Graph permissions:

- `User.Read` - To read the current user's profile
- `Application.Read.All` - To read all enterprise applications

## Contributing

We welcome contributions from the community! Here are some ways you can help:

- **Translations**: Help translate the webpart into additional languages. The current supported languages are English and Dutch. If you'd like to contribute translations, please submit a pull request with the updated localization files in the `loc` folder.
- **Feature Suggestions**: Have an idea for a new feature or improvement? Please open an issue to share your suggestion. We'd love to hear about features you'd like to see in the Entra Enterprise Apps webpart.

## References

- [Getting started with SharePoint Framework](https://docs.microsoft.com/sharepoint/dev/spfx/set-up-your-developer-tenant)
- [Building for Microsoft Teams](https://docs.microsoft.com/sharepoint/dev/spfx/build-for-teams-overview)
- [Use Microsoft Graph in your solution](https://docs.microsoft.com/sharepoint/dev/spfx/web-parts/get-started/using-microsoft-graph-apis)
- [Publish SharePoint Framework applications to the Marketplace](https://docs.microsoft.com/sharepoint/dev/spfx/publish-to-marketplace-overview)
- [Microsoft 365 Patterns and Practices](https://aka.ms/m365pnp) - Guidance, tooling, samples and open-source controls for your Microsoft 365 development
- [Heft Documentation](https://heft.rushstack.io/)
- [Microsoft Entra ID Documentation](https://learn.microsoft.com/en-us/entra/)