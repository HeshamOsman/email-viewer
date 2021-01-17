# lazy-email-viewer

## Summary

This react webpart app lists the last 10 emails in your inbox using microsoft graph api, then connects to a third party api -spring boot app- for analyses.

Technologies and frameworks uesed in frontend:
1- Microsoft sharepoint framework
2- React with typescript
3- Office-fabric-react-ui library
4- CDN to serve html\css and js when deployed to sharepoint app-catalog

## Used SharePoint Framework Version

![version](https://img.shields.io/badge/version-1.11-green.svg)

## Applies to

- [SharePoint Framework](https://aka.ms/spfx)
- [Microsoft 365 tenant](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)

## Prerequisites

To run it localy:

1- Run (docker run --name email-ana -p 8080:8080 heshamosman28/email-analyser:latest) to get the backend app up and running
2- Install nodejs lts version 10 (the last supported verion in sharepoint -don't use version14-)
3- Install gulp (npm install gulp --global)
4- Clone this repository
5- Go to the application folder and run (npm install)
6- Run (gulp trust-dev-cert)
7- Run (gulp serve)
8- A new browser tab will open click the plus (+) button
9- Choose (lazy-email-viewer)
10- Now you should see a table with the emails and two boxes with the analysis

To run it on Sharepoint:

TBD.....

## References

- [Getting started with SharePoint Framework](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)
- [Building for Microsoft teams](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/build-for-teams-overview)
- [Use Microsoft Graph in your solution](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/get-started/using-microsoft-graph-apis)
- [Publish SharePoint Framework applications to the Marketplace](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/publish-to-marketplace-overview)
- [Microsoft 365 Patterns and Practices](https://aka.ms/m365pnp) - Guidance, tooling, samples and open-source controls for your Microsoft 365 development
