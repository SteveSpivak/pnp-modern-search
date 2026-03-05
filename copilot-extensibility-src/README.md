# PnP Modern Search Copilot Extensibility Library

This SharePoint Framework (SPFx) Library Component acts as a bridge between PnP Modern Search v4 and Microsoft 365 Copilot Graph APIs.

## Architecture

This extensibility library provides:
1. **CopilotSearchDataSource**: Connects PnP Modern Search directly to `/beta/copilot/search`, allowing users to natively search OneDrive/SharePoint content using Copilot's semantic search.
2. **CopilotChatComponent**: A `<contoso-copilot-chat>` Web Component. Pass the `{searchTerms}` token to it, and it initializes a Copilot conversation based on the current search query.
3. **CopilotMeetingInsightsComponent**: A `<contoso-copilot-meeting-insights>` Web Component. Pass a `userId` and `meetingId` to display AI-generated summaries and action items.
4. **CopilotUsageReportsComponent**: A `<contoso-copilot-usage-reports>` Web Component. Pass a `period` token (`D7`, `D30`, `D90`, `D180`) to retrieve an administrative report of Copilot usage statistics. Requires `Reports.Read.All` permissions.

All connections use the `MSGraphClientV3` running under the delegated permissions of the current logged-in user, natively injected via the PnP Extensibility Library ServiceScope.

## Pre-Commit Verification Checklist

Before committing and submitting changes to this library, ensure the following steps are passed:

1. **Linting & Compilation**:
   - Run `gulp build` to ensure all TypeScript and React code compiles cleanly.
   - Run `gulp lint` to verify that there are no standard SPFx linting errors.

2. **Testing Locally with PnP Modern Search**:
   - Serve the extensibility library locally: `gulp serve --nobrowser`.
   - In your SharePoint tenant, open the Hosted Workbench (`https://<tenant>.sharepoint.com/_layouts/15/workbench.aspx`).
   - Use the `?debugManifestsFile=https://localhost:4321/temp/manifests.js` query string parameter.
   - Add the PnP "Search Results" or "Search Box" web parts to the workbench.
   - Open the web part Property Pane -> Extensibility configuration.
   - Add the Library ID (`c60a4f5b-11d8-4f5c-a521-cf7b98d28c31`) and enable it.

3. **Verify Graph Context and Dynamic Data**:
   - Add `<contoso-copilot-chat data-search-terms="{searchTerms}"></contoso-copilot-chat>` to the Search Results custom Handlebars template.
   - Ensure you do not see "MS Graph Client is not initialized" errors on load.
   - Perform a search and verify that the `{searchTerms}` dynamic data flows into the chat component.

## Deployment Strategy

1. **API Permissions Approval**:
   - After packaging (`gulp package-solution --ship`), an administrator must deploy the `.sppkg` file to the tenant App Catalog.
   - The administrator must go to the SharePoint Admin Center -> Advanced -> API Access and approve the requested Graph scopes (`Chat.Read`, `ChannelMessage.Read.All`, `ExternalItem.Read.All`, `Reports.Read.All`, etc.).

2. **Site Deployment**:
   - Because `skipFeatureDeployment` is set to `true`, the library does not need to be installed on individual site collections. It is available tenant-wide to be consumed by any PnP Modern Search web part.

## Note on Extensibility Initialization

Do not resolve `MSGraphClientFactory` globally inside the `IExtensibilityLibrary`'s `onInit` without wrapping it in a Promise, or without resolving it inside Web Components separately. The `onInit` method returns `void` or `Promise<void>`, so returning a Promise for Data Sources properly pauses initialization to await setup. For Web Components, mapping the `ServiceScope` directly into React ensures proper lifecycle handling and prevents "MS Graph Client is not initialized" race conditions on load.
