# LinkedIn Feed - PnP Modern Search Extensibility

This project provides a Custom Data Source and a Custom Web Component for [PnP Modern Search v4](https://microsoft-search.github.io/pnp-modern-search/), enabling you to display a live LinkedIn feed within SharePoint Online.

## Important Note on LinkedIn APIs

LinkedIn actively blocks direct scraping from the browser via CORS policies and authentication requirements.
**To use this extensibility component, you must host a backend proxy API** (e.g., Azure Function, AWS Lambda, Node.js server) that performs the actual data fetching from LinkedIn and returns JSON to this component.

This component is configured by default to point to: `https://www.linkedin.com/company/cellebrite/posts/?feedView=all` (as a placeholder), but you must point it to your working proxy API via the Property Pane.

---

## Installation & Deployment Guide

### Prerequisites
* [Node.js v18.17.1+](https://nodejs.org/) (Required for SPFx 1.18.2)
* Gulp CLI (`npm install -g gulp-cli`)

### 1. Build the Package
Open your terminal, navigate to the `linkedin-feed-ext` folder, and run:
```bash
npm install
gulp bundle --ship
gulp package-solution --ship
```
This will generate the SharePoint package file at: `sharepoint/solution/linkedin-feed-ext.sppkg`.

### 2. Deploy to SharePoint App Catalog
1. Go to your SharePoint Tenant Admin Center > **More features** > **Apps** (App Catalog).
2. Click **Upload** and select the `linkedin-feed-ext.sppkg` file.
3. Check the box to **Make this solution available to all sites in the organization** (since `skipFeatureDeployment` is set to true).
4. Click **Deploy**.

### 3. Configure PnP Modern Search
1. Navigate to the SharePoint site where you have PnP Modern Search installed.
2. Edit the page and add the **PnP - Search Results** web part.
3. Open the Property Pane for the Search Results web part.
4. Go to Page 4: **Extensibility configuration**.
5. Find the **Extensibility libraries** section and paste the manifest GUID of this library:
   `4fdd2d20-bb92-4137-9bf1-7327ea4b6bb6` (You can confirm this ID in `config/package-solution.json`).
6. Click **Add** and **Apply**.

### 4. Setup the Data Source
1. In the Search Results Property Pane, go to Page 1: **Data source**.
2. From the dropdown, select **LinkedIn Feed**.
3. A new field will appear: **API / Proxy URL**. Enter the URL of your backend proxy that returns the LinkedIn JSON feed.

### 5. Setup the Custom Layout
1. In the Search Results Property Pane, go to Page 2: **Layouts**.
2. Select **Custom** and click **Edit results template**.
3. Modify the Handlebars template to loop over the results and use the custom web component tag: `<linkedin-post-card>`.

Example Template:
```html
<style>
  .linkedin-feed-container { display: grid; gap: 16px; }
</style>
<div class="linkedin-feed-container">
  {{#each items}}
    <linkedin-post-card
      data-author-name="{{authorName}}"
      data-author-title="{{authorTitle}}"
      data-post-text="{{postText}}"
      data-image-url="{{imageUrl}}"
      data-post-url="{{postUrl}}"
      data-date="{{date}}">
    </linkedin-post-card>
  {{/each}}
</div>
```
*(Make sure the attributes map exactly to the JSON keys returned by your proxy API).*
