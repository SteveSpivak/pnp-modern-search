# LinkedIn Feed Extensibility - Architecture

This SPFx Library component provides extensibility to PnP Modern Search v4 specifically for displaying a LinkedIn company feed. It is designed to overcome client-side CORS and authentication limitations by leveraging a proxy API.

## Core Components

The solution consists of two primary extensibility points:

1. **Custom Data Source** (`LinkedinFeedDataSource.ts`)
2. **Custom Web Component** (`LinkedinPostCardComponent.tsx` and `LinkedinPostCard.tsx`)

These components are registered and exposed to PnP Modern Search via the central entry point: `LinkedinFeedExtLibrary.ts`.

---

### 1. The Proxy Architecture (Mandatory)

**Why is a proxy required?**
Client-side JavaScript running in a SharePoint Online browser context (`yourtenant.sharepoint.com`) cannot make direct HTTP requests to LinkedIn's `voyager` internal API due to:
* **CORS (Cross-Origin Resource Sharing):** LinkedIn strictly denies cross-origin requests.
* **Authentication:** The Voyager API requires active user session cookies (`li_at`) and CSRF tokens. A SharePoint user's browser does not share these cookies with the SharePoint domain.

**The Solution:**
You must host a secure backend server (e.g., Azure Function) that:
1. Securely stores a LinkedIn `li_at` authentication cookie.
2. Accepts a request from the SharePoint Custom Data Source.
3. Makes the `fetch` request to LinkedIn (`voyager/api/graphql...`).
4. Parses the complex Voyager JSON response.
5. Returns a clean, simplified JSON array of posts back to SharePoint.

---

### 2. Custom Data Source (`LinkedinFeedDataSource`)

The Custom Data Source is responsible for fetching the clean JSON from your Proxy API.

* **Configuration:** It exposes an `apiUrl` setting in the PnP Modern Search Property Pane.
* **Caching Strategy:** To prevent hammering the proxy API (and by extension, LinkedIn) on every page load, it implements a **10-minute client-side cache** using `sessionStorage`.
  * *Mechanism:* Before fetching, it checks `sessionStorage` for the specific API URL key. If valid data exists and is less than 10 minutes old, it returns the cached data immediately.
* **Data Mapping:** PnP Modern Search requires every item in the `IDataSourceData` array to have a unique `Key` property. The data source automatically maps the proxy response to ensure this requirement is met.

---

### 3. Custom Web Component (`LinkedinPostCardComponent`)

This component handles the UI rendering of individual feed items. It follows the standard PnP Modern Search "Two-Layer Pattern" for Web Components:

* **Layer 1: The Web Component Wrapper (`LinkedinPostCardComponent.tsx`)**
  * Extends `BaseWebComponent` (an `HTMLElement` subclass).
  * Its primary job is lifecycle management. In `connectedCallback()`, it resolves HTML `data-*` attributes (passed from the Handlebars template) into React props.
  * Critically, it calls `ReactDOM.unmountComponentAtNode(this)` inside `disconnectedCallback()` to prevent memory leaks when users navigate away or reconfigure the web part.

* **Layer 2: The React Component (`LinkedinPostCard.tsx`)**
  * A standard stateless React Functional Component.
  * Receives properties (`authorName`, `postText`, `imageUrl`, etc.) and renders them using scoped SCSS Modules (`LinkedinPostCard.module.scss`) to prevent CSS bleed into other SharePoint elements.

---

### Summary Workflow

1. User loads the SharePoint page.
2. PnP Modern Search initializes the `LinkedinFeedDataSource`.
3. The Data Source checks `sessionStorage`. If empty, it calls your Proxy API.
4. The Proxy API fetches from LinkedIn and returns JSON posts.
5. The Data Source maps the JSON and caches it in `sessionStorage`.
6. PnP Modern Search passes the items to the Custom Layout (Handlebars).
7. The Handlebars template loops over the items and generates `<linkedin-post-card>` HTML tags with `data-*` attributes.
8. The browser parses these tags, instantiating the `LinkedinPostCardComponent` wrapper.
9. The wrapper mounts the `LinkedinPostCard` React component, rendering the final UI.
