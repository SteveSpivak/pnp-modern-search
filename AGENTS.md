# PnP Modern Search - AI Agent Skill Guide

This `AGENTS.md` file serves as a reference guide for AI agents interacting with the [PnP Modern Search](https://github.com/microsoft-search/pnp-modern-search) repository. It contains essential context, architectural understanding, and best practices for developing, reviewing, and extending this codebase.

## Overview

PnP Modern Search is an open-source solution that allows you to build engaging search-based solutions in the SharePoint modern experience. The codebase is primarily built using **SharePoint Framework (SPFx)**, **React**, and **TypeScript**.

The repository consists of two main packages:
1. `search-parts`: Contains the core web parts for building a search experience.
2. `search-extensibility`: A base library for creating PnP Modern Search extensions (custom layouts, web components, data sources).

## Architectural Structure

### 1. `search-parts`
This package contains the end-user facing web parts.
- **Location:** `/search-parts/`
- **Core Web Parts:** Located in `/search-parts/src/webparts/`.
  - `searchBox`: A configurable search input box that passes queries to other web parts via Dynamic Data.
  - `searchResults`: Displays search results based on queries. Supports Handlebars templating and custom React components.
  - `searchFilters`: Displays refiners based on the search results.
  - `searchVerticals`: Allows switching context/scopes for search results.
- **Data Sources:** Located in `/search-parts/src/dataSources/`. These components abstract the APIs used to fetch search results (e.g., SharePoint Search, Microsoft Search API).
- **Layouts:** Located in `/search-parts/src/layouts/`. Contains different presentation options (Cards, Lists, Details, etc.).

### 2. `search-extensibility`
This package defines the interfaces and base classes necessary for extending PnP Modern Search functionalities.
- **Location:** `/search-extensibility/`
- Use this library if you are creating a new custom layout, custom web component, custom suggestions provider, or a custom data source.
- Extensions created using this package can be registered in SPFx using the component ID.
- **Microsoft 365 Copilot APIs**: This package also includes service wrappers for interacting with Microsoft 365 Copilot APIs through Microsoft Graph. These services are located in `/search-extensibility/src/services/copilot/` and include functionality for:
  - **Copilot Search API** (`CopilotSearchService.ts`): Performs hybrid semantic and lexical search across OneDrive.
  - **Copilot Chat API** (`CopilotChatService.ts`): Queries past interactions with Copilot.
  - **Copilot Meeting Insights API** (`CopilotMeetingInsightsService.ts`): Retrieves AI-generated insights from meetings.
  - **Copilot Usage Reports API** (`CopilotUsageReportsService.ts`): Accesses usage reporting for M365 Copilot.
  *Note: These APIs might require specific Microsoft Graph scopes (e.g., `Files.Read.All`, `Chat.Read`) and a properly initialized `MSGraphClientV3`.*

## Development Stack & Tooling

- **SharePoint Framework (SPFx):** The project uses SPFx v1.22+. Pay attention to `@microsoft/sp-*` package versions when updating or adding dependencies.
- **Node.js:** Requires Node.js version `>=22.14.0 < 23.0.0` (as defined in `package.json` engines).
- **Package Manager:** `pnpm` is the preferred package manager.
- **UI Framework:** **React** (`^17.0.1`) and **Fluent UI** (`@fluentui/react` `^8.106.4`).
- **Templating:** **Handlebars** (`^4.7.7`) is used heavily in `searchResults` for custom presentation layers.
- **Build System:**
  - `search-parts` uses Gulp (`gulp bundle`, `gulp clean`).
  - `search-extensibility` uses Heft (`heft build`).

## Best Practices & Guidelines

When generating, editing, or reviewing code for this repository, you must adhere to the following best practices:

### 1. Web Part State & Properties
- PnP Modern Search web parts heavily rely on complex properties panes and dynamic data.
- Ensure that web part properties are properly typed and defaults are securely managed.
- Utilize `@pnp/spfx-property-controls` and custom controls located in `search-parts/src/controls/` for rich property pane configurations.
- Use **SPFx Dynamic Data** to allow web parts to communicate with each other (e.g., Search Box sending queries to Search Results, Search Results sending available refiners to Search Filters).

### 2. UI / UX Design
- All components must align with the Microsoft Fluent Design language.
- Use `@fluentui/react` components whenever possible instead of native HTML elements or raw CSS.
- Ensure all custom UI is fully accessible (WCAG compliant). Pay attention to `aria-` labels, role attributes, and keyboard navigability.

### 3. Asynchronous Operations & Data Fetching
- For SharePoint specific calls, utilize `@pnp/sp` (PnPjs).
- Ensure all external calls handle failures gracefully. Use proper error boundaries in React and provide localized error messages to users.

### 4. Custom Extensibility (Handlebars & Web Components)
- When adding Handlebars helpers (in `search-parts/src/helpers/`), ensure they are pure, fast, and safely handle `null`/`undefined` inputs.
- For more complex UI, favor creating standard Web Components over inline Handlebars logic. Handlebars should only be used for structural templating and simple presentation.

### 5. Localization
- Always use the localized strings found in the `loc` folders (`/search-parts/src/loc/`). Do not hardcode user-facing strings in TypeScript or Handlebars files.

### 6. File Naming and Formatting
- Follow strict camelCase for variables and PascalCase for React components and class names.
- Interfaces should generally be prefixed with `I` (e.g., `ISearchResultsProps`).
- Run the SPFx specific ESLint configurations to validate code (`eslint`).

## Actionable Instructions for AI Agents

When tasked with modifying this codebase:
1. **Identify the Scope:** Determine if the change belongs in a core web part (`search-parts/src/webparts`), an extensibility component (`search-extensibility`), or a shared helper (`search-parts/src/helpers`).
2. **Review Existing Patterns:** Look for similar implementations (e.g., if adding a property pane field, check how existing fields are built).
3. **Respect Node Versions:** Always run commands using the specified Node version to avoid build/restore errors. Use `nvm` or the applicable Node manager if testing locally via `run_in_bash_session`.
4. **Compile Frequently:** Run `gulp bundle` in `search-parts` or `heft build` in `search-extensibility` to catch TypeScript/Linting errors early.
5. **Verify Dynamic Connections:** If altering web part properties, ensure the `getPropertyDefinitions()` and Dynamic Data logic remains intact so that interconnected web parts (Search Box + Results + Filters) do not break.