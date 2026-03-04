# Architecture

## Overview
The PnP Modern Search (v4) solution includes multiple SPFX Web Parts (Search Box, Search Results, Search Filters, Search Verticals). It incorporates a robust search extensibility model that allows developers to write custom queries, layout templates, data sources, and query modifiers.

## M365 Copilot Integration
The `search-extensibility` and `copilot-extensibility-src` packages introduce deep integration with M365 Copilot services.
Using the `@pnp/modern-search-extensibility` API layer, it connects with the `/beta/copilot/conversations` Graph API. A fallback mechanism guarantees backward compatibility by falling back to files and `webSearchEnabled` flags.

**Key Components:**
- `CopilotChatService`
- `CopilotSearchDataSource`
- `CopilotApiConnectionsLibrary`

## Unified Design Language (Orange/Slate)
The project utilizes a modern Unified Design Language based on Orange and Slate colors.
- It uses gradient toggle buttons for views (List/Grid view) through pure CSS sibling selectors (`#view-toggle:checked ~ .app-catalogue`), avoiding costly React state re-renders.
- It employs high-fidelity hover states and shadow micro-animations (`transform: scale(1.05)`, `translateY(-2px)`) for a premium dynamic feel.

## Performance Optimizations
- **DataSourceHelper**: `parseAndCleanOptions` now relies on highly performant native `for` loops instead of complex higher-order functions like `flatMap`. It gracefully parses comma-separated keys without removing non-comma options, eliminating UI bugs in dropdown filters.
