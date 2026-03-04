# Learned Lessons

Here are the key lessons and takeaways from recent development cycles, focusing on the UI overhaul and M365 Copilot extensibility integration:

## 1. Handling TypeScript Cyclical Dependencies
During the integration of M365 Copilot API Connections (`CopilotApiConnectionsLibrary.ts`), a build error occurred: `TS2307`. This was caused by importing directly from the package name (`@pnp/modern-search-extensibility`) inside the very package being compiled.
**Lesson Learned:** Use proper relative paths (`../../index`) when importing internal types/modules inside the core library to prevent Heft/TypeScript compilation failures.

## 2. Performance Optimizations in Data Parsing
In the `search-parts` UI filter controls, a bug frequently surfaced regarding parsed options. We relied heavily on higher-order functional chains like `flatMap`.
**Lesson Learned:** Refactoring `parseAndCleanOptions` in `DataSourceHelper.ts` to use performant native `for` loops successfully mitigated the issue, demonstrating that raw iteration can be safer and more predictable than deeply chained functional paradigms when evaluating mixed inputs.

## 3. Pure CSS State Management vs. React State
The new Orange/Slate Unified Design Language required complex toggles (List/Grid view) and animations. Initially, doing this through React state triggered heavy re-renders across the component tree.
**Lesson Learned:** Utilizing pure CSS sibling selectors (`#view-toggle:checked ~ .app-catalogue`) effectively removed the need for React state. Combined with CSS transforms (`transform: scale(1.05)`), we achieved high-fidelity interactions with almost zero performance penalty.
