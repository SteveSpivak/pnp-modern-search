import * as React from "react";
import * as ReactDOM from "react-dom";
import { BaseWebComponent } from "@pnp/modern-search-extensibility";
import { CopilotChat } from "./CopilotChat";

/**
 * BaseWebComponent wrapper that bridges PnP Modern Search's custom HTML tags
 * into a fully-fledged React component instance.
 */
export class CopilotChatComponent extends BaseWebComponent {
  public async connectedCallback(): Promise<void> {
    // 1. Resolve Attributes passed from Handlebars templates, e.g., <contoso-copilot-chat data-search-terms="{searchTerms}"></contoso-copilot-chat>
    const props = this.resolveAttributes();

    // 2. We extract specific attributes needed for the Chat Context.
    // 'data-search-terms' -> 'searchTerms' through PnP Modern Search conversion
    const searchTerms = props.searchTerms as string || "";
    const conversationId = props.conversationId as string || undefined;

    // 3. Render the Component in this HTMLElement.
    // Always load the inner React logic separately to keep the custom element footprint tiny.
    ReactDOM.render(
      React.createElement(CopilotChat, {
        initialMessage: searchTerms,
        existingConversationId: conversationId
      }),
      this
    );
  }

  protected disconnectedCallback(): void {
    // REQUIRED: PnP Modern Search recreates components when attributes change.
    // Calling unmount prevents massive memory leaks inside the SPFx host environment.
    ReactDOM.unmountComponentAtNode(this);
  }
}
