import * as React from "react";
import * as ReactDOM from "react-dom";
import { BaseWebComponent } from "@pnp/modern-search-extensibility";
import { CopilotMeetingInsights } from "./CopilotMeetingInsights";

/**
 * Custom web component mapping <contoso-copilot-meeting-insights> to React.
 */
export class CopilotMeetingInsightsComponent extends BaseWebComponent {
  public async connectedCallback(): Promise<void> {
    const props = this.resolveAttributes();

    const meetingId = props.meetingId as string;
    const userId = props.userId as string;

    ReactDOM.render(
      React.createElement(CopilotMeetingInsights, {
        meetingId: meetingId,
        userId: userId
      }),
      this
    );
  }

  protected disconnectedCallback(): void {
    ReactDOM.unmountComponentAtNode(this);
  }
}
