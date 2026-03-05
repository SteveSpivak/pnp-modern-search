import * as React from "react";
import * as ReactDOM from "react-dom";
import { BaseWebComponent } from "@pnp/modern-search-extensibility";
import { CopilotUsageReports } from "./CopilotUsageReports";
import { CopilotApiLibrary } from "../../CopilotApiLibrary";

/**
 * Custom web component mapping <contoso-copilot-usage-reports> to React.
 */
export class CopilotUsageReportsComponent extends BaseWebComponent {
  public async connectedCallback(): Promise<void> {
    const props = this.resolveAttributes();

    const period = props.period as "D7" | "D30" | "D90" | "D180" || "D7";

    ReactDOM.render(
      React.createElement(CopilotUsageReports, {
        period: period,
        serviceScope: CopilotApiLibrary.serviceScope
      }),
      this
    );
  }

  protected disconnectedCallback(): void {
    ReactDOM.unmountComponentAtNode(this);
  }
}
