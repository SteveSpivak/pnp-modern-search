import { ServiceScope } from "@microsoft/sp-core-library";
import { PageContext } from "@microsoft/sp-page-context";
import { MSGraphClientFactory, MSGraphClientV3 } from "@microsoft/sp-http";
import { IExtensibilityLibrary, IComponentDefinition, IDataSourceDefinition } from "@pnp/modern-search-extensibility";
import { CopilotSearchDataSource } from "./dataSources/CopilotSearchDataSource";
import { CopilotChatComponent } from "./components/chat/CopilotChatComponent";
import { CopilotMeetingInsightsComponent } from "./components/meetingInsights/CopilotMeetingInsightsComponent";
import { CopilotUsageReportsComponent } from "./components/usageReports/CopilotUsageReportsComponent";

export class CopilotApiLibrary implements IExtensibilityLibrary {
  /**
   * By keeping a reference to the global ServiceScope provided by PnP Modern Search,
   * we can pass it down to our Custom Web Components. They can then independently
   * consume the MSGraphClientFactory inside their React components, avoiding race conditions.
   */
  public static serviceScope: ServiceScope;

  public onInit(serviceScope: ServiceScope): void {
    // ALWAYS consume services inside whenFinished() as the ServiceScope is not sealed before this.
    serviceScope.whenFinished(() => {
      // We store the serviceScope statically so that Web Components
      // (which are instantiated by the browser, not by us) can access it.
      CopilotApiLibrary.serviceScope = serviceScope;
    });
  }

  public getCustomWebComponents(): IComponentDefinition<any>[] {
    return [
      {
        componentName: "contoso-copilot-chat",
        componentClass: CopilotChatComponent
      },
      {
        componentName: "contoso-copilot-meeting-insights",
        componentClass: CopilotMeetingInsightsComponent
      },
      {
        componentName: "contoso-copilot-usage-reports",
        componentClass: CopilotUsageReportsComponent
      }
    ];
  }

  public getCustomDataSources(): IDataSourceDefinition[] {
    return [
      {
        name: "Microsoft 365 Copilot Search",
        iconName: "Robot",
        key: "CopilotSearchDataSource",
        serviceKey: CopilotSearchDataSource.ServiceKey
      }
    ];
  }
}
