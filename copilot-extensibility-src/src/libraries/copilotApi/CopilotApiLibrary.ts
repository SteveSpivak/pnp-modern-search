import { ServiceScope } from "@microsoft/sp-core-library";
import { PageContext } from "@microsoft/sp-page-context";
import { MSGraphClientFactory, MSGraphClientV3 } from "@microsoft/sp-http";
import {
  IExtensibilityLibrary,
  IComponentDefinition,
  IDataSourceDefinition,
  ILayoutDefinition,
  ISuggestionProviderDefinition
} from "@pnp/modern-search-extensibility";
import { CopilotSearchDataSource } from "./dataSources/CopilotSearchDataSource";
import { CopilotChatComponent } from "./components/chat/CopilotChatComponent";
import { CopilotMeetingInsightsComponent } from "./components/meetingInsights/CopilotMeetingInsightsComponent";
import { CopilotUsageReportsComponent } from "./components/usageReports/CopilotUsageReportsComponent";

export class CopilotApiLibrary implements IExtensibilityLibrary {
  public static serviceScope: ServiceScope;

  public onInit(serviceScope: ServiceScope): void {
    serviceScope.whenFinished(() => {
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
        serviceKey: CopilotSearchDataSource.ServiceKey as any // Type override due to dependency mismatch
      }
    ];
  }

  // Required by PnP Modern Search interface
  public getCustomLayouts(): ILayoutDefinition[] {
    return [];
  }

  // Required by PnP Modern Search interface
  public getCustomSuggestionProviders(): ISuggestionProviderDefinition[] {
    return [];
  }

  // Required by PnP Modern Search interface
  public invokeCardAction(action: any): void {
  }
}
