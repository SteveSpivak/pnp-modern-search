import { ServiceScope } from "@microsoft/sp-core-library";
import { PageContext } from "@microsoft/sp-page-context";
import { MSGraphClientFactory, MSGraphClientV3 } from "@microsoft/sp-http";
import { IExtensibilityLibrary, IComponentDefinition, IDataSourceDefinition } from "@pnp/modern-search-extensibility";
import { CopilotSearchDataSource } from "./dataSources/CopilotSearchDataSource";
import { CopilotChatComponent } from "./components/chat/CopilotChatComponent";
import { CopilotMeetingInsightsComponent } from "./components/meetingInsights/CopilotMeetingInsightsComponent";

export class CopilotApiLibrary implements IExtensibilityLibrary {
  /**
   * Statically store the MSGraphClientV3 instance so components and data sources
   * can use the exact same initialized client with the current user's delegated permissions.
   */
  public static msGraphClient: MSGraphClientV3;

  /**
   * Statically store the PageContext for additional SharePoint context if needed.
   */
  public static pageContext: PageContext;

  public onInit(serviceScope: ServiceScope): void {
    // ALWAYS initialize Graph/PnPjs or consume services inside whenFinished()
    // The service scope is not sealed before this callback fires.
    serviceScope.whenFinished(async () => {
      CopilotApiLibrary.pageContext = serviceScope.consume(PageContext.serviceKey);

      const msGraphClientFactory = serviceScope.consume(MSGraphClientFactory.serviceKey);

      try {
        // Initialize the Graph Client for V3 (which supports continuous access evaluation)
        CopilotApiLibrary.msGraphClient = await msGraphClientFactory.getClient("3");
        console.log("[CopilotApiLibrary] Successfully initialized MSGraphClientV3.");
      } catch (err) {
        console.error("[CopilotApiLibrary] Failed to initialize MSGraphClientV3.", err);
      }
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
