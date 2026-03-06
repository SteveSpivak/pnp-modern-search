import { ServiceScope } from "@microsoft/sp-core-library";
import {
  IExtensibilityLibrary,
  IComponentDefinition,
  ILayoutDefinition,
  IDataSourceDefinition,
  IQueryModifierDefinition,
  ISuggestionProviderDefinition,
} from "@pnp/modern-search-extensibility";
import { LinkedinFeedDataSource } from "./dataSources/LinkedinFeedDataSource";
import { LinkedinPostCardComponent } from "./components/LinkedinPostCardComponent";

export class LinkedinFeedExtLibrary implements IExtensibilityLibrary {
  public onInit(serviceScope: ServiceScope): void {
    // Initialization code runs here
  }

  public getCustomWebComponents(): IComponentDefinition<any>[] {
    return [
      {
        componentName: "linkedin-post-card",
        componentClass: LinkedinPostCardComponent
      }
    ];
  }

  public getCustomDataSources(): IDataSourceDefinition[] {
    return [
      {
        name: "LinkedIn Feed",
        iconName: "Share",
        key: "LinkedinFeedDataSource",
        serviceKey: LinkedinFeedDataSource.ServiceKey
      }
    ];
  }

  public getCustomLayouts(): ILayoutDefinition[] {
    return [];
  }

  public getCustomQueryModifiers(): IQueryModifierDefinition[] {
    return [];
  }

  public getCustomSuggestionProviders(): ISuggestionProviderDefinition[] {
    return [];
  }

  public registerHandlebarsCustomizations(hbs: any): void {
    // Register Handlebars helpers here if needed
  }

  public getCustomAdaptiveCardsActions(): any[] {
    return [];
  }
}
