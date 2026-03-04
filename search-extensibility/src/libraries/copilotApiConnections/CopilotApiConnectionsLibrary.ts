import { ServiceScope } from "@microsoft/sp-core-library";
import { PageContext } from "@microsoft/sp-page-context";
import { IExtensibilityLibrary, IComponentDefinition, IDataSourceDefinition } from "../../index";

export class CopilotApiConnectionsLibrary implements IExtensibilityLibrary {
  public onInit(serviceScope: ServiceScope): void {
  }

  public getCustomWebComponents(): IComponentDefinition<any>[] {
    return [];
  }

  public getCustomDataSources(): IDataSourceDefinition[] {
    return [];
  }

  public getCustomLayouts(): any[] {
    return [];
  }

  public getCustomSuggestionProviders(): any[] {
    return [];
  }

  public invokeCardAction(): void {
  }
}
