import { ServiceScope } from "@microsoft/sp-core-library";
import { PageContext } from "@microsoft/sp-page-context";
import { IExtensibilityLibrary, IComponentDefinition, IDataSourceDefinition } from "@pnp/modern-search-extensibility";

export class CopilotApiConnectionsLibrary implements IExtensibilityLibrary {
  public onInit(serviceScope: ServiceScope): void {
  }

  public getCustomWebComponents(): IComponentDefinition<any>[] {
    return [];
  }

  public getCustomDataSources(): IDataSourceDefinition[] {
    return [];
  }
}
