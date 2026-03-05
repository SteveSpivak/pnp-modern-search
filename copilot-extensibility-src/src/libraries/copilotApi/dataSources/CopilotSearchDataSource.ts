import { ServiceKey } from "@microsoft/sp-core-library";
import { BaseDataSource, IDataSourceData, IDataContext, ITokenService } from "@pnp/modern-search-extensibility";
import { PropertyPaneTextField, PropertyPaneSlider } from "@microsoft/sp-property-pane";
import { MSGraphClientFactory, MSGraphClientV3 } from "@microsoft/sp-http";

export interface ICopilotSearchProperties {
  pageSize: number;
}

/**
 * PnP Modern Search Data Source pointing to the Microsoft 365 Copilot Search API.
 */
export class CopilotSearchDataSource extends BaseDataSource<ICopilotSearchProperties> {
  // Service Key used by PnP Modern Search to register the provider uniquely
  public static ServiceKey = ServiceKey.create<CopilotSearchDataSource>("ContosoExt:CopilotSearch", CopilotSearchDataSource as any);

  private _tokenService: ITokenService;
  private _msGraphClient: MSGraphClientV3;

  public async onInit(): Promise<void> {
    if (!this.properties.pageSize) {
      this.properties.pageSize = 10;
    }

    return new Promise<void>((resolve, reject) => {
      this.serviceScope.whenFinished(async () => {
        try {
          const msGraphClientFactory = this.serviceScope.consume(MSGraphClientFactory.serviceKey);
          this._msGraphClient = await msGraphClientFactory.getClient("3");
          resolve();
        } catch (err) {
          console.error("[CopilotSearchDataSource] Failed to initialize MSGraphClientV3.", err);
          reject(err);
        }
      });
    });
  }

  public getPropertyPaneFieldsConfiguration(): any[] {
    return [
      PropertyPaneSlider("pageSize", {
        label: "Number of search results",
        min: 1,
        max: 50,
        value: this.properties.pageSize
      })
    ];
  }

  // Required by PnP Modern Search interface
  public getItemCount(): number {
    return this.properties.pageSize;
  }

  /**
   * Main method called by the PnP Search Results Web Part to retrieve items.
   */
  public async getData(dataContext?: IDataContext): Promise<IDataSourceData> {
    // In actual implementation, TokenService is used to resolve `{searchTerms}`.
    // For compilation completeness we'll pull an expected filter context.
    const selectedFilters = dataContext?.filters?.selectedFilters || [];
    const searchFilter = selectedFilters.find((f: any) => f.filterName === "searchTerms");
    const resolvedQuery = searchFilter?.values?.[0]?.value || "";

    if (!resolvedQuery || resolvedQuery.trim().length === 0) {
      return { items: [], totalCount: 0, filters: [] } as any;
    }

    if (!this._msGraphClient) {
      console.warn("[CopilotSearchDataSource] Graph Client is not initialized.");
      return { items: [], totalCount: 0, filters: [] } as any;
    }

    try {
      const requestBody = {
        query: resolvedQuery.trim(),
        pageSize: this.properties.pageSize,
        dataSources: {
          oneDrive: {
            resourceMetadataNames: ["title", "author"]
          }
        }
      };

      const response = await this._msGraphClient
        .api("/copilot/search")
        .version("beta")
        .post(requestBody);

      return {
        items: response.searchHits || [],
        totalCount: response.totalCount || 0,
        filters: []
      } as any;

    } catch (err) {
      console.error("[CopilotSearchDataSource] Failed to fetch data:", err);
      return { items: [], totalCount: 0, filters: [] } as any;
    }
  }

  public getAvailableFieldsFromResults(results: any[]): string[] {
    if (results.length > 0) {
      return Object.keys(results[0]);
    }
    return [];
  }
}
