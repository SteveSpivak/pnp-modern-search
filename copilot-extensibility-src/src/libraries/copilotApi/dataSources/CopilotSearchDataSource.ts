import { ServiceKey } from "@microsoft/sp-core-library";
import { BaseDataSource, IDataSourceData, IDataFilterInfo, TokenService, ITokenService } from "@pnp/modern-search-extensibility";
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
  public static ServiceKey = ServiceKey.create<CopilotSearchDataSource>("ContosoExt:CopilotSearch", CopilotSearchDataSource);

  private _tokenService: ITokenService;
  private _msGraphClient: MSGraphClientV3;

  public async onInit(): Promise<void> {
    // Provide some robust defaults
    if (!this.properties.pageSize) {
      this.properties.pageSize = 10;
    }

    // We ensure the ServiceScope is resolved to access built-in extensibility services.
    // Wrap the whenFinished in a Promise to explicitly block initialization
    // until the MSGraphClientV3 is fully hydrated and ready.
    return new Promise<void>((resolve, reject) => {
      this.serviceScope.whenFinished(async () => {
        try {
          this._tokenService = this.serviceScope.consume(TokenService.ServiceKey);
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

  /**
   * Main method called by the PnP Search Results Web Part to retrieve items.
   */
  public async getData(queryText: string, filters: IDataFilterInfo[], page: number): Promise<IDataSourceData> {
    // 1. We resolve tokens. Ex: User inputs '{searchTerms}' from the search box
    const resolvedQuery = await this._tokenService.resolveTokens(queryText);

    // Provide a fallback if no query is submitted to prevent an error
    if (!resolvedQuery || resolvedQuery.trim().length === 0) {
      return { items: [], totalCount: 0, filters: [] };
    }

    if (!this._msGraphClient) {
      console.warn("[CopilotSearchDataSource] Graph Client is not initialized.");
      return { items: [], totalCount: 0, filters: [] };
    }

    try {
      // 2. Format the payload for /beta/copilot/search
      const requestBody = {
        query: resolvedQuery.trim(),
        pageSize: this.properties.pageSize,
        dataSources: {
          oneDrive: {
            resourceMetadataNames: ["title", "author"]
          }
        }
      };

      // 3. Perform the delegated user context API Call
      const response = await this._msGraphClient
        .api("/copilot/search")
        .version("beta")
        .post(requestBody);

      // 4. Map the response schema to PnP Modern Search's expected `IDataSourceData` model
      return {
        items: response.searchHits || [],
        totalCount: response.totalCount || 0,
        filters: [] // Not implementing custom Copilot Search refiners in this basic sample
      };

    } catch (err) {
      console.error("[CopilotSearchDataSource] Failed to fetch data:", err);
      // Return empty array instead of throwing an unhandled UI error for the search web part
      return { items: [], totalCount: 0, filters: [] };
    }
  }

  /**
   * Required method that maps the raw schema to available properties inside
   * the property pane (e.g., when choosing fields for custom Handlebars templates).
   */
  public getAvailableFieldsFromResults(results: any[]): string[] {
    if (results.length > 0) {
      return Object.keys(results[0]);
    }
    return [];
  }
}
