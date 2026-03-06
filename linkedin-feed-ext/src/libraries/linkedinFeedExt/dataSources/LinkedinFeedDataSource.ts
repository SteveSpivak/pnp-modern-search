import { ServiceKey } from "@microsoft/sp-core-library";
import {
  BaseDataSource,
  IDataSourceData,
  IDataFilterInfo
} from "@pnp/modern-search-extensibility";
import { IPropertyPaneField, PropertyPaneTextField } from "@microsoft/sp-property-pane";

export interface ILinkedinFeedDataSourceProperties {
  apiUrl: string;
}

export class LinkedinFeedDataSource extends BaseDataSource<ILinkedinFeedDataSourceProperties> {
  public static ServiceKey = ServiceKey.create<LinkedinFeedDataSource>(
    "LinkedinExt:LinkedinFeedDataSource",
    LinkedinFeedDataSource
  );

  public async onInit(): Promise<void> {
    if (!this.properties.apiUrl) {
      this.properties.apiUrl = "https://www.linkedin.com/company/cellebrite/posts/?feedView=all";
    }
  }

  public getPropertyPaneFieldsConfiguration(): IPropertyPaneField<any>[] {
    return [
      PropertyPaneTextField("apiUrl", {
        label: "API / Proxy URL",
        description: "The URL to fetch the LinkedIn feed from."
      })
    ];
  }

  public async getData(queryText: string, filters: IDataFilterInfo[], page: number): Promise<IDataSourceData> {
    const defaultData: IDataSourceData = { items: [], totalCount: 0, filters: [] };

    try {
      const cacheKey = `linkedin_feed_${this.properties.apiUrl}`;
      const cached = sessionStorage.getItem(cacheKey);

      if (cached) {
        const parsed = JSON.parse(cached);
        // Simple 10-minute cache expiration
        if (parsed.timestamp && Date.now() - parsed.timestamp < 10 * 60 * 1000) {
          return { items: parsed.items, totalCount: parsed.items.length, filters: [] };
        }
      }

      const response = await fetch(this.properties.apiUrl, {
        headers: { Accept: "application/json" }
      });

      if (!response.ok) {
        throw new Error(`Failed to fetch LinkedIn feed: ${response.status}`);
      }

      const json = await response.json();

      // We assume the proxy returns an array of posts.
      // If the proxy returns a different structure, we map it here.
      const items = Array.isArray(json) ? json : (json.data || json.items || []);

      const mappedItems = items.map((item: any, index: number) => {
        return {
          ...item,
          // PnP Modern Search requires a Key property
          Key: item.id || `linkedin-post-${index}`,
        };
      });

      // Cache the result
      sessionStorage.setItem(cacheKey, JSON.stringify({ items: mappedItems, timestamp: Date.now() }));

      return { items: mappedItems, totalCount: mappedItems.length, filters: [] };

    } catch (error) {
      console.error("LinkedinFeedDataSource getData error:", error);
      return defaultData;
    }
  }

  public getAvailableFieldsFromResults(results: any[]): string[] {
    if (results && results.length > 0) {
      return Object.keys(results[0]);
    }
    return [];
  }
}
