import { BaseLayout } from "@pnp/modern-search-extensibility";
import { IPropertyPaneField } from '@microsoft/sp-property-pane';

export interface IUnifiedCardsLayoutProperties {
    /**
     * Determines whether we enable file preview
     */
    enablePreview: boolean;

    /**
     * Number of prefered cards
     */
    preferedCardNumberPerRow: number;
}

export class UnifiedCardsLayout extends BaseLayout<IUnifiedCardsLayoutProperties> {

    public async onInit(): Promise<void> {
        // Setup default values
        this.properties.enablePreview = this.properties.enablePreview !== null && this.properties.enablePreview !== undefined ? this.properties.enablePreview : true;
        this.properties.preferedCardNumberPerRow = this.properties.preferedCardNumberPerRow !== null && this.properties.preferedCardNumberPerRow !== undefined ? this.properties.preferedCardNumberPerRow : 3;
    }

    public getPropertyPaneFieldsConfiguration(availableFields: string[]): IPropertyPaneField<any>[] {
        return [];
    }
}
