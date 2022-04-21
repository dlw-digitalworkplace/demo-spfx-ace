import { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseAdaptiveCardExtension } from '@microsoft/sp-adaptive-card-extension-base';
import { CardView } from './cardView/CardView';
import { MediumCardView } from './cardView/MediumCardView';
import { QuickView } from './quickView/QuickView';
import { CardViewAdvancedPropertyPane } from './CardViewAdvancedPropertyPane';
import { SPHttpClient } from '@microsoft/sp-http';

export interface ICardViewAdvancedAdaptiveCardExtensionProps {
  title: string;
  listId: string;
}

export interface ICardViewAdvancedAdaptiveCardExtensionState {
  currentIndex: number;
  items: IListItem[];
}

export interface IListItem {
  title: string;
  description: string;
}

const CARD_VIEW_REGISTRY_ID: string = 'CardViewAdvanced_CARD_VIEW';
const MEDIUM_VIEW_REGISTRY_ID: string = 'CardViewAdvanced_MEDIUM_VIEW';
export const QUICK_VIEW_REGISTRY_ID: string = 'CardViewAdvanced_QUICK_VIEW';

export default class CardViewAdvancedAdaptiveCardExtension extends BaseAdaptiveCardExtension<
  ICardViewAdvancedAdaptiveCardExtensionProps,
  ICardViewAdvancedAdaptiveCardExtensionState
> {
  private _deferredPropertyPane: CardViewAdvancedPropertyPane | undefined;

  public onInit(): Promise<void> {
    this.state = {
      currentIndex: 0,
      items: []
    };

    this.cardNavigator.register(CARD_VIEW_REGISTRY_ID, () => new CardView());
    this.cardNavigator.register(MEDIUM_VIEW_REGISTRY_ID, () => new MediumCardView());
    this.quickViewNavigator.register(QUICK_VIEW_REGISTRY_ID, () => new QuickView());

    return this._fetchData();
  }

  protected loadPropertyPaneResources(): Promise<void> {
    return import(
      /* webpackChunkName: 'CardViewAdvanced-property-pane'*/
      './CardViewAdvancedPropertyPane'
    )
      .then(
        (component) => {
          this._deferredPropertyPane = new component.CardViewAdvancedPropertyPane();
        }
      );
  }

  protected renderCard(): string | undefined {
    return this.cardSize === "Medium" ? MEDIUM_VIEW_REGISTRY_ID : CARD_VIEW_REGISTRY_ID;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return this._deferredPropertyPane!.getPropertyPaneConfiguration();
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    if (propertyPath === 'listId' && newValue !== oldValue) {
      if (newValue) {
        this._fetchData();
      } else {
        this.setState({ items: [] });
      }
    }
  }

  // protected getCacheSettings(): Partial<ICacheSettings> {
  //   return {
  //     isEnabled: true, // can be set to false to disable caching
  //     expiryTimeInSeconds: 86400, // controls how long until the cached card and state are stale
  //     cachedCardView: () => new CardView() // function that returns the custom Card view that will be used to generate the cached card
  //   };
  // }

  private async _fetchData(): Promise<void> {
    if (this.properties.listId) {
      const response = await this.context.spHttpClient.get(
        `${this.context.pageContext.web.absoluteUrl}` +
        `/_api/web/lists/GetById(id='${this.properties.listId}')/items`,
        SPHttpClient.configurations.v1
      );
      const jsonResponse = await response.json();
      const items = jsonResponse.value.map((item) => { return { title: item.Title, description: item.Description }; });
      this.setState({ items });
    }
  }
}
