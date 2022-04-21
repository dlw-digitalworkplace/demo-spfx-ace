import { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseAdaptiveCardExtension, RenderType } from '@microsoft/sp-adaptive-card-extension-base';
import { CardView } from './cardView/CardView';
import { QuickView } from './quickView/QuickView';
import { QuickViewAdvancedPropertyPane } from './QuickViewAdvancedPropertyPane';
import { SPHttpClient } from '@microsoft/sp-http';
import { MediumCardView } from './cardView/MediumCardView';
import { DetailedView } from './quickView/DetailedQuickView';


export interface IQuickViewAdvancedAdaptiveCardExtensionProps {
  title: string;
  listId: string;
}

export interface IQuickViewAdvancedAdaptiveCardExtensionState {
  items: IListItem[];
  currentIndex: number;
}

export interface IListItem {
  title: string;
  description: string;
  index: number;
}

const CARD_VIEW_REGISTRY_ID: string = 'QuickViewAdvanced_CARD_VIEW';
const MEDIUM_VIEW_REGISTRY_ID: string = 'QuickViewAdvanced_MEDIUM_VIEW';
export const QUICK_VIEW_REGISTRY_ID: string = 'QuickViewAdvanced_QUICK_VIEW';
export const DETAILED_QUICK_VIEW_REGISTRY_ID: string = 'QuickViewAdvanced_DETAILED_QUICK_VIEW';

export default class QuickViewAdvancedAdaptiveCardExtension extends BaseAdaptiveCardExtension<
  IQuickViewAdvancedAdaptiveCardExtensionProps,
  IQuickViewAdvancedAdaptiveCardExtensionState
> {
  private _cardIndex: number;
  private _deferredPropertyPane: QuickViewAdvancedPropertyPane | undefined;

  public onInit(): Promise<void> {
    this.state = {
      items: [],
      currentIndex: 0
    };

    this.cardNavigator.register(CARD_VIEW_REGISTRY_ID, () => new CardView());
    this.cardNavigator.register(MEDIUM_VIEW_REGISTRY_ID, () => new MediumCardView());
    this.quickViewNavigator.register(QUICK_VIEW_REGISTRY_ID, () => new QuickView());
    this.quickViewNavigator.register(DETAILED_QUICK_VIEW_REGISTRY_ID, () => new DetailedView());
    return this._fetchData();
  }

  protected loadPropertyPaneResources(): Promise<void> {
    return import(
      /* webpackChunkName: 'QuickViewAdvanced-property-pane'*/
      './QuickViewAdvancedPropertyPane'
    )
      .then(
        (component) => {
          this._deferredPropertyPane = new component.QuickViewAdvancedPropertyPane();
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

  protected onRenderTypeChanged(oldRenderType: RenderType): void {
    if (oldRenderType === 'QuickView') {
      // Reset to the Card state when the Quick View was opened.
      this.setState({ currentIndex: this._cardIndex });
    } else {
      // The Quick View is opened, save the current index.
      this._cardIndex = this.state.currentIndex;
    }
  }

  private async _fetchData(): Promise<void> {
    if (this.properties.listId) {
      const response = await this.context.spHttpClient.get(
        `${this.context.pageContext.web.absoluteUrl}` +
        `/_api/web/lists/GetById(id='${this.properties.listId}')/items`,
        SPHttpClient.configurations.v1
      );
      const jsonResponse = await response.json();
      const items = jsonResponse.value.map((item, index) => { return { title: item.Title, description: item.Description, index }; });
      this.setState({ items });
    }
  }
}
