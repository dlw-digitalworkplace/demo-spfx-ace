import { ISPFxAdaptiveCard, BaseAdaptiveCardView, ISubmitActionArguments } from '@microsoft/sp-adaptive-card-extension-base';
import * as strings from 'QuickViewAdvancedAdaptiveCardExtensionStrings';
import { IQuickViewAdvancedAdaptiveCardExtensionProps, IQuickViewAdvancedAdaptiveCardExtensionState } from '../QuickViewAdvancedAdaptiveCardExtension';
import { IListItem } from '../QuickViewAdvancedAdaptiveCardExtension';
import { DETAILED_QUICK_VIEW_REGISTRY_ID } from '../QuickViewAdvancedAdaptiveCardExtension';


export interface IQuickViewData {
  items: IListItem[];
}

export class QuickView extends BaseAdaptiveCardView<
  IQuickViewAdvancedAdaptiveCardExtensionProps,
  IQuickViewAdvancedAdaptiveCardExtensionState,
  IQuickViewData
> {
  public get data(): IQuickViewData {
    return {
      items: this.state.items
    };
  }

  public get template(): ISPFxAdaptiveCard {
    return require('./template/QuickViewTemplate.json');
  }

  public onAction(action: ISubmitActionArguments): void {
    if (action.type === 'Submit') {
      const { id, newIndex } = action.data;
      if (id === 'selectAction') {
        this.quickViewNavigator.push(DETAILED_QUICK_VIEW_REGISTRY_ID, true);
        this.setState({ currentIndex: newIndex });
      }
    }
  }
}