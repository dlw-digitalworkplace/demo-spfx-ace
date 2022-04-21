import { BaseAdaptiveCardView, IActionArguments, ISPFxAdaptiveCard, ISubmitActionArguments } from '@microsoft/sp-adaptive-card-extension-base';
import { IQuickViewAdvancedAdaptiveCardExtensionProps, IQuickViewAdvancedAdaptiveCardExtensionState } from '../QuickViewAdvancedAdaptiveCardExtension';


export interface IDetailedViewData {
  title: string;
  description: string;
  details: string;
}

export class DetailedView extends BaseAdaptiveCardView<
  IQuickViewAdvancedAdaptiveCardExtensionProps,
  IQuickViewAdvancedAdaptiveCardExtensionState,
  IDetailedViewData
> {
  public get data(): IDetailedViewData {
    const { description, title } = this.state.items[this.state.currentIndex];
    return {
      description,
      title,
      details: 'More details'
    };
  }

  public get template(): ISPFxAdaptiveCard {
    return require('./template/DetailedQuickViewTemplate.json');
  }

  public onAction(action: ISubmitActionArguments): void {
    if (action.type === 'Submit') {
      const { id } = action.data;
      if (id === 'back') {
        this.quickViewNavigator.pop();
      }
    }
  }
}