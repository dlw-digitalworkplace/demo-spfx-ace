import {
  BasePrimaryTextCardView,
  IPrimaryTextCardParameters,
  IExternalLinkCardAction,
  IQuickViewCardAction,
  ICardButton,
  BaseBasicCardView,
  IBasicCardParameters,
  ISubmitActionArguments
} from '@microsoft/sp-adaptive-card-extension-base';
import * as strings from 'QuickViewAdvancedAdaptiveCardExtensionStrings';
import { IQuickViewAdvancedAdaptiveCardExtensionProps, IQuickViewAdvancedAdaptiveCardExtensionState, QUICK_VIEW_REGISTRY_ID } from '../QuickViewAdvancedAdaptiveCardExtension';

export class CardView extends BaseBasicCardView<IQuickViewAdvancedAdaptiveCardExtensionProps, IQuickViewAdvancedAdaptiveCardExtensionState> {
  public get cardButtons(): [ICardButton] | [ICardButton, ICardButton] | undefined {
    const buttons: ICardButton[] = [];

    // Hide the Previous button if at Step 1
    if (this.state.currentIndex > 0) {
      buttons.push({
        title: 'Previous',
        action: {
          type: 'Submit',
          parameters: {
            id: 'previous',
            op: -1 // Decrement the index
          }
        }
      });
    }

    // Hide the Next button if at the end
    if (this.state.currentIndex < this.state.items.length - 1) {
      buttons.push({
        title: 'Next',
        action: {
          type: 'Submit',
          parameters: {
            id: 'next',
            op: 1 // Increment the index
          }
        }
      });
    }

    return buttons as [ICardButton] | [ICardButton, ICardButton];
  }

  public get data(): IBasicCardParameters {
    const { title, description } = this.state.items[this.state.currentIndex];
    return {
      title: title,
      primaryText: description
    };
  }

  public get onCardSelection(): IQuickViewCardAction | IExternalLinkCardAction | undefined {
    return {
      type: 'QuickView',
      parameters: {
        view: QUICK_VIEW_REGISTRY_ID
      }
    };
  }

  public onAction(action: ISubmitActionArguments): void {
    if (action.type === 'Submit') {
      const { id, op } = action.data;
      switch (id) {
        case 'previous':
        case 'next':
          this.setState({ currentIndex: this.state.currentIndex + op });
          break;
      }
    }
  }
}
