import {
  BasePrimaryTextCardView,
  IPrimaryTextCardParameters,
  IExternalLinkCardAction,
  IQuickViewCardAction,
  ICardButton,
  BaseImageCardView,
  IImageCardParameters
} from '@microsoft/sp-adaptive-card-extension-base';
import * as strings from 'HelloWorldAdaptiveCardExtensionStrings';
import { IHelloWorldAdaptiveCardExtensionProps, IHelloWorldAdaptiveCardExtensionState, QUICK_VIEW_REGISTRY_ID } from '../HelloWorldAdaptiveCardExtension';

export class CardView extends BasePrimaryTextCardView<IHelloWorldAdaptiveCardExtensionProps, IHelloWorldAdaptiveCardExtensionState> {
  public get cardButtons(): [ICardButton] | [ICardButton, ICardButton] | undefined {
    return [
      {
        title: strings.QuickViewButton,
        action: {
          type: 'QuickView',
          parameters: {
            view: QUICK_VIEW_REGISTRY_ID
          }
        }
      },
      {
        title: "Hello delaware",
        action: {
          type: 'ExternalLink',
          parameters: {
            target: 'https://www.delaware.pro'
          }
        }
      }
    ];
  }

  public get data(): IPrimaryTextCardParameters {
    return {
      primaryText: this.properties.title,
      description: strings.Description,
    };

    //https://dci-prd-corpweb-frontend-web.azurewebsites.net/getattachment/0043ee7d-73d3-4dba-b94c-04cc769373cb/delaware-logo-opengraph.jpg
  }

  public get onCardSelection(): IQuickViewCardAction | IExternalLinkCardAction | undefined {
    return {
      type: 'QuickView',
      parameters: {
        view: QUICK_VIEW_REGISTRY_ID
      }
    };
  }
}
