import { ISPFxAdaptiveCard, BaseAdaptiveCardView, ISubmitActionArguments } from '@microsoft/sp-adaptive-card-extension-base';
import { IHelloWorldAdaptiveCardExtensionProps, IHelloWorldAdaptiveCardExtensionState } from '../HelloWorldAdaptiveCardExtension';

export interface IQuickViewData {
  title: string;
  subTitle: string;
}

export class QuickView extends BaseAdaptiveCardView<
  IHelloWorldAdaptiveCardExtensionProps,
  IHelloWorldAdaptiveCardExtensionState,
  IQuickViewData
> {
  public get data(): IQuickViewData {
    return {
      title: this.properties.title,
      subTitle: this.state.subTitle
    };
  }

  public get template(): ISPFxAdaptiveCard {
    return require('./template/QuickViewTemplate.json');
  }

  public onAction(action: ISubmitActionArguments): void {
    if (action.type === 'Submit') {
      const { id, message } = action.data;
      switch (id) {
        case 'button1':
        case 'button2':
          this.setState({
            subTitle: message
          });
          break;
      }
    }
  }
}