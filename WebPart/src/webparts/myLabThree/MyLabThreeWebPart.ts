import { 
  Log,
  Version,
  Environment, 
  EnvironmentType  
} from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneLink,
  PropertyPaneLabel,
  PropertyPaneButton
} from '@microsoft/sp-webpart-base';
import { PropertyPaneButtonType } from '@microsoft/sp-webpart-base/lib/propertyPane/propertyPaneFields/propertyPaneButton/IPropertyPaneButton';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './MyLabThreeWebPart.module.scss';
import * as strings from 'MyLabThreeWebPartStrings';

export interface IMyLabThreeWebPartProps {
}

export default class MyLabThreeWebPartWebPart extends BaseClientSideWebPart<IMyLabThreeWebPartProps> {

  public render(): void {
    try {
      Log.verbose('SPLabThree', 'Invoking render DOM elements');
      this.domElement.innerHTML = `
        <div class="${ styles.myLabThree }">
          <div className="${ styles.container }">
            <div className="${ styles.row }">
              <div className="${ styles.column }">
                <span className="${ styles.title }">${escape(strings.WelcomeMessage)}</span>
                <p className="${ styles.subTitle }">${escape(strings.IntroductionMessage)}</p>
                <a href="${ strings.LearnMoreLinkAddress }" className="${ styles.button }">
                  <span className="${ styles.label }">${escape(strings.LearnLocaleName)}</span>
                </a>
                <div id="spContainer" />
              </div>
            </div>
          </div>
        </div>`;

        this.renderMessage();
    }
    catch(e) {
      Log.error('SPLabThree', e);
    }
    
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.DisplayGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                }),
                PropertyPaneLink('link', {
                  disabled: false,
                  href: strings.LinkAddress,
                  text: strings.LinkTextDisplay
                }),
                PropertyPaneLabel('label', {
                  required: false,
                  text: strings.LabelLocaleText
                }),
                PropertyPaneButton('click', {
                  disabled: false,
                  text: strings.ButtonLocaleName,
                  ariaDescription: 'description',
                  ariaLabel: 'label',
                  buttonType: PropertyPaneButtonType.Normal,
                  onClick: () => alert(strings.AlertMessage)
                })
              ]
            }
          ]
        }
      ]
    };
  }

  private renderMessage(): void {
    try {
      Log.verbose('SPLabThree', 'Invoking render additional message');

      let html: string = '';
      // Local environment
      html = this.checkEnvironment(Environment.type) ? 
      `<p class="${styles.description}">${strings.LocalMessage}</p>` : 
      `<p class="${styles.description}">${strings.OnlineMessage}</p>`;

      const container: Element = this.domElement.querySelector('#spContainer');
      container.innerHTML = html;
      
    }
    catch(e) {
      Log.error('SPLabThree', e);
    }
    
  }

  private checkEnvironment(value: EnvironmentType): boolean {
   return (value == EnvironmentType.Local) ? true : false;
  }
}
