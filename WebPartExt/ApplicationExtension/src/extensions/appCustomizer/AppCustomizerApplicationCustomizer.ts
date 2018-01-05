import styles from './AppCustomizerApplicationCustomizer.module.scss';
import { escape } from '@microsoft/sp-lodash-subset'; 
import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName
} from '@microsoft/sp-application-base';
import { Dialog } from '@microsoft/sp-dialog';

import * as strings from 'AppCustomizerApplicationCustomizerStrings';

const LOG_SOURCE: string = 'AppCustomizerApplicationCustomizer';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IAppCustomizerApplicationCustomizerProperties {
  // This is an example; replace with your own property
  top: string;
  bottom: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class AppCustomizerApplicationCustomizer
  extends BaseApplicationCustomizer<IAppCustomizerApplicationCustomizerProperties> {

    private topPlaceholder: PlaceholderContent | undefined;
    private bottomPlaceholder: PlaceholderContent | undefined;

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    // Added to handle possible changes on the existence of placeholders.
    this.context.placeholderProvider.changedEvent.add(this, this._renderPlaceHolders);

    // Call render method for generating the HTML elements.
    this._renderPlaceHolders();
    return Promise.resolve<void>();
  }

  private _renderPlaceHolders(): void {
    console.log('HelloWorldApplicationCustomizer.renderPlaceHolders()');
    console.log('Available placeholders: ', 
      this.context.placeholderProvider
        .placeholderNames
        .map(name => PlaceholderName[name]).join(', '));

      if (!this.topPlaceholder) {
        this.topPlaceholder =
          this.context.placeholderProvider.tryCreateContent(
            PlaceholderName.Top,
            { onDispose: this._onDispose });
      
        // The extension should not assume that the expected placeholder is available.
        if (!this.topPlaceholder) {
          console.error('The expected placeholder (Top) was not found.');
          return;
        }
      
        if (this.properties) {
          let topString: string = this.properties.top;
          if (!topString) {
            topString = '(Top property was not defined.)';
          }
      
          if (this.topPlaceholder.domElement) {
            this.topPlaceholder.domElement.innerHTML = `
              <div class="${styles.app}">
                <div class="ms-bgColor-themeDark ms-fontColor-white ${styles.top}">
                  <i class="ms-Icon ms-Icon--Info" aria-hidden="true"></i> ${escape(topString)}
                </div>
              </div>`;
          }
        }
      }

      // Handling the bottom placeholder
    if (!this.bottomPlaceholder) {
      this.bottomPlaceholder =
        this.context.placeholderProvider.tryCreateContent(
          PlaceholderName.Bottom,
          { onDispose: this._onDispose });
    
      // The extension should not assume that the expected placeholder is available.
      if (!this.bottomPlaceholder) {
        console.error('The expected placeholder (Bottom) was not found.');
        return;
      }
    
      if (this.properties) {
        let bottomString: string = this.properties.bottom;
        if (!bottomString) {
          bottomString = '(Bottom property was not defined.)';
        }
    
        if (this.bottomPlaceholder.domElement) {
          this.bottomPlaceholder.domElement.innerHTML = `
            <div class="${styles.app}">
              <div class="ms-bgColor-themeDark ms-fontColor-white ${styles.bottom}">
                <i class="ms-Icon ms-Icon--Info" aria-hidden="true"></i> ${escape(bottomString)}
              </div>
            </div>`;
        }
      }
    }

  }

  private _onDispose(): void {
    console.log('[HelloWorldApplicationCustomizer._onDispose] Disposed custom top and bottom placeholders.');
  }
}
