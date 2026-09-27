/**
 * SPO-Hub Site Footer – Application Customizer
 * Renders into PlaceholderName.Bottom on modern SharePoint pages.
 */

import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName
} from '@microsoft/sp-application-base';
import * as React from 'react';
import * as ReactDom from 'react-dom';

import { SiteFooter, ISiteFooterProps, IFooterLink } from '../../components/footer/SiteFooter';
import { SharePointListDataService } from '../../services/SharePointListDataService';

const LOG_SOURCE = 'FooterApplicationCustomizer';

export interface IFooterApplicationCustomizerProperties {
  listTitle?: string;
  copyrightText?: string;
  showCopyright?: boolean;
  backgroundColor?: string;
  textColor?: string;
  enableListLinks?: boolean;
  maxLinks?: number;
  layout?: 'columns' | 'inline';
  staticLinksJson?: string;
}

export default class FooterApplicationCustomizer
  extends BaseApplicationCustomizer<IFooterApplicationCustomizerProperties> {

  private _bottomPlaceholder: PlaceholderContent | undefined;
  private _footerElement: HTMLElement | undefined;

  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized FooterApplicationCustomizer`);
    this.context.placeholderProvider.changedEvent.add(this, this._renderPlaceHolders);
    return this._renderPlaceHolders();
  }

  private async _renderPlaceHolders(): Promise<void> {
    if (!this._bottomPlaceholder) {
      this._bottomPlaceholder = this.context.placeholderProvider.tryCreateContent(
        PlaceholderName.Bottom,
        { onDispose: this._onDispose }
      );
      if (!this._bottomPlaceholder) {
        Log.error(LOG_SOURCE, new Error('Bottom placeholder was not found.'));
        return;
      }
    }

    if (!this._bottomPlaceholder.domElement) {
      return;
    }

    this._footerElement = this._bottomPlaceholder.domElement;
    const links = await this._loadLinks();

    const props: ISiteFooterProps = {
      links,
      copyrightText: this.properties.copyrightText
        || `\u00a9 ${new Date().getFullYear()}. All rights reserved.`,
      showCopyright: this.properties.showCopyright !== false,
      backgroundColor: this.properties.backgroundColor,
      textColor: this.properties.textColor,
      layout: this.properties.layout === 'inline' ? 'inline' : 'columns',
      siteTitle: this.context.pageContext?.web?.title,
      siteUrl: this.context.pageContext?.web?.absoluteUrl
    };

    ReactDom.render(React.createElement(SiteFooter, props), this._footerElement);
  }

  private async _loadLinks(): Promise<IFooterLink[]> {
    if (this.properties.staticLinksJson) {
      try {
        const parsed = JSON.parse(this.properties.staticLinksJson);
        if (Array.isArray(parsed)) {
          return parsed.map((x: any, i: number) => ({
            id: i,
            title: x.title || x.Title || '',
            url: x.url || x.Url || x.link || '#',
            category: x.category || x.Category,
            openInNewTab: x.openInNewTab !== false
          })).filter((l: IFooterLink) => !!l.title);
        }
      } catch {
        Log.warn(LOG_SOURCE, 'staticLinksJson is not valid JSON');
      }
    }

    if (this.properties.enableListLinks !== false && this.properties.listTitle) {
      try {
        const svc = new SharePointListDataService(this.context as any);
        const top = this.properties.maxLinks || 12;
        const result = await svc.getItems<any>({
          listTitle: this.properties.listTitle,
          orderBy: 'DisplayOrder',
          orderDirection: 'asc',
          top,
          select: ['Id', 'Title', 'Link', 'Url', 'Category', 'DisplayOrder', 'IsActive', 'OpenInNewTab'],
          filter: `(IsActive eq 1 or IsActive eq null)`
        });

        return result.items.map((raw: any) => {
          const urlField = raw.Link || raw.Url || raw.link;
          const url = typeof urlField === 'object' && urlField?.Url
            ? urlField.Url
            : (urlField || '#');
          return {
            id: raw.Id || raw.ID || 0,
            title: raw.Title || '',
            url,
            category: raw.Category,
            openInNewTab: raw.OpenInNewTab !== false && raw.OpenInNewTab !== 0
          } as IFooterLink;
        }).filter((l: IFooterLink) => !!l.title);
      } catch {
        Log.warn(LOG_SOURCE, `Could not load footer links from list "${this.properties.listTitle}".`);
      }
    }

    return [];
  }

  private _onDispose = (): void => {
    if (this._footerElement) {
      ReactDom.unmountComponentAtNode(this._footerElement);
      this._footerElement = undefined;
    }
  }

  protected onDispose(): void {
    this._onDispose();
    super.onDispose();
  }
}
