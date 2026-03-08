import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import HwpxCollabApp from './components/HwpxCollabApp';
import { IHwpxCollabAppProps } from './components/IHwpxCollabAppProps';

export default class HwpxCollabWebPart extends BaseClientSideWebPart<{}> {

  public render(): void {
    const userName: string = this.context.pageContext.user.displayName || '편집자';
    const userEmail: string = this.context.pageContext.user.email || '';
    const userLoginName: string = this.context.pageContext.user.loginName || '';
    const userColor: string = this._generateColor(userEmail || userName);

    const element: React.ReactElement<IHwpxCollabAppProps> = React.createElement(
      HwpxCollabApp,
      {
        userName: userName,
        userEmail: userEmail,
        userLoginName: userLoginName,
        userColor: userColor,
        wsUrl: 'wss://collab.gfea-scratch.org',
        spHttpClient: this.context.spHttpClient,
        siteUrl: this.context.pageContext.web.absoluteUrl,
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  private _generateColor(seed: string): string {
    const COLORS: string[] = [
      '#4F8EF7', '#10B981', '#F59E0B', '#EF4444', '#8B5CF6',
      '#EC4899', '#14B8A6', '#F97316', '#6366F1', '#84CC16',
    ];
    let hash: number = 0;
    for (let i: number = 0; i < seed.length; i++) {
      hash = seed.charCodeAt(i) + ((hash << 5) - hash);
    }
    return COLORS[Math.abs(hash) % COLORS.length];
  }
}
