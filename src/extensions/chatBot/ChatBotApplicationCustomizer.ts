import { override } from '@microsoft/decorators';
import { BaseApplicationCustomizer, PlaceholderContent, PlaceholderName } from '@microsoft/sp-application-base';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import * as React from 'react';
import * as ReactDom from 'react-dom';
import ChatBot from '../../components/chatbot/ChatBot';
import { AppScreen } from '../../common/types/common';

// Flag toàn cục — web part set = true khi nó tự render chatbot rồi
declare global { interface Window { __vbChatBotLoaded?: boolean; } }

export interface IChatBotApplicationCustomizerProperties {
  claudeApiKey:   string;
  bookingPageUrl: string; // URL trang Vehicle Booking để navigate khi mở ticket từ trang khác
}

export default class ChatBotApplicationCustomizer
  extends BaseApplicationCustomizer<IChatBotApplicationCustomizerProperties> {

  private _container: HTMLDivElement | undefined;

  @override
  public onInit(): Promise<void> {
    // Đợi web part load xong rồi mới kiểm tra — tránh flash trùng
    setTimeout(() => this._render(), 1500);
    return Promise.resolve();
  }

  private _render(): void {
    // Nếu web part đã tự render chatbot → không render nữa
    if (window.__vbChatBotLoaded) return;

    const apiKey = this.properties.claudeApiKey;
    if (!apiKey) return;

    // Tạo container cố định ngoài flow SharePoint
    if (!this._container) {
      this._container = document.createElement('div');
      this._container.id = 'vb-global-chatbot';
      document.body.appendChild(this._container);
    }

    const ctx = this.context;
    const user = ctx.pageContext.user;
    const bookingPageUrl = this.properties.bookingPageUrl;

    // onNavigate: nếu có bookingPageUrl thì redirect với ?requestId=
    const onNavigate = (_screen: AppScreen, params?: unknown): void => {
      if (!bookingPageUrl) return;
      const p = params as Record<string, unknown> | undefined;
      const id = p?.id as number | undefined;
      const url = id ? `${bookingPageUrl}?requestId=${id}` : bookingPageUrl;
      window.location.href = url;
    };

    ReactDom.render(
      React.createElement(ChatBot, {
        apiKey,
        userName:       user.displayName,
        userEmail:      user.email,
        userRole:       'requester',
        userDepartment: '',
        context:        ctx as unknown as WebPartContext,
        onNavigate,
      }),
      this._container
    );
  }

  @override
  public onDispose(): void {
    if (this._container) {
      ReactDom.unmountComponentAtNode(this._container);
      this._container.remove();
    }
  }
}
