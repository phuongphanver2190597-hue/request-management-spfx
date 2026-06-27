import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient } from '@microsoft/sp-http';
import { buildApprovalLink } from '../common/helpers/linkHelper';

interface ISendApprovalEmailParams {
  toEmail:      string;
  toName:       string;
  requestId:    number;
  requestCode:  string;
  requestTitle: string;
  requesterName: string;
}

export class EmailService {
  private context: WebPartContext;

  constructor(context: WebPartContext) {
    this.context = context;
  }

  async sendApprovalRequest(params: ISendApprovalEmailParams): Promise<void> {
    const link = buildApprovalLink(params.requestId);
    const siteUrl = this.context.pageContext.web.absoluteUrl;

    const body = `<div style="font-family:Segoe UI,sans-serif;max-width:600px">
      <div style="background:#0078D4;padding:20px 28px;border-radius:8px 8px 0 0">
        <h2 style="color:#fff;margin:0;font-size:18px">Yêu cầu phê duyệt mới</h2>
      </div>
      <div style="background:#fff;padding:20px 28px;border:1px solid #E1E4E8;border-top:none;border-radius:0 0 8px 8px">
        <p>Xin chào <strong>${params.toName}</strong>,</p>
        <p>Bạn có một yêu cầu đặt xe cần phê duyệt:</p>
        <table style="width:100%;border-collapse:collapse;margin:12px 0">
          <tr style="background:#F3F2F1">
            <td style="padding:9px 12px;font-weight:600;width:130px">Mã yêu cầu</td>
            <td style="padding:9px 12px">${params.requestCode}</td>
          </tr>
          <tr>
            <td style="padding:9px 12px;font-weight:600">Nội dung</td>
            <td style="padding:9px 12px">${params.requestTitle}</td>
          </tr>
          <tr style="background:#F3F2F1">
            <td style="padding:9px 12px;font-weight:600">Người yêu cầu</td>
            <td style="padding:9px 12px">${params.requesterName}</td>
          </tr>
        </table>
        <div style="text-align:center;margin:20px 0">
          <a href="${link}" style="background:#0078D4;color:#fff;padding:11px 28px;border-radius:6px;text-decoration:none;font-weight:600">
            Xem &amp; Phê duyệt
          </a>
        </div>
        <p style="color:#6A737D;font-size:12px">Link: <a href="${link}">${link}</a></p>
      </div>
    </div>`;

    const res = await this.context.spHttpClient.post(
      `${siteUrl}/_api/SP.Utilities.Utility.SendEmail`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          'Accept': 'application/json;odata=nometadata',
          'Content-type': 'application/json;odata=nometadata',
          'odata-version': '',
        },
        body: JSON.stringify({
          properties: {
            '__metadata': { 'type': 'SP.Utilities.EmailProperties' },
            'To':      { 'results': [params.toEmail] },
            'Subject': `[Phê duyệt] ${params.requestCode} — ${params.requestTitle}`,
            'Body':    body,
          },
        }),
      }
    );

    if (!res.ok && res.status !== 200) {
      const err = await res.json().catch(() => ({}));
      throw new Error((err as { error?: { message?: string } })?.error?.message || `HTTP ${res.status}`);
    }
  }
}
