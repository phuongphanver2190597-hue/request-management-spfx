import * as React from 'react';
import { IconButton, TextField, Spinner, SpinnerSize } from '@fluentui/react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { VehicleBookingRequestService } from '../../services/vehicleBookingRequestService';
import { IVehicleBookingRequest } from '../../models/VehicleBookingRequest';
import { AppScreen } from '../../common/types/common';

// ─── Types ────────────────────────────────────────────────────────────────────

interface IMessage { role: 'user' | 'assistant'; content: string; }

// Claude raw API message (may contain tool_use / tool_result blocks)
type ClaudeMsg =
  | { role: 'user' | 'assistant'; content: string }
  | { role: 'assistant'; content: ClaudeBlock[] }
  | { role: 'user';      content: ClaudeToolResult[] };

type ClaudeBlock =
  | { type: 'text'; text: string }
  | { type: 'tool_use'; id: string; name: string; input: Record<string, unknown> };

type ClaudeToolResult = { type: 'tool_result'; tool_use_id: string; content: string };

// OpenAI raw API message
type OAIMsg =
  | { role: 'system' | 'user' | 'assistant'; content: string; tool_calls?: OAIToolCall[] }
  | { role: 'tool'; tool_call_id: string; content: string };

type OAIToolCall = { id: string; type: 'function'; function: { name: string; arguments: string } };

// ─── Props / State ────────────────────────────────────────────────────────────

export interface IChatBotProps {
  apiKey:          string;
  userName:        string;
  userEmail:       string;
  userRole:        string;
  userDepartment:  string;
  context:         WebPartContext;
  onNavigate:      (screen: AppScreen, params?: unknown) => void;
}

interface IChatBotState {
  open:     boolean;
  messages: IMessage[];
  input:    string;
  loading:  boolean;
  error:    string | null;
}

// ─── Constants ────────────────────────────────────────────────────────────────

const PRIMARY = '#0078D4';
const BG      = '#F0F2F5';
const BORDER  = '#E1E4E8';
const MUTED   = '#6A737D';

const SYSTEM_PROMPT = `Bạn là trợ lý ảo của hệ thống đặt xe nội bộ công ty.
Bạn CÓ THỂ tra cứu dữ liệu thực tế, tạo yêu cầu đặt xe và điều hướng app bằng các công cụ được cung cấp.
Trả lời ngắn gọn, thân thiện, bằng tiếng Việt.

Khi người dùng muốn TẠO yêu cầu đặt xe:
- Hỏi lần lượt: điểm đón, điểm đến, thời gian đón (ngày giờ cụ thể), mục đích chuyến đi, số hành khách.
- Hỏi thêm: có về không (khứ hồi)? nếu có thì thời gian về?
- Sau khi có đủ thông tin, gọi tool create_booking_request.
- pickup_datetime phải ở dạng ISO: "2026-06-28T08:00:00"

QUAN TRỌNG - phân biệt rõ 2 loại:
- "yêu cầu của tôi" / "tôi đã tạo" → dùng get_my_requests (yêu cầu DO TÔI TẠO)
- "chờ tôi duyệt" / "cần tôi phê duyệt" / "chờ phê duyệt" → dùng get_pending_approvals (yêu cầu của NGƯỜI KHÁC gửi lên cho tôi duyệt)
Khi người dùng muốn xem / mở chi tiết một ticket:
- Nếu có mã VBR (VBR-YYYY-NNNNN) → dùng open_request_by_code
- Nếu có ID số → dùng navigate_to_request`;

// Tool definitions — dùng chung cho cả Claude và OpenAI
const TOOLS_SCHEMA = [
  {
    name: 'get_my_requests',
    description: 'Lấy danh sách yêu cầu đặt xe của người dùng hiện tại. Có thể lọc theo trạng thái.',
    parameters: {
      type: 'object',
      properties: {
        status: { type: 'string', description: 'Lọc theo trạng thái: Draft, Pending, Approved, Rejected... (bỏ trống = tất cả)' },
      },
    },
  },
  {
    name: 'open_request_by_code',
    description: 'Xem chi tiết hoặc mở một yêu cầu theo MÃ VBR. LUÔN dùng tool này khi người dùng nhắc đến mã VBR (ví dụ VBR-2026-70280, detail VBR-xxx, chi tiết VBR-xxx, mở VBR-xxx). Tìm trong cả yêu cầu của tôi lẫn yêu cầu chờ tôi duyệt.',
    parameters: {
      type: 'object',
      properties: {
        code: { type: 'string', description: 'Mã yêu cầu dạng VBR-YYYY-NNNNN' },
      },
      required: ['code'],
    },
  },
  {
    name: 'navigate_to_request',
    description: 'Mở màn hình chi tiết theo ID số (chỉ dùng khi có ID số, KHÔNG phải mã VBR).',
    parameters: {
      type: 'object',
      properties: {
        request_id: { type: 'number', description: 'ID số của yêu cầu' },
      },
      required: ['request_id'],
    },
  },
  {
    name: 'get_pending_approvals',
    description: 'Lấy danh sách yêu cầu đang chờ NGƯỜI DÙNG HIỆN TẠI phê duyệt (với tư cách trưởng phòng/manager). Khác với get_my_requests — đây là yêu cầu của người KHÁC gửi lên cho mình duyệt.',
    parameters: { type: 'object', properties: {} },
  },
  {
    name: 'create_booking_request',
    description: 'Tạo yêu cầu đặt xe mới dưới dạng bản nháp (Draft). Gọi khi đã có đủ thông tin từ người dùng.',
    parameters: {
      type: 'object',
      properties: {
        pickup_location:     { type: 'string',  description: 'Điểm đón' },
        dropoff_location:    { type: 'string',  description: 'Điểm đến' },
        pickup_datetime:     { type: 'string',  description: 'Thời gian đón, ISO format: 2026-06-28T08:00:00' },
        purpose:             { type: 'string',  description: 'Mục đích chuyến đi' },
        number_of_passengers:{ type: 'number',  description: 'Số hành khách' },
        is_round_trip:       { type: 'boolean', description: 'Có khứ hồi không' },
        return_datetime:     { type: 'string',  description: 'Thời gian về (nếu khứ hồi), ISO format' },
        vehicle_type:        { type: 'string',  description: 'Loại xe: Sedan, SUV, Van, Bus (mặc định Sedan)' },
        special_requirement: { type: 'string',  description: 'Yêu cầu đặc biệt nếu có' },
        phone_number:        { type: 'string',  description: 'Số điện thoại liên hệ' },
      },
      required: ['pickup_location', 'dropoff_location', 'pickup_datetime', 'purpose', 'number_of_passengers'],
    },
  },
];

// ─── Component ────────────────────────────────────────────────────────────────

export default class ChatBot extends React.Component<IChatBotProps, IChatBotState> {
  private messagesEndRef = React.createRef<HTMLDivElement>();
  private bookingSvc: VehicleBookingRequestService;

  constructor(props: IChatBotProps) {
    super(props);
    this.bookingSvc = new VehicleBookingRequestService(props.context);
    this.state = {
      open: false,
      messages: [{ role: 'assistant', content: `Xin chào ${props.userName}! Tôi có thể tra cứu, **tạo mới** yêu cầu đặt xe và mở ticket cho bạn. Bạn cần hỗ trợ gì?` }],
      input: '', loading: false, error: null,
    };
  }

  public componentDidUpdate(_: IChatBotProps, prev: IChatBotState): void {
    if (this.state.messages.length !== prev.messages.length)
      this.messagesEndRef.current?.scrollIntoView({ behavior: 'smooth' });
  }

  // ─── Tool execution ──────────────────────────────────────────────────────────

  private async _executeTool(name: string, args: Record<string, unknown>): Promise<string> {
    const { userEmail, userName, userDepartment, onNavigate } = this.props;
    try {
      if (name === 'get_my_requests') {
        const all = await this.bookingSvc.getMyRequests(userEmail);
        const filtered = args.status
          ? all.filter(r => r.Status.toLowerCase().includes(String(args.status).toLowerCase()))
          : all;
        if (!filtered.length) return 'Không có yêu cầu nào.';
        return filtered.slice(0, 10).map(r =>
          `- ${r.RequestCode} | ${r.Status} | ${r.PickupLocation} → ${r.DropoffLocation} | ${r.PickupDateTime?.slice(0, 10) || 'N/A'} | ID:${r.ID}`
        ).join('\n');
      }



      if (name === 'get_pending_approvals') {
        const list = await this.bookingSvc.getPendingApproval(userEmail);
        if (!list.length) return 'Không có yêu cầu nào đang chờ bạn phê duyệt.';
        return `Có ${list.length} yêu cầu chờ bạn phê duyệt:\n` + list.map(r =>
          `- ${r.RequestCode} | ${r.RequesterName} | ${r.PickupLocation} → ${r.DropoffLocation} | ${r.PickupDateTime?.slice(0, 10) || 'N/A'} | ID:${r.ID}`
        ).join('\n');
      }

      if (name === 'navigate_to_request') {
        const id = Number(args.request_id);
        onNavigate('request-detail', { id });
        return `Đã mở ticket ID ${id}.`;
      }

      if (name === 'open_request_by_code') {
        const [mine, pending] = await Promise.all([
          this.bookingSvc.getMyRequests(userEmail),
          this.bookingSvc.getPendingApproval(userEmail),
        ]);
        const code = String(args.code).toLowerCase().trim();
        const r = [...mine, ...pending].find(x => x.RequestCode?.toLowerCase() === code);
        if (!r) return `Không tìm thấy yêu cầu ${args.code}.`;
        onNavigate('request-detail', { id: r.ID });
        return `Đã mở ticket ${r.RequestCode} (ID: ${r.ID}).`;
      }

      if (name === 'create_booking_request') {
        const created = await this.bookingSvc.createDraft(
          {
            pickupLocation:      String(args.pickup_location || ''),
            dropoffLocation:     String(args.dropoff_location || ''),
            pickupDateTime:      String(args.pickup_datetime || ''),
            returnDateTime:      String(args.return_datetime || ''),
            isRoundTrip:         Boolean(args.is_round_trip),
            numberOfPassengers:  Number(args.number_of_passengers || 1),
            purpose:             String(args.purpose || ''),
            vehicleType:         String(args.vehicle_type || 'Sedan'),
            specialRequirement:  String(args.special_requirement || ''),
            phoneNumber:         String(args.phone_number || ''),
            department:          userDepartment,
          },
          { id: userEmail, name: userName, email: userEmail }
        );
        onNavigate('request-detail', { id: created.ID });
        return `Đã tạo yêu cầu **${created.RequestCode}** thành công (trạng thái: Nháp).\nĐang mở chi tiết để bạn kiểm tra và gửi phê duyệt.`;
      }

      return 'Tool không xác định.';
    } catch (err) {
      return `Lỗi: ${err instanceof Error ? err.message : String(err)}`;
    }
  }

  // ─── Claude API với tool use ──────────────────────────────────────────────────

  private async _runClaude(history: ClaudeMsg[]): Promise<string> {
    const claudeTools = TOOLS_SCHEMA.map(t => ({
      name: t.name,
      description: t.description,
      input_schema: t.parameters,
    }));

    for (let turn = 0; turn < 5; turn++) {
      const res = await fetch('https://api.anthropic.com/v1/messages', {
        method: 'POST',
        headers: {
          'x-api-key': this.props.apiKey,
          'anthropic-version': '2023-06-01',
          'content-type': 'application/json',
          'anthropic-dangerous-direct-browser-access': 'true',
        },
        body: JSON.stringify({
          model: 'claude-haiku-4-5-20251001',
          max_tokens: 1024,
          system: SYSTEM_PROMPT,
          tools: claudeTools,
          messages: history,
        }),
      });

      if (!res.ok) {
        const err = await res.json().catch(() => ({}));
        throw new Error((err as { error?: { message?: string } }).error?.message || `HTTP ${res.status}`);
      }

      const data = await res.json() as { stop_reason: string; content: ClaudeBlock[] };

      if (data.stop_reason === 'tool_use') {
        // Thêm assistant message chứa tool_use blocks
        history.push({ role: 'assistant', content: data.content });

        // Thực hiện tất cả tool calls trong lần này
        const toolResults: ClaudeToolResult[] = [];
        for (const block of data.content) {
          if (block.type === 'tool_use') {
            const result = await this._executeTool(block.name, block.input);
            toolResults.push({ type: 'tool_result', tool_use_id: block.id, content: result });
          }
        }
        history.push({ role: 'user', content: toolResults });
        continue;
      }

      // stop_reason === 'end_turn'
      const textBlock = data.content.find(b => b.type === 'text') as { type: 'text'; text: string } | undefined;
      return textBlock?.text || '(Không có phản hồi)';
    }
    return 'Đã thử quá nhiều lần, vui lòng thử lại.';
  }

  // ─── OpenAI API với function calling ─────────────────────────────────────────

  private async _runOpenAI(history: OAIMsg[]): Promise<string> {
    const oaiTools = TOOLS_SCHEMA.map(t => ({
      type: 'function' as const,
      function: { name: t.name, description: t.description, parameters: t.parameters },
    }));

    for (let turn = 0; turn < 5; turn++) {
      const res = await fetch('https://api.openai.com/v1/chat/completions', {
        method: 'POST',
        headers: { 'Authorization': `Bearer ${this.props.apiKey}`, 'content-type': 'application/json' },
        body: JSON.stringify({ model: 'gpt-4o-mini', max_tokens: 1024, tools: oaiTools, messages: history }),
      });

      if (!res.ok) {
        const err = await res.json().catch(() => ({}));
        throw new Error((err as { error?: { message?: string } }).error?.message || `HTTP ${res.status}`);
      }

      const data = await res.json() as { choices: { finish_reason: string; message: OAIMsg & { tool_calls?: OAIToolCall[] } }[] };
      const msg = data.choices[0].message;
      history.push(msg);

      if (data.choices[0].finish_reason === 'tool_calls' && msg.tool_calls) {
        for (const tc of msg.tool_calls) {
          const args = JSON.parse(tc.function.arguments) as Record<string, unknown>;
          const result = await this._executeTool(tc.function.name, args);
          history.push({ role: 'tool', tool_call_id: tc.id, content: result });
        }
        continue;
      }

      return (msg as { content: string }).content || '(Không có phản hồi)';
    }
    return 'Đã thử quá nhiều lần, vui lòng thử lại.';
  }

  // ─── Send ─────────────────────────────────────────────────────────────────────

  private _isOpenAI(): boolean {
    return this.props.apiKey.startsWith('sk-') && !this.props.apiKey.startsWith('sk-ant-');
  }

  private async _send(): Promise<void> {
    const { input, messages } = this.state;
    const text = input.trim();
    if (!text || this.state.loading) return;

    const newMessages: IMessage[] = [...messages, { role: 'user', content: text }];
    this.setState({ messages: newMessages, input: '', loading: true, error: null });

    try {
      let reply: string;
      if (this._isOpenAI()) {
        const history: OAIMsg[] = [
          { role: 'system', content: SYSTEM_PROMPT },
          ...newMessages.map(m => ({ role: m.role as 'user' | 'assistant', content: m.content })),
        ];
        reply = await this._runOpenAI(history);
      } else {
        const history: ClaudeMsg[] = newMessages.map(m => ({ role: m.role, content: m.content }));
        reply = await this._runClaude(history);
      }

      this.setState({ messages: [...newMessages, { role: 'assistant', content: reply }], loading: false });
    } catch (err) {
      this.setState({ loading: false, error: err instanceof Error ? err.message : 'Lỗi không xác định' });
    }
  }

  private _onKeyDown = (e: React.KeyboardEvent): void => {
    if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); this._send().catch(console.error); }
  }

  // ─── Render ───────────────────────────────────────────────────────────────────

  public render(): React.ReactElement {
    const { open, messages, input, loading, error } = this.state;
    if (!this.props.apiKey) return <></>;

    const provider = this._isOpenAI() ? 'GPT-4o Mini' : 'Claude Haiku';

    return (
      <div style={{ position: 'fixed', bottom: 24, right: 24, zIndex: 9999 }}>
        {open && (
          <div style={{
            position: 'absolute', bottom: 64, right: 0, width: 340, height: 500,
            background: '#fff', borderRadius: 14, border: `1px solid ${BORDER}`,
            boxShadow: '0 8px 32px rgba(0,0,0,0.18)',
            display: 'flex', flexDirection: 'column', overflow: 'hidden',
          }}>
            {/* Header */}
            <div style={{ background: PRIMARY, padding: '12px 16px', display: 'flex', alignItems: 'center', gap: 10 }}>
              <div style={{ width: 32, height: 32, borderRadius: '50%', background: 'rgba(255,255,255,0.2)', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 16 }}>🤖</div>
              <div>
                <div style={{ color: '#fff', fontWeight: 700, fontSize: 14 }}>Trợ lý đặt xe</div>
                <div style={{ color: 'rgba(255,255,255,0.75)', fontSize: 10 }}>Powered by {provider}</div>
              </div>
              <IconButton iconProps={{ iconName: 'Cancel' }} onClick={() => this.setState({ open: false })}
                styles={{ root: { marginLeft: 'auto' }, icon: { color: '#fff', fontSize: 14 }, rootHovered: { background: 'rgba(255,255,255,0.15)' } }} />
            </div>

            {/* Messages */}
            <div style={{ flex: 1, overflowY: 'auto', padding: '12px 12px 4px', display: 'flex', flexDirection: 'column', gap: 8 }}>
              {messages.map((msg, i) => (
                <div key={i} style={{ display: 'flex', justifyContent: msg.role === 'user' ? 'flex-end' : 'flex-start' }}>
                  <div style={{
                    maxWidth: '82%', padding: '8px 12px', fontSize: 13, lineHeight: 1.5,
                    borderRadius: msg.role === 'user' ? '14px 14px 4px 14px' : '14px 14px 14px 4px',
                    background: msg.role === 'user' ? PRIMARY : BG,
                    color: msg.role === 'user' ? '#fff' : '#1B1F23',
                    whiteSpace: 'pre-wrap' as const, wordBreak: 'break-word' as const,
                    boxShadow: '0 1px 2px rgba(0,0,0,0.07)',
                  }}>
                    {msg.content}
                  </div>
                </div>
              ))}
              {loading && (
                <div style={{ display: 'flex', justifyContent: 'flex-start' }}>
                  <div style={{ padding: '10px 14px', background: BG, borderRadius: '14px 14px 14px 4px', display: 'flex', alignItems: 'center', gap: 8 }}>
                    <Spinner size={SpinnerSize.xSmall} />
                    <span style={{ fontSize: 12, color: MUTED }}>Đang xử lý...</span>
                  </div>
                </div>
              )}
              {error && (
                <div style={{ fontSize: 12, color: '#C53030', background: '#FFF5F5', padding: '8px 12px', borderRadius: 8, textAlign: 'center' as const }}>
                  Lỗi: {error}
                </div>
              )}
              <div ref={this.messagesEndRef} />
            </div>

            {/* Quick prompts */}
            {messages.length === 1 && (
              <div style={{ padding: '0 12px 8px', display: 'flex', flexWrap: 'wrap' as const, gap: 6 }}>
                {['Tạo yêu cầu mới', 'Yêu cầu của tôi', 'Chờ tôi phê duyệt'].map(q => (
                  <button key={q} onClick={() => this.setState({ input: q }, () => this._send().catch(console.error))}
                    style={{ fontSize: 11, padding: '4px 10px', borderRadius: 12, border: `1px solid ${PRIMARY}`, background: '#fff', color: PRIMARY, cursor: 'pointer' }}>
                    {q}
                  </button>
                ))}
              </div>
            )}

            {/* Input */}
            <div style={{ padding: '10px 12px', borderTop: `1px solid ${BORDER}`, display: 'flex', gap: 8, alignItems: 'flex-end' }}>
              <TextField value={input} onChange={(_, v) => this.setState({ input: v || '' })} onKeyDown={this._onKeyDown}
                placeholder="Nhập câu hỏi... (Enter để gửi)" multiline autoAdjustHeight disabled={loading}
                styles={{ root: { flex: 1 }, field: { fontSize: 13, minHeight: 36, maxHeight: 80, resize: 'none' }, fieldGroup: { borderRadius: 8 } }} />
              <IconButton iconProps={{ iconName: 'Send' }} disabled={loading || !input.trim()} onClick={() => this._send().catch(console.error)}
                styles={{ root: { background: PRIMARY, borderRadius: 8, width: 36, height: 36, opacity: (!input.trim() || loading) ? 0.5 : 1 }, icon: { color: '#fff', fontSize: 14 }, rootHovered: { background: '#006CBE' }, rootDisabled: { background: PRIMARY } }} />
            </div>
          </div>
        )}

        {/* Toggle button */}
        <button onClick={() => this.setState(prev => ({ open: !prev.open }))}
          style={{ width: 52, height: 52, borderRadius: '50%', background: PRIMARY, border: 'none', cursor: 'pointer', boxShadow: '0 4px 16px rgba(0,120,212,0.4)', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 22 }}
          title="Trợ lý đặt xe">
          {open ? '✕' : '🤖'}
        </button>
      </div>
    );
  }
}

// Helper để format request thành text ngắn gọn
// eslint-disable-next-line @typescript-eslint/no-unused-vars
function _fmt(r: IVehicleBookingRequest): string {
  return `${r.RequestCode} | ${r.Status} | ${r.PickupLocation}→${r.DropoffLocation}`;
}
