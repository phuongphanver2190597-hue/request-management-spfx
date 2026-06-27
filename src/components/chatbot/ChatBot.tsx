import * as React from 'react';
import { IconButton, TextField, Spinner, SpinnerSize } from '@fluentui/react';

interface IMessage {
  role: 'user' | 'assistant';
  content: string;
}

interface IChatBotProps {
  apiKey: string;
  userName: string;
  userRole: string;
}

interface IChatBotState {
  open: boolean;
  messages: IMessage[];
  input: string;
  loading: boolean;
  error: string | null;
}

const SYSTEM_PROMPT = `Bạn là trợ lý ảo của hệ thống đặt xe nội bộ công ty. Hỗ trợ nhân viên với các câu hỏi về:
- Quy trình đặt xe: tạo yêu cầu, gửi phê duyệt, theo dõi trạng thái
- Các trạng thái yêu cầu: Nháp, Chờ phê duyệt, Đã phê duyệt, Từ chối, Cần bổ sung
- Hướng dẫn sử dụng app
- Chính sách đặt xe của công ty
Trả lời ngắn gọn, thân thiện, bằng tiếng Việt. Nếu không biết thì hướng dẫn liên hệ bộ phận hành chính.`;

const COLORS = {
  primary: '#0078D4',
  bg: '#F0F2F5',
  white: '#FFFFFF',
  border: '#E1E4E8',
  muted: '#6A737D',
  userBubble: '#0078D4',
  botBubble: '#FFFFFF',
};

export default class ChatBot extends React.Component<IChatBotProps, IChatBotState> {
  private messagesEndRef = React.createRef<HTMLDivElement>();

  constructor(props: IChatBotProps) {
    super(props);
    this.state = {
      open: false,
      messages: [{
        role: 'assistant',
        content: `Xin chào ${props.userName}! Tôi là trợ lý đặt xe. Bạn cần hỗ trợ gì?`,
      }],
      input: '',
      loading: false,
      error: null,
    };
  }

  public componentDidUpdate(_: IChatBotProps, prev: IChatBotState): void {
    if (this.state.messages.length !== prev.messages.length) {
      this.messagesEndRef.current?.scrollIntoView({ behavior: 'smooth' });
    }
  }

  private _isOpenAI(): boolean {
    const key = this.props.apiKey;
    return key.startsWith('sk-') && !key.startsWith('sk-ant-');
  }

  private async _callClaude(messages: IMessage[]): Promise<string> {
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
        messages: messages.map(m => ({ role: m.role, content: m.content })),
      }),
    });
    if (!res.ok) {
      const err = await res.json().catch(() => ({}));
      throw new Error((err as { error?: { message?: string } }).error?.message || `HTTP ${res.status}`);
    }
    const data = await res.json() as { content: { text: string }[] };
    return data.content?.[0]?.text || '(Không có phản hồi)';
  }

  private async _callOpenAI(messages: IMessage[]): Promise<string> {
    const res = await fetch('https://api.openai.com/v1/chat/completions', {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${this.props.apiKey}`,
        'content-type': 'application/json',
      },
      body: JSON.stringify({
        model: 'gpt-4o-mini',
        max_tokens: 1024,
        messages: [
          { role: 'system', content: SYSTEM_PROMPT },
          ...messages.map(m => ({ role: m.role, content: m.content })),
        ],
      }),
    });
    if (!res.ok) {
      const err = await res.json().catch(() => ({}));
      throw new Error((err as { error?: { message?: string } }).error?.message || `HTTP ${res.status}`);
    }
    const data = await res.json() as { choices: { message: { content: string } }[] };
    return data.choices?.[0]?.message?.content || '(Không có phản hồi)';
  }

  private async _send(): Promise<void> {
    const { input, messages } = this.state;
    const text = input.trim();
    if (!text || this.state.loading) return;

    const newMessages: IMessage[] = [...messages, { role: 'user', content: text }];
    this.setState({ messages: newMessages, input: '', loading: true, error: null });

    try {
      const reply = this._isOpenAI()
        ? await this._callOpenAI(newMessages)
        : await this._callClaude(newMessages);

      this.setState({
        messages: [...newMessages, { role: 'assistant', content: reply }],
        loading: false,
      });
    } catch (err) {
      this.setState({
        loading: false,
        error: err instanceof Error ? err.message : 'Lỗi không xác định',
      });
    }
  }

  private _onKeyDown = (e: React.KeyboardEvent): void => {
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault();
      this._send().catch(console.error);
    }
  }

  public render(): React.ReactElement {
    const { open, messages, input, loading, error } = this.state;
    const { apiKey } = this.props;

    if (!apiKey) return <></>;

    return (
      <div style={{ position: 'fixed', bottom: 24, right: 24, zIndex: 9999 }}>
        {/* Chat panel */}
        {open && (
          <div style={{
            position: 'absolute', bottom: 64, right: 0,
            width: 340, height: 480,
            background: COLORS.white,
            borderRadius: 14,
            boxShadow: '0 8px 32px rgba(0,0,0,0.18)',
            display: 'flex', flexDirection: 'column',
            overflow: 'hidden',
            border: `1px solid ${COLORS.border}`,
          }}>
            {/* Header */}
            <div style={{
              background: COLORS.primary, padding: '12px 16px',
              display: 'flex', alignItems: 'center', gap: 10,
            }}>
              <div style={{
                width: 32, height: 32, borderRadius: '50%',
                background: 'rgba(255,255,255,0.2)',
                display: 'flex', alignItems: 'center', justifyContent: 'center',
                fontSize: 16,
              }}>🤖</div>
              <div>
                <div style={{ color: '#fff', fontWeight: 700, fontSize: 14 }}>Trợ lý đặt xe</div>
                <div style={{ color: 'rgba(255,255,255,0.8)', fontSize: 11 }}>Powered by Claude</div>
              </div>
              <IconButton
                iconProps={{ iconName: 'Cancel' }}
                onClick={() => this.setState({ open: false })}
                styles={{
                  root: { marginLeft: 'auto', color: '#fff', background: 'transparent' },
                  icon: { color: '#fff', fontSize: 14 },
                  rootHovered: { background: 'rgba(255,255,255,0.15)' },
                }}
              />
            </div>

            {/* Messages */}
            <div style={{
              flex: 1, overflowY: 'auto', padding: '12px 12px 4px',
              display: 'flex', flexDirection: 'column', gap: 8,
            }}>
              {messages.map((msg, i) => (
                <div key={i} style={{
                  display: 'flex',
                  justifyContent: msg.role === 'user' ? 'flex-end' : 'flex-start',
                }}>
                  <div style={{
                    maxWidth: '80%',
                    padding: '8px 12px',
                    borderRadius: msg.role === 'user' ? '14px 14px 4px 14px' : '14px 14px 14px 4px',
                    background: msg.role === 'user' ? COLORS.userBubble : COLORS.bg,
                    color: msg.role === 'user' ? '#fff' : '#1B1F23',
                    fontSize: 13,
                    lineHeight: 1.5,
                    boxShadow: '0 1px 2px rgba(0,0,0,0.08)',
                    whiteSpace: 'pre-wrap' as const,
                    wordBreak: 'break-word' as const,
                  }}>
                    {msg.content}
                  </div>
                </div>
              ))}
              {loading && (
                <div style={{ display: 'flex', justifyContent: 'flex-start' }}>
                  <div style={{
                    padding: '10px 14px', background: COLORS.bg,
                    borderRadius: '14px 14px 14px 4px',
                    display: 'flex', alignItems: 'center', gap: 8,
                  }}>
                    <Spinner size={SpinnerSize.xSmall} />
                    <span style={{ fontSize: 12, color: COLORS.muted }}>Đang trả lời...</span>
                  </div>
                </div>
              )}
              {error && (
                <div style={{
                  fontSize: 12, color: '#C53030', background: '#FFF5F5',
                  padding: '8px 12px', borderRadius: 8, textAlign: 'center' as const,
                }}>
                  Lỗi: {error}
                </div>
              )}
              <div ref={this.messagesEndRef} />
            </div>

            {/* Input */}
            <div style={{
              padding: '10px 12px',
              borderTop: `1px solid ${COLORS.border}`,
              display: 'flex', gap: 8, alignItems: 'flex-end',
            }}>
              <TextField
                value={input}
                onChange={(_, v) => this.setState({ input: v || '' })}
                onKeyDown={this._onKeyDown}
                placeholder="Nhập câu hỏi... (Enter để gửi)"
                multiline autoAdjustHeight
                disabled={loading}
                styles={{
                  root: { flex: 1 },
                  field: { fontSize: 13, minHeight: 36, maxHeight: 80, resize: 'none' },
                  fieldGroup: { borderRadius: 8 },
                }}
              />
              <IconButton
                iconProps={{ iconName: 'Send' }}
                disabled={loading || !input.trim()}
                onClick={() => this._send().catch(console.error)}
                styles={{
                  root: {
                    background: COLORS.primary, borderRadius: 8,
                    width: 36, height: 36,
                    opacity: (!input.trim() || loading) ? 0.5 : 1,
                  },
                  icon: { color: '#fff', fontSize: 14 },
                  rootHovered: { background: '#006CBE' },
                  rootDisabled: { background: COLORS.primary },
                }}
              />
            </div>
          </div>
        )}

        {/* Toggle button */}
        <button
          onClick={() => this.setState(prev => ({ open: !prev.open }))}
          style={{
            width: 52, height: 52, borderRadius: '50%',
            background: COLORS.primary,
            border: 'none', cursor: 'pointer',
            boxShadow: '0 4px 16px rgba(0,120,212,0.4)',
            display: 'flex', alignItems: 'center', justifyContent: 'center',
            fontSize: 22,
            transition: 'transform 0.2s',
          }}
          title="Trợ lý đặt xe"
        >
          {open ? '✕' : '🤖'}
        </button>
      </div>
    );
  }
}
