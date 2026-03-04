import * as React from "react";
import { CopilotApiLibrary } from "../../CopilotApiLibrary";

export interface ICopilotChatProps {
  initialMessage?: string;
  existingConversationId?: string;
}

/**
 * Inner React component that performs the actual rendering and Graph API calls
 * while encapsulating all mutable state to adhere to PnP BaseWebComponent rules.
 */
export const CopilotChat: React.FC<ICopilotChatProps> = ({ initialMessage, existingConversationId }) => {
  const [conversationId, setConversationId] = React.useState<string | undefined>(existingConversationId);
  const [messages, setMessages] = React.useState<any[]>([]);
  const [loading, setLoading] = React.useState<boolean>(false);
  const [error, setError] = React.useState<string | undefined>();
  const [input, setInput] = React.useState<string>(initialMessage || "");

  // Fetch or initialize the conversation
  const sendMessage = async () => {
    if (!input || input.trim().length === 0) return;

    setLoading(true);
    setError(undefined);

    try {
      const client = CopilotApiLibrary.msGraphClient;
      if (!client) {
        throw new Error("MS Graph Client is not initialized.");
      }

      let activeConvoId = conversationId;

      // 1. Create a conversation if one doesn't exist
      if (!activeConvoId) {
        const convoResponse = await client.api("/copilot/conversations").version("beta").post({});
        activeConvoId = convoResponse.id;
        setConversationId(activeConvoId);
      }

      // 2. Add the user's message to local state
      const userMessage = { author: "User", text: input };
      setMessages((prev) => [...prev, userMessage]);

      // 3. Send message to the specific conversation endpoint
      const chatResponse = await client
        .api(`/copilot/conversations/${activeConvoId}/chat`)
        .version("beta")
        .post({
          message: { text: input.trim() },
          locationHint: { timeZone: Intl.DateTimeFormat().resolvedOptions().timeZone || "UTC" }
        });

      // 4. Record Copilot's response
      // The Graph API returns the full conversation turn
      const copilotResponse = chatResponse.messages?.find((m: any) => m.author === "Copilot")?.text || "No response generated.";

      setMessages((prev) => [...prev, { author: "Copilot", text: copilotResponse }]);
      setInput(""); // clear input box

    } catch (err: any) {
      console.error("[CopilotChat] Error processing message.", err);
      setError(err.message || "Failed to communicate with Copilot.");
    } finally {
      setLoading(false);
    }
  };

  return (
    <div style={{ padding: "16px", border: "1px solid #ccc", borderRadius: "8px", maxWidth: "600px", fontFamily: "Segoe UI, sans-serif" }}>
      <h3 style={{ marginTop: 0 }}>Copilot Search Chat</h3>

      {/* Chat Messages */}
      <div style={{ maxHeight: "300px", overflowY: "auto", marginBottom: "16px", padding: "8px", backgroundColor: "#f9f9f9", borderRadius: "4px" }}>
        {messages.length === 0 && <span style={{ color: "#666" }}>Send a message to start talking with Copilot...</span>}
        {messages.map((msg, idx) => (
          <div key={idx} style={{ marginBottom: "12px", textAlign: msg.author === "User" ? "right" : "left" }}>
            <div style={{
              display: "inline-block",
              padding: "8px 12px",
              borderRadius: "16px",
              backgroundColor: msg.author === "User" ? "#0078d4" : "#e1dfdd",
              color: msg.author === "User" ? "white" : "black"
            }}>
              <strong>{msg.author}:</strong><br />
              {msg.text}
            </div>
          </div>
        ))}
      </div>

      {/* Error Output */}
      {error && <div style={{ color: "red", marginBottom: "8px" }}>{error}</div>}

      {/* Input Area */}
      <div style={{ display: "flex", gap: "8px" }}>
        <input
          type="text"
          value={input}
          onChange={(e) => setInput(e.target.value)}
          onKeyDown={(e) => e.key === "Enter" && !loading && sendMessage()}
          placeholder="Ask Copilot a question..."
          disabled={loading}
          style={{ flexGrow: 1, padding: "8px", borderRadius: "4px", border: "1px solid #ccc" }}
        />
        <button
          onClick={sendMessage}
          disabled={loading || !input.trim()}
          style={{ padding: "8px 16px", backgroundColor: "#0078d4", color: "white", border: "none", borderRadius: "4px", cursor: loading ? "not-allowed" : "pointer" }}
        >
          {loading ? "Sending..." : "Send"}
        </button>
      </div>
    </div>
  );
};
