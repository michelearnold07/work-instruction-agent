import { useState, useRef, useEffect } from "react";
import * as XLSX from "xlsx";
import mammoth from "mammoth";

const SYSTEM_PROMPT = `You are a Work Instruction Assistant for an electrical manufacturing planning department.
You have been given the full text of multiple work instruction documents uploaded by the user.

Your job:
- Answer questions about how to build, assemble, or complete manufacturing tasks
- Reference the specific document name when you use it as a source
- If the answer spans multiple documents, cite each one
- If the answer is not in any document, say so clearly — do not guess or hallucinate steps
- Be concise and practical — your users are planners and supervisors on the floor
- Use numbered steps when describing assembly or build procedures
- Flag any safety-related steps prominently

You will receive the document contents in the user's first message.`;

function FileIcon({ type }) {
  const isExcel = type === "excel";
  return (
    <div style={{
      width: 32, height: 32, borderRadius: 6,
      background: isExcel ? "#1a6b3a" : "#1a3a6b",
      display: "flex", alignItems: "center", justifyContent: "center",
      fontSize: 10, fontWeight: 800, color: "#fff", letterSpacing: "0.04em",
      flexShrink: 0,
    }}>
      {isExcel ? "XLS" : "DOC"}
    </div>
  );
}

async function extractTextFromFile(file) {
  const name = file.name.toLowerCase();
  if (name.endsWith(".xlsx") || name.endsWith(".xls") || name.endsWith(".csv")) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const data = new Uint8Array(e.target.result);
          const wb = XLSX.read(data, { type: "array" });
          let text = "";
          wb.SheetNames.forEach(sheetName => {
            const ws = wb.Sheets[sheetName];
            const csv = XLSX.utils.sheet_to_csv(ws);
            if (csv.trim()) text += `[Sheet: ${sheetName}]\n${csv}\n\n`;
          });
          resolve({ name: file.name, type: "excel", text: text.trim() });
        } catch (e) { reject(e); }
      };
      reader.onerror = reject;
      reader.readAsArrayBuffer(file);
    });
  }
  if (name.endsWith(".docx")) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = async (e) => {
        try {
          const result = await mammoth.extractRawText({ arrayBuffer: e.target.result });
          resolve({ name: file.name, type: "word", text: result.value.trim() });
        } catch (e) { reject(e); }
      };
      reader.onerror = reject;
      reader.readAsArrayBuffer(file);
    });
  }
  if (name.endsWith(".txt")) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => resolve({ name: file.name, type: "word", text: e.target.result.trim() });
      reader.onerror = reject;
      reader.readAsText(file);
    });
  }
  throw new Error(`Unsupported file type: ${file.name}`);
}

function buildDocContext(docs) {
  return docs.map((d, i) => `=== DOCUMENT ${i + 1}: ${d.name} ===\n${d.text}\n`).join("\n");
}

export default function App() {
  const [apiKey, setApiKey] = useState(() => localStorage.getItem("wi_api_key") || "");
  const [apiKeySaved, setApiKeySaved] = useState(!!localStorage.getItem("wi_api_key"));
  const [docs, setDocs] = useState([]);
  const [processing, setProcessing] = useState(false);
  const [processErrors, setProcessErrors] = useState([]);
  const [messages, setMessages] = useState([]);
  const [input, setInput] = useState("");
  const [thinking, setThinking] = useState(false);
  const [phase, setPhase] = useState("upload");
  const [dragging, setDragging] = useState(false);
  const [apiError, setApiError] = useState("");
  const fileRef = useRef(null);
  const bottomRef = useRef(null);
  const inputRef = useRef(null);
  const historyRef = useRef([]);

  useEffect(() => {
    bottomRef.current?.scrollIntoView({ behavior: "smooth" });
  }, [messages, thinking]);

  const saveApiKey = () => {
    const trimmed = apiKey.trim();
    if (!trimmed.startsWith("sk-ant-")) {
      setApiError("That doesn't look right — it should start with sk-ant-");
      return;
    }
    localStorage.setItem("wi_api_key", trimmed);
    setApiKeySaved(true);
    setApiError("");
  };

  const clearApiKey = () => {
    localStorage.removeItem("wi_api_key");
    setApiKey("");
    setApiKeySaved(false);
  };

  const handleFiles = async (files) => {
    const arr = Array.from(files).filter(f => f.name.match(/\.(xlsx|xls|csv|docx|txt)$/i));
    if (!arr.length) return;
    setProcessing(true);
    setProcessErrors([]);
    const results = [], errors = [];
    for (const file of arr) {
      try { results.push(await extractTextFromFile(file)); }
      catch (e) { errors.push(file.name); }
    }
    setDocs(prev => {
      const names = new Set(prev.map(d => d.name));
      return [...prev, ...results.filter(r => !names.has(r.name))];
    });
    setProcessErrors(errors);
    setProcessing(false);
  };

  const removeDoc = (name) => setDocs(d => d.filter(x => x.name !== name));

  const startSession = () => {
    historyRef.current = [];
    setMessages([{
      role: "assistant",
      text: `I've loaded ${docs.length} work instruction${docs.length !== 1 ? "s" : ""}:\n${docs.map(d => `• ${d.name}`).join("\n")}\n\nAsk me anything about how to build or assemble — I'll pull the answer straight from your documents.`,
    }]);
    setPhase("chat");
    setTimeout(() => inputRef.current?.focus(), 100);
  };

  const sendMessage = async () => {
    const q = input.trim();
    if (!q || thinking) return;
    setInput("");

    setMessages(m => [...m, { role: "user", text: q }]);
    setThinking(true);

    const docContext = buildDocContext(docs);
    const history = historyRef.current;

    const apiMessages = history.length === 0
      ? [{ role: "user", content: `Here are the work instruction documents for this session:\n\n${docContext}\n\n---\n\nMy first question: ${q}` }]
      : [...history, { role: "user", content: q }];

    try {
      // Call our Vercel serverless function instead of Anthropic directly
      const res = await fetch("/api/chat", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "x-api-key": localStorage.getItem("wi_api_key") || "",
        },
        body: JSON.stringify({ messages: apiMessages, system: SYSTEM_PROMPT }),
      });
      const data = await res.json();
      if (data.error) throw new Error(data.error.message || data.error);
      const reply = data.content?.map(b => b.text || "").join("").trim();
      historyRef.current = [...apiMessages, { role: "assistant", content: reply }];
      setMessages(m => [...m, { role: "assistant", text: reply }]);
    } catch (e) {
      setMessages(m => [...m, { role: "assistant", text: `Error: ${e.message}. Try clicking "Change API Key" and re-entering your key.` }]);
    }
    setThinking(false);
  };

  const reset = () => { setPhase("upload"); setMessages([]); setDocs([]); historyRef.current = []; };
  const totalWords = docs.reduce((sum, d) => sum + d.text.split(/\s+/).length, 0);

  // ── API KEY SCREEN ──
  if (!apiKeySaved) return (
    <div style={{ minHeight: "100vh", background: "#f7f5f0", fontFamily: "'IBM Plex Sans', 'Helvetica Neue', sans-serif", display: "flex", alignItems: "center", justifyContent: "center", padding: 24 }}>
      <div style={{ background: "#fff", border: "1px solid #e5e2da", borderRadius: 14, padding: "36px 32px", maxWidth: 440, width: "100%" }}>
        <div style={{ marginBottom: 24 }}>
          <div style={{ display: "inline-block", background: "#1c2b1e", borderRadius: 8, padding: "6px 12px", marginBottom: 16 }}>
            <span style={{ color: "#4ade80", fontSize: 10, fontWeight: 700, letterSpacing: "0.14em" }}>WORK INSTRUCTION AGENT</span>
          </div>
          <h1 style={{ fontSize: 20, fontWeight: 700, color: "#1c2b1e", margin: "0 0 8px" }}>Enter your API key</h1>
          <p style={{ fontSize: 13, color: "#888", lineHeight: 1.6, margin: 0 }}>
            Get your key from <a href="https://console.anthropic.com/keys" target="_blank" rel="noreferrer" style={{ color: "#2d6a35" }}>console.anthropic.com/keys</a>. It's stored only in your browser.
          </p>
        </div>
        <div style={{ marginBottom: 12 }}>
          <label style={{ fontSize: 11, fontWeight: 700, color: "#888", letterSpacing: "0.08em", display: "block", marginBottom: 6 }}>ANTHROPIC API KEY</label>
          <input type="password" value={apiKey} onChange={e => { setApiKey(e.target.value); setApiError(""); }}
            onKeyDown={e => e.key === "Enter" && saveApiKey()}
            placeholder="sk-ant-api03-..."
            style={{ width: "100%", padding: "11px 14px", border: "1px solid #d5d2ca", borderRadius: 8, fontSize: 13, fontFamily: "monospace", color: "#1c2b1e", background: "#f7f5f0", outline: "none", boxSizing: "border-box" }}
          />
          {apiError && <p style={{ color: "#b91c1c", fontSize: 12, margin: "6px 0 0" }}>{apiError}</p>}
        </div>
        <button onClick={saveApiKey} style={{ width: "100%", padding: "13px", background: "#1c2b1e", color: "#4ade80", border: "none", borderRadius: 8, fontSize: 13, fontWeight: 700, letterSpacing: "0.06em", cursor: "pointer", fontFamily: "inherit" }}>
          SAVE KEY AND CONTINUE →
        </button>
      </div>
    </div>
  );

  // ── UPLOAD SCREEN ──
  if (phase === "upload") return (
    <div style={{ minHeight: "100vh", background: "#f7f5f0", fontFamily: "'IBM Plex Sans', 'Helvetica Neue', sans-serif", display: "flex", flexDirection: "column" }}>
      <div style={{ background: "#1c2b1e", padding: "22px 28px 18px", borderBottom: "3px solid #2d6a35" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <div style={{ width: 8, height: 8, borderRadius: "50%", background: "#4ade80", boxShadow: "0 0 6px #4ade80" }} />
          <span style={{ color: "#4ade80", fontSize: 10, fontWeight: 700, letterSpacing: "0.16em" }}>PLANNING DEPT · WORK INSTRUCTION AGENT</span>
          <button onClick={clearApiKey} style={{ marginLeft: "auto", background: "none", border: "1px solid #2d6a35", borderRadius: 6, color: "#6b8f72", fontSize: 10, fontWeight: 700, padding: "4px 10px", cursor: "pointer", fontFamily: "inherit" }}>CHANGE API KEY</button>
        </div>
        <h1 style={{ margin: "8px 0 2px", color: "#fff", fontSize: 22, fontWeight: 700 }}>Load Your Work Instructions</h1>
        <p style={{ margin: 0, color: "#6b8f72", fontSize: 13 }}>Upload your Excel and Word files — then ask anything about how to build or assemble</p>
      </div>
      <div style={{ flex: 1, padding: "28px", maxWidth: 700, width: "100%", margin: "0 auto", boxSizing: "border-box" }}>
        <div onDragOver={e => { e.preventDefault(); setDragging(true); }} onDragLeave={() => setDragging(false)}
          onDrop={e => { e.preventDefault(); setDragging(false); handleFiles(e.dataTransfer.files); }}
          onClick={() => fileRef.current?.click()}
          style={{ border: `2px dashed ${dragging ? "#2d6a35" : "#ccc8be"}`, borderRadius: 12, padding: "40px 24px", textAlign: "center", cursor: "pointer", background: dragging ? "#edf7ee" : "#fff", transition: "all 0.2s", marginBottom: 20 }}>
          <input ref={fileRef} type="file" multiple accept=".xlsx,.xls,.csv,.docx,.txt" style={{ display: "none" }} onChange={e => handleFiles(e.target.files)} />
          <div style={{ fontSize: 36, marginBottom: 12 }}>📂</div>
          <p style={{ margin: "0 0 6px", fontSize: 15, fontWeight: 600, color: "#1c2b1e" }}>{processing ? "Processing..." : "Drop your folder here or click to browse"}</p>
          <p style={{ margin: 0, fontSize: 12, color: "#999" }}>Supports .xlsx · .xls · .csv · .docx · .txt</p>
        </div>
        {processErrors.length > 0 && <div style={{ background: "#fff5f5", border: "1px solid #fca5a5", borderRadius: 8, padding: "12px 16px", marginBottom: 16, fontSize: 12, color: "#7f1d1d" }}>Could not read: {processErrors.join(", ")}</div>}
        {docs.length > 0 && (
          <div style={{ marginBottom: 24 }}>
            <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 10 }}>
              <span style={{ fontSize: 11, fontWeight: 700, color: "#888", letterSpacing: "0.08em" }}>{docs.length} DOCUMENT{docs.length !== 1 ? "S" : ""} · ~{totalWords.toLocaleString()} WORDS</span>
              <button onClick={() => setDocs([])} style={{ background: "none", border: "none", cursor: "pointer", fontSize: 11, color: "#aaa", fontFamily: "inherit" }}>clear all</button>
            </div>
            {docs.map(doc => (
              <div key={doc.name} style={{ background: "#fff", border: "1px solid #e5e2da", borderRadius: 8, padding: "10px 14px", display: "flex", alignItems: "center", gap: 12, marginBottom: 8 }}>
                <FileIcon type={doc.type} />
                <div style={{ flex: 1, minWidth: 0 }}>
                  <p style={{ margin: 0, fontSize: 13, fontWeight: 600, color: "#1c2b1e", whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{doc.name}</p>
                  <p style={{ margin: 0, fontSize: 11, color: "#aaa" }}>~{doc.text.split(/\s+/).length.toLocaleString()} words</p>
                </div>
                <button onClick={() => removeDoc(doc.name)} style={{ background: "none", border: "none", cursor: "pointer", color: "#ccc", fontSize: 16, flexShrink: 0 }}>×</button>
              </div>
            ))}
          </div>
        )}
        <button onClick={startSession} disabled={docs.length === 0 || processing}
          style={{ width: "100%", padding: "15px", background: docs.length > 0 ? "#1c2b1e" : "#e0ddd6", color: docs.length > 0 ? "#4ade80" : "#aaa", border: "none", borderRadius: 10, fontSize: 14, fontWeight: 700, letterSpacing: "0.06em", cursor: docs.length > 0 ? "pointer" : "not-allowed", fontFamily: "inherit" }}>
          {docs.length === 0 ? "UPLOAD DOCUMENTS TO CONTINUE" : `START SESSION WITH ${docs.length} DOCUMENT${docs.length !== 1 ? "S" : ""} →`}
        </button>
      </div>
    </div>
  );

  // ── CHAT SCREEN ──
  return (
    <div style={{ minHeight: "100vh", background: "#f7f5f0", fontFamily: "'IBM Plex Sans', 'Helvetica Neue', sans-serif", display: "flex", flexDirection: "column" }}>
      <div style={{ background: "#1c2b1e", padding: "14px 20px", borderBottom: "2px solid #2d6a35", display: "flex", alignItems: "center", gap: 14, flexShrink: 0 }}>
        <div>
          <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
            <div style={{ width: 7, height: 7, borderRadius: "50%", background: "#4ade80", boxShadow: "0 0 5px #4ade80" }} />
            <span style={{ color: "#4ade80", fontSize: 10, fontWeight: 700, letterSpacing: "0.14em" }}>LIVE SESSION</span>
          </div>
          <p style={{ margin: "2px 0 0", color: "#6b8f72", fontSize: 12 }}>{docs.length} doc{docs.length !== 1 ? "s" : ""} loaded</p>
        </div>
        <button onClick={clearApiKey} style={{ marginLeft: "auto", background: "none", border: "1px solid #2d6a35", borderRadius: 6, color: "#6b8f72", fontSize: 11, fontWeight: 700, padding: "5px 12px", cursor: "pointer", fontFamily: "inherit", marginRight: 8 }}>KEY</button>
        <button onClick={reset} style={{ background: "none", border: "1px solid #2d6a35", borderRadius: 6, color: "#6b8f72", fontSize: 11, fontWeight: 700, padding: "5px 12px", cursor: "pointer", fontFamily: "inherit" }}>← NEW SESSION</button>
      </div>
      <div style={{ flex: 1, overflowY: "auto", padding: "20px 16px" }}>
        <div style={{ maxWidth: 680, margin: "0 auto", display: "flex", flexDirection: "column", gap: 16 }}>
          {messages.map((msg, i) => {
            const isUser = msg.role === "user";
            return (
              <div key={i} style={{ display: "flex", flexDirection: isUser ? "row-reverse" : "row", gap: 10, alignItems: "flex-start" }}>
                <div style={{ width: 32, height: 32, borderRadius: "50%", flexShrink: 0, background: isUser ? "#e5f0e7" : "#1c2b1e", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 13 }}>{isUser ? "👤" : "🏭"}</div>
                <div style={{ maxWidth: "80%", background: isUser ? "#1c2b1e" : "#fff", color: isUser ? "#e5f0e7" : "#1c2b1e", border: isUser ? "none" : "1px solid #e5e2da", borderRadius: isUser ? "14px 4px 14px 14px" : "4px 14px 14px 14px", padding: "12px 16px", fontSize: 14, lineHeight: 1.7, whiteSpace: "pre-wrap" }}>{msg.text}</div>
              </div>
            );
          })}
          {thinking && (
            <div style={{ display: "flex", gap: 10, alignItems: "flex-start" }}>
              <div style={{ width: 32, height: 32, borderRadius: "50%", background: "#1c2b1e", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 13 }}>🏭</div>
              <div style={{ background: "#fff", border: "1px solid #e5e2da", borderRadius: "4px 14px 14px 14px", padding: "14px 18px", display: "flex", gap: 5, alignItems: "center" }}>
                {[0,1,2].map(i => <div key={i} style={{ width: 7, height: 7, borderRadius: "50%", background: "#4ade80", animation: "bounce 1.2s ease-in-out infinite", animationDelay: `${i*0.2}s` }} />)}
                <style>{`@keyframes bounce{0%,80%,100%{transform:scale(0.7);opacity:0.4}40%{transform:scale(1);opacity:1}}`}</style>
              </div>
            </div>
          )}
          <div ref={bottomRef} />
        </div>
      </div>
      <div style={{ borderTop: "1px solid #e5e2da", padding: "14px 16px", background: "#fff", flexShrink: 0 }}>
        <div style={{ maxWidth: 680, margin: "0 auto", display: "flex", gap: 10 }}>
          <input ref={inputRef} value={input} onChange={e => setInput(e.target.value)} onKeyDown={e => e.key === "Enter" && !e.shiftKey && sendMessage()}
            placeholder="Ask about assembly steps, build procedures, materials..."
            style={{ flex: 1, padding: "12px 16px", border: "1px solid #d5d2ca", borderRadius: 10, fontSize: 14, fontFamily: "inherit", color: "#1c2b1e", background: "#f7f5f0", outline: "none" }}
          />
          <button onClick={sendMessage} disabled={!input.trim() || thinking}
            style={{ background: input.trim() && !thinking ? "#1c2b1e" : "#e5e2da", color: input.trim() && !thinking ? "#4ade80" : "#aaa", border: "none", borderRadius: 10, padding: "12px 20px", cursor: input.trim() ? "pointer" : "not-allowed", fontSize: 18, transition: "all 0.2s", flexShrink: 0 }}>→</button>
        </div>
        <p style={{ textAlign: "center", margin: "8px 0 0", fontSize: 11, color: "#bbb" }}>Answers sourced from your uploaded documents only</p>
      </div>
    </div>
  );
}
