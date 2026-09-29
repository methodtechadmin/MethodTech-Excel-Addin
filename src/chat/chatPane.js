import { applyExcelActions } from "./excelActions";
import { readSheetContext } from "./sheetContext";

function appendMessage(role, text) {
  const list = document.getElementById("copilot-messages");
  if (!list) {
    return;
  }
  const item = document.createElement("div");
  item.className = `methodtech-chat-msg methodtech-chat-msg--${role}`;
  item.textContent = text;
  list.appendChild(item);
  list.scrollTop = list.scrollHeight;
}

function setChatStatus(message, isError = false) {
  const el = document.getElementById("copilot-status");
  if (!el) {
    return;
  }
  el.textContent = message || "";
  el.style.color = isError ? "#a4262c" : "#605e5c";
}

export async function sendChatMessage(message) {
  const text = String(message || "").trim();
  if (!text) {
    return;
  }
  appendMessage("user", text);
  setChatStatus("Reading the selection…");

  let sheet;
  try {
    sheet = await readSheetContext();
  } catch (error) {
    setChatStatus((error && error.message) || "Could not read the sheet.", true);
    return;
  }

  const summary = sheet.summary || "No sheet context.";
  appendMessage(
    "assistant",
    `I can see ${summary}. The answer from MethodTech is connected in the next step.`
  );
  setChatStatus("Sheet context captured. Django connection comes next.");
  return { message: text, context: sheet };
}

export function initChatPane() {
  const form = document.getElementById("copilot-form");
  const input = document.getElementById("copilot-input");
  if (!form || !input) {
    return;
  }

  form.addEventListener("submit", (event) => {
    event.preventDefault();
    const message = input.value;
    input.value = "";
    sendChatMessage(message).catch((error) => {
      setChatStatus((error && error.message) || "Chat failed.", true);
    });
  });
}

export { applyExcelActions };
