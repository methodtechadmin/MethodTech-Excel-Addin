import { djangoFetch } from "../auth/http";
import { ensureSignedIn } from "../auth/signIn";
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

  setChatStatus("Sending to MethodTech…");
  await ensureSignedIn({ force: false, interactive: true });
  const data = await djangoFetch("/api/microsoft/excel/chat/", {
    method: "POST",
    body: JSON.stringify({
      message: text,
      context: sheet,
    }),
  });

  const reply = (data && data.reply) || "No reply returned.";
  appendMessage("assistant", reply);

  const actions = data && data.excel_actions;
  if (actions && actions.length) {
    const done = await applyExcelActions(actions);
    setChatStatus(done.length ? `Updated the sheet: ${done.join(", ")}.` : "");
  } else {
    setChatStatus("");
  }
  return data;
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
    const button = form.querySelector("button");
    input.value = "";
    if (button) {
      button.disabled = true;
    }
    sendChatMessage(message)
      .catch((error) => {
        appendMessage("assistant", (error && error.message) || "Chat failed.");
        setChatStatus((error && error.message) || "Chat failed.", true);
      })
      .finally(() => {
        if (button) {
          button.disabled = false;
        }
      });
  });
}

export { applyExcelActions };
