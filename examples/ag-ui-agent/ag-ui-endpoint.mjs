/**
 * A minimal AG-UI 1.0 agent endpoint: it reads a RunAgentInput, asks a "responder" what the
 * agent says, and streams the answer as AG-UI events (JSON over server-sent events).
 *
 * A responder is `async (input, { signal }) => ({ text, toolCalls })`, where `toolCalls` is a
 * list of `{ id, name, args }` (args as an object). The tools an agent may call are the frontend
 * tools the client sent in `input.tools`: the client runs them, and answers in the messages of
 * its next run. So a run that calls a tool ends there, as AG-UI 1.0 describes for frontend tools.
 *
 * The endpoint has no sign-in. It only refuses requests that a page on another site could send
 * (see `refusal`). A real host must authenticate each caller and tie `threadId` to that caller's
 * session before it runs the agent.
 */
import { randomUUID } from "node:crypto";
import { isIP } from "node:net";

const maxBodyBytes = 2 * 1024 * 1024;

/**
 * Why a request may not start a run, as [status, message], or null if it may.
 *
 * Any web page open in the same browser can send a POST to this server, and a run may be a
 * billed model call. So the endpoint answers only a request that such a page cannot send:
 *  - Its Host is this machine (localhost or an IP address) or is listed in `allowedHosts`. A
 *    page that reaches the server under a DNS name of its own (DNS rebinding) is refused.
 *  - Its Origin, when a browser sends one, is the server's own or is listed in `allowedOrigins`.
 *  - Its body is application/json. A page on another site cannot send that without asking
 *    first (a CORS preflight), and this server never says yes.
 */
function refusal(request, { allowedOrigins = [], allowedHosts = [] }) {
  const host = (request.headers.host ?? "").toLowerCase();
  const hostname = host.replace(/:\d+$/, "").replace(/^\[(.*)\]$/, "$1");
  if (
    hostname !== "localhost" &&
    !isIP(hostname) &&
    !allowedHosts.includes(hostname)
  )
    return [403, "This server does not answer to that Host."];
  const { origin } = request.headers;
  const ownOrigin = `${request.socket?.encrypted ? "https" : "http"}://${host}`;
  if (
    origin !== undefined &&
    origin.toLowerCase() !== ownOrigin &&
    !allowedOrigins.includes(origin)
  )
    return [403, "Requests from that origin are not allowed."];
  if (!/^application\/json\b/i.test(request.headers["content-type"] ?? ""))
    return [415, "Send application/json."];
  return null;
}

async function readJson(request) {
  const chunks = [];
  let size = 0;
  for await (const chunk of request) {
    size += chunk.length;
    if (size > maxBodyBytes) throw new Error("The request is too large.");
    chunks.push(chunk);
  }
  return JSON.parse(Buffer.concat(chunks).toString("utf8"));
}

function isRunInput(input) {
  return (
    input &&
    typeof input === "object" &&
    typeof input.threadId === "string" &&
    typeof input.runId === "string" &&
    Array.isArray(input.messages)
  );
}

/**
 * Returns `async (request, response)` for a Node HTTP server.
 *
 * @param options.allowedOrigins  Origins besides the server's own that may call the endpoint,
 *   such as "https://app.example.com". Default: none.
 * @param options.allowedHosts    Host names besides localhost and IP addresses that the server
 *   is reached under, in lower case and without a port. Default: none.
 */
export function createAgUiEndpoint(respond, options = {}) {
  return async function handleRun(request, response) {
    const refused = refusal(request, options);
    if (refused) {
      response.writeHead(refused[0], { "Content-Type": "application/json" });
      response.end(JSON.stringify({ error: refused[1] }));
      return;
    }
    let input;
    try {
      input = await readJson(request);
      if (!isRunInput(input)) throw new Error("Not an AG-UI run input.");
    } catch (error) {
      response.writeHead(400, { "Content-Type": "application/json" });
      response.end(JSON.stringify({ error: error.message }));
      return;
    }
    response.writeHead(200, {
      "Content-Type": "text/event-stream",
      "Cache-Control": "no-store",
      Connection: "keep-alive",
    });
    const send = (event) =>
      response.write(`data: ${JSON.stringify(event)}\n\n`);
    const abort = new AbortController();
    response.on("close", () => abort.abort());
    const run = { threadId: input.threadId, runId: input.runId };
    send({ type: "RUN_STARTED", ...run, protocolVersion: "1.0" });
    try {
      const { text = "", toolCalls = [] } = await respond(
        { ...input, tools: input.tools ?? [], context: input.context ?? [] },
        { signal: abort.signal },
      );
      const messageId = randomUUID();
      if (text) {
        send({ type: "TEXT_MESSAGE_START", messageId, role: "assistant" });
        send({ type: "TEXT_MESSAGE_CONTENT", messageId, delta: text });
        send({ type: "TEXT_MESSAGE_END", messageId });
      }
      for (const call of toolCalls) {
        send({
          type: "TOOL_CALL_START",
          toolCallId: call.id,
          toolCallName: call.name,
          parentMessageId: messageId,
        });
        send({
          type: "TOOL_CALL_ARGS",
          toolCallId: call.id,
          delta: JSON.stringify(call.args),
        });
        send({ type: "TOOL_CALL_END", toolCallId: call.id });
      }
      send({ type: "RUN_FINISHED", ...run });
    } catch (error) {
      // Say what failed without passing internal details to the browser.
      console.error("Agent run failed:", error);
      send({
        type: "RUN_ERROR",
        message: "The agent could not complete this request.",
        code: "agent_error",
      });
    }
    response.end();
  };
}
