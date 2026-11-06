import { MailtrapClient } from "mailtrap";

const EMAIL_PATTERN = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
const MAX_BODY_BYTES = 8 * 1024 * 1024;
const MAX_ATTACHMENT_BYTES = 5 * 1024 * 1024;
const MAX_REQUESTS_PER_MINUTE = 5;
const rateLimitByIp = new Map();

const sendJson = (response, statusCode, payload) => {
  response.statusCode = statusCode;
  response.setHeader("Content-Type", "application/json; charset=utf-8");
  response.setHeader("Cache-Control", "no-store");
  response.end(JSON.stringify(payload));
};

const readJsonBody = (request) =>
  new Promise((resolve, reject) => {
    let body = "";
    let bodySize = 0;
    let bodyTooLarge = false;
    request.setEncoding("utf8");
    request.on("data", (chunk) => {
      bodySize += Buffer.byteLength(chunk);
      if (bodySize > MAX_BODY_BYTES) {
        if (!bodyTooLarge) reject(new Error("Request body is too large."));
        bodyTooLarge = true;
        return;
      }
      if (bodyTooLarge) return;
      body += chunk;
    });
    request.on("end", () => {
      try {
        resolve(JSON.parse(body));
      } catch {
        reject(new Error("Request body must be valid JSON."));
      }
    });
    request.on("error", reject);
  });

const isAllowedAddress = (address, envValue) => {
  const allowedAddresses = (envValue ?? "")
    .split(",")
    .map((value) => value.trim().toLowerCase())
    .filter(Boolean);
  return allowedAddresses.includes(address.toLowerCase());
};

const isRateLimited = (ipAddress) => {
  const now = Date.now();
  const requestTimes = (rateLimitByIp.get(ipAddress) ?? []).filter((time) => now - time < 60_000);
  if (requestTimes.length >= MAX_REQUESTS_PER_MINUTE) {
    rateLimitByIp.set(ipAddress, requestTimes);
    return true;
  }
  requestTimes.push(now);
  rateLimitByIp.set(ipAddress, requestTimes);
  return false;
};

export const mailtrapTestSendMiddleware = async (request, response, next) => {
  if (request.url?.split("?")[0] !== "/api/mailtrap/test-send") {
    next();
    return;
  }

  if (request.method !== "POST") {
    sendJson(response, 405, { error: "Method not allowed." });
    return;
  }

  if (!request.headers["content-type"]?.toLowerCase().startsWith("application/json")) {
    sendJson(response, 415, { error: "Content-Type must be application/json." });
    return;
  }

  const ipAddress = request.socket.remoteAddress ?? "unknown";
  if (isRateLimited(ipAddress)) {
    sendJson(response, 429, { error: "Test email rate limit reached. Try again in a minute." });
    return;
  }

  const apiToken = process.env.MAILTRAP_API_TOKEN?.trim();
  const allowedRecipients = process.env.MAILTRAP_ALLOWED_RECIPIENTS;
  const allowedSenders = process.env.MAILTRAP_ALLOWED_FROM_EMAILS;
  if (!apiToken || /^(replace-with|configure-on-server)/i.test(apiToken)) {
    sendJson(response, 503, {
      error: "Mailtrap API token is missing or still a placeholder. Set MAILTRAP_API_TOKEN in the project-root .env file and restart the server.",
    });
    return;
  }
  if (!allowedRecipients?.trim() || !allowedSenders?.trim()) {
    sendJson(response, 503, {
      error: "Mailtrap sender or recipient allowlist is missing. Set MAILTRAP_ALLOWED_FROM_EMAILS and MAILTRAP_ALLOWED_RECIPIENTS in the project-root .env file.",
    });
    return;
  }

  let payload;
  try {
    payload = await readJsonBody(request);
  } catch (error) {
    sendJson(response, 400, { error: error instanceof Error ? error.message : "Invalid request body." });
    return;
  }

  const senderEmail = typeof payload?.from?.email === "string" ? payload.from.email.trim() : "";
  const senderName = typeof payload?.from?.name === "string" ? payload.from.name.trim() : "";
  const recipientEmail = typeof payload?.to === "string" ? payload.to.trim() : "";
  const replyToEmail = typeof payload?.replyTo === "string" ? payload.replyTo.trim() : "";
  const subject = typeof payload?.subject === "string" ? payload.subject.trim() : "";
  const text = typeof payload?.text === "string" ? payload.text.trim() : "";
  const attachment = payload?.attachment;

  if (
    !EMAIL_PATTERN.test(senderEmail) ||
    !EMAIL_PATTERN.test(recipientEmail) ||
    (replyToEmail && !EMAIL_PATTERN.test(replyToEmail)) ||
    senderName.length > 200 ||
    !subject ||
    subject.length > 200 ||
    !text ||
    text.length > 20_000 ||
    (attachment !== undefined &&
      (!attachment ||
        typeof attachment.filename !== "string" ||
        !attachment.filename.trim() ||
        attachment.filename.length > 255 ||
        /[\u0000-\u001f\u007f\\/]/.test(attachment.filename) ||
        typeof attachment.content !== "string" ||
        !attachment.content ||
        !/^(?:[A-Za-z0-9+/]{4})*(?:[A-Za-z0-9+/]{2}==|[A-Za-z0-9+/]{3}=)?$/.test(attachment.content) ||
        Buffer.from(attachment.content, "base64").length > MAX_ATTACHMENT_BYTES ||
        (attachment.type !== undefined &&
          (typeof attachment.type !== "string" ||
            attachment.type.length > 150 ||
            /[\u0000-\u001f\u007f]/.test(attachment.type)))))
  ) {
    sendJson(response, 400, { error: "Provide valid sender/recipient details, a subject, message text, and an optional attachment up to 5 MB." });
    return;
  }

  if (!isAllowedAddress(senderEmail, allowedSenders)) {
    sendJson(response, 403, { error: "This sender address is not allowed for Mailtrap test sends." });
    return;
  }
  if (!isAllowedAddress(recipientEmail, allowedRecipients)) {
    sendJson(response, 403, { error: "This recipient is not allowed for Mailtrap test sends." });
    return;
  }

  try {
    const client = new MailtrapClient({ token: apiToken, bulk: true });
    const result = await client.send({
      from: { email: senderEmail, name: senderName || undefined },
      to: [{ email: recipientEmail }],
      ...(replyToEmail ? { reply_to: { email: replyToEmail } } : {}),
      subject,
      text,
      category: "MsgFlow test",
      ...(attachment
        ? {
            attachments: [{
              filename: attachment.filename,
              type: attachment.type || undefined,
              content: Buffer.from(attachment.content, "base64"),
            }],
          }
        : {}),
    });
    sendJson(response, 200, { success: true, messageIds: result.message_ids });
  } catch {
    sendJson(response, 502, { error: "Mailtrap could not accept this test email. Check the server configuration and Mailtrap sender/domain setup." });
  }
};
