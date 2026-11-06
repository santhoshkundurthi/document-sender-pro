import { useEffect, useMemo, useRef, useState, type ChangeEvent, type ReactNode } from "react";
import * as XLSX from "xlsx";

type Role = "admin" | "user";
type Channel = "email" | "sms" | "whatsapp";
type WorkflowTab = "upload" | "csv-preview" | "template" | "provider" | "preview" | "dashboard";
type CampaignStatus = "sent" | "delivered" | "opened" | "clicked" | "bounced" | "failed" | "unsubscribed";
type ContactStatus = "active" | "unsubscribed" | "invalid";

type Contact = {
  id: string;
  name: string;
  email: string;
  phone: string;
  company: string;
  segment: string;
  status: ContactStatus;
  custom: Record<string, string>;
};

type TemplateItem = {
  id: string;
  name: string;
  subject: string;
  body: string;
};

type CampaignResult = {
  id: string;
  name: string;
  recipient: string;
  channel: Channel;
  status: CampaignStatus;
  reason: string;
};

type MappedFields = {
  name: string;
  email: string;
  phone: string;
  company: string;
  segment: string;
};

type IconProps = {
  className?: string;
  children?: ReactNode;
};

const MAX_ROWS_PER_UPLOAD = 10000;
const workflowTabs: { id: WorkflowTab; label: string }[] = [
  { id: "upload", label: "1. Upload Contacts" },
  { id: "csv-preview", label: "2. Verify Data" },
  { id: "template", label: "3. Craft Message" },
  { id: "provider", label: "4. Choose Channel" },
  { id: "preview", label: "5. Preview & Test" },
  { id: "dashboard", label: "6. Launch & Track" },
];
const flashMessages = [
  "Connect. Send. Grow.",
  "One app for all your messaging gateways.",
  "Fast, reliable SMS and email delivery made simple.",
  "Power your business communication with multi-gateway messaging.",
  "Reach customers instantly, across every channel.",
];
const footerFlashMessages = [
  "Send personalized messages to thousands of customers in just a few clicks.",
  "Upload your customer list and start your campaign in minutes.",
  "Send individual or bulk SMS, Email, and WhatsApp messages effortlessly.",
  "Save time by automating repetitive customer communications.",
  "Personalize every message using customer-specific information.",
  "Manage multiple communication channels from one centralized dashboard.",
  "Create professional email campaigns and reach your audience at scale.",
  "Connect your preferred email gateway and start sending instantly.",
  "Use templates to create consistent and professional customer communications.",
  "Send documents, images, and personalized content with your emails.",
  "Reach customers instantly with fast and reliable SMS communication.",
  "Perfect for alerts, reminders, OTPs, promotions, and notifications.",
  "Send personalized SMS campaigns directly from your customer list.",
  "Choose the SMS gateway that fits your business requirements.",
  "Connect with customers through one of the world's most popular messaging channels.",
  "Send personalized WhatsApp campaigns to your customer list.",
  "Share offers, notifications, documents, images, and updates with ease.",
  "Bring your customer conversations closer with WhatsApp messaging.",
  "🔄 Smart Failover Active: If one gateway drops, we auto-reroute to keep delivery at 99.9%.",
  "📊 Unified Analytics: Track open rates, bounces, and costs across all providers in one dashboard.",
  "📝 Dynamic Templates: Save time with pre-built, variable-ready SMS & Email templates.",
  "⚖️ Cost Optimizer: We automatically route messages through the most cost-effective gateway.",
  "🔒 Compliance Built-in: Auto-handles opt-outs, unsubscribes, and GDPR/TCPA rules.",
  "⚡ Welcome! Connect your first SMS or Email gateway in under 60 seconds.",
  "🔗 You’re all set! Your communication hub is ready. Add a gateway to start sending.",
  "🌍 Go Global: Connect multiple gateways to reach your audience anywhere, anytime.",
  "Never Miss an OTP — Deliver verification codes instantly with reliable SMS gateway integration.",
  "Urgent Alerts, Instant Reach — Send emergency and service alerts directly to your customers’ screens.",
  "Boost Sales With Flash Offers — Launch time-limited deals and promotions that customers see immediately.",
  "Reminders That Convert — Send appointment, payment, and event reminders automatically.",
  "Engage Customers Instantly — Share offers, updates, and announcements across SMS and email.",
  "Grow With Bulk Messaging — Run large-scale marketing campaigns with speed, control, and detailed reports.",
  "Better Customer Experience — Keep customers informed with timely transaction and delivery updates.",
  "Perfect for Every Business — Ideal for e-commerce, banking, education, healthcare, logistics, and service businesses.",
  "Multiple Gateways, One App — Connect SMS, email, and WhatsApp gateways in minutes. Start sending smarter today.",
  "Send Without Limits — Choose from leading SMS and email gateways for fast, reliable delivery worldwide.",
  "One Dashboard, Every Channel — Manage SMS, email, and campaigns from a single powerful platform.",
  "Reliable Delivery, Every Time — Smart gateway routing helps your messages reach customers when it matters.",
  "Integrate in Minutes — Connect your favorite SMS or email gateway with simple setup and powerful APIs.",
  "Scale With Confidence — From 10 messages to 10 million, send bulk campaigns without delays.",
  "Your Messages, Your Way — Personalize SMS and email campaigns for higher engagement and better results.",
  "Track Every Message — Monitor delivery, opens, clicks, and campaign performance in real time.",
  "Automate Customer Communication — Send OTPs, alerts, reminders, and updates automatically.",
  "Choose the Right Gateway — Flexible gateway options designed to match your business needs and budget.",
];

const Icon = ({ className, children }: IconProps) => (
  <svg
    viewBox="0 0 24 24"
    fill="none"
    stroke="currentColor"
    strokeWidth="1.8"
    strokeLinecap="round"
    strokeLinejoin="round"
    className={className}
    aria-hidden="true"
  >
    {children}
  </svg>
);

const UploadCloudIcon = (props: IconProps) => (
  <Icon {...props}>
    <path d="M7 18a4 4 0 0 1-.85-7.9A6 6 0 0 1 18.1 8.5 4.5 4.5 0 1 1 18 18H7Z" />
    <path d="M12 9v8" />
    <path d="m9.5 11.5 2.5-2.5 2.5 2.5" />
  </Icon>
);

const SendIcon = (props: IconProps) => (
  <Icon {...props}>
    <path d="M22 2 11 13" />
    <path d="M22 2 15 22l-4-9-9-4 20-7Z" />
  </Icon>
);

const FileTextIcon = (props: IconProps) => (
  <Icon {...props}>
    <path d="M14 2H7a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V7Z" />
    <path d="M14 2v6h6" />
    <path d="M9 13h6" />
    <path d="M9 17h6" />
  </Icon>
);

const CheckCircleIcon = (props: IconProps) => (
  <Icon {...props}>
    <circle cx="12" cy="12" r="9" />
    <path d="m9 12 2 2 4-5" />
  </Icon>
);

const XCircleIcon = (props: IconProps) => (
  <Icon {...props}>
    <circle cx="12" cy="12" r="9" />
    <path d="m15 9-6 6" />
    <path d="m9 9 6 6" />
  </Icon>
);

const ChartIcon = (props: IconProps) => (
  <Icon {...props}>
    <path d="M4 19V5" />
    <path d="M4 19h16" />
    <path d="M8 15l3-4 3 2 5-7" />
  </Icon>
);

const ClockIcon = (props: IconProps) => (
  <Icon {...props}>
    <circle cx="12" cy="12" r="9" />
    <path d="M12 7v5l3 2" />
  </Icon>
);

const templateLibrary: TemplateItem[] = [
  {
    id: "welcome-email",
    name: "Welcome email",
    subject: "Welcome to NorthStar, {{name}}",
    body: "Hi {{name}},\n\nThanks for joining {{company}}. We are excited to share updates with you.\n\nYour plan is ready and we will keep you informed on every milestone.\n\nRegards,\n{{sender_name}}"
  },
  {
    id: "invoice-reminder",
    name: "Invoice reminder",
    subject: "Payment reminder for {{company}}",
    body: "Hello {{name}},\n\nThis is a reminder that your invoice is due soon. Please review the details and confirm.\n\nCustomer segment: {{segment}}\n\nRegards,\n{{sender_name}}"
  },
  {
    id: "order-confirmation",
    name: "Order confirmation",
    subject: "Your order has been confirmed",
    body: "Hello {{name}},\n\nYour purchase with {{company}} has been confirmed and is now in processing.\n\nWe will keep you updated with a delivery status.\n\nBest,\n{{sender_name}}"
  },
];

const sampleContacts: Contact[] = [
  {
    id: "1",
    name: "Aisha Patel",
    email: "aisha@example.com",
    phone: "+14155550111",
    company: "NorthStar Labs",
    segment: "Enterprise",
    status: "active",
    custom: { plan: "Pro", city: "Austin" },
  },
  {
    id: "2",
    name: "Marcus Chen",
    email: "marcus@example.com",
    phone: "+14155550112",
    company: "Summit CRM",
    segment: "SMB",
    status: "active",
    custom: { plan: "Starter", city: "Seattle" },
  },
  {
    id: "3",
    name: "Priya Nair",
    email: "priya@example.com",
    phone: "+14155550113",
    company: "BluePeak",
    segment: "Enterprise",
    status: "unsubscribed",
    custom: { plan: "Scale", city: "Boston" },
  },
  {
    id: "4",
    name: "Daniel Gomez",
    email: "daniel@invalid",
    phone: "+14155550114",
    company: "Luma Works",
    segment: "Growth",
    status: "invalid",
    custom: { plan: "Trial", city: "Denver" },
  },
];

const normalizeHeader = (value: string) => value.trim().toLowerCase().replace(/[^a-z0-9]+/g, " ").trim();
const isValidEmail = (value: string) => /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(value.trim());
const isValidIndianMobile = (value: string) => {
  const phone = value.trim();
  if (!/^\+?[\d\s()-]+$/.test(phone)) return false;
  let digits = phone.replace(/\D/g, "");
  if (digits.length === 12 && digits.startsWith("91")) digits = digits.slice(2);
  if (digits.length === 11 && digits.startsWith("0")) digits = digits.slice(1);
  return /^[6-9]\d{9}$/.test(digits);
};
const isContactValidForChannel = (contact: Contact, channel: Channel) => {
  if (contact.status === "unsubscribed") return false;
  if (channel === "email") return isValidEmail(contact.email);
  return isValidIndianMobile(contact.phone);
};
const formatFileSize = (bytes: number) => {
  if (!bytes) return "0 B";
  const units = ["B", "KB", "MB", "GB"];
  const index = Math.min(Math.floor(Math.log(bytes) / Math.log(1024)), units.length - 1);
  const value = bytes / 1024 ** index;
  return `${value.toFixed(value >= 10 || index === 0 ? 0 : 1)} ${units[index]}`;
};

interface ImportMetaEnv {
  [key: string]: string | boolean | undefined;
}

declare global {
  interface ImportMeta {
    env?: ImportMetaEnv;
  }
}

const getEnvValue = (name: string, fallback = "") => {
  const clientConfig: Record<string, string | boolean | undefined> =
    typeof import.meta === "undefined"
      ? {}
      : {
          VITE_EMAIL_PROVIDER_PREFERRED: import.meta.env?.VITE_EMAIL_PROVIDER_PREFERRED,
          VITE_SMS_PROVIDER_PREFERRED: import.meta.env?.VITE_SMS_PROVIDER_PREFERRED,
          VITE_WHATSAPP_PROVIDER_PREFERRED: import.meta.env?.VITE_WHATSAPP_PROVIDER_PREFERRED,
          VITE_SENDGRID_FROM_NAME: import.meta.env?.VITE_SENDGRID_FROM_NAME,
          VITE_SENDGRID_FROM_EMAIL: import.meta.env?.VITE_SENDGRID_FROM_EMAIL,
          VITE_SENDGRID_REPLY_TO: import.meta.env?.VITE_SENDGRID_REPLY_TO,
        };
  const runtimeValue = clientConfig[name];
  return runtimeValue && String(runtimeValue).trim() ? String(runtimeValue).trim() : fallback;
};

const providerOptions: Record<Channel, string[]> = {
  email: ["mailtrap", "sendgrid", "brevo", "smtp"],
  sms: ["twilio", "msg91"],
  whatsapp: ["gupshup", "wati"],
};

const encodeAttachment = async (file: File) => {
  if (file.size > 5 * 1024 * 1024) {
    throw new Error("Mailtrap test attachments must be 5 MB or smaller.");
  }
  const bytes = new Uint8Array(await file.arrayBuffer());
  let binary = "";
  const chunkSize = 0x8000;
  for (let offset = 0; offset < bytes.length; offset += chunkSize) {
    binary += String.fromCharCode(...bytes.subarray(offset, offset + chunkSize));
  }
  return {
    filename: file.name,
    type: file.type,
    content: btoa(binary),
  };
};

const getPreferredProvider = (channel: Channel) => {
  const envName = channel === "email"
    ? "VITE_EMAIL_PROVIDER_PREFERRED"
    : channel === "sms"
      ? "VITE_SMS_PROVIDER_PREFERRED"
      : "VITE_WHATSAPP_PROVIDER_PREFERRED";
  const defaultProvider = providerOptions[channel][0];
  const preferredProvider = getEnvValue(envName, defaultProvider).toLowerCase();
  return providerOptions[channel].includes(preferredProvider) ? preferredProvider : defaultProvider;
};

const firstMatchingValue = (row: Record<string, string>, keys: string[]) => {
  for (const key of keys) {
    const value = row[key];
    if (value && value.trim()) return value.trim();
  }
  return "";
};

const parseCsvLine = (line: string) => {
  const cells: string[] = [];
  let current = "";
  let inQuotes = false;

  for (let index = 0; index < line.length; index += 1) {
    const char = line[index];
    if (char === '"') {
      if (inQuotes && line[index + 1] === '"') {
        current += '"';
        index += 1;
      } else {
        inQuotes = !inQuotes;
      }
      continue;
    }

    if (char === "," && !inQuotes) {
      cells.push(current);
      current = "";
      continue;
    }

    current += char;
  }

  cells.push(current);
  return cells.map((cell) => cell.trim());
};

const parseCsvText = (text: string) => {
  const rows = text
    .split(/\r?\n/)
    .filter((line) => line.trim().length > 0);

  if (rows.length === 0) return [];

  const headers = parseCsvLine(rows[0]).map((header) => normalizeHeader(header));
  return rows.slice(1).map((line) => {
    const values = parseCsvLine(line);
    const row: Record<string, string> = {};
    headers.forEach((header, index) => {
      row[header] = values[index] ?? "";
    });
    return row;
  });
};

const detectDefaultMapping = (headers: string[]): MappedFields => {
  const match = (values: string[]) => {
    for (const value of values) {
      const found = headers.find((header) => normalizeHeader(header) === normalizeHeader(value));
      if (found) return found;
    }
    return "";
  };

  return {
    name: match(["name", "full name", "customer name", "contact name"]),
    email: match(["email", "email address", "e-mail", "e mail", "customer email", "email id"]),
    phone: match(["mobile", "phone", "phone number", "mobile number", "cellphone", "mobile no", "contact number", "whatsapp", "whatsapp number", "telephone"]),
    company: match(["company", "organization", "brand", "account name"]),
    segment: match(["segment", "customer type", "group", "category"]),
  };
};

const mapRowsToContacts = (rows: Record<string, string>[], mapping: MappedFields): Contact[] => {
  const mappedContacts = rows.map((row, index) => {
    const name = firstMatchingValue(row, [mapping.name, "name", "full name", "customer name", "contact name"]) || `Contact ${index + 1}`;
    const email = firstMatchingValue(row, [mapping.email, "email", "email address", "e mail", "customer email", "email id"]);
    const phone = firstMatchingValue(row, [mapping.phone, "phone", "phone number", "mobile", "mobile number", "mobile no", "contact number", "whatsapp", "whatsapp number", "telephone"]);

    return {
      id: `${name}-${index}`,
      name,
      email,
      phone,
      company: firstMatchingValue(row, [mapping.company, "company", "organization", "brand", "account name"]),
      segment: firstMatchingValue(row, [mapping.segment, "segment", "customer type", "group", "category"]) || "General",
      status: "active" as ContactStatus,
      custom: { city: row.city || "N/A", plan: row.plan || "Standard" },
    };
  });

  const seenAddresses = new Set<string>();
  return mappedContacts.filter((contact) => {
    const address = contact.email
      ? `email:${contact.email.toLowerCase()}`
      : contact.phone
        ? `phone:${contact.phone.replace(/\D/g, "")}`
        : "";
    if (!address) return true;
    if (seenAddresses.has(address)) return false;
    seenAddresses.add(address);
    return true;
  });
};

const renderTemplate = (template: string, contact: Contact) => {
  const sample: Record<string, string> = {
    name: contact?.name || "Customer",
    company: contact?.company || "your company",
    segment: contact?.segment || "customer",
    sender_name: "NorthStar Support",
    ...contact.custom,
  };

  return template.replace(/{{\s*([a-zA-Z0-9_\s.-]+)\s*}}/g, (_, key: string) => {
    const rawKey = key.trim();
    return sample[rawKey] ?? `{{${rawKey}}}`;
  });
};

const downloadCsv = (rows: CampaignResult[]) => {
  const csvRows = [
    ["name", "recipient", "channel", "status", "reason"],
    ...rows.map((row) => [row.name, row.recipient, row.channel, row.status, row.reason]),
  ];

  const csv = csvRows
    .map((row) => row.map((value) => `"${String(value).replace(/"/g, '""')}"`).join(","))
    .join("\n");

  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = "campaign-report.csv";
  link.click();
  URL.revokeObjectURL(link.href);
};

function App() {
  const fileInputRef = useRef<HTMLInputElement | null>(null);
  const attachmentInputRef = useRef<HTMLInputElement | null>(null);
  const [role, setRole] = useState<Role>("admin");
  const [channel, setChannel] = useState<Channel>("email");
  const [provider, setProvider] = useState(() => getPreferredProvider("email"));
  const [fromName, setFromName] = useState(() => getEnvValue("VITE_SENDGRID_FROM_NAME", "My Application"));
  const [fromEmail, setFromEmail] = useState(() => getEnvValue("VITE_SENDGRID_FROM_EMAIL", "noreply@example.com"));
  const [replyTo, setReplyTo] = useState(() => getEnvValue("VITE_SENDGRID_REPLY_TO", "support@example.com"));
  const [customTemplates, setCustomTemplates] = useState<TemplateItem[]>([]);
  const [customTemplateName, setCustomTemplateName] = useState("");
  const [customTemplateSubject, setCustomTemplateSubject] = useState("Customer update");
  const [customTemplateBody, setCustomTemplateBody] = useState("Hi {{name}},\n\nThis is a personalized message for {{company}}.\n\nRegards,\n{{sender_name}}");
  const [batchSize, setBatchSize] = useState(25);
  const [schedule, setSchedule] = useState<"now" | "today" | "tomorrow" | "custom">("now");
  const [customScheduleDate, setCustomScheduleDate] = useState(() => new Date().toISOString().slice(0, 10));
  const [scheduleCustomerField, setScheduleCustomerField] = useState("name");
  const [activeDataTab, setActiveDataTab] = useState<"summary" | "rows">("summary");
  const [activeWorkflowTab, setActiveWorkflowTab] = useState<WorkflowTab>("upload");
  const [retryFailed, setRetryFailed] = useState(true);
  const [maxRetries, setMaxRetries] = useState(2);
  const [trackingEnabled, setTrackingEnabled] = useState(true);
  const [subject, setSubject] = useState("Welcome to NorthStar");
  const [senderName, setSenderName] = useState("NorthStar Support");
  const [templateText, setTemplateText] = useState(templateLibrary[0].body);
  const [selectedTemplateId, setSelectedTemplateId] = useState(templateLibrary[0].id);
  const [fileName, setFileName] = useState("customers.csv");
  const [emailAttachment, setEmailAttachment] = useState<File | null>(null);
  const [attachmentError, setAttachmentError] = useState("");
  const [featureSpotlightIndex, setFeatureSpotlightIndex] = useState(0);
  const [footerFlashIndex, setFooterFlashIndex] = useState(0);
  const [headers, setHeaders] = useState<string[]>(["name", "email", "phone", "company", "segment"]);
  const [mapping, setMapping] = useState<MappedFields>({
    name: "name",
    email: "email",
    phone: "phone",
    company: "company",
    segment: "segment",
  });
  const [contacts, setContacts] = useState<Contact[]>(sampleContacts);
  const [sourceRows, setSourceRows] = useState<Record<string, string>[]>([]);
  const [selectedContactIds, setSelectedContactIds] = useState<string[]>(() => sampleContacts.map((contact) => contact.id));
  const [resultRows, setResultRows] = useState<CampaignResult[]>([]);
  const [isSending, setIsSending] = useState(false);
  const [isTestSending, setIsTestSending] = useState(false);
  const [testSendFeedback, setTestSendFeedback] = useState("");
  const [progress, setProgress] = useState(0);
  const [uploadError, setUploadError] = useState("");

  const canEdit = role === "admin";
  const activeWorkflowIndex = workflowTabs.findIndex((tab) => tab.id === activeWorkflowTab);

  const navigateWorkflow = (tab: WorkflowTab) => {
    setActiveWorkflowTab(tab);
    document.getElementById("campaign-workflow")?.scrollIntoView({ behavior: "smooth", block: "start" });
  };

  const toggleAllValidContacts = () => {
    setSelectedContactIds((previous) => {
      if (allValidContactsSelected) {
        const validIds = new Set(validContacts.map((contact) => contact.id));
        return previous.filter((id) => !validIds.has(id));
      }
      return [...new Set([...previous, ...validContacts.map((contact) => contact.id)])];
    });
  };

  const toggleContactSelection = (contactId: string) => {
    setSelectedContactIds((previous) => previous.includes(contactId)
      ? previous.filter((id) => id !== contactId)
      : [...previous, contactId]);
  };

  useEffect(() => {
    const timerId = window.setInterval(() => {
      setFeatureSpotlightIndex((index) => (index + 1) % flashMessages.length);
    }, 5000);
    return () => window.clearInterval(timerId);
  }, []);

  useEffect(() => {
    const timerId = window.setInterval(() => {
      setFooterFlashIndex((index) => (index + 1) % footerFlashMessages.length);
    }, 7000);
    return () => window.clearInterval(timerId);
  }, []);

  const validContacts = useMemo(
    () => contacts.filter((contact) => isContactValidForChannel(contact, channel)),
    [contacts, channel],
  );
  const selectedValidContacts = useMemo(
    () => validContacts.filter((contact) => selectedContactIds.includes(contact.id)),
    [validContacts, selectedContactIds],
  );
  const allValidContactsSelected = validContacts.length > 0
    && validContacts.every((contact) => selectedContactIds.includes(contact.id));

  const validationSummary = useMemo(() => {
    const delivered = contacts.filter((item) => item.status !== "unsubscribed" && isContactValidForChannel(item, channel)).length;
    const unsubscribed = contacts.filter((item) => item.status === "unsubscribed").length;
    const invalid = contacts.filter((item) => item.status !== "unsubscribed" && !isContactValidForChannel(item, channel)).length;
    const duplicates = 0;
    return { delivered, unsubscribed, invalid, duplicates };
  }, [contacts, channel]);

  const dashboardStats = useMemo(() => {
    const stats: Record<CampaignStatus, number> = {
      sent: 0,
      delivered: 0,
      opened: 0,
      clicked: 0,
      bounced: 0,
      failed: 0,
      unsubscribed: 0,
    };

    resultRows.forEach((row) => {
      stats[row.status] += 1;
    });

    return {
      ...stats,
      sent: stats.sent + stats.delivered + stats.opened + stats.clicked,
    };
  }, [resultRows]);

  const templateVariables = useMemo(() => {
    const matches = templateText.match(/{{\s*([a-zA-Z0-9_\s.-]+)\s*}}/g) ?? [];
    return [...new Set(matches.map((match) => match.replace(/[{}]/g, "").trim()))];
  }, [templateText]);

  const activityFeed = [
    { name: "Aisha Patel", note: "Welcome flow synced", tone: "text-emerald-300" },
    { name: "Marcus Chen", note: "Invoice reminder opened", tone: "text-cyan-300" },
    { name: "Priya Nair", note: "Unsubscribe rule matched", tone: "text-amber-300" },
    { name: "Daniel Gomez", note: "Delivery failed retry queued", tone: "text-rose-300" },
  ];

  const previewContact = selectedValidContacts[0] ?? validContacts[0] ?? contacts[0] ?? sampleContacts[0];
  const previewBody = renderTemplate(templateText, previewContact);
  const attachmentExtensions = ["pdf", "doc", "docx", "xls", "xlsx", "ppt", "pptx", "jpg", "jpeg", "png", "gif", "webp", "mp4", "mov", "avi", "mkv", "webm"];

  const handleAttachmentUpload = (event: ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    const extension = file.name.split(".").pop()?.toLowerCase();
    if (!extension || !attachmentExtensions.includes(extension)) {
      setAttachmentError("Choose a document, spreadsheet, presentation, image, or video file.");
      event.currentTarget.value = "";
      return;
    }

    setEmailAttachment(file);
    setAttachmentError("");
  };

  const removeEmailAttachment = () => {
    setEmailAttachment(null);
    setAttachmentError("");
    if (attachmentInputRef.current) attachmentInputRef.current.value = "";
  };

  const handleFileUpload = (event: ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    const fileType = file.name.toLowerCase();
    const isCsv = fileType.endsWith(".csv");
    const isExcel = fileType.endsWith(".xlsx") || fileType.endsWith(".xls");

    if (!isCsv && !isExcel) {
      setUploadError("Please upload a CSV or Excel sheet.");
      return;
    }

    const reader = new FileReader();
    reader.onload = () => {
      try {
        const rawContent = reader.result;
        let rows: Record<string, string>[] = [];

        if (isCsv && typeof rawContent === "string") {
          rows = parseCsvText(rawContent);
        } else if (rawContent instanceof ArrayBuffer) {
          const workbook = XLSX.read(rawContent, { type: "array" });
          const worksheet = workbook.Sheets[workbook.SheetNames[0]];
          const jsonRows = XLSX.utils.sheet_to_json<Record<string, unknown>>(worksheet, { defval: "" });
          rows = jsonRows.map((row) => {
            const normalized: Record<string, string> = {};
            Object.entries(row).forEach(([key, value]) => {
              normalized[normalizeHeader(key)] = String(value ?? "").trim();
            });
            return normalized;
          });
        }

        if (rows.length === 0) {
          setUploadError("The uploaded file does not contain rows to process.");
          return;
        }

        if (rows.length > MAX_ROWS_PER_UPLOAD) {
          setUploadError(`Upload limit reached: ${MAX_ROWS_PER_UPLOAD.toLocaleString()} rows maximum.`);
          return;
        }

        const sourceHeaders = Object.keys(rows[0]).filter(Boolean);
        const nextHeaders = sourceHeaders.length ? sourceHeaders : ["name", "email", "phone", "company", "segment"];
        const nextMapping = detectDefaultMapping(nextHeaders);
        const mappedContacts = mapRowsToContacts(rows, nextMapping);
        setHeaders(nextHeaders);
        setMapping(nextMapping);
        setSourceRows(rows.slice(0, MAX_ROWS_PER_UPLOAD));
        setContacts(mappedContacts);
        setSelectedContactIds(mappedContacts.map((contact) => contact.id));
        setFileName(file.name);
        setUploadError("");
        setResultRows([]);
        setProgress(0);
      } catch {
        setUploadError("This file could not be parsed. Check the content and try again.");
      }
    };

    reader.readAsArrayBuffer(file);
  };

  const handleTemplateChange = (templateId: string) => {
    const template = [...templateLibrary, ...customTemplates].find((item) => item.id === templateId);
    if (!template) return;

    setSelectedTemplateId(templateId);
    setSubject(template.subject);
    setTemplateText(template.body);
  };

  const handleFieldMapping = (field: keyof MappedFields, value: string) => {
    const updatedMapping = { ...mapping, [field]: value };
    setMapping(updatedMapping);
    if (sourceRows.length > 0) {
      const mappedContacts = mapRowsToContacts(sourceRows, updatedMapping);
      setContacts(mappedContacts);
      setSelectedContactIds(mappedContacts.map((contact) => contact.id));
    }
  };

  const handleContactFieldChange = (
    contactId: string,
    field: "name" | "email" | "phone" | "company" | "segment",
    value: string,
  ) => {
    if (!canEdit) return;
    setContacts((previous) =>
      previous.map((contact) => (contact.id === contactId ? { ...contact, [field]: value } : contact)),
    );
  };

  const handleCreateCustomTemplate = () => {
    if (!customTemplateName.trim()) return;

    const newTemplate: TemplateItem = {
      id: `custom-${Date.now()}`,
      name: customTemplateName.trim(),
      subject: customTemplateSubject.trim() || "Custom notification",
      body: customTemplateBody.trim() || "Hi {{name}}",
    };

    setCustomTemplates((previous) => [newTemplate, ...previous]);
    setSelectedTemplateId(newTemplate.id);
    setSubject(newTemplate.subject);
    setTemplateText(newTemplate.body);
    setCustomTemplateName("");
    setCustomTemplateSubject("Customer update");
    setCustomTemplateBody("Hi {{name}},\n\nThis is a personalized message for {{company}}.\n\nRegards,\n{{sender_name}}");
  };

  const triggerTestSend = async () => {
    if (!selectedValidContacts.length) {
      setTestSendFeedback("Select at least one valid recipient in Choose Channel before sending a test.");
      return;
    }
    const contact = previewContact;
    if (channel === "email" && provider === "mailtrap") {
      setIsTestSending(true);
      setTestSendFeedback("");
      try {
        const response = await fetch("/api/mailtrap/test-send", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            from: { email: fromEmail, name: fromName },
            to: contact.email,
            replyTo,
            subject,
            text: previewBody,
            attachment: emailAttachment ? await encodeAttachment(emailAttachment) : undefined,
          }),
        });
        const payload: { error?: string; messageIds?: string[] } = await response.json();
        if (!response.ok) throw new Error(payload.error || "Mailtrap test send failed.");

        const messageId = payload.messageIds?.[0];
        const result: CampaignResult = {
          id: `test-${Date.now()}`,
          name: contact.name,
          recipient: contact.email,
          channel: "email",
          status: "sent",
          reason: `Mailtrap accepted the test email${messageId ? ` (${messageId})` : ""}.`,
        };
        setResultRows((previous) => [result, ...previous].slice(0, 12));
        setTestSendFeedback(`Mailtrap accepted the test email for ${contact.email}.`);
      } catch (error) {
        const reason = error instanceof Error ? error.message : "Mailtrap test send failed.";
        const result: CampaignResult = {
          id: `test-${Date.now()}`,
          name: contact.name,
          recipient: contact.email,
          channel: "email",
          status: "failed",
          reason,
        };
        setResultRows((previous) => [result, ...previous].slice(0, 12));
        setTestSendFeedback(reason);
      } finally {
        setIsTestSending(false);
      }
      return;
    }

    const result: CampaignResult = {
      id: `test-${Date.now()}`,
      name: contact.name,
      recipient: channel === "email" ? contact.email : contact.phone,
      channel,
      status: "sent",
      reason: `Test send simulated via ${provider.toUpperCase()}${channel === "email" && emailAttachment ? `; attachment: ${emailAttachment.name}` : ""}`,
    };
    setResultRows((previous) => [result, ...previous].slice(0, 12));
    setTestSendFeedback(`Test send simulated via ${provider.toUpperCase()}.`);
  };

  const triggerSend = async () => {
    if (!selectedValidContacts.length) {
      setResultRows([
        {
          id: `failed-${Date.now()}`,
          name: "System",
          recipient: "n/a",
          channel,
          status: "failed",
          reason: "No valid recipients are selected for this channel.",
        },
      ]);
      return;
    }

    setIsSending(true);
    setProgress(0);
    const results: CampaignResult[] = [];
    const recipients = selectedValidContacts.slice(0, Math.max(1, batchSize));

    for (let index = 0; index < recipients.length; index += 1) {
      const contact = recipients[index];
      const recipient = channel === "email" ? contact.email : contact.phone;
      const outcome: CampaignStatus = index % 9 === 0 ? "bounced" : index % 6 === 0 ? "opened" : index % 4 === 0 ? "clicked" : "delivered";

      const providerLabel = provider.toUpperCase();
      const scheduleLabel =
        schedule === "now"
          ? "Sent immediately"
          : schedule === "today"
            ? "Scheduled for today"
            : schedule === "tomorrow"
              ? "Scheduled for tomorrow"
              : `Scheduled for ${customScheduleDate}`;

      await new Promise((resolve) => {
        window.setTimeout(() => {
          results.push({
            id: `${contact.id}-${index}`,
            name: contact.name,
            recipient,
            channel,
            status: outcome,
            reason: `${outcome === "bounced"
              ? `Rejected by ${providerLabel}; ${scheduleLabel}`
              : outcome === "clicked"
                ? `Link clicked via ${providerLabel}; ${scheduleLabel}`
                : outcome === "opened"
                  ? `Opened by recipient via ${providerLabel}; ${scheduleLabel}`
                  : `Message delivered via ${providerLabel}; ${scheduleLabel}`}${channel === "email" && emailAttachment ? `; attachment: ${emailAttachment.name}` : ""}`,
          });
          setResultRows([...results]);
          setProgress(Math.round(((index + 1) / recipients.length) * 100));
          resolve(undefined);
        }, 250);
      });
    }

    if (retryFailed) {
      const failedRows = results.filter((entry) => entry.status === "bounced").map((entry) => ({
        ...entry,
        id: `retry-${entry.id}`,
        status: "sent" as CampaignStatus,
        reason: `Retry scheduled (${maxRetries} retries allowed)`,
      }));
      setResultRows((previous) => [...previous, ...failedRows]);
    }

    setIsSending(false);
    setProgress(100);
  };

  return (
    <div className="min-h-screen text-slate-100" onContextMenu={(event) => event.preventDefault()}>
      <div className="mx-auto max-w-7xl px-4 py-8 sm:px-6 lg:px-8">
        <header className="glass-panel glow-card fade-in mb-8 overflow-hidden rounded-[28px] border border-violet-500/30 bg-gradient-to-br from-slate-900/95 via-slate-900/80 to-violet-950/60 p-5 shadow-[0_25px_80px_rgba(76,29,149,0.35)]">
          <div className="flex flex-col gap-4 lg:flex-row lg:items-center lg:justify-between">
            <div>
              <p className="text-xs font-semibold uppercase tracking-[0.38em] text-violet-300">MsgFlow</p>
              <h1 className="mt-3 text-3xl font-black tracking-tight text-white sm:text-4xl">
                Campaign <span className="shimmer-text">workspace</span>
              </h1>
            </div>

            <div className="flex flex-wrap items-center gap-3">
              <div className="rounded-full border border-slate-700 bg-slate-950/70 px-3 py-2 text-sm text-slate-200 shadow-lg shadow-slate-950/40">
                {role === "admin" ? "Admin access" : "User view"}
              </div>
              <div className="rounded-full border border-violet-500/30 bg-violet-500/10 px-3 py-2 text-sm text-violet-100 shadow-lg shadow-violet-900/30">
                {selectedValidContacts.length} selected / {validContacts.length} valid
              </div>
            </div>
          </div>

          <div
            key={featureSpotlightIndex}
            className="feature-spotlight relative mt-5 flex min-h-28 items-center gap-4 overflow-hidden rounded-2xl border border-cyan-400/20 bg-gradient-to-r from-slate-950/90 via-violet-950/35 to-slate-950/80 p-5 sm:p-6"
            aria-live="polite"
            aria-atomic="true"
          >
            <div className="absolute -right-10 -top-20 h-48 w-48 rounded-full bg-violet-500/10 blur-3xl" aria-hidden="true" />
            <span className="relative flex h-3 w-3 shrink-0" aria-hidden="true">
              <span className="absolute inline-flex h-full w-full animate-ping rounded-full bg-emerald-400 opacity-60" />
              <span className="relative inline-flex h-3 w-3 rounded-full bg-emerald-400 shadow-[0_0_16px_rgba(52,211,153,0.8)]" />
            </span>
            <div className="relative min-w-0 flex-1">
              <p className="text-[10px] font-semibold uppercase tracking-[0.24em] text-cyan-300">
                MsgFlow <span className="mx-1 text-slate-600">/</span> Channel spotlight
              </p>
              <p className="flash-message-title mt-2 text-xl font-bold tracking-tight text-white sm:text-2xl">
                {flashMessages[featureSpotlightIndex]}
              </p>
              <div className="mt-3 flex items-center gap-1.5" aria-hidden="true">
                {flashMessages.map((_, index) => (
                  <span
                    key={index}
                    className={`h-1 rounded-full transition-all duration-500 ${
                      index === featureSpotlightIndex ? "w-7 bg-cyan-300" : "w-1.5 bg-slate-600"
                    }`}
                  />
                ))}
              </div>
            </div>
          </div>

          <div className="mt-6 grid gap-3 md:grid-cols-3">
            {[
              { label: "Upload limit", value: "10,000 rows" },
              { label: "Channels", value: "Email · SMS · WhatsApp" },
              { label: "Edit access", value: role === "admin" ? "Enabled" : "Restricted to admin" },
            ].map((item) => (
              <div key={item.label} className="rounded-2xl border border-slate-800/90 bg-slate-950/50 p-3 shadow-inner shadow-slate-950/40">
                <p className="text-xs uppercase tracking-[0.25em] text-slate-400">{item.label}</p>
                <p className="mt-2 text-base font-semibold text-white">{item.value}</p>
              </div>
            ))}
          </div>
        </header>

        <div className="mb-8 rounded-2xl border border-slate-800/80 bg-slate-900/70 p-4 shadow-xl shadow-slate-950/30 backdrop-blur-sm">
          <div className="flex flex-col gap-4 lg:flex-row lg:items-center lg:justify-between">
            <div>
              <p className="text-sm text-slate-400">User role</p>
              <div className="mt-2 inline-flex rounded-full border border-slate-700 bg-slate-950 p-1">
                {(["admin", "user"] as Role[]).map((item) => (
                  <button
                    key={item}
                    type="button"
                    onClick={() => setRole(item)}
                    className={`rounded-full px-4 py-2 text-sm font-medium transition ${
                      role === item ? "bg-violet-500 text-white shadow-lg shadow-violet-600/30" : "text-slate-300 hover:text-white"
                    }`}
                  >
                    {item === "admin" ? "Admin" : "User"}
                  </button>
                ))}
              </div>
            </div>

            <div className="text-sm text-slate-300">
              {canEdit
                ? "Admin can create, update, and schedule campaigns."
                : "User can upload contacts and view campaign status, but cannot edit templates or settings."}
            </div>
          </div>
        </div>

        <div id="campaign-workflow" className="mt-6 overflow-hidden rounded-2xl border border-slate-800 bg-slate-900/70 shadow-xl shadow-slate-950/30">
          <div role="tablist" aria-label="Campaign setup steps" className="flex gap-2 overflow-x-auto border-b border-slate-800 p-2">
            {workflowTabs.map((tab) => (
              <button
                key={tab.id}
                type="button"
                role="tab"
                aria-selected={activeWorkflowTab === tab.id}
                onClick={() => setActiveWorkflowTab(tab.id)}
                className={`shrink-0 rounded-xl px-4 py-3 text-sm font-medium transition ${
                  activeWorkflowTab === tab.id
                    ? "bg-violet-500 text-white shadow-lg shadow-violet-600/20"
                    : "text-slate-300 hover:bg-slate-800 hover:text-white"
                }`}
              >
                {tab.label}
              </button>
            ))}
          </div>

          <div className="p-4 sm:p-6">
            <nav aria-label="Workflow step navigation" className="mb-5 flex items-center justify-between gap-4 border-b border-slate-800 pb-5">
              <button
                type="button"
                onClick={() => navigateWorkflow(workflowTabs[activeWorkflowIndex - 1].id)}
                disabled={activeWorkflowIndex === 0}
                className="rounded-xl border border-slate-700 bg-slate-900 px-4 py-2.5 text-sm font-medium text-slate-200 transition hover:border-violet-400 hover:text-white disabled:cursor-not-allowed disabled:opacity-40"
              >
                ← Previous
              </button>
              <span className="text-center text-xs text-slate-400 sm:text-sm">
                Step {activeWorkflowIndex + 1} of {workflowTabs.length}: {workflowTabs[activeWorkflowIndex].label.replace(/^\d+\.\s*/, "")}
              </span>
              <button
                type="button"
                onClick={() => navigateWorkflow(workflowTabs[activeWorkflowIndex + 1].id)}
                disabled={activeWorkflowIndex === workflowTabs.length - 1}
                className="rounded-xl bg-violet-600 px-4 py-2.5 text-sm font-semibold text-white transition hover:bg-violet-500 disabled:cursor-not-allowed disabled:opacity-40"
              >
                Next →
              </button>
            </nav>
            {activeWorkflowTab === "upload" ? (
            <section className="glass-panel fade-in rounded-2xl border border-slate-800/80 bg-slate-900/80 p-5 shadow-[0_20px_50px_rgba(15,23,42,0.7)]">
              <div className="mb-5 flex items-center justify-between gap-3">
                <div className="flex items-center gap-3">
                  <div className="rounded-xl bg-violet-500/15 p-2 text-violet-200 ring-1 ring-violet-400/30">
                    <UploadCloudIcon className="h-5 w-5" />
                  </div>
                  <div>
                    <h2 className="text-lg font-semibold text-white">1. Upload Contacts</h2>
                    <p className="text-sm text-slate-400">Initial limit: {MAX_ROWS_PER_UPLOAD.toLocaleString()} rows</p>
                  </div>
                </div>
                <button
                  type="button"
                  className="rounded-xl border border-slate-700 bg-slate-800 px-3 py-2 text-sm text-slate-200 transition hover:border-slate-500 hover:bg-slate-700"
                  onClick={() => fileInputRef.current?.click()}
                >
                  Upload CSV
                </button>
              </div>

              <input
                ref={fileInputRef}
                type="file"
                accept=".csv,.xlsx,.xls"
                onChange={handleFileUpload}
                className="hidden"
              />

              <label className="flex cursor-pointer flex-col items-center justify-center rounded-2xl border border-dashed border-violet-500/40 bg-slate-950/70 p-6 text-center transition hover:border-violet-400 hover:bg-violet-500/5">
                <UploadCloudIcon className="mb-3 h-10 w-10 text-violet-300" />
                <span className="text-base font-medium text-white">{fileName}</span>
                <span className="mt-2 text-xs text-slate-400">
                  {fileName ? `Size: ${formatFileSize(240000)}` : "Drop a CSV or Excel file here"}
                </span>
              </label>

              {uploadError ? (
                <div className="mt-4 rounded-xl border border-rose-500/40 bg-rose-500/10 p-3 text-sm text-rose-200">{uploadError}</div>
              ) : null}
            </section>
            ) : null}

            {activeWorkflowTab === "csv-preview" ? (
            <section className="glass-panel fade-in rounded-2xl border border-slate-800/80 bg-slate-900/80 p-5 shadow-[0_20px_50px_rgba(15,23,42,0.7)]">
              <div className="mb-5">
                <h2 className="text-lg font-semibold text-white">2. Verify Data</h2>
                <p className="text-sm text-slate-400">{fileName} · {contacts.length.toLocaleString()} contacts loaded</p>
              </div>
              <div className="mt-5 grid gap-3 md:grid-cols-4">
                <div className="rounded-2xl border border-slate-800 bg-slate-950/50 p-3 shadow-lg shadow-emerald-950/20">
                  <p className="text-xs uppercase tracking-[0.2em] text-slate-400">Valid</p>
                  <p className="mt-2 text-2xl font-bold text-emerald-400">{validationSummary.delivered}</p>
                </div>
                <div className="rounded-2xl border border-slate-800 bg-slate-950/50 p-3 shadow-lg shadow-amber-950/20">
                  <p className="text-xs uppercase tracking-[0.2em] text-slate-400">Unsubscribed</p>
                  <p className="mt-2 text-2xl font-bold text-amber-400">{validationSummary.unsubscribed}</p>
                </div>
                <div className="rounded-2xl border border-slate-800 bg-slate-950/50 p-3 shadow-lg shadow-rose-950/20">
                  <p className="text-xs uppercase tracking-[0.2em] text-slate-400">Invalid</p>
                  <p className="mt-2 text-2xl font-bold text-rose-400">{validationSummary.invalid}</p>
                </div>
                <div className="rounded-2xl border border-slate-800 bg-slate-950/50 p-3 shadow-lg shadow-sky-950/20">
                  <p className="text-xs uppercase tracking-[0.2em] text-slate-400">Duplicate</p>
                  <p className="mt-2 text-2xl font-bold text-sky-400">{validationSummary.duplicates}</p>
                </div>
              </div>

              <div className="mt-5 overflow-hidden rounded-2xl border border-slate-800">
                <div className="overflow-x-auto">
                  <table className="min-w-full text-left text-sm">
                    <thead className="bg-slate-950/70 text-slate-300">
                      <tr>
                        <th className="px-4 py-3 font-medium">Field</th>
                        <th className="px-4 py-3 font-medium">CSV column</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-800 bg-slate-900/70 text-slate-100">
                      {Object.entries(mapping).map(([field, value]) => (
                        <tr key={field}>
                          <td className="px-4 py-3 capitalize">{field}</td>
                          <td className="px-4 py-3">
                            <select
                              value={value}
                              disabled={!canEdit}
                              onChange={(event) => handleFieldMapping(field as keyof MappedFields, event.target.value)}
                              className="w-full rounded-lg border border-slate-700 bg-slate-950 px-3 py-2 text-sm text-slate-100 outline-none focus:border-violet-500 disabled:cursor-not-allowed disabled:opacity-60"
                            >
                              <option value="">— Select column —</option>
                              {headers.map((header) => (
                                <option key={header} value={header}>
                                  {header}
                                </option>
                              ))}
                            </select>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
              <div className="mt-6 overflow-hidden rounded-2xl border border-slate-800">
                <div className="overflow-x-auto">
                  <table className="min-w-full text-left text-sm">
                    <thead className="bg-slate-950/70 text-slate-300">
                      <tr>
                        <th className="px-4 py-3 font-medium">Name{canEdit ? " (editable)" : ""}</th>
                        <th className="px-4 py-3 font-medium">Email{canEdit ? " (editable)" : ""}</th>
                        <th className="px-4 py-3 font-medium">Mobile{canEdit ? " (editable)" : ""}</th>
                        <th className="px-4 py-3 font-medium">Company{canEdit ? " (editable)" : ""}</th>
                        <th className="px-4 py-3 font-medium">Segment{canEdit ? " (editable)" : ""}</th>
                        <th className="px-4 py-3 font-medium">Status</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-800 bg-slate-900/70 text-slate-100">
                      {contacts.map((contact) => (
                        <tr key={contact.id}>
                          {([
                            ["name", contact.name],
                            ["email", contact.email],
                            ["phone", contact.phone],
                            ["company", contact.company],
                            ["segment", contact.segment],
                          ] as const).map(([field, value]) => (
                            <td key={field} className="px-2 py-2">
                              <input
                                aria-label={`${field} for ${contact.name}`}
                                type={field === "email" ? "email" : "text"}
                                value={value}
                                disabled={!canEdit}
                                onChange={(event) => handleContactFieldChange(contact.id, field, event.target.value)}
                                className="w-full min-w-28 rounded-lg border border-transparent bg-transparent px-2 py-1.5 text-slate-100 outline-none transition focus:border-violet-500 focus:bg-slate-950 disabled:cursor-default disabled:text-slate-300"
                              />
                            </td>
                          ))}
                          <td className="px-4 py-3 capitalize text-slate-300">
                            {contact.status === "unsubscribed"
                              ? "Unsubscribed"
                              : isContactValidForChannel(contact, channel)
                                ? "Valid"
                                : channel === "email"
                                  ? "Invalid email"
                                  : "Invalid Indian mobile"}
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            </section>
            ) : null}

            {activeWorkflowTab === "template" ? (
            <section className="glass-panel fade-in rounded-2xl border border-slate-800/80 bg-slate-900/80 p-5 shadow-[0_20px_50px_rgba(15,23,42,0.7)]">
              <div className="mb-4 flex items-center justify-between gap-3">
                <div className="flex items-center gap-3">
                  <div className="rounded-xl bg-emerald-500/15 p-2 text-emerald-200 ring-1 ring-emerald-400/30">
                    <FileTextIcon className="h-5 w-5" />
                  </div>
                  <div>
                    <h2 className="text-lg font-semibold text-white">3. Craft Message</h2>
                    <p className="text-sm text-slate-400">Use {'{{name}}'}, {'{{company}}'}, {'{{segment}}'} and more</p>
                  </div>
                </div>
                <div className="text-xs uppercase tracking-[0.2em] text-slate-400">Template variables</div>
              </div>

              <div className="mb-4 grid gap-3 sm:grid-cols-3">
                {[...templateLibrary, ...customTemplates].map((template) => (
                  <button
                    key={template.id}
                    type="button"
                    onClick={() => handleTemplateChange(template.id)}
                    disabled={!canEdit}
                    className={`rounded-xl border px-3 py-2 text-left text-sm transition ${
                      selectedTemplateId === template.id
                        ? "border-violet-500 bg-violet-500/10 text-violet-100"
                        : "border-slate-700 bg-slate-800 text-slate-300"
                    } disabled:cursor-not-allowed disabled:opacity-60`}
                  >
                    {template.name}
                  </button>
                ))}
              </div>

              <div className="mb-4 rounded-2xl border border-slate-800 bg-slate-950/50 p-4">
                <div className="mb-3 text-sm font-medium text-slate-200">Create custom template</div>
                <div className="grid gap-3 md:grid-cols-2">
                  <input
                    value={customTemplateName}
                    onChange={(event) => setCustomTemplateName(event.target.value)}
                    placeholder="Template name"
                    disabled={!canEdit}
                    className="w-full rounded-xl border border-slate-700 bg-slate-950 px-3 py-2.5 text-sm text-white outline-none transition focus:border-violet-500 disabled:cursor-not-allowed disabled:opacity-60"
                  />
                  <input
                    value={customTemplateSubject}
                    onChange={(event) => setCustomTemplateSubject(event.target.value)}
                    placeholder="Subject line"
                    disabled={!canEdit}
                    className="w-full rounded-xl border border-slate-700 bg-slate-950 px-3 py-2.5 text-sm text-white outline-none transition focus:border-violet-500 disabled:cursor-not-allowed disabled:opacity-60"
                  />
                </div>
                <textarea
                  rows={4}
                  value={customTemplateBody}
                  onChange={(event) => setCustomTemplateBody(event.target.value)}
                  placeholder="Write your custom template body with {{name}}, {{company}}, {{segment}} etc."
                  disabled={!canEdit}
                  className="mt-3 w-full rounded-xl border border-slate-700 bg-slate-950 px-3 py-2.5 text-sm text-white outline-none transition focus:border-violet-500 disabled:cursor-not-allowed disabled:opacity-60"
                />
                <div className="mt-3 flex justify-end">
                  <button
                    type="button"
                    onClick={handleCreateCustomTemplate}
                    disabled={!canEdit || !customTemplateName.trim()}
                    className="rounded-xl bg-violet-600 px-3 py-2 text-sm font-semibold text-white transition hover:bg-violet-500 disabled:cursor-not-allowed disabled:opacity-60"
                  >
                    Save template
                  </button>
                </div>
              </div>

              <div className="space-y-4">
                <div>
                  <label className="mb-2 block text-sm font-medium text-slate-300">Subject line</label>
                  <input
                    value={subject}
                    onChange={(event) => setSubject(event.target.value)}
                    disabled={!canEdit}
                    className="w-full rounded-xl border border-slate-700 bg-slate-950 px-3 py-2.5 text-sm text-white outline-none transition focus:border-violet-500 disabled:cursor-not-allowed disabled:opacity-60"
                  />
                </div>

                <div>
                  <label className="mb-2 block text-sm font-medium text-slate-300">Sender name</label>
                  <input
                    value={senderName}
                    onChange={(event) => setSenderName(event.target.value)}
                    disabled={!canEdit}
                    className="w-full rounded-xl border border-slate-700 bg-slate-950 px-3 py-2.5 text-sm text-white outline-none transition focus:border-violet-500 disabled:cursor-not-allowed disabled:opacity-60"
                  />
                </div>

                <div>
                  <label className="mb-2 block text-sm font-medium text-slate-300">Message body</label>
                  <textarea
                    rows={7}
                    value={templateText}
                    onChange={(event) => setTemplateText(event.target.value)}
                    disabled={!canEdit}
                    className="w-full rounded-xl border border-slate-700 bg-slate-950 px-3 py-2.5 text-sm text-white outline-none transition focus:border-violet-500 disabled:cursor-not-allowed disabled:opacity-60"
                  />
                </div>

                {channel === "email" ? (
                  <div className="rounded-xl border border-slate-800 bg-slate-950/50 p-4">
                    <div className="mb-2">
                      <div className="text-sm font-medium text-slate-200">Optional email attachment</div>
                      <p className="mt-1 text-xs text-slate-400">Attach a document, brochure, image, or promotional video to each email.</p>
                    </div>
                    <input
                      ref={attachmentInputRef}
                      type="file"
                      accept=".pdf,.doc,.docx,.xls,.xlsx,.ppt,.pptx,.jpg,.jpeg,.png,.gif,.webp,.mp4,.mov,.avi,.mkv,.webm"
                      onChange={handleAttachmentUpload}
                      className="hidden"
                    />
                    {emailAttachment ? (
                      <div className="flex flex-wrap items-center justify-between gap-3 rounded-lg border border-slate-700 bg-slate-900 px-3 py-2">
                        <div className="min-w-0">
                          <div className="truncate text-sm text-white">{emailAttachment.name}</div>
                          <div className="text-xs text-slate-400">{formatFileSize(emailAttachment.size)}</div>
                        </div>
                        <div className="flex gap-2">
                          <button
                            type="button"
                            onClick={() => attachmentInputRef.current?.click()}
                            className="rounded-lg border border-slate-700 px-3 py-1.5 text-xs text-slate-200 transition hover:border-violet-400"
                          >
                            Change
                          </button>
                          <button
                            type="button"
                            onClick={removeEmailAttachment}
                            className="rounded-lg border border-slate-700 px-3 py-1.5 text-xs text-slate-200 transition hover:border-rose-400 hover:text-rose-200"
                          >
                            Remove
                          </button>
                        </div>
                      </div>
                    ) : (
                      <button
                        type="button"
                        onClick={() => attachmentInputRef.current?.click()}
                        className="flex w-full items-center justify-center gap-2 rounded-lg border border-dashed border-slate-700 px-3 py-3 text-sm text-slate-300 transition hover:border-violet-400 hover:text-white"
                      >
                        <UploadCloudIcon className="h-4 w-4" />
                        Choose a file
                      </button>
                    )}
                    {attachmentError ? <p role="alert" className="mt-2 text-sm text-rose-300">{attachmentError}</p> : null}
                  </div>
                ) : null}

                <div className="rounded-xl border border-slate-800 bg-slate-950/50 p-3">
                  <div className="mb-2 flex items-center justify-between text-xs uppercase tracking-[0.2em] text-slate-400">
                    <span>Available variables</span>
                    <span>{templateVariables.length}</span>
                  </div>
                  <div className="flex flex-wrap gap-2">
                    {templateVariables.length === 0 ? (
                      <span className="text-sm text-slate-400">No variables found</span>
                    ) : (
                      templateVariables.map((variable) => (
                        <span key={variable} className="rounded-full bg-slate-800 px-2 py-1 text-xs text-violet-200">
                          {`{{${variable}}}`}
                        </span>
                      ))
                    )}
                  </div>
                </div>
              </div>
            </section>
            ) : null}

            {activeWorkflowTab === "provider" ? (
            <section className="glass-panel fade-in rounded-2xl border border-slate-800/80 bg-slate-900/80 p-5 shadow-[0_20px_50px_rgba(15,23,42,0.7)]">
              <div className="mb-5 flex items-center gap-3">
                <div className="rounded-xl bg-sky-500/15 p-2 text-sky-200 ring-1 ring-sky-400/30">
                  <SendIcon className="h-5 w-5" />
                </div>
                <div>
                  <h2 className="text-lg font-semibold text-white">4. Choose Channel</h2>
                </div>
              </div>

              <div className="space-y-4">
                <div>
                  <label className="mb-2 block text-sm font-medium text-slate-300">Channel</label>
                  <div className="grid grid-cols-3 gap-2">
                    {(["email", "sms", "whatsapp"] as Channel[]).map((item) => (
                      <button
                        key={item}
                        type="button"
                        onClick={() => {
                          setChannel(item);
                          setProvider(getPreferredProvider(item));
                        }}
                        className={`rounded-xl border px-3 py-2 text-sm capitalize transition ${
                          channel === item
                            ? "border-violet-500 bg-violet-500/10 text-violet-100"
                            : "border-slate-700 bg-slate-800 text-slate-300"
                        }`}
                      >
                        {item}
                      </button>
                    ))}
                  </div>
                  <p className="mt-2 text-xs text-slate-400">
                    {channel === "email"
                      ? "Email channel validates each recipient's email address."
                      : "SMS and WhatsApp validate Indian mobile numbers (10 digits starting 6–9, optionally with +91 or 0 prefix)."}
                  </p>
                </div>

                <div className="rounded-xl border border-slate-800 bg-slate-950/50 p-4">
                  <div className="mb-3 flex flex-wrap items-center justify-between gap-3">
                    <div>
                      <h3 className="text-sm font-medium text-slate-200">
                        {sourceRows.length ? `Recipients from ${fileName}` : "Sample recipients"}
                      </h3>
                      <p className="mt-1 text-xs text-slate-400">
                        {selectedValidContacts.length} selected · {validContacts.length} valid for {channel}
                      </p>
                    </div>
                    <button
                      type="button"
                      onClick={toggleAllValidContacts}
                      disabled={!canEdit || validContacts.length === 0}
                      className="rounded-lg border border-violet-500/40 px-3 py-1.5 text-xs font-medium text-violet-200 transition hover:bg-violet-500/10 disabled:cursor-not-allowed disabled:opacity-50"
                    >
                      {allValidContactsSelected ? "Clear channel selection" : `Select all ${channel} contacts`}
                    </button>
                  </div>
                  {contacts.length ? (
                    <div className="max-h-64 space-y-2 overflow-y-auto pr-1">
                      {contacts.map((contact) => {
                        const isValid = isContactValidForChannel(contact, channel);
                        const address = channel === "email" ? contact.email : contact.phone;
                        const isSelected = selectedContactIds.includes(contact.id);
                        return (
                          <label
                            key={contact.id}
                            className={`flex items-center gap-3 rounded-lg border px-3 py-2 ${
                              isValid ? "border-slate-800 bg-slate-900/70" : "border-slate-800/70 bg-slate-900/30 opacity-70"
                            }`}
                          >
                            <input
                              type="checkbox"
                              checked={isSelected}
                              disabled={!canEdit || !isValid}
                              onChange={() => toggleContactSelection(contact.id)}
                              aria-label={`Select ${contact.name} for ${channel}`}
                              className="h-4 w-4 accent-violet-500 disabled:cursor-not-allowed"
                            />
                            <span className="min-w-0 flex-1">
                              <span className="block truncate text-sm text-slate-100">{contact.name}</span>
                              <span className="block truncate text-xs text-slate-400">
                                {address || `No ${channel === "email" ? "email address" : "phone number"} in CSV`}
                              </span>
                            </span>
                            <span className={`shrink-0 text-xs ${isValid ? "text-emerald-300" : "text-rose-300"}`}>
                              {contact.status === "unsubscribed" ? "Unsubscribed" : isValid ? "Valid" : "Invalid"}
                            </span>
                          </label>
                        );
                      })}
                    </div>
                  ) : (
                    <p className="text-sm text-slate-400">
                      No contacts loaded. Upload a CSV or Excel file in step 1.
                    </p>
                  )}
                  {!validContacts.length && contacts.length ? (
                    <p className="mt-3 text-xs text-amber-300">
                      No usable {channel === "email" ? "email addresses" : "Indian mobile numbers"} found. Check the CSV column mapping in step 2.
                    </p>
                  ) : null}
                </div>

                <div className="grid gap-4 md:grid-cols-2">
                  <div>
                    <label className="mb-2 block text-sm font-medium text-slate-300">Provider</label>
                    <select
                      value={provider}
                      disabled={!canEdit}
                      onChange={(event) => setProvider(event.target.value)}
                      className="w-full rounded-xl border border-slate-700 bg-slate-950 px-3 py-2.5 text-sm text-white outline-none transition focus:border-violet-500 disabled:cursor-not-allowed disabled:opacity-60"
                    >
                      {channel === "email" ? (
                        <>
                          <option value="smtp">SMTP / Email provider</option>
                          <option value="sendgrid">SendGrid</option>
                          <option value="brevo">Brevo</option>
                          <option value="mailtrap">Mailtrap</option>
                        </>
                      ) : null}
                      {channel === "sms" ? (
                        <>
                          <option value="msg91">MSG91</option>
                          <option value="twilio">Twilio</option>
                        </>
                      ) : null}
                      {channel === "whatsapp" ? (
                        <>
                          <option value="gupshup">Gupshup</option>
                          <option value="wati">WATI</option>
                        </>
                      ) : null}
                    </select>
                  </div>

                  <div>
                    <label className="mb-2 block text-sm font-medium text-slate-300">Schedule by</label>
                    <select
                      value={scheduleCustomerField}
                      onChange={(event) => setScheduleCustomerField(event.target.value)}
                      disabled={!canEdit}
                      className="w-full rounded-xl border border-slate-700 bg-slate-950 px-3 py-2.5 text-sm text-white outline-none transition focus:border-violet-500 disabled:cursor-not-allowed disabled:opacity-60"
                    >
                      <option value="name">Customer name</option>
                      <option value="email">Email</option>
                      <option value="phone">Phone</option>
                      <option value="company">Company</option>
                      <option value="segment">Segment</option>
                    </select>
                  </div>
                </div>

                {channel === "email" ? (
                  <div className="space-y-4 rounded-xl border border-slate-800 bg-slate-950/50 p-4">
                    <div className="grid gap-4 md:grid-cols-2">
                      <div>
                        <label className="mb-2 block text-sm font-medium text-slate-300">From Name</label>
                        <input
                          value={fromName}
                          onChange={(event) => setFromName(event.target.value)}
                          disabled={!canEdit}
                          className="w-full rounded-xl border border-slate-700 bg-slate-950 px-3 py-2.5 text-sm text-white outline-none transition focus:border-violet-500 disabled:cursor-not-allowed disabled:opacity-60"
                        />
                      </div>
                      <div>
                        <label className="mb-2 block text-sm font-medium text-slate-300">From Email</label>
                        <input
                          value={fromEmail}
                          onChange={(event) => setFromEmail(event.target.value)}
                          disabled={!canEdit}
                          className="w-full rounded-xl border border-slate-700 bg-slate-950 px-3 py-2.5 text-sm text-white outline-none transition focus:border-violet-500 disabled:cursor-not-allowed disabled:opacity-60"
                        />
                      </div>
                    </div>

                    <div>
                      <label className="mb-2 block text-sm font-medium text-slate-300">Reply-To</label>
                      <input
                        value={replyTo}
                        onChange={(event) => setReplyTo(event.target.value)}
                        disabled={!canEdit}
                        className="w-full rounded-xl border border-slate-700 bg-slate-950 px-3 py-2.5 text-sm text-white outline-none transition focus:border-violet-500 disabled:cursor-not-allowed disabled:opacity-60"
                      />
                    </div>
                  </div>
                ) : null}

                {channel === "sms" ? (
                  <div className="rounded-xl border border-slate-800 bg-slate-950/50 p-4 text-sm text-slate-300">
                    SMS gateway: <span className="font-medium text-white">{provider.toUpperCase()}</span>
                  </div>
                ) : null}

                {channel === "whatsapp" ? (
                  <div className="rounded-xl border border-slate-800 bg-slate-950/50 p-4 text-sm text-slate-300">
                    WhatsApp gateway: <span className="font-medium text-white">{provider.toUpperCase()}</span>
                  </div>
                ) : null}

                <div className="grid gap-4 md:grid-cols-2">
                  <div>
                    <label className="mb-2 block text-sm font-medium text-slate-300">Batch size</label>
                    <input
                      type="number"
                      min={20}
                      max={50}
                      value={batchSize}
                      onChange={(event) => setBatchSize(Number(event.target.value) || 25)}
                      disabled={!canEdit}
                      className="w-full rounded-xl border border-slate-700 bg-slate-950 px-3 py-2.5 text-sm text-white outline-none transition focus:border-violet-500 disabled:cursor-not-allowed disabled:opacity-60"
                    />
                  </div>
                  <div>
                    <label className="mb-2 block text-sm font-medium text-slate-300">Schedule</label>
                    <select
                      value={schedule}
                      onChange={(event) => setSchedule(event.target.value as "now" | "today" | "tomorrow" | "custom")}
                      disabled={!canEdit}
                      className="w-full rounded-xl border border-slate-700 bg-slate-950 px-3 py-2.5 text-sm text-white outline-none transition focus:border-violet-500 disabled:cursor-not-allowed disabled:opacity-60"
                    >
                      <option value="now">Send now</option>
                      <option value="today">Today</option>
                      <option value="tomorrow">Tomorrow</option>
                      <option value="custom">Custom date</option>
                    </select>
                  </div>
                </div>

                {schedule === "custom" ? (
                  <div className="rounded-xl border border-slate-800 bg-slate-950/50 p-4">
                    <label className="mb-2 block text-sm font-medium text-slate-300">Custom send date</label>
                    <input
                      type="date"
                      value={customScheduleDate}
                      min={new Date().toISOString().slice(0, 10)}
                      onChange={(event) => setCustomScheduleDate(event.target.value)}
                      disabled={!canEdit}
                      className="w-full rounded-xl border border-slate-700 bg-slate-950 px-3 py-2.5 text-sm text-white outline-none transition focus:border-violet-500 disabled:cursor-not-allowed disabled:opacity-60"
                    />
                  </div>
                ) : null}

                <div className="rounded-xl border border-slate-800 bg-slate-950/50 p-4">
                  <div className="mb-3 flex items-center justify-between">
                    <div className="text-sm font-medium text-slate-200">Customer scheduling data</div>
                    <div className="rounded-full border border-slate-700 bg-slate-900 px-2 py-1 text-[10px] uppercase tracking-[0.2em] text-slate-400">
                      {schedule === "now" ? "Send now" : schedule === "today" ? "Today" : schedule === "tomorrow" ? "Tomorrow" : `Custom: ${customScheduleDate}`}
                    </div>
                  </div>
                  <div className="flex flex-wrap gap-2">
                    {contacts.slice(0, 6).map((contact) => (
                      <button
                        key={contact.id}
                        type="button"
                        onClick={() => setScheduleCustomerField((previous) => previous || contact.name)}
                        className="rounded-full border border-slate-700 bg-slate-900 px-2.5 py-1 text-xs text-slate-200 transition hover:border-violet-400 hover:text-white"
                      >
                        {contact.name}
                      </button>
                    ))}
                  </div>
                  <div className="mt-3 text-xs text-slate-400">
                    Selected data field: <span className="text-slate-200">{scheduleCustomerField}</span>
                  </div>
                </div>

                <div className="grid gap-4 md:grid-cols-2">
                  <label className="flex items-center gap-3 rounded-xl border border-slate-800 bg-slate-950/50 px-3 py-2 text-sm text-slate-200">
                    <input type="checkbox" checked={retryFailed} onChange={() => setRetryFailed((value) => !value)} disabled={!canEdit} />
                    Retry failed messages
                  </label>
                  <label className="flex items-center gap-3 rounded-xl border border-slate-800 bg-slate-950/50 px-3 py-2 text-sm text-slate-200">
                    <input type="checkbox" checked={trackingEnabled} onChange={() => setTrackingEnabled((value) => !value)} disabled={!canEdit} />
                    Track opens/clicks
                  </label>
                </div>

                <div>
                  <label className="mb-2 block text-sm font-medium text-slate-300">Retry attempts</label>
                  <input
                    type="number"
                    min={1}
                    max={5}
                    value={maxRetries}
                    onChange={(event) => setMaxRetries(Number(event.target.value) || 2)}
                    disabled={!canEdit || !retryFailed}
                    className="w-full rounded-xl border border-slate-700 bg-slate-950 px-3 py-2.5 text-sm text-white outline-none transition focus:border-violet-500 disabled:cursor-not-allowed disabled:opacity-60"
                  />
                </div>

                <div className="rounded-xl border border-slate-800 bg-slate-950/50 p-4">
                  <div className="mb-2 flex items-center gap-2 text-sm font-medium text-slate-200">
                    <CheckCircleIcon className="h-4 w-4 text-emerald-400" />
                    Pre-send validation
                  </div>
                  <ul className="space-y-2 text-sm text-slate-300">
                    <li>• Email format validated before sending</li>
                    <li>• Provider credentials must be configured on the backend for live delivery</li>
                    <li>• Duplicates removed automatically</li>
                    <li>• Unsubscribed contacts excluded</li>
                    <li>• All message logs retained for audit and support</li>
                  </ul>
                </div>
              </div>
            </section>
            ) : null}

            {activeWorkflowTab === "preview" ? (
            <section className="glass-panel fade-in rounded-2xl border border-slate-800/80 bg-slate-900/80 p-5 shadow-[0_20px_50px_rgba(15,23,42,0.7)]">
              <div className="mb-4 flex items-center justify-between gap-3">
                <div className="flex items-center gap-3">
                  <div className="rounded-xl bg-cyan-500/15 p-2 text-cyan-200 ring-1 ring-cyan-400/30">
                    <ChartIcon className="h-5 w-5" />
                  </div>
                  <div>
                    <h2 className="text-lg font-semibold text-white">5. Preview & Test</h2>
                  </div>
                </div>
              </div>

              <div className="rounded-2xl border border-slate-800 bg-slate-950/60 p-4">
                <div className="mb-2 text-xs uppercase tracking-[0.2em] text-slate-400">Preview</div>
                <p className="mb-3 text-xs text-slate-400">
                  {selectedValidContacts.length
                    ? `Test recipient: ${previewContact.name} · ${channel === "email" ? previewContact.email : previewContact.phone}`
                    : "No valid recipients selected. Choose recipients in step 4."}
                </p>
                <div className="mb-3 rounded-xl border border-slate-800 bg-slate-900/60 p-3 text-sm text-slate-300">
                  <div className="mb-2 font-medium text-white">{subject}</div>
                  <div className="whitespace-pre-line text-slate-200">{previewBody}</div>
                  {(channel === "email" || channel === "whatsapp") && emailAttachment ? (
                    <div className="mt-3 border-t border-slate-800 pt-3 text-xs text-slate-300">
                      Attachment: <span className="text-white">{emailAttachment.name}</span> ({formatFileSize(emailAttachment.size)})
                    </div>
                  ) : null}
                </div>
                <div className="flex gap-3">
                  <button
                    type="button"
                    onClick={triggerTestSend}
                    disabled={isTestSending || selectedValidContacts.length === 0}
                    className="flex-1 rounded-xl border border-violet-500/40 bg-violet-500/10 px-3 py-2 text-sm font-medium text-violet-100 transition hover:bg-violet-500/20 disabled:cursor-not-allowed disabled:opacity-70"
                  >
                    {isTestSending ? "Sending test..." : provider === "mailtrap" && channel === "email" ? "Send test email" : "Test send"}
                  </button>
                  <button
                    type="button"
                    disabled={isSending || selectedValidContacts.length === 0}
                    onClick={triggerSend}
                    className="flex-1 rounded-xl bg-violet-600 px-3 py-2 text-sm font-semibold text-white transition hover:bg-violet-500 disabled:cursor-not-allowed disabled:opacity-70"
                  >
                    {isSending ? "Sending..." : "Send campaign"}
                  </button>
                </div>
                {testSendFeedback ? (
                  <p className="mt-3 text-xs text-slate-300" role="status" aria-live="polite">
                    {testSendFeedback}
                  </p>
                ) : null}
                {provider === "mailtrap" && channel === "email" ? (
                  <p className="mt-2 text-xs text-slate-500">
                    Sends only to the preview contact. Configure server-side sender and recipient allowlists before testing.
                  </p>
                ) : null}
              </div>
            </section>
            ) : null}

          {activeWorkflowTab === "dashboard" ? (
        <section className="glass-panel fade-in mt-8 rounded-2xl border border-slate-800/80 bg-slate-900/80 p-5 shadow-[0_20px_50px_rgba(15,23,42,0.7)]">
          <div className="mb-5 flex flex-col gap-3 md:flex-row md:items-center md:justify-between">
            <div className="flex items-center gap-3">
              <div className="rounded-xl bg-indigo-500/15 p-2 text-indigo-200 ring-1 ring-indigo-400/30">
                <ClockIcon className="h-5 w-5" />
              </div>
              <div>
                <h2 className="text-lg font-semibold text-white">6. Launch & Track</h2>
              </div>
            </div>
            <div className="flex items-center gap-3">
              <div className="text-sm text-slate-400">Progress</div>
              <div className="w-44 overflow-hidden rounded-full bg-slate-800">
                <div className="h-2.5 rounded-full bg-gradient-to-r from-violet-500 via-fuchsia-500 to-cyan-400" style={{ width: `${progress}%` }} />
              </div>
              <div className="text-sm text-slate-200">{progress}%</div>
            </div>
          </div>

          <div className="mb-5 flex gap-2 rounded-xl border border-slate-800 bg-slate-950/60 p-1">
            <button
              type="button"
              onClick={() => setActiveDataTab("summary")}
              className={`rounded-lg px-3 py-2 text-sm transition ${activeDataTab === "summary" ? "bg-violet-500 text-white" : "text-slate-300 hover:text-white"}`}
            >
              Summary
            </button>
            <button
              type="button"
              onClick={() => setActiveDataTab("rows")}
              className={`rounded-lg px-3 py-2 text-sm transition ${activeDataTab === "rows" ? "bg-violet-500 text-white" : "text-slate-300 hover:text-white"}`}
            >
              Uploaded data preview
            </button>
          </div>

          {activeDataTab === "summary" ? (
            <>
             <div className="mb-6 grid gap-4 md:grid-cols-6">
               {[
                 { label: "Sent", value: dashboardStats.sent, tone: "text-violet-300" },
                 { label: "Delivered", value: dashboardStats.delivered, tone: "text-emerald-300" },
                 { label: "Opened", value: dashboardStats.opened, tone: "text-cyan-300" },
                 { label: "Clicked", value: dashboardStats.clicked, tone: "text-sky-300" },
                 { label: "Bounced", value: dashboardStats.bounced, tone: "text-rose-300" },
                 { label: "Unsubscribed", value: dashboardStats.unsubscribed, tone: "text-amber-300" },
               ].map((stat) => (
                 <div key={stat.label} className="rounded-2xl border border-slate-800 bg-slate-950/60 p-4 shadow-lg shadow-slate-950/40">
                   <p className="text-xs uppercase tracking-[0.2em] text-slate-400">{stat.label}</p>
                   <p className={`mt-3 text-2xl font-bold ${stat.tone}`}>{stat.value}</p>
                 </div>
               ))}
             </div>

             <div className="mb-4 grid gap-4 lg:grid-cols-[1.4fr_0.6fr]">
               <div className="rounded-2xl border border-slate-800 bg-slate-950/60 p-4">
                 <div className="mb-3 text-xs uppercase tracking-[0.2em] text-slate-400">Live activity</div>
                 <div className="space-y-3">
                   {activityFeed.map((item) => (
                     <div key={item.name} className="flex items-center justify-between gap-3 rounded-xl border border-slate-800 bg-slate-900/60 px-3 py-2">
                       <div>
                         <div className="text-sm font-medium text-white">{item.name}</div>
                         <div className="text-xs text-slate-400">{item.note}</div>
                       </div>
                       <div className={`h-2.5 w-2.5 rounded-full ${item.tone.replace("text-", "bg-")}`} />
                     </div>
                   ))}
                 </div>
               </div>

               <div className="rounded-2xl border border-slate-800 bg-slate-950/60 p-4">
                 <div className="mb-3 text-xs uppercase tracking-[0.2em] text-slate-400">Delivery health</div>
                 <div className="space-y-3">
                   <div>
                     <div className="mb-1 flex justify-between text-xs text-slate-300"><span>Success rate</span><span>92%</span></div>
                     <div className="h-2 rounded-full bg-slate-800"><div className="h-2 w-[92%] rounded-full bg-gradient-to-r from-emerald-400 to-cyan-400" /></div>
                   </div>
                   <div>
                     <div className="mb-1 flex justify-between text-xs text-slate-300"><span>Open rate</span><span>67%</span></div>
                     <div className="h-2 rounded-full bg-slate-800"><div className="h-2 w-[67%] rounded-full bg-gradient-to-r from-violet-400 to-fuchsia-400" /></div>
                   </div>
                   <div>
                     <div className="mb-1 flex justify-between text-xs text-slate-300"><span>Click rate</span><span>41%</span></div>
                     <div className="h-2 rounded-full bg-slate-800"><div className="h-2 w-[41%] rounded-full bg-gradient-to-r from-sky-400 to-indigo-400" /></div>
                   </div>
                 </div>
               </div>
             </div>

             <div className="mb-4 flex justify-end">
               <button
                 type="button"
                 onClick={() => downloadCsv(resultRows)}
                 className="rounded-xl border border-slate-700 bg-slate-800 px-3 py-2 text-sm text-slate-100 transition hover:border-slate-500 hover:bg-slate-700"
               >
                 Download CSV report
               </button>
             </div>

             <div className="overflow-hidden rounded-2xl border border-slate-800">
               <div className="overflow-x-auto">
                 <table className="min-w-full text-left text-sm">
                   <thead className="bg-slate-950/80 text-slate-300">
                     <tr>
                       <th className="px-4 py-3 font-medium">Name</th>
                       <th className="px-4 py-3 font-medium">Recipient</th>
                       <th className="px-4 py-3 font-medium">Channel</th>
                       <th className="px-4 py-3 font-medium">Status</th>
                       <th className="px-4 py-3 font-medium">Reason</th>
                     </tr>
                   </thead>
                   <tbody className="divide-y divide-slate-800 bg-slate-900/70 text-slate-100">
                     {resultRows.length === 0 ? (
                       <tr>
                         <td colSpan={5} className="px-4 py-8 text-center text-slate-400">
                           No campaign activity yet. Upload contacts and send a test or live campaign.
                         </td>
                       </tr>
                     ) : (
                       resultRows.map((row) => (
                         <tr key={row.id}>
                           <td className="px-4 py-3">{row.name}</td>
                           <td className="px-4 py-3 text-slate-300">{row.recipient}</td>
                           <td className="px-4 py-3 capitalize text-slate-300">{row.channel}</td>
                           <td className="px-4 py-3">
                             <span className="rounded-full border border-slate-700 bg-slate-800 px-2 py-1 text-xs uppercase tracking-wide text-slate-200">
                               {row.status}
                             </span>
                           </td>
                           <td className="px-4 py-3 text-slate-300">{row.reason}</td>
                         </tr>
                       ))
                     )}
                   </tbody>
                 </table>
               </div>
             </div>
           </>
         ) : (
           <div className="rounded-2xl border border-slate-800 bg-slate-950/60 p-4">
             <div className="mb-4 text-xs uppercase tracking-[0.2em] text-slate-400">Uploaded data preview</div>
             <div className="overflow-x-auto">
               <table className="min-w-full text-left text-sm">
                 <thead className="bg-slate-950/80 text-slate-300">
                   <tr>
                     <th className="px-4 py-3 font-medium">Name</th>
                     <th className="px-4 py-3 font-medium">Email</th>
                     <th className="px-4 py-3 font-medium">Mobile</th>
                     <th className="px-4 py-3 font-medium">Status</th>
                   </tr>
                 </thead>
                 <tbody className="divide-y divide-slate-800 bg-slate-900/70 text-slate-100">
                   {contacts.map((contact) => (
                     <tr key={contact.id}>
                       <td className="px-4 py-3">{contact.name}</td>
                       <td className="px-4 py-3 text-slate-300">{contact.email}</td>
                       <td className="px-4 py-3 text-slate-300">{contact.phone}</td>
                       <td className="px-4 py-3">
                         <span className="rounded-full border border-slate-700 bg-slate-800 px-2 py-1 text-[10px] uppercase tracking-wide text-slate-200">
                           {contact.status}
                         </span>
                       </td>
                     </tr>
                   ))}
                 </tbody>
               </table>
             </div>
           </div>
         )}
        </section>
          ) : null}
          </div>
        </div>

        <footer className="mt-10 overflow-hidden rounded-3xl border border-slate-800/80 bg-slate-900/75 shadow-[0_20px_60px_rgba(2,6,23,0.35)]">
          <div className="grid gap-8 p-6 sm:grid-cols-2 sm:p-8 lg:grid-cols-[1.4fr_1fr_1fr_1fr]">
            <div>
              <a href="#campaign-workflow" className="text-lg font-black tracking-tight text-white transition hover:text-violet-200">
                Msg<span className="text-violet-300">Flow</span>
              </a>
              <p className="mt-3 max-w-xs text-sm leading-6 text-slate-400">
                Thoughtful tools for creating, sending, and tracking customer communications.
              </p>
            </div>
            <div>
              <h2 className="text-xs font-semibold uppercase tracking-[0.2em] text-slate-300">Explore</h2>
              <a href="#campaign-workflow" className="mt-3 block text-sm text-slate-400 transition hover:text-white">Campaign workspace</a>
              <a href="#campaign-workflow" className="mt-2 block text-sm text-slate-400 transition hover:text-white">Features & reports</a>
            </div>
            <div>
              <h2 className="text-xs font-semibold uppercase tracking-[0.2em] text-slate-300">Get in touch</h2>
              <a href="mailto:feedback@MsgFlow.world?subject=MsgFlow%20feedback" className="mt-3 block text-sm text-slate-400 transition hover:text-white">Feedback</a>
              <a href="mailto:support@msgflow.world?subject=MsgFlow%20support%20request" className="mt-2 block text-sm text-slate-400 transition hover:text-white">Support</a>
              <a href="mailto:contactus@msgflow.world?subject=MsgFlow%20contact" className="mt-2 block text-sm text-slate-400 transition hover:text-white">Contact us</a>
              <a href="mailto:careers@msgflow.world?subject=Careers%20at%20MsgFlow" className="mt-2 block text-sm text-slate-400 transition hover:text-white">Careers</a>
            </div>
            <div>
              <h2 className="text-xs font-semibold uppercase tracking-[0.2em] text-slate-300">Built for better outreach</h2>
              <p className="mt-3 text-sm leading-6 text-slate-400">Validate contacts, prepare personalized messages, and keep campaign activity organized.</p>
            </div>
          </div>
          <div className="border-t border-slate-800/80 bg-slate-950/40 px-5 py-5 sm:px-8">
            <div className="footer-flash mx-auto max-w-5xl rounded-2xl border border-violet-400/20 bg-gradient-to-r from-violet-500/10 via-slate-900/80 to-cyan-500/10 p-4 sm:p-5" aria-live="polite" aria-atomic="true">
              <div className="flex items-center justify-between gap-3">
                <p className="text-[10px] font-semibold uppercase tracking-[0.22em] text-violet-200">MsgFlow spotlight</p>
                <span className="text-xs tabular-nums text-slate-500">{footerFlashIndex + 1} / {footerFlashMessages.length}</span>
              </div>
              <div key={footerFlashIndex} className="footer-flash-message mt-3 min-h-12 text-left text-sm font-medium leading-6 text-white sm:text-base">
                {footerFlashMessages[footerFlashIndex]}
              </div>
              <div className="mt-3 flex items-center justify-between gap-3">
                <div className="flex gap-1" aria-hidden="true">
                  {footerFlashMessages.map((_, index) => (
                    <span
                      key={index}
                      className={`h-1 rounded-full transition-all duration-500 ${
                        index === footerFlashIndex ? "w-6 bg-violet-300" : "w-1.5 bg-slate-700"
                      }`}
                    />
                  ))}
                </div>
                <div className="flex gap-2">
                  <button
                    type="button"
                    aria-label="Previous message"
                    onClick={() => setFooterFlashIndex((index) => (index - 1 + footerFlashMessages.length) % footerFlashMessages.length)}
                    className="rounded-lg border border-slate-700 px-2.5 py-1 text-xs text-slate-300 transition hover:border-violet-400 hover:text-white"
                  >
                    Previous
                  </button>
                  <button
                    type="button"
                    aria-label="Next message"
                    onClick={() => setFooterFlashIndex((index) => (index + 1) % footerFlashMessages.length)}
                    className="rounded-lg border border-slate-700 px-2.5 py-1 text-xs text-slate-300 transition hover:border-violet-400 hover:text-white"
                  >
                    Next
                  </button>
                </div>
              </div>
            </div>
            <p className="mx-auto mt-3 max-w-5xl text-left text-[11px] leading-5 text-slate-500">
              Promotional message examples only. This demo simulates campaign sending; gateway failover, delivery guarantees, cost optimization, compliance automation, live analytics, and automatic campaigns require production integrations and are not active here.
            </p>
          </div>
          <div className="flex flex-wrap items-center justify-center gap-x-6 gap-y-2 border-t border-slate-800/80 bg-slate-950/30 px-6 py-4 text-xs sm:text-sm">
            <a href="mailto:support@msgflow.world" className="text-slate-400 transition hover:text-violet-200">
              Support: support@msgflow.world
            </a>
            <a href="mailto:contactus@msgflow.world" className="text-slate-400 transition hover:text-violet-200">
              Contact: contactus@msgflow.world
            </a>
            <a href="mailto:careers@msgflow.world" className="text-slate-400 transition hover:text-violet-200">
              Careers: careers@msgflow.world
            </a>
          </div>
          <div className="flex flex-col gap-2 border-t border-slate-800 px-6 py-4 text-xs text-slate-500 sm:flex-row sm:items-center sm:justify-between sm:px-8">
            <span>© {new Date().getFullYear()} MsgFlow. All rights reserved.</span>
            <span>Made for clear, connected customer communication.</span>
          </div>
        </footer>
      </div>
    </div>
  );
}

export default App;
