import { config } from "./config.js";

const HAL_INTERPRETATION_SCHEMA = {
  type: "object",
  additionalProperties: false,
  properties: {
    intentType: {
      type: "string",
      enum: ["use_local", "task", "idea", "meeting", "reminder"],
    },
    title: { type: "string" },
    formatted: { type: "string" },
    person: { type: "string" },
    confidence: { type: "number" },
    rationale: { type: "string" },
  },
  required: ["intentType", "title", "formatted", "person", "confidence", "rationale"],
};

let openAiClientPromise = null;

export async function createHalInterpretation(payload) {
  const client = await getOpenAiClient();
  const response = await client.responses.create({
    model: config.openai.model,
    reasoning: { effort: "low" },
    text: {
      format: {
        type: "json_schema",
        name: "hal_interpretation",
        strict: true,
        schema: HAL_INTERPRETATION_SCHEMA,
      },
    },
    input: [
      {
        role: "system",
        content: [
          {
            type: "input_text",
            text: buildSystemPrompt(),
          },
        ],
      },
      {
        role: "user",
        content: [
          {
            type: "input_text",
            text: JSON.stringify(payload),
          },
        ],
      },
    ],
  });

  return JSON.parse(response.output_text);
}

async function getOpenAiClient() {
  if (openAiClientPromise) {
    return openAiClientPromise;
  }

  openAiClientPromise = (async () => {
    try {
      const { default: OpenAI } = await import("openai");
      return new OpenAI({ apiKey: config.openai.apiKey });
    } catch (error) {
      throw new Error(
        `HAL could not load the OpenAI SDK. Run npm install after updating package.json. ${error.message}`
      );
    }
  })();

  return openAiClientPromise;
}

function buildSystemPrompt() {
  return [
    "You are HAL's interpretation layer for a personal assistant app.",
    "Return only structured JSON that matches the schema.",
    "Your job is to improve classification, titles, and formatting while respecting HAL's local business rules.",
    "Choose intentType='use_local' when the local interpretation is already strong or when you are not confident.",
    "Prefer concise, human-readable titles that match how David would want to see them later.",
    "For task titles, remove command wording and leave just the action.",
    "For idea titles, create a short clear title based on the concept.",
    "For meeting notes, you may provide a cleaner title and a talking-point formatted version in formatted.",
    "For reminders, keep the title focused on the reminder text only. Do not invent dates or times.",
    "Use the provided preference memory and correction history as a personalization hint.",
  ].join(" ");
}
