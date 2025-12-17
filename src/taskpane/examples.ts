/**
 * Octagon API Agents Definitions
 * These interfaces define the available agents in Excel Add-in
 */

import { AgentType } from "../api/types";
import Logger from "../utils/logger";

interface AgentInfo {
  id: AgentType;
  displayName: string;
  excelFormulaName: string;
  description: string;
  category: AgentCategory;
  usageExamples?: UsageExample[];
}

interface UsageExample {
  topic: string;
  prompt: string;
}

enum AgentCategory {
  MarketIntelligence = "Market Intelligence",
}

const OCTAGON_AGENTS: AgentInfo[] = [
  {
    id: AgentType.OctagonAgent,
    displayName: "Octagon Agent",
    excelFormulaName: "OCTAGON.AGENT",
    description:
      "Public and Private market intelligence agent that optimally routes requests to appropriate specialized agents",
    category: AgentCategory.MarketIntelligence,
    usageExamples: [
      {
        topic: "Financial Metrics Analysis",
        prompt:
          '=OCTAGON.AGENT("Retrieve year-over-year growth in key income-statement items for AAPL, limited to 5 records and filtered by period FY.")',
      },
      {
        topic: "SEC Filing Analysis",
        prompt:
          '=OCTAGON.AGENT("Analyze the latest 10-K filing for AAPL and extract key financial metrics and risk factors.", "raw")',
      },
      {
        topic: "Stock Performance",
        prompt:
          '=OCTAGON.AGENT("Retrieve the daily closing prices for AAPL over the last 30 days.")',
      },
      {
        topic: "Earnings Call Insights",
        prompt:
          '=OCTAGON.AGENT("In one concise sentence, summarize AAPL\'s latest earnings call with a focus on revenue growth and future guidance.", "cell")',
      },
      {
        topic: "Company Overview",
        prompt:
          '=OCTAGON.AGENT("Provide a comprehensive overview of Stripe, including its business model and key metrics.", "raw")',
      },
      {
        topic: "Funding History",
        prompt:
          '=OCTAGON.AGENT("Create a table of Stripe\'s funding history with columns Round, Date, AmountUSD, LeadInvestor, OtherInvestors, and PostMoneyValuation.", "table")',
      },
      {
        topic: "M&A Activity",
        prompt: '=OCTAGON.AGENT("List all M&A transactions involving Stripe in the last 2 years.")',
      },
      {
        topic: "Investor Profile",
        prompt:
          '=OCTAGON.AGENT("Provide a detailed profile of Sequoia Capital\'s investment strategy and portfolio.", "raw")',
      },
      {
        topic: "Debt Analysis",
        prompt:
          '=OCTAGON.AGENT("Analyze Stripe\'s debt financing history and current debt structure.", "raw")',
      },
      {
        topic: "Institutional Holdings",
        prompt:
          '=OCTAGON.AGENT("Retrieve the most recent Form 13F and related filings submitted by institutional investors.", "raw")',
      },
    ],
  },
];

// Precompute agents by category once at module load time since the data is static
const AGENTS_BY_CATEGORY: Record<string, AgentInfo[]> = OCTAGON_AGENTS.reduce(
  (categories, agent) => {
    if (!categories[agent.category]) {
      categories[agent.category] = [];
    }
    categories[agent.category].push(agent);
    return categories;
  },
  {} as Record<string, AgentInfo[]>
);

/**
 * Populate the agents list with categories and agent cards.
 * This only builds the DOM once since the underlying data is static.
 */
function createAgentExamplesFragment(): DocumentFragment {
  const fragment = document.createDocumentFragment();

  // Create a list for the examples
  const examplesList = document.createElement("div");
  examplesList.className = "examples-list";
  fragment.appendChild(examplesList);

  // Add examples for each category
  for (const categoryAgents of Object.values(AGENTS_BY_CATEGORY)) {
    // Create agent cards for this category
    categoryAgents.forEach((agent) => {
      addAgentExamples(agent, examplesList);
    });
  }
  return fragment;
}

/**
 * Create an agent card element
 * @param agent Agent information
 * @returns HTMLElement The agent card
 */
function addAgentExamples(agent: AgentInfo, examplesList: HTMLElement) {
  // Add each example to the list
  agent.usageExamples.forEach((example) => {
    const exampleItem = document.createElement("div");
    exampleItem.className = "example-item";

    const topicElement = document.createElement("div");
    topicElement.className = "example-topic";
    topicElement.textContent = example.topic;
    exampleItem.appendChild(topicElement);

    // Create a container for the prompt to allow positioning the copy button
    const promptContainer = document.createElement("div");
    promptContainer.className = "example-prompt-container";
    promptContainer.style.position = "relative";

    const promptElement = document.createElement("div");
    promptElement.className = "example-prompt code";
    promptElement.textContent = example.prompt;
    promptContainer.appendChild(promptElement);

    // Add copy button
    const copyButton = document.createElement("button");
    copyButton.className = "copy-button";
    copyButton.title = "Copy example";
    copyButton.innerHTML = '<i class="ms-Icon ms-Icon--Copy"></i>';
    copyButton.onclick = (e) => {
      e.stopPropagation();
      navigator.clipboard
        .writeText(example.prompt)
        .then(() => {
          // Show success feedback
          copyButton.innerHTML = '<i class="ms-Icon ms-Icon--CheckMark copy-success"></i>';
          setTimeout(() => {
            copyButton.innerHTML = '<i class="ms-Icon ms-Icon--Copy"></i>';
          }, 1500);
        })
        .catch((err) => {
          Logger.error("Could not copy text: ", err);
        });
    };

    promptContainer.appendChild(copyButton);
    exampleItem.appendChild(promptContainer);

    examplesList.appendChild(exampleItem);
  });
}

export const AGENTS_EXAMPLES_FRAGMENT: DocumentFragment = createAgentExamplesFragment();
