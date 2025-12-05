/**
 * Octagon API Agents Definitions
 * These interfaces define the available agents in Excel Add-in
 */

import { AgentType } from "./types";

interface AgentInfo {
  id: AgentType;
  displayName: string;
  excelFormulaName: string;
  description: string;
  category: AgentCategory;
  examplePrompt?: string;
  usageExamples?: UsageExample[];
}

interface UsageExample {
  topic: string;
  prompt: string;
}

enum AgentCategory {
  MarketIntelligence = "Market Intelligence",
}

export const OCTAGON_AGENTS: AgentInfo[] = [
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
          '=OCTAGON.AGENT("Analyze the latest 10-K filing for AAPL and extract key financial metrics and risk factors.")',
      },
      {
        topic: "Stock Performance",
        prompt:
          '=OCTAGON.AGENT("Retrieve the daily closing prices for AAPL over the last 30 days.")',
      },
      {
        topic: "Earnings Call Insights",
        prompt:
          '=OCTAGON.AGENT("Analyze AAPL\'s latest earnings call transcript and extract key insights about future guidance.")',
      },
      {
        topic: "Company Overview",
        prompt:
          '=OCTAGON.AGENT("Provide a comprehensive overview of Stripe, including its business model and key metrics.")',
      },
      {
        topic: "Funding History",
        prompt:
          '=OCTAGON.AGENT("Retrieve the funding history for Stripe, including all rounds and investors.")',
      },
      {
        topic: "M&A Activity",
        prompt: '=OCTAGON.AGENT("List all M&A transactions involving Stripe in the last 2 years.")',
      },
      {
        topic: "Investor Profile",
        prompt:
          '=OCTAGON.AGENT("Provide a detailed profile of Sequoia Capital\'s investment strategy and portfolio.")',
      },
      {
        topic: "Debt Analysis",
        prompt:
          '=OCTAGON.AGENT("Analyze Stripe\'s debt financing history and current debt structure.")',
      },
      {
        topic: "Institutional Holdings",
        prompt:
          '=OCTAGON.AGENT("Retrieve the most recent Form 13F and related filings submitted by institutional investors.")',
      },
    ],
  },
];
