/**
 * Octagon API Agents Definitions
 * These interfaces define the available agents in the Octagon API
 */

import { AgentInfo, AgentCategory } from './types';

export const OCTAGON_AGENTS: AgentInfo[] = [  
  // Deep Research Agents
  {
    id: 'octagon-deep-research-agent',
    displayName: 'Deep Research Agent',
    excelFormulaName: 'OCTAGON.DEEP_RESEARCH_AGENT',
    description: 'A comprehensive agent that can utilize multiple sources for deep research analysis.',
    category: AgentCategory.DeepResearch,
    examplePrompt: '=OCTAGON.DEEP_RESEARCH_AGENT("Research the financial impact of Apple privacy changes on digital advertising companies revenue and margins")'
  },
  {
    id: 'octagon-scraper-agent',
    displayName: 'Scraper Agent',
    excelFormulaName: 'OCTAGON.SCRAPER_AGENT',
    description: 'Specialized agent for financial data extraction from investor websites.',
    category: AgentCategory.DeepResearch,
    examplePrompt: '=OCTAGON.SCRAPER_AGENT("Extract all data fields from zillow.com/san-francisco-ca/ max_pages:2, country:us")'
  },
  // Market Intelligence Agents
  {
    id: 'octagon-agent',
    displayName: 'Octagon Agent',
    excelFormulaName: 'OCTAGON.OCTAGON_AGENT',
    description: 'Public and Private market intelligence agent that optimally routes requests to appropriate specialized agents',
    category: AgentCategory.MarketIntelligence,
    usageExamples: [
      { 
        topic: 'Financial Metrics Analysis',
        prompt: '=OCTAGON.OCTAGON_AGENT("Retrieve year-over-year growth in key income-statement items for AAPL, limited to 5 records and filtered by period FY.")'
      },
      { 
        topic: 'SEC Filing Analysis',
        prompt: '=OCTAGON.OCTAGON_AGENT("Analyze the latest 10-K filing for AAPL and extract key financial metrics and risk factors.")'
      },
      { 
        topic: 'Stock Performance',
        prompt: '=OCTAGON.OCTAGON_AGENT("Retrieve the daily closing prices for AAPL over the last 30 days.")'
      },
      { 
        topic: 'Earnings Call Insights',
        prompt: '=OCTAGON.OCTAGON_AGENT("Analyze AAPL\'s latest earnings call transcript and extract key insights about future guidance.")'
      },
      { 
        topic: 'Company Overview',
        prompt: '=OCTAGON.OCTAGON_AGENT("Provide a comprehensive overview of Stripe, including its business model and key metrics.")'
      },
      { 
        topic: 'Funding History',
        prompt: '=OCTAGON.OCTAGON_AGENT("Retrieve the funding history for Stripe, including all rounds and investors.")'
      },
      { 
        topic: 'M&A Activity',
        prompt: '=OCTAGON.OCTAGON_AGENT("List all M&A transactions involving Stripe in the last 2 years.")'
      },
      { 
        topic: 'Investor Profile',
        prompt: '=OCTAGON.OCTAGON_AGENT("Provide a detailed profile of Sequoia Capital\'s investment strategy and portfolio.")'
      },
      { 
        topic: 'Debt Analysis',
        prompt: '=OCTAGON.OCTAGON_AGENT("Analyze Stripe\'s debt financing history and current debt structure.")'
      },
      { 
        topic: 'Institutional Holdings',
        prompt: '=OCTAGON.OCTAGON_AGENT("Retrieve the most recent Form 13F and related filings submitted by institutional investors.")'
      }
    ]
  },
];
