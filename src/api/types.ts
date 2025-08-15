/**
 * Octagon API TypeScript Interfaces
 * These interfaces define the contract between our Excel Add-in and the Octagon API.
 */

// ==================== COMMON API TYPES ====================

// Common API response structure
export interface ApiResponse<T> {
  success: boolean;
  data?: T;
  error?: string;
}

// Generic agent response base
export interface AgentResponse {
  query: string;
  timestamp: string;
  agent: AgentInfo | AgentListing;
}

// ==================== OPENAI COMPATIBLE TYPES ====================

// OpenAI-style message structure
export interface ChatMessage {
  role: 'system' | 'user' | 'assistant';
  content: string;
}

// OpenAI-style chat completion request
export interface StreamRequest {
  model: string;
  messages: ChatMessage[];
  temperature?: number;
  max_tokens?: number;
  stream?: boolean;
}

// OpenAI-style choice structure
export interface StreamChoice {
  index: number;
  message: ChatMessage;
  finish_reason: 'stop' | 'length' | 'function_call' | 'content_filter' | null;
}

// OpenAI-style usage statistics
export interface StreamUsage {
  prompt_tokens: number;
  completion_tokens: number;
  total_tokens: number;
}

// OpenAI-style responses endpoint response
export interface StreamResponse {
  id: string;
  object: 'responses';
  created: number;
  model: string;
  choices: StreamChoice[];
  usage?: StreamUsage;
  // Adding new fields for Octagon specific response structure
  content?: string;
  output?: any[];
  rawResponse?: any;
  status?: string;
  text?: any;
}

// ==================== AGENT TYPES ====================

export interface AgentInfo {
  id: string;
  displayName: string;
  excelFormulaName: string;
  description: string;
  category: AgentCategory;
  examplePrompt?: string;
  usageExamples?: UsageExample[];
}

export interface UsageExample {
  topic: string;
  prompt: string;
}

export enum AgentCategory {
  MarketIntelligence = 'Market Intelligence',
  DeepResearch = 'Deep Research'
}

export enum AgentListing {
  OctagonAgent = 'octagon-agent',
  ScraperAgent = 'octagon-scraper-agent',
  DeepResearchAgent = 'octagon-deep-research-agent'
}

// ==================== REQUEST/RESPONSE TYPES ====================

// Detailed agent response with content and structured data
export interface AgentFullResponse {
  content: string;
  // Additional fields that might be in the Octagon API response
  model?: string;
  created?: number;
  id?: string;
  [key: string]: any; // Allow for additional unknown properties
}

// ==================== AUTHENTICATION & MANAGEMENT TYPES ====================

// API Authentication types
export interface AuthCredentials {
  apiKey: string;
}

// API Key interface
export interface APIKey {
  id?: string | null;
  created_at: string;
  updated_at?: string | null;
  name: string;
  truncated_secret: string;
}

// ==================== UTILITY TYPES ====================

// Type aliases for common data structures
export type Option<T = string> = { value: T; label: string };
export type Maybe<T> = T | null | undefined;

// ==================== EXCEL INTEGRATION TYPES ====================

// Excel table formatting options
export interface ExcelTableOptions {
  hasHeaders?: boolean;
  showTotals?: boolean;
  tableName?: string;
  tableStyle?: string;
}

// Excel insertion result
export interface ExcelInsertionResult {
  success: boolean;
  rowsInserted: number;
  columnsInserted: number;
  range?: string;
  error?: string;
}

// ==================== ERROR HANDLING TYPES ====================

// Standardized error response
export interface ErrorResponse {
  error: {
    code: string;
    message: string;
    type: 'api_error' | 'rate_limit_exceeded' | 'invalid_request' | 'authentication_error';
    param?: string;
  };
}

// ==================== EXPORTED CONSTANTS ====================

// Default responses endpoint completion options
export const DEFAULT_CHAT_OPTIONS = {
  temperature: 0.7,
  max_tokens: 1500,
  stream: true
} as const;