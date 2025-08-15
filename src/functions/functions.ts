/* global console, CustomFunctions, Office */

/**
 * Octagon Excel Add-in Custom Functions
 * This file defines custom functions that expose Octagon's AI agents as Excel formulas
 */

import { octagonApiService } from '../api/octagonApi';
import Logger from '../utils/logger';

// Initialize the Octagon API client
const octagonApi = octagonApiService;

// Flag to track if we've initialized the API
let isApiInitialized = false;

/**
 * Initializes the Octagon API and registers custom functions
 * This is automatically called when the add-in is initialized with SharedRuntime
 */
Office.onReady(() => {
  try {
    Logger.info("Office is ready - initializing Octagon functions");
    // Initialize the API service
    octagonApi.initialize();
    isApiInitialized = true;
    Logger.info(`API initialization complete. Authentication status: ${octagonApi.isAuthenticated() ? 'Authenticated' : 'Not authenticated'}`);
    
    // Register functions with the Excel namespace (OCTAGON)
    registerCustomFunctions();
  } catch (error) {
    Logger.error("Error during initialization:", error);
  }
});

/**
 * Register all custom functions with Excel
 * This is called after Office is ready and the API is initialized
 */
function registerCustomFunctions() {
  try {    
    // Register all Octagon agents with the OCTAGON namespace
    CustomFunctions.associate("OCTAGON.OCTAGON_AGENT", OCTAGON_AGENT);
    CustomFunctions.associate("OCTAGON.DEEP_RESEARCH_AGENT", DEEP_RESEARCH_AGENT);
    CustomFunctions.associate("OCTAGON.SCRAPER_AGENT", SCRAPER_AGENT);
    
    Logger.info("Custom functions registered successfully");
  } catch (error) {
    Logger.error("Error registering custom functions:", error);
  }
}

/**
 * Map short agent IDs to their full model names
 */
const AGENT_ID_MAP: Record<string, string> = {
  'research': 'octagon-deep-research-agent',
  'scraper': 'octagon-scraper-agent',
  'octagon': 'octagon-agent'
};

/**
 * Base implementation for Octagon agent functions
 * @param agentId - The ID of the Octagon agent to use
 * @param prompt - The prompt to send to the agent
 * @param invocation - Optional streaming invocation object for cancellation support
 * @returns The agent's response
 */
async function callOctagonAgent(
  agentId: string, 
  prompt: string, 
  invocation?: CustomFunctions.StreamingInvocation<string>
): Promise<string> {
  try {
    // Validate the prompt
    if (!prompt || prompt.trim() === "") {
      return "Error: Please provide a valid prompt";
    }

    // Check if API is initialized
    if (!isApiInitialized) {
      return "Error: API not initialized. Please try again in a few moments.";
    }

    // Check if authenticated
    if (!octagonApi.isAuthenticated()) {
      return "Error: Not authenticated. Please enter your API key in the Login Screen.";
    }

    // Map agent ID to full model name
    const modelName = AGENT_ID_MAP[agentId.toLowerCase()] || agentId;

    // Call the agent
    const response = await octagonApi.callAgent(modelName, prompt);
    
    // Extract the text content from the response object
    return response.data.content || "No response content";
    
  } catch (error) {
    Logger.error(`Error calling Octagon agent (${agentId}):`, error);
    if (error instanceof Error) {
      return `Error: ${error.message}`;
    }
    
    // Register all Octagon agents with the OCTAGON namespace
    CustomFunctions.associate("OCTAGON_AGENT", OCTAGON_AGENT);
    CustomFunctions.associate("DEEP_RESEARCH_AGENT", DEEP_RESEARCH_AGENT);
    CustomFunctions.associate("SCRAPER_AGENT", SCRAPER_AGENT);
    
    return "Error: An unexpected error occurred";
  }
}

// ===============================================================
// Market Intelligence Agent
// ===============================================================

/**
 * Call the Market Intelligence agent that routes to appropriate specialized agents
 * @customfunction OCTAGON_AGENT
 * @param prompt The question or prompt for the Octagon agent
 * @helpUrl https://docs.octagonagents.com/guide/agents/octagon-agent.html
 * @returns A string containing the agent's response
 */
export function OCTAGON_AGENT(prompt: string): Promise<string> {
  return callOctagonAgent('octagon', prompt);
}

// ===============================================================
// Deep Research Agent
// ===============================================================

/**
 * Call the Research agent for deep, comprehensive research
 * @customfunction DEEP_RESEARCH_AGENT
 * @param prompt The question or prompt for the Research agent
 * @helpUrl https://docs.octagonagents.com/guide/agents/deep-research-agent.html
 * @returns Comprehensive research based on the prompt
 */
export function DEEP_RESEARCH_AGENT(prompt: string): Promise<string> {
  return callOctagonAgent('research', prompt);
}

// ===============================================================
// Scraper Agent
// ===============================================================

/**
 * Call the Scraper agent for web data extraction
 * @customfunction SCRAPER_AGENT
 * @param prompt The question or prompt for the Scraper agent
 * @helpUrl https://docs.octagonagents.com/guide/agents/scraper-agent.html
 * @returns Web data extraction based on the prompt
 */
export function SCRAPER_AGENT(prompt: string): Promise<string> {
  return callOctagonAgent('scraper', prompt);
}