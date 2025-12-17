/* global console, CustomFunctions, Office */

/**
 * Octagon Excel Add-in Custom Functions
 * This file defines custom functions that expose Octagon's AI agents as Excel formulas
 */

import { AgentType, OctagonApiService } from "../api";

const octagonApi = new OctagonApiService();

/**
 * Call the Market Intelligence agent that routes to appropriate specialized agents
 * @customfunction AGENT
 * @param prompt The question or prompt for the Octagon agent
 * @param format optional with possible values of 'table' (default), 'cell', or 'raw'.
 * @helpUrl https://docs.octagonagents.com/guide/agents/octagon-agent.html
 * @returns array of arrays of strings or numbers
 */
export async function OCTAGON_AGENT(
  prompt: string,
  format?: string
): Promise<Array<Array<string | number>>> {
  try {
    // Default to table format if no format is provided
    const agentFormat = format ? format.toLowerCase() : "table";
    return await octagonApi.callAgent(AgentType.OctagonAgent, prompt, agentFormat);
  } catch (error) {
    if (error instanceof CustomFunctions.Error) {
      throw error;
    }
    // Throw a custom error with the error message
    throw new CustomFunctions.Error(
      CustomFunctions.ErrorCode.notAvailable,
      error instanceof Error ? error.message : "An unexpected error occurred"
    );
  }
}
