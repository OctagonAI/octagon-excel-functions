/* global console, CustomFunctions, Office */

/**
 * Octagon Excel Add-in Custom Functions
 * This file defines custom functions that expose Octagon's AI agents as Excel formulas
 */

import { AgentType, OctagonApiService } from "../api";
import Logger from "../utils/logger";

const octagonApi = new OctagonApiService();

/**
 * Initializes the Octagon API and registers custom functions
 * This is automatically called when the add-in is initialized with SharedRuntime
 */
Office.onReady(async () => {
  try {
    Logger.debug("Office is ready - initializing Octagon functions");

    // Register functions namespaced under OCTAGON
    CustomFunctions.associate("OCTAGON.AGENT", OCTAGON_AGENT);
  } catch (error) {
    Logger.error("Error during initialization:", error);
  }
});

// ===============================================================
// Market Intelligence Agent
// ===============================================================

/**
 * Call the Market Intelligence agent that routes to appropriate specialized agents
 * @customfunction AGENT
 * @param prompt The question or prompt for the Octagon agent
 * @param format The format of the response, one of "raw", "table", or "cell". Defaults to "table".
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
