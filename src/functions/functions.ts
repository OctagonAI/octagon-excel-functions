/* global console, CustomFunctions, Office */

/**
 * Octagon Excel Add-in Custom Functions
 * This file defines custom functions that expose Octagon's AI agents as Excel formulas
 */

import { AgentType, octagonApi } from "../api";
import Logger from "../utils/logger";

/**
 * Initializes the Octagon API and registers custom functions
 * This is automatically called when the add-in is initialized with SharedRuntime
 */
Office.onReady(() => {
  try {
    Logger.debug("Office is ready - initializing Octagon functions");

    // Register functions namespaced under OCTAGON
    CustomFunctions.associate("OCTAGON.OCTAGON_AGENT", OCTAGON_AGENT);
  } catch (error) {
    Logger.error("Error during initialization:", error);
  }
});

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
export async function OCTAGON_AGENT(prompt: string): Promise<string> {
  try {
    return await octagonApi.callAgent(AgentType.OctagonAgent, prompt);
  } catch (error) {
    // Throw a custom error with the error message
    throw new CustomFunctions.Error(
      CustomFunctions.ErrorCode.notAvailable,
      error instanceof Error ? error.message : "An unexpected error occurred"
    );
  }
}
