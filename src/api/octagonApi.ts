/**
 * Octagon API Service Defintions
 * This service handles all communication with the Octagon API using OpenAI-compatible endpoints.
 * It provides methods to query agents using the chat/completions pattern.
 */

import Logger from "../utils/logger";
import { getTextFormat, parseTextFormat } from "./format";
import { AgentRequest, AgentResponse, AgentType, OutputFormat } from "./types";

// API Configuration
const DEFAULT_API_URL = "https://api-gateway.octagonagents.com";

// Storage key constants (used with OfficeRuntime.storage)
const API_KEY_STORAGE_NAME = "octagon_api_key";

/**
 * OctagonApiService class handles all API interactions
 */
export class OctagonApiService {
  private apiKey: string | null = null;
  private apiUrl: string;

  constructor(apiUrl: string = DEFAULT_API_URL) {
    this.apiUrl = apiUrl;
    // Don't automatically initialize in constructor to avoid race conditions
    // Instead, require explicit initialization after Office.onReady
  }

  /**
   * Set the API key for authentication
   */
  public async setApiKey(apiKey: string): Promise<void> {
    const trimmedKey = apiKey.trim() || null;
    this.apiKey = trimmedKey;

    if (!trimmedKey) {
      Logger.warn("Attempted to set an empty API key; clearing stored key instead");
      await this.clearApiKey();
      return;
    }

    // Save the API key to OfficeRuntime.storage
    try {
      await OfficeRuntime.storage.setItem(API_KEY_STORAGE_NAME, trimmedKey);
      Logger.debug("API key saved to OfficeRuntime.storage");
    } catch (error) {
      Logger.error("Failed to save API key to OfficeRuntime.storage", error);
    }
  }

  /**
   * Call an Octagon agent with a prompt
   *
   * @param agentId - The ID of the agent to call
   * @param prompt - The prompt to send to the agent
   * @returns A promise that resolves to the agent's response
   */
  public async callAgent(
    model: string,
    prompt: string,
    format: string
  ): Promise<Array<Array<string | number>>> {
    Logger.debug(`callAgent invoked with model: ${model}, prompt: ${prompt.substring(0, 50)}...`);

    // Validate the prompt
    if (!prompt || prompt.trim() === "") {
      throw new Error("Please provide a valid prompt");
    }

    // Ensure add-in has been authenticated
    if (!(await this.isAuthenticated())) {
      Logger.error("API request failed: Not authenticated");
      throw new Error("Not authenticated. Please set your API key first.");
    }

    const response = await this.createResponse({
      model,
      input: prompt,
      text: getTextFormat(format),
    });

    // Parse the content of the response as a JSON object if it is a table or single cell
    return parseTextFormat(response.content || "No response content", format as OutputFormat);
  }

  /**
   * Load API key from storage (OfficeRuntime.storage)
   */
  private async loadApiKey(): Promise<void> {
    const storedKey = await this.getStoredApiKey();

    if (storedKey) {
      this.apiKey = storedKey;
      return;
    }

    // If we get here, no API key was found
    this.apiKey = null;
  }

  /**
   * Returns true if the API key is set and not empty.
   * Always reloads from OfficeRuntime.storage to avoid stale cached keys across runtimes.
   */
  public async isAuthenticated(): Promise<boolean> {
    Logger.info("isAuthenticated invoked - synchronizing API key from storage");

    await this.loadApiKey();

    // Return true if the API key is set
    return !!this.apiKey;
  }

  /**
   * Clear stored API key from OfficeRuntime.storage
   */
  public async clearApiKey(): Promise<void> {
    this.apiKey = null;

    try {
      await OfficeRuntime.storage.removeItem(API_KEY_STORAGE_NAME);
      Logger.info("API key cleared from OfficeRuntime.storage");
    } catch (error) {
      Logger.error("Failed to clear API key from OfficeRuntime.storage", error);
    }
  }

  /**
   * Get the stored API key from OfficeRuntime.storage.
   */
  private async getStoredApiKey(): Promise<string | null> {
    try {
      const storedValue = await OfficeRuntime.storage.getItem(API_KEY_STORAGE_NAME);
      if (typeof storedValue === "string") {
        return storedValue;
      }
    } catch (error) {
      Logger.error("Failed to read API key from OfficeRuntime.storage", error);
    }
    return null;
  }

  /**
   * Generic method to make authenticated API requests
   * Handles both regular JSON responses and streamed responses
   */
  private async createResponse(data: AgentRequest, apiKey?: string): Promise<AgentResponse> {
    const requestUrl = new URL("/v1/responses", this.apiUrl);

    const token = apiKey ?? this.apiKey;
    if (!token) {
      throw new Error("No API key provided");
    }

    const headers = new Headers({
      "Content-Type": "application/json",
      Authorization: `Bearer ${token}`,
      "User-Agent": "Octagon-Excel-AddIn/1.2.0",
    });

    const response = await fetch(requestUrl, {
      method: "POST",
      headers,
      body: JSON.stringify(data),
    });

    if (!response.ok) {
      // Throw an error with HTTP status and message from the API
      await this.handleApiError(response);
    }

    // Bubble up any errors from the parseAgentResponse method to the caller
    return await this.parseAgentResponse(response);
  }

  /**
   * Parse the agent response from the Octagon API
   * @param response - The fetch response object with streamed data
   * @returns Parsed response data
   */
  private async parseAgentResponse(response: Response): Promise<AgentResponse> {
    try {
      // Get the reader for the stream
      const data = await response.json();

      // Ensure the response is in the expected format
      if (!data.output || !Array.isArray(data.output)) {
        Logger.warn("Unexpected response format: missing output array", data);
        return {
          content: "",
          id: data?.id,
          model: data?.model,
        } as AgentResponse;
      }

      // Create an empty content string, and append the text of each output message to the content
      let content = "";

      for (const output of data.output) {
        if (output.type !== "message") continue;
        // Ensure the content is in the expected format
        if (!output.content || !Array.isArray(output.content)) continue;

        // Add the text of each output message to the content
        for (const message of output.content) {
          if (message.type !== "output_text") continue;
          content += message.text;
        }
      }

      return {
        content,
        id: data.id,
        model: data.model,
      } as AgentResponse;
    } catch (error) {
      Logger.error("Error processing agent response:", error);
      throw new Error("Failed to process agent response");
    }
  }

  /**
   * Test API key validity with a simple request
   */
  public async testConnection(apiKey: string): Promise<boolean> {
    const request = {
      model: AgentType.OctagonAgent,
      input: "Test connection",
      max_tokens: 10,
    };

    try {
      // Test the connection with the provided API key
      await this.createResponse(request, apiKey);
      return true;
    } catch (error) {
      Logger.error("Test connection threw an exception", error);
      return false;
    }
  }

  private async handleApiError(response: Response): Promise<void> {
    // Parse error message from response
    let data: { detail: string } | undefined;
    try {
      data = await response.json();
    } catch (error) {
      // Fallback to unknown error
      data = { detail: "Unknown error" };
    }

    const status = response.status;
    const errorMessage = data?.detail ?? "Unknown error";
    Logger.error("Agent request failed", { status, message: errorMessage });
    throw new Error(`HTTP ${status}: ${errorMessage}`);
  }
}
