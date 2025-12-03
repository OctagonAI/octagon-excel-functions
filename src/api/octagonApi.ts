/**
 * Octagon API Service Defintions
 * This service handles all communication with the Octagon API using OpenAI-compatible endpoints.
 * It provides methods to query agents using the chat/completions pattern.
 */

import Logger from "../utils/logger";
import { AgentRequest, AgentResponse, AgentType } from "./types";

// API Configuration
const DEFAULT_API_URL = "https://api-gateway.octagonagents.com/v1";

// SessionStorage key constants
const API_KEY_STORAGE_NAME = "octagon_api_key";
const CACHED_SESSION_SETTINGS_KEY = "CachedSessionSettings";

/**
 * OctagonApiService class handles all API interactions
 */
export class OctagonApiService {
  private apiKey: string | null = null;
  private apiUrl: string;
  private isOfficeInitialized: boolean = false;

  constructor(apiUrl: string = DEFAULT_API_URL) {
    this.apiUrl = apiUrl;
    // Don't automatically initialize in constructor to avoid race conditions
    // Instead, require explicit initialization after Office.onReady
    Logger.info(`OctagonApiService created with API URL: ${this.apiUrl}`);
  }

  /**
   * Initialize the service after Office is ready
   * This should be called after Office.onReady completes
   */
  public initialize(): void {
    this.isOfficeInitialized = true;
    this.loadApiKey();
    Logger.info(
      `OctagonApiService fully initialized. API key status: ${this.isAuthenticated() ? "Available" : "Not available"}`
    );
  }

  /**
   * Check if API key exists in SessionStorage
   * @returns boolean True if API key is found in SessionStorage
   */
  public checkForStoredApiKey(): boolean {
    try {
      // Check if we have a cached session settings with an API key
      const cachedSessionData = sessionStorage.getItem(CACHED_SESSION_SETTINGS_KEY);
      if (cachedSessionData) {
        const sessionData = JSON.parse(cachedSessionData);
        if (sessionData && sessionData.octagon_api_key) {
          Logger.info("Found API key in CachedSessionSettings");
          return true;
        }
      }

      // Also check our direct storage
      const directApiKey = sessionStorage.getItem(API_KEY_STORAGE_NAME);
      if (directApiKey) {
        Logger.info("Found API key in direct SessionStorage");
        return true;
      }
    } catch (error) {
      Logger.error("Error checking for stored API key", error);
    }

    return false;
  }

  /**
   * Set the API key for authentication
   */
  public setApiKey(apiKey: string): void {
    this.apiKey = apiKey;
    this.saveApiKey(apiKey);
  }

  /**
   * Save API key to SessionStorage
   */
  private saveApiKey(apiKey: string): void {
    try {
      // Save to SessionStorage
      sessionStorage.setItem(API_KEY_STORAGE_NAME, apiKey);

      // Also save to CachedSessionSettings for compatibility
      try {
        const cachedSessionData = sessionStorage.getItem(CACHED_SESSION_SETTINGS_KEY);
        if (cachedSessionData) {
          const sessionData = JSON.parse(cachedSessionData);
          sessionData.octagon_api_key = apiKey;
          sessionStorage.setItem(CACHED_SESSION_SETTINGS_KEY, JSON.stringify(sessionData));
        } else {
          // Create a new cached session settings object
          const newSessionData = { octagon_api_key: apiKey };
          sessionStorage.setItem(CACHED_SESSION_SETTINGS_KEY, JSON.stringify(newSessionData));
        }
      } catch (error) {
        Logger.warn("Failed to save to CachedSessionSettings", error);
      }

      Logger.info("API key saved to SessionStorage");
    } catch (error) {
      Logger.error("Failed to save API key to SessionStorage", error);
    }
  }

  /**
   * Call an Octagon agent with a prompt
   *
   * @param agentId - The ID of the agent to call
   * @param prompt - The prompt to send to the agent
   * @returns A promise that resolves to the agent's response
   */
  public async callAgent(model: string, prompt: string): Promise<string> {
    Logger.debug(`callAgent invoked with model: ${model}, prompt: ${prompt.substring(0, 50)}...`);

    // Validate the prompt
    if (!prompt || prompt.trim() === "") {
      throw new Error("Please provide a valid prompt");
    }

    // Ensure add-in has been authenticated
    if (!this.isAuthenticated()) {
      Logger.error("API request failed: Not authenticated");
      throw new Error("Not authenticated. Please set your API key first.");
    }

    const response = await this.createResponse({
      model,
      input: prompt,
    });
    return response.content || "No response content";
  }

  /**
   * Load API key from SessionStorage with a simplified approach
   */
  private loadApiKey(): void {
    Logger.info("Attempting to load API key from SessionStorage");

    try {
      // Try CachedSessionSettings first
      const cachedSessionData = sessionStorage.getItem(CACHED_SESSION_SETTINGS_KEY);
      if (cachedSessionData) {
        try {
          const sessionData = JSON.parse(cachedSessionData);
          if (sessionData && sessionData.octagon_api_key) {
            this.apiKey = sessionData.octagon_api_key;
            Logger.info("API key loaded from CachedSessionSettings");
            return;
          }
        } catch (parseError) {
          Logger.warn("Failed to parse CachedSessionSettings", parseError);
        }
      }

      // Fallback to direct storage
      const directApiKey = sessionStorage.getItem(API_KEY_STORAGE_NAME);
      if (directApiKey) {
        this.apiKey = directApiKey;
        Logger.info("API key loaded from direct SessionStorage");
        return;
      }
    } catch (error) {
      Logger.error("Failed to load API key from SessionStorage", error);
    }

    // If we get here, no API key was found
    Logger.info("No API key found in SessionStorage");
  }

  /**
   * Returns true if the API key is set and not empty. Loads the API key from SessionStorage if not set in current instance.
   */
  public isAuthenticated(): boolean {
    if (this.apiKey) {
      return true;
    }
    // Try to load the API key from SessionStorage
    this.loadApiKey();

    // Return true if the API key is set and not empty
    return !!this.apiKey && this.apiKey.trim() !== "";
  }

  /**
   * Clear stored API key from all SessionStorage locations
   */
  public clearApiKey(): void {
    this.apiKey = null;

    try {
      // Clear direct API key storage
      sessionStorage.removeItem(API_KEY_STORAGE_NAME);

      // Clear from CachedSessionSettings if it exists
      const cachedSessionData = sessionStorage.getItem(CACHED_SESSION_SETTINGS_KEY);
      if (cachedSessionData) {
        try {
          const sessionData = JSON.parse(cachedSessionData);
          if (sessionData) {
            delete sessionData.octagon_api_key;
            sessionStorage.setItem(CACHED_SESSION_SETTINGS_KEY, JSON.stringify(sessionData));
          }
        } catch (parseError) {
          Logger.warn("Failed to parse CachedSessionSettings for clearing", parseError);
        }
      }

      Logger.info("API key cleared from all SessionStorage locations");
    } catch (error) {
      Logger.error("Failed to clear API key from SessionStorage", error);
    }
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

    try {
      const headers = new Headers({
        "Content-Type": "application/json",
        Authorization: `Bearer ${token}`,
      });

      const response = await fetch(requestUrl, {
        method: "POST",
        headers,
        body: JSON.stringify(data),
      });

      if (!response.ok) {
        Logger.error(
          `Agent request failed: ${response.status} ${response.statusText}`,
          response.json()
        );
        throw new Error(`Agent request failed: ${response.status} ${response.statusText}`);
      }

      const streamedResponse = await this.handleStreamedResponse(response);
      return streamedResponse as AgentResponse;
    } catch (error) {
      throw new Error(
        `Failed to create agent response: ${error instanceof Error ? error.message : "Unknown error"}`
      );
    }
  }

  /**
   * Handle streamed responses from the Octagon API
   * @param response - The fetch response object with streamed data
   * @returns Parsed response data
   */
  private async handleStreamedResponse(response: Response): Promise<AgentResponse> {
    try {
      // Get the reader for the stream
      const reader = response.body?.getReader();
      if (!reader) {
        throw new Error("Response stream could not be read");
      }

      Logger.debug("Processing streamed response");

      // Variables to collect the response
      let fullText = "";
      let finalResponse = null;
      let parseErrorCount = 0; // Track number of parse errors

      // Process the stream
      while (true) {
        try {
          const { done, value } = await reader.read();
          if (done) break;

          // Convert the chunks to text
          const chunk = new TextDecoder().decode(value);
          Logger.debug(
            "Received stream chunk:",
            chunk.substring(0, 100) + (chunk.length > 100 ? "..." : "")
          );

          // Process each line that starts with "data: "
          const lines = chunk.split("\n");
          for (const line of lines) {
            if (line.startsWith("data: ")) {
              const jsonStr = line.substring(6); // Remove 'data: ' prefix

              // Check for the end of stream marker
              if (jsonStr.trim() === "[DONE]") {
                Logger.debug("Stream completed [DONE] marker received");
                continue;
              }

              try {
                const data = JSON.parse(jsonStr);
                Logger.debug("Parsed stream data type:", data.type);

                // Process based on the type of data
                if (data.type === "response.completed") {
                  // This is the final response with all data
                  finalResponse = data.response;
                  Logger.debug("Final response received");
                } else if (
                  data.type === "response.output_text.delta" ||
                  data.type === "response.output_text.done"
                ) {
                  // This is a text delta/update, append to our content
                  if (data.delta) {
                    fullText += data.delta;
                  } else if (data.text) {
                    fullText = data.text; // This is the complete text
                  }
                } else if (data.type === "response.content_part.done" && data.part?.text) {
                  // Content part with complete text
                  fullText = data.part.text;
                } else if (data.response && data.response.output) {
                  // Store the full response data for direct access
                  finalResponse = data.response;
                }
              } catch (parseError) {
                parseErrorCount++;
                Logger.warn(
                  `Error parsing stream JSON (${parseErrorCount}): ${parseError} - ${jsonStr.substring(0, 100)}`
                );

                // If we hit too many parse errors, we might be dealing with a bad stream
                if (parseErrorCount > 5) {
                  Logger.error("Too many JSON parse errors in stream, aborting");
                  throw new Error("Stream parsing failed: Too many JSON parse errors");
                }
              }
            }
          }
        } catch (streamError) {
          Logger.error("Error processing stream chunk:", streamError);
          throw streamError; // Rethrow to be caught by the outer try/catch
        }
      }

      // Ensure we have some content to return, even if it's empty
      const transformedResponse = {
        content: fullText || "",
        id: finalResponse?.id || "",
        model: finalResponse?.model || "",
        created: finalResponse?.created_at || Date.now(),
        // For compatibility with functions.ts, include the raw response
        output: finalResponse?.output || [],
        // Include full response data for more advanced processing if needed
        rawResponse: finalResponse,
      };

      Logger.info("Streamed response processed successfully");

      return transformedResponse as AgentResponse;
    } catch (error) {
      Logger.error("Error processing streamed response:", error);
      throw new Error(
        `Failed to process streamed response: ${error instanceof Error ? error.message : "Unknown error"}`
      );
    }
  }

  /**
   * Test API key validity with a simple request
   */
  public async testConnection(apiKey: string): Promise<boolean> {
    Logger.info("Testing API connection with stored key");

    const request = {
      model: AgentType.OctagonAgent,
      input: "Test connection",
      max_tokens: 10,
    };

    try {
      await this.createResponse(request, apiKey);
      return true;
    } catch (error) {
      Logger.error("Test connection threw an exception", error);
      return false;
    }
  }
}

// Export singleton instance
export const octagonApi = new OctagonApiService();
