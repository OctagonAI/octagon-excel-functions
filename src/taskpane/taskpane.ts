/**
 * Octagon Excel Add-in Taskpane
 * This file handles the UI and interaction for the Octagon Excel Add-in taskpane
 */

import { OctagonApiService } from "../api";
import { checkRequiredApiSupport, detectIE } from "../utils/browserSupport";
import Logger from "../utils/logger";
import { AGENTS_EXAMPLES_FRAGMENT } from "./examples";

// Track the application state
const octagonApi = new OctagonApiService();

// Initialize the taskpane when Office is ready
Office.onReady(async () => {
  try {
    Logger.info("Office is ready - initializing Octagon taskpane");

    // Add event listeners to UI elements
    setupEventListeners();

    // Check authentication status and show the appropriate view
    await checkAuthentication();

    // Check browser compatibility
    if (detectIE()) {
      showBrowserWarning(
        "Internet Explorer is not fully supported. For the best experience, please use Microsoft Edge, Chrome, or another modern browser."
      );
    }

    // Check if all required APIs are supported
    const apiSupportIssues = checkRequiredApiSupport();
    if (apiSupportIssues.length > 0) {
      showApiSupportWarning(apiSupportIssues);
    }

    Logger.info("Taskpane initialized");
  } catch (error) {
    Logger.error("Error during taskpane initialization:", error);
  }
});

/**
 * Set up event listeners for UI interactions
 */
function setupEventListeners() {
  // Submit button
  const submitButton = document.getElementById("submit-button");
  if (submitButton) {
    submitButton.addEventListener("click", handleSubmit);
  }

  // Test Connection button
  const testConnectionButton = document.getElementById("test-connection-button");
  if (testConnectionButton) {
    testConnectionButton.addEventListener("click", handleTestConnection);
  }

  // Back to Login button
  const backToLoginButton = document.getElementById("back-to-login-button");
  if (backToLoginButton) {
    backToLoginButton.addEventListener("click", handleLogout);
  }

  // Enter key in API key input
  const apiKeyInput = document.getElementById("api-key-input") as HTMLInputElement;
  if (apiKeyInput) {
    apiKeyInput.addEventListener("keydown", (event) => {
      if (event.key === "Enter") {
        handleSubmit();
      }
    });

    // Clear the authentication message and update the button states when the user types or backspaces
    apiKeyInput.addEventListener("input", () => {
      clearAuthError();
      updateButtonStates();
    });
  }

  // Set initial button states
  updateButtonStates();
}

/**
 * Update button states based on API key input
 */
function updateButtonStates() {
  const apiKeyInput = document.getElementById("api-key-input") as HTMLInputElement;
  const submitButton = document.getElementById("submit-button");
  const testConnectionButton = document.getElementById("test-connection-button");

  const hasValue = apiKeyInput?.value?.trim().length > 0;

  // Enable/disable buttons based on whether there's input
  if (submitButton) {
    if (hasValue) {
      submitButton.removeAttribute("disabled");
    } else {
      submitButton.setAttribute("disabled", "true");
    }
  }

  if (testConnectionButton) {
    if (hasValue) {
      testConnectionButton.removeAttribute("disabled");
    } else {
      testConnectionButton.setAttribute("disabled", "true");
    }
  }
}

/**
 * Show the authentication view
 */
function showAuthView() {
  const authView = document.getElementById("auth-view");
  const agentsView = document.getElementById("agents-view");

  if (authView && agentsView) {
    authView.style.display = "block";
    authView.classList.add("fade-in");
    agentsView.style.display = "none";
  }
}

/**
 * Check the status of the API key and update the UI accordingly
 */
async function checkAuthentication() {
  const isAuthenticated = await octagonApi.isAuthenticated();

  if (isAuthenticated) {
    // If the api service is authenticated, show the agents view
    Logger.info("Authentication successful");

    // Automatically proceed to the Main Menu after a brief delay
    setTimeout(() => showAgentsView(), 500);
  } else {
    // If the api service is not authenticated, show the authentication view
    showAuthView();
  }
}

/**
 * Handle the Submit button click
 */
async function handleSubmit() {
  try {
    // Get the API key from the input field
    const apiKeyInput = document.getElementById("api-key-input") as HTMLInputElement;
    const apiKey = apiKeyInput?.value?.trim();

    if (!apiKey) {
      showAuthError("Please enter an API Key.");
      return;
    }

    // TODO: Test connection before saving
    // Set and persist the API key
    await octagonApi.setApiKey(apiKey);

    // Success - show success message and enable continue button
    clearAuthError();
    showAuthSuccess("API Key saved successfully!");

    // Automatically proceed to the Main Menu after a brief delay
    // This gives the user a chance to see the success message first
    setTimeout(() => {
      // Clear the input field before showing the agents view
      apiKeyInput.value = "";
      updateButtonStates();

      showAgentsView();
    }, 800);
  } catch (error) {
    // Show error message
    showAuthError("An error occurred while saving the API key. Please try again.");
    Logger.error("Error saving API key:", error);
  }
}

/**
 * Handle the Test Connection button click
 */
async function handleTestConnection() {
  try {
    // Get the API key from the input field (if visible)
    const apiKeyInput = document.getElementById("api-key-input") as HTMLInputElement;
    let apiKey = apiKeyInput?.value?.trim();

    // If input is not visible or empty, raise an error immediately
    if (!apiKeyInput || apiKeyInput.style.display === "none" || !apiKey) {
      showAuthError("Please enter an API Key.");
      return;
    }

    // Show loading state
    clearAuthError();
    toggleAuthLoadingState(true);

    // Test the connection
    const isConnected = await octagonApi.testConnection(apiKey);

    // Hide loading state
    toggleAuthLoadingState(false);

    if (isConnected) {
      // Success - show success message and enable continue button
      showAuthSuccess("Connection successful! API Key verified.");
    } else {
      // Failed - show error message
      showAuthError("Invalid API key. Please check and try again.");
    }
  } catch (error) {
    // Hide loading state
    toggleAuthLoadingState(false);

    // Show error message
    showAuthError("An error occurred. Please try again.");
    Logger.error("Authentication error:", error);
  }
}

/**
 * Handle the Back to Login button click
 */
async function handleLogout() {
  // Clear the API key
  await octagonApi.clearApiKey();
  Logger.info("API Key cleared");

  // Redirect to the login page after a brief delay
  setTimeout(() => showAuthView(), 500);
}

/**
 * Show authentication error message
 * @param message Error message to display
 */
function showAuthError(message: string) {
  const errorElement = document.getElementById("auth-error");
  if (errorElement) {
    errorElement.textContent = message;
    errorElement.style.display = "block";
  }
}

/**
 * Show authentication success message (temporarily)
 * @param message Success message to display
 */
function showAuthSuccess(message: string) {
  const errorElement = document.getElementById("auth-error");
  if (errorElement) {
    errorElement.textContent = message;
    errorElement.style.color = "green";
    errorElement.style.display = "block";
  }
}

/**
 * Clear authentication error message
 */
function clearAuthError() {
  const errorElement = document.getElementById("auth-error");
  if (errorElement) {
    errorElement.textContent = "";
    errorElement.style.display = "none";
    errorElement.style.color = "#a80000"; // Reset to error color
  }
}

/**
 * Toggle the loading state during authentication
 * @param isLoading Whether to show or hide the loading state
 */
function toggleAuthLoadingState(isLoading: boolean) {
  const submitButton = document.getElementById("submit-button");
  const testConnectionButton = document.getElementById("test-connection-button");
  const spinner = document.getElementById("auth-spinner");

  if (spinner) {
    if (isLoading) {
      if (submitButton) submitButton.setAttribute("disabled", "true");
      if (testConnectionButton) testConnectionButton.setAttribute("disabled", "true");
      spinner.style.display = "block";
    } else {
      if (submitButton) submitButton.removeAttribute("disabled");
      if (testConnectionButton) testConnectionButton.removeAttribute("disabled");
      spinner.style.display = "none";
    }
  }
}

/**
 * Show the agents view and populate it with agent data
 */
function showAgentsView() {
  const authView = document.getElementById("auth-view");
  const agentsView = document.getElementById("agents-view");

  if (authView && agentsView) {
    authView.style.display = "none";
    agentsView.style.display = "block";
    agentsView.classList.add("fade-in");

    // Hide the old back-to-auth button in the header (if it exists)
    const oldBackButton = document.getElementById("back-to-auth-button-header");
    if (oldBackButton) {
      oldBackButton.style.display = "none";
    }

    // Populate the agents list
    populateAgentsList();
  }
}

function populateAgentsList() {
  const agentCard = document.getElementById("agent-card");
  console.info("agentCard", agentCard);
  if (agentCard) {
    agentCard.appendChild(AGENTS_EXAMPLES_FRAGMENT);
  }
}

/**
 * Shows a browser compatibility warning to the user
 */
function showBrowserWarning(message: string) {
  const warningDiv = document.createElement("div");
  warningDiv.className = "browser-warning";
  warningDiv.innerHTML = `
    <div class="warning-icon"><i class="ms-Icon ms-Icon--Warning"></i></div>
    <div class="warning-message">${message}</div>
  `;

  // Insert at the top of the body or in a specific container
  const container = document.querySelector(".content-container") || document.body;
  container.insertBefore(warningDiv, container.firstChild);
}

/**
 * Shows API support warnings to the user
 */
function showApiSupportWarning(issues: string[]) {
  const warningDiv = document.querySelector(".api-support-warning");
  if (!warningDiv) return;

  (warningDiv as HTMLElement).style.display = "inline-block";

  if (issues.length > 0) {
    const warningMessage = warningDiv.firstElementChild;
    const issueList = document.createElement("ul");
    issues.forEach((issue) => {
      const issueItem = document.createElement("li");
      issueItem.textContent = issue;
      issueList.appendChild(issueItem);
    });
    warningMessage.appendChild(issueList);
  }
}
