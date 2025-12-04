/**
 * Octagon Excel Add-in Taskpane
 * This file handles the UI and interaction for the Octagon Excel Add-in taskpane
 */

import { OCTAGON_AGENTS, octagonApi } from "../api";
import { checkRequiredApiSupport, detectIE } from "../utils/browserSupport";
import Logger from "../utils/logger";

// Track the application state
let isAuthenticated = false;

// Initialize the taskpane when Office is ready
Office.onReady(async (info) => {
  try {
    Logger.info("Office is ready - initializing Octagon taskpane");

    // Initialize the API service
    octagonApi.initialize();

    // Add event listeners to UI elements
    setupEventListeners();

    // Show auth view first (this ensures UI is visible)
    checkApiKeyStatus();

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
function checkApiKeyStatus() {
  const hasStoredKey = octagonApi.checkForStoredApiKey();

  if (hasStoredKey) {
    // If we have a stored key, hide the status message but log it
    Logger.info("API Key detected from previous sessions");

    // Set authenticated state
    isAuthenticated = true;

    // Automatically proceed to the Main Menu after a brief delay
    setTimeout(() => showAgentsView(), 500);
  } else {
    // No API key detected, show the authentication view
    showAuthView();
  }
}

/**
 * Handle the Submit button click
 */
function handleSubmit() {
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
    octagonApi.setApiKey(apiKey);

    // Success - show success message and enable continue button
    clearAuthError();
    showAuthSuccess("API Key saved successfully!");
    isAuthenticated = true;

    // Automatically proceed to the Main Menu after a brief delay
    // This gives the user a chance to see the success message first
    setTimeout(() => showAgentsView(), 800);
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
function handleLogout() {
  // Clear the API key
  octagonApi.clearApiKey();
  isAuthenticated = false;
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

/**
 * Populate the agents list with categories and agent cards
 */
function populateAgentsList() {
  const container = document.getElementById("agent-categories-container");

  if (!container) return;

  // Clear the container
  container.innerHTML = "";

  // Group agents by category
  const agentsByCategory = groupAgentsByCategory();

  // Create a section for each category
  for (const [category, agents] of Object.entries(agentsByCategory)) {
    // Create category container
    const categoryElement = document.createElement("div");
    categoryElement.className = "agent-category";

    // Create category title
    const titleElement = document.createElement("h3");
    titleElement.className = "category-title";
    titleElement.textContent = category;
    categoryElement.appendChild(titleElement);

    // Create agent cards for this category
    agents.forEach((agent) => {
      const agentCard = createAgentCard(agent);
      categoryElement.appendChild(agentCard);
    });

    // Add the category to the container
    container.appendChild(categoryElement);
  }
}

/**
 * Group agents by their category
 * @returns Record<string, typeof OCTAGON_AGENTS[0][]>
 */
function groupAgentsByCategory() {
  const categories: Record<string, (typeof OCTAGON_AGENTS)[0][]> = {};

  OCTAGON_AGENTS.forEach((agent) => {
    if (!categories[agent.category]) {
      categories[agent.category] = [];
    }
    categories[agent.category].push(agent);
  });

  return categories;
}

/**
 * Create an agent card element
 * @param agent Agent information
 * @returns HTMLElement The agent card
 */
function createAgentCard(agent: (typeof OCTAGON_AGENTS)[0]): HTMLElement {
  const card = document.createElement("div");
  card.className = "agent-card";

  // Agent title
  const title = document.createElement("h4");
  title.className = "agent-title";
  title.textContent = agent.displayName;
  card.appendChild(title);

  // Agent description
  const description = document.createElement("p");
  description.className = "agent-description";
  description.textContent = agent.description;
  card.appendChild(description);

  // Agent metadata
  const meta = document.createElement("div");
  meta.className = "agent-meta";

  // Formula name
  const formula = document.createElement("div");
  formula.innerHTML = `<strong>Excel Formula:</strong> <span class="formula-name">${agent.excelFormulaName}("your prompt")</span>`;
  meta.appendChild(formula);

  // Example prompt
  if (agent.examplePrompt) {
    const example = document.createElement("div");
    example.innerHTML = `<strong>Example:</strong>`;

    // Create a container for the prompt to allow positioning the copy button
    const promptContainer = document.createElement("div");
    promptContainer.style.position = "relative";
    promptContainer.style.display = "inline-block";
    promptContainer.style.width = "100%";

    const promptElement = document.createElement("span");
    promptElement.className = "example-prompt";
    promptElement.textContent = agent.examplePrompt;
    promptContainer.appendChild(promptElement);

    // Add copy button for the example prompt
    const copyButton = document.createElement("button");
    copyButton.className = "copy-button";
    copyButton.title = "Copy example";
    copyButton.innerHTML = '<i class="ms-Icon ms-Icon--Copy"></i>';
    copyButton.onclick = (e) => {
      e.stopPropagation();
      navigator.clipboard
        .writeText(agent.examplePrompt)
        .then(() => {
          // Show success feedback
          copyButton.innerHTML = '<i class="ms-Icon ms-Icon--CheckMark copy-success"></i>';
          setTimeout(() => {
            copyButton.innerHTML = '<i class="ms-Icon ms-Icon--Copy"></i>';
          }, 1500);
        })
        .catch((err) => {
          console.error("Could not copy text: ", err);
        });
    };

    promptContainer.appendChild(copyButton);
    example.appendChild(promptContainer);
    meta.appendChild(example);
  }

  card.appendChild(meta);

  // Usage examples section (if available)
  if (agent.usageExamples && agent.usageExamples.length > 0) {
    const usageSection = document.createElement("div");
    usageSection.className = "usage-examples-section";

    // Usage examples heading
    const usageHeading = document.createElement("h5");
    usageHeading.className = "usage-heading";
    usageHeading.textContent = "Usage Examples:";
    usageSection.appendChild(usageHeading);

    // Create a list for the examples
    const examplesList = document.createElement("div");
    examplesList.className = "examples-list";

    // Add each example to the list
    agent.usageExamples.forEach((example) => {
      const exampleItem = document.createElement("div");
      exampleItem.className = "example-item";

      const topicElement = document.createElement("div");
      topicElement.className = "example-topic";
      topicElement.textContent = example.topic;
      exampleItem.appendChild(topicElement);

      // Create a container for the prompt to allow positioning the copy button
      const promptContainer = document.createElement("div");
      promptContainer.className = "example-prompt-container";
      promptContainer.style.position = "relative";

      const promptElement = document.createElement("div");
      promptElement.className = "example-prompt code";
      promptElement.textContent = example.prompt;
      promptContainer.appendChild(promptElement);

      // Add copy button
      const copyButton = document.createElement("button");
      copyButton.className = "copy-button";
      copyButton.title = "Copy example";
      copyButton.innerHTML = '<i class="ms-Icon ms-Icon--Copy"></i>';
      copyButton.onclick = (e) => {
        e.stopPropagation();
        navigator.clipboard
          .writeText(example.prompt)
          .then(() => {
            // Show success feedback
            copyButton.innerHTML = '<i class="ms-Icon ms-Icon--CheckMark copy-success"></i>';
            setTimeout(() => {
              copyButton.innerHTML = '<i class="ms-Icon ms-Icon--Copy"></i>';
            }, 1500);
          })
          .catch((err) => {
            console.error("Could not copy text: ", err);
          });
      };

      promptContainer.appendChild(copyButton);
      exampleItem.appendChild(promptContainer);

      examplesList.appendChild(exampleItem);
    });

    usageSection.appendChild(examplesList);
    card.appendChild(usageSection);
  }

  return card;
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

  console.log("showing api support warning", issues);

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
