/**
 * Octagon Excel Add-in Taskpane
 * This file handles the UI and interaction for the Octagon Excel Add-in taskpane
 */

import { octagonApiService } from '../api/octagonApi';
import { OCTAGON_AGENTS } from '../api/agents';
import Logger from '../utils/logger';
import { detectIE, checkRequiredApiSupport } from '../utils/browserSupport';

// Initialize API service
const apiService = octagonApiService;

// Track the application state
let isAuthenticated = false;
let isCheckingApiKey = false;
let hasAutoRedirected = false;

// Initialize the taskpane when Office is ready
Office.onReady(async (info) => {
  try {
    Logger.info("Office is ready - initializing Octagon taskpane");
    
    // Initialize the API service
    apiService.initialize();
    
    // Hide the sideload message
    document.getElementById("sideload-msg").style.display = "none";
    
    // Add event listeners to UI elements
    setupEventListeners();
    
    // Show auth view first (this ensures UI is visible)
    showAuthView();
        
    // Check browser compatibility
    if (detectIE()) {
      showBrowserWarning('Internet Explorer is not fully supported. For the best experience, please use Microsoft Edge, Chrome, or another modern browser.');
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
  
  // Clear API Keys button
  const clearApiKeysButton = document.getElementById("clear-api-keys-button");
  if (clearApiKeysButton) {
    clearApiKeysButton.addEventListener("click", handleClearApiKeys);
  }
  
  // Continue to Main Menu button
  const continueToMenuButton = document.getElementById("continue-to-menu-button");
  if (continueToMenuButton) {
    continueToMenuButton.addEventListener("click", showAgentsView);
  }
  
  // Note: We no longer need to set up the back button listener here
  // since we add it dynamically when creating the agents view
  
  // Enter key in API key input
  const apiKeyInput = document.getElementById("api-key-input") as HTMLInputElement;
  if (apiKeyInput) {
    apiKeyInput.addEventListener("keypress", (event) => {
      if (event.key === "Enter") {
        handleSubmit();
      }
    });
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
  
  // Only check API key status if this function wasn't called during initialization
  // or as part of another API key status check
  if (!isCheckingApiKey) {
    isCheckingApiKey = true;
    
    // Uses setTimeout to break the potential call stack chain
    // This ensures UI renders before we continue with more logic
    setTimeout(() => {
      checkApiKeyStatus();
      isCheckingApiKey = false;
    }, 0);
  }
}

/**
 * Check the status of the API key and update the UI accordingly
 */
function checkApiKeyStatus() {
  const hasStoredKey = apiService.checkForStoredApiKey();
  const statusMessageContainer = document.getElementById("api-key-status-message");
  const continueToMenuContainer = document.getElementById("continue-to-menu-container");
  const apiKeyInputContainer = document.getElementById("api-key-input-container");
  const clearApiKeysButton = document.getElementById("clear-api-keys-button");
  
  // Get the authentication header and instruction text
  const authHeader = document.querySelector(".auth-container h3.ms-font-l");
  const authInstructions = document.querySelector(".auth-container p.ms-font-m");
  
  if (hasStoredKey) {
    // If we have a stored key, hide the status message but log it
    Logger.info("API Key detected from previous sessions");
    
    if (statusMessageContainer) {
      statusMessageContainer.style.display = "none";
    }
    
    // Hide the authentication header and instructions
    if (authHeader) {
      authHeader.textContent = "API Key Detected";
    }
    
    if (authInstructions) {
      authInstructions.textContent = "Your API key has been detected. You can proceed to the main menu or verify your connection.";
    }
    
    // Show the continue button and style it at the bottom center
    if (continueToMenuContainer) {
      continueToMenuContainer.style.display = "flex";
      continueToMenuContainer.innerHTML = "";
      
      const continueButton = document.createElement("button");
      continueButton.id = "continue-to-menu-button";
      continueButton.className = "ms-Button ms-Button--icon continue-to-menu-button";
      continueButton.innerHTML = `
        <span class="ms-Button-label">Main Menu</span>
        <span class="ms-Button-icon">
          <i class="ms-Icon ms-Icon--ChevronRight"></i>
        </span>
      `;
      continueButton.addEventListener("click", showAgentsView);
      
      continueToMenuContainer.appendChild(continueButton);
    }
    
    // Hide the API key input container
    if (apiKeyInputContainer) {
      apiKeyInputContainer.style.display = "none";
    }
    
    // Show Clear API Keys button since there's a stored key
    if (clearApiKeysButton) {
      clearApiKeysButton.style.display = "inline-block";
    }
    
    // Rename the action buttons
    updateActionButtonLabels(true);
    
    // Set authenticated state
    isAuthenticated = true;
    
    // Only auto-redirect if this is the first time we're checking the API key
    // and we haven't auto-redirected before
    if (!hasAutoRedirected) {
      hasAutoRedirected = true;
      // Automatically proceed to the Main Menu after a brief delay
      setTimeout(() => {
        showAgentsView();
      }, 1000);
    }
  } else {
    // If no key is stored, show a message
    if (statusMessageContainer) {
      statusMessageContainer.textContent = "No API Key detected. Please provide an API Key to continue.";
      statusMessageContainer.style.display = "block";
    }
    
    // Show the default authentication header and instructions
    if (authHeader) {
      authHeader.textContent = "Authentication Required";
    }
    
    if (authInstructions) {
      authInstructions.textContent = "Please enter your Octagon API key to continue:";
    }
    
    if (continueToMenuContainer) {
      continueToMenuContainer.style.display = "none";
    }
    
    // Show the API key input container
    if (apiKeyInputContainer) {
      apiKeyInputContainer.style.display = "block";
    }
    
    // Hide Clear API Keys button since there's no stored key
    if (clearApiKeysButton) {
      clearApiKeysButton.style.display = "none";
    }
    
    // Update button labels for new users
    updateActionButtonLabels(false);
  }
}

/**
 * Update action button labels based on authentication state
 * @param isAuthenticated Whether the user is authenticated
 */
function updateActionButtonLabels(isAuthenticated: boolean) {
  const submitButton = document.getElementById("submit-button");
  const testConnectionButton = document.getElementById("test-connection-button");
  
  if (submitButton) {
    if (isAuthenticated) {
      submitButton.style.display = "none";
    } else {
      submitButton.style.display = "inline-block";
    }
  }
  
  if (testConnectionButton) {
    // Test Connection button label stays the same regardless of auth state
    testConnectionButton.innerHTML = '<span class="ms-Button-label">Test Connection</span>';
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
    
    // Set and persist the API key
    apiService.setApiKey(apiKey);
    
    // Success - show success message and enable continue button
    clearAuthError();
    showAuthSuccess("API Key saved successfully!");
    
    // Get the authentication header and instruction text
    const authHeader = document.querySelector(".auth-container h3.ms-font-l");
    const authInstructions = document.querySelector(".auth-container p.ms-font-m");
    
    // Update text for authenticated users
    if (authHeader) {
      authHeader.textContent = "API Key Detected";
    }
    
    if (authInstructions) {
      authInstructions.textContent = "Your API key has been saved. You can proceed to the main menu or verify your connection.";
    }
    
    // Update the UI to authenticated state
    const apiKeyInputContainer = document.getElementById("api-key-input-container");
    if (apiKeyInputContainer) {
      apiKeyInputContainer.style.display = "none";
    }
    
    // Show the continue button
    const continueToMenuContainer = document.getElementById("continue-to-menu-container");
    if (continueToMenuContainer) {
      continueToMenuContainer.style.display = "flex";
      continueToMenuContainer.innerHTML = "";
      
      const continueButton = document.createElement("button");
      continueButton.id = "continue-to-menu-button";
      continueButton.className = "ms-Button ms-Button--icon continue-to-menu-button";
      continueButton.innerHTML = `
        <span class="ms-Button-label">Main Menu</span>
        <span class="ms-Button-icon">
          <i class="ms-Icon ms-Icon--ChevronRight"></i>
        </span>
      `;
      continueButton.addEventListener("click", showAgentsView);
      
      continueToMenuContainer.appendChild(continueButton);
    }
    
    // Hide status message
    const statusMessageContainer = document.getElementById("api-key-status-message");
    if (statusMessageContainer) {
      statusMessageContainer.style.display = "none";
    }
    
    // Show Clear API Keys button
    const clearApiKeysButton = document.getElementById("clear-api-keys-button");
    if (clearApiKeysButton) {
      clearApiKeysButton.style.display = "inline-block";
    }
    
    updateActionButtonLabels(true);
    isAuthenticated = true;
    
    // Automatically proceed to the Main Menu after a brief delay
    // This gives the user a chance to see the success message first
    setTimeout(() => {
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
    
    // If input is not visible or empty, we need to use the stored key
    if (!apiKeyInput || apiKeyInput.style.display === "none" || !apiKey) {
      // Check if a stored API Key exists
      const hasStoredKey = apiService.checkForStoredApiKey();
      
      // If we have a stored key but don't have access to it directly,
      // we can just proceed with testing the connection using whatever
      // the API service already has
      if (!hasStoredKey) {
        showAuthError("No API Key available. Please enter an API Key.");
        return;
      }
      
      // We'll just test with the key that's already set in the service
    } else {
      // We have a new key from the input, set it before testing
      apiService.setApiKey(apiKey);
    }
    
    // Show loading state
    toggleAuthLoadingState(true);
    
    // Test the connection
    const isValid = await testApiConnection();
    
    // Hide loading state
    toggleAuthLoadingState(false);
    
    if (isValid) {
      // Success - show success message and enable continue button
      clearAuthError();
      showAuthSuccess("Connection successful! API Key verified.");
      
      // Get the authentication header and instruction text
      const authHeader = document.querySelector(".auth-container h3.ms-font-l");
      const authInstructions = document.querySelector(".auth-container p.ms-font-m");
      
      // Update text for authenticated users
      if (authHeader) {
        authHeader.textContent = "API Key Detected";
      }
      
      if (authInstructions) {
        authInstructions.textContent = "Your API key has been detected. You can proceed to the main menu or verify your connection.";
      }
      
      // Update the UI to authenticated state
      const apiKeyInputContainer = document.getElementById("api-key-input-container");
      if (apiKeyInputContainer) {
        apiKeyInputContainer.style.display = "none";
      }
      
      // Show the continue button
      const continueToMenuContainer = document.getElementById("continue-to-menu-container");
      if (continueToMenuContainer) {
        continueToMenuContainer.style.display = "flex";
        continueToMenuContainer.innerHTML = "";
        
        const continueButton = document.createElement("button");
        continueButton.id = "continue-to-menu-button";
        continueButton.className = "ms-Button ms-Button--icon continue-to-menu-button";
        continueButton.innerHTML = `
          <span class="ms-Button-label">Main Menu</span>
          <span class="ms-Button-icon">
            <i class="ms-Icon ms-Icon--ChevronRight"></i>
          </span>
        `;
        continueButton.addEventListener("click", showAgentsView);
        
        continueToMenuContainer.appendChild(continueButton);
      }
      
      // Hide status message
      const statusMessageContainer = document.getElementById("api-key-status-message");
      if (statusMessageContainer) {
        statusMessageContainer.style.display = "none";
      }
      
      // Show Clear API Keys button
      const clearApiKeysButton = document.getElementById("clear-api-keys-button");
      if (clearApiKeysButton) {
        clearApiKeysButton.style.display = "inline-block";
      }
      
      updateActionButtonLabels(true);
      isAuthenticated = true;
      
      // Automatically proceed to the Main Menu after a brief delay
      // This gives the user a chance to see the success message first
      setTimeout(() => {
        showAgentsView();
      }, 1000);

    } else {
      
      // Failed - show error message
      showAuthError("Invalid API key. Please check and try again.");
      apiService.clearApiKey();
      
      // Reset UI text
      const authHeader = document.querySelector(".auth-container h3.ms-font-l");
      const authInstructions = document.querySelector(".auth-container p.ms-font-m");
      
      if (authHeader) {
        authHeader.textContent = "Authentication Required";
      }
      
      if (authInstructions) {
        authInstructions.textContent = "Please enter your Octagon API key to continue:";
      }
      
      // Show the API key input
      const apiKeyInputContainer = document.getElementById("api-key-input-container");
      if (apiKeyInputContainer) {
        apiKeyInputContainer.style.display = "block";
      }
      
      updateActionButtonLabels(false);
      isAuthenticated = false;
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
 * Handle the Clear API Keys button click
 */
function handleClearApiKeys() {
  apiService.clearApiKey();
  isAuthenticated = false;
  hasAutoRedirected = false; // Reset the auto-redirect flag
  
  // Reset the API key input
  const apiKeyInput = document.getElementById("api-key-input") as HTMLInputElement;
  if (apiKeyInput) {
    apiKeyInput.value = "";
  }
  
  // Update the authentication header and instruction text
  const authHeader = document.querySelector(".auth-container h3.ms-font-l");
  const authInstructions = document.querySelector(".auth-container p.ms-font-m");
  
  if (authHeader) {
    authHeader.textContent = "Authentication Required";
  }
  
  if (authInstructions) {
    authInstructions.textContent = "Please enter your Octagon API key to continue:";
  }
  
  // Show the API key input container
  const apiKeyInputContainer = document.getElementById("api-key-input-container");
  if (apiKeyInputContainer) {
    apiKeyInputContainer.style.display = "block";
  }
  
  // Hide the continue button
  const continueToMenuContainer = document.getElementById("continue-to-menu-container");
  if (continueToMenuContainer) {
    continueToMenuContainer.style.display = "none";
  }
  
  // Hide the Clear API Keys button since there's no stored key anymore
  const clearApiKeysButton = document.getElementById("clear-api-keys-button");
  if (clearApiKeysButton) {
    clearApiKeysButton.style.display = "none";
  }
  
  // Update status message
  const statusMessageContainer = document.getElementById("api-key-status-message");
  if (statusMessageContainer) {
    statusMessageContainer.textContent = "All stored API Keys have been cleared. Please provide a new API Key to continue.";
    statusMessageContainer.style.display = "block";
  }
  
  // Update button labels
  updateActionButtonLabels(false);
  
  Logger.info("All API Keys cleared");
}

/**
 * Test the API connection
 * @returns Promise<boolean> True if connection is valid, false otherwise
 */
async function testApiConnection(): Promise<boolean> {
  try {
    Logger.info("Testing API connection");
    const response = await apiService.testConnection();
    
    Logger.info(`API test response: ${JSON.stringify(response)}`);
    
    return response.success;
  } catch (error) {
    Logger.error("API test connection error:", error);
    return false;
  }
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
    
    // Hide the message after 3 seconds
    setTimeout(() => {
      clearAuthError();
    }, 3000);
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
    agents.forEach(agent => {
      const agentCard = createAgentCard(agent);
      categoryElement.appendChild(agentCard);
    });
    
    // Add the category to the container
    container.appendChild(categoryElement);
  }
  
  // Add a small back button at the bottom of the container
  const backButtonContainer = document.createElement("div");
  backButtonContainer.className = "back-button-container";
  
  const backButton = document.createElement("button");
  backButton.id = "back-to-auth-button";
  backButton.className = "ms-Button ms-Button--icon back-to-login-button";
  backButton.innerHTML = `
    <span class="ms-Button-icon">
      <i class="ms-Icon ms-Icon--ChevronLeft"></i>
    </span>
    <span class="ms-Button-label">Back to Log In</span>
  `;
  backButton.addEventListener("click", showAuthView);
  
  backButtonContainer.appendChild(backButton);
  container.appendChild(backButtonContainer);
}

/**
 * Group agents by their category
 * @returns Record<string, typeof OCTAGON_AGENTS[0][]>
 */
function groupAgentsByCategory() {
  const categories: Record<string, typeof OCTAGON_AGENTS[0][]> = {};
  
  OCTAGON_AGENTS.forEach(agent => {
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
function createAgentCard(agent: typeof OCTAGON_AGENTS[0]): HTMLElement {
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
      navigator.clipboard.writeText(agent.examplePrompt)
        .then(() => {
          // Show success feedback
          copyButton.innerHTML = '<i class="ms-Icon ms-Icon--CheckMark copy-success"></i>';
          setTimeout(() => {
            copyButton.innerHTML = '<i class="ms-Icon ms-Icon--Copy"></i>';
          }, 1500);
        })
        .catch(err => {
          console.error('Could not copy text: ', err);
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
    agent.usageExamples.forEach(example => {
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
        navigator.clipboard.writeText(example.prompt)
          .then(() => {
            // Show success feedback
            copyButton.innerHTML = '<i class="ms-Icon ms-Icon--CheckMark copy-success"></i>';
            setTimeout(() => {
              copyButton.innerHTML = '<i class="ms-Icon ms-Icon--Copy"></i>';
            }, 1500);
          })
          .catch(err => {
            console.error('Could not copy text: ', err);
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
  const warningDiv = document.createElement('div');
  warningDiv.className = 'browser-warning';
  warningDiv.innerHTML = `
    <div class="warning-icon"><i class="ms-Icon ms-Icon--Warning"></i></div>
    <div class="warning-message">${message}</div>
  `;
  
  // Insert at the top of the body or in a specific container
  const container = document.querySelector('.content-container') || document.body;
  container.insertBefore(warningDiv, container.firstChild);
}

/**
 * Shows API support warnings to the user
 */
function showApiSupportWarning(issues: string[]) {
  const warningDiv = document.createElement('div');
  warningDiv.className = 'api-support-warning';
  
  let warningHtml = `
    <div class="warning-icon"><i class="ms-Icon ms-Icon--Warning"></i></div>
    <div class="warning-message">
      <p>Some features may not work correctly in your current environment:</p>
      <ul>
  `;
  
  issues.forEach(issue => {
    warningHtml += `<li>${issue}</li>`;
  });
  
  warningHtml += `
      </ul>
    </div>
  `;
  
  warningDiv.innerHTML = warningHtml;
  
  // Insert after the browser warning or at the top
  const container = document.querySelector('.content-container') || document.body;
  const existingWarning = document.querySelector('.browser-warning');
  
  if (existingWarning) {
    container.insertBefore(warningDiv, existingWarning.nextSibling);
  } else {
    container.insertBefore(warningDiv, container.firstChild);
  }
}