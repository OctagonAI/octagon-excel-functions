/**
 * Extension methods for OctagonApi to handle Excel-specific formatting and data processing
 */

import { octagonApiService } from './octagonApi';
import Logger from '../utils/logger';

/**
 * Extends the base OctagonApi with additional methods specifically for Excel integration
 */
export class OctagonApiExcel {
  private api: typeof octagonApiService;

  constructor(api: typeof octagonApiService) {
    this.api = api;
  }
  
  /**
   * Call an agent and return the results formatted for Excel
   * 
   * @param agentId - The ID of the agent to call
   * @param prompt - The prompt to send to the agent
   * @param returnType - The desired return type ('text', 'table', or 'auto')
   * @returns A promise that resolves to the agent's response, formatted according to returnType
   */
  async callAgentForExcel(
    agentId: string, 
    prompt: string, 
    returnType: 'text' | 'table' | 'auto' = 'auto'
  ): Promise<string | string[][] | Error> {
    
    try {
      // Call the base API method
      const response = await octagonApiService.callAgent(agentId, prompt);
      
      if (!response.success || !response.data) {
        return new Error(response.error || 'Failed to get response from Octagon');
      }
      
      // Get the content from the response
      const content = response.data.content;
      
      // For text return type, just return the content as a string
      if (returnType === 'text' || !content) {
        return content || '';
      }
      
    } catch (error) {
      Logger.error('Error in callAgentForExcel', error);
      return new Error(error instanceof Error ? error.message : 'Unknown error occurred');
    }
  }
}