import axios from 'axios';

export interface Slide {
  title: string;
  content: string;
}

export interface ApiLog {
  timestamp: string;
  request: any;
  response?: any;
  error?: any;
}

export class MistralService {
  private apiKey: string;
  private baseUrl = 'https://api.mistral.ai/v1';
  private logs: ApiLog[] = [];

  constructor(apiKey: string) {
    this.apiKey = apiKey;
  }

  getLogs(): ApiLog[] {
    return this.logs;
  }

  clearLogs(): void {
    this.logs = [];
  }

  async generatePresentation(prompt: string): Promise<Slide[]> {
    const request = {
      model: 'mistral-large-latest',
      messages: [{
        role: 'user',
        content: `Crée une présentation powerpoint sur le sujet suivant :  [${prompt}]. Ta réponse contiendra un texte brut au Format JSON array contenant title and content, sans aucun autre caractère de mise forme ni retour chariot`
      }],
      temperature: 0.7
    };

    const log: ApiLog = {
      timestamp: new Date().toISOString(),
      request: request
    };

    try {
      const response = await axios.post(
        `${this.baseUrl}/chat/completions`,
        request,
        {
          headers: {
            'Authorization': `Bearer ${this.apiKey}`,
            'Content-Type': 'application/json'
          }
        }
      );

      log.response = response.data;
      this.logs.unshift(log);
      
      // Clean the response content by removing escape characters
      let content = response.data.choices[0].message.content;
      
      // Try to extract JSON if it's wrapped in backticks or other markers
      const jsonMatch = content.match(/\[[\s\S]*\]/);
      if (jsonMatch) {
        content = jsonMatch[0];
      }
      
      // Clean the content by replacing escaped characters
      content = content
        .replace(/\\"/g, '"')  // Replace escaped quotes
        .replace(/\\n/g, ' ')  // Replace newlines with spaces
        .replace(/\\\\/g, '\\'); // Replace double backslashes
      
      try {
        return JSON.parse(content);
      } catch (parseError) {
        console.error('JSON Parse Error:', parseError);
        console.log('Content that failed to parse:', content);
        throw new Error('Erreur lors du traitement de la réponse JSON');
      }
    } catch (error) {
      log.error = error;
      this.logs.unshift(log);
      console.error('Mistral API Error:', error);
      throw new Error('Erreur lors de la génération de la présentation');
    }
  }
}