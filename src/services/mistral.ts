import axios from 'axios';

export interface Slide {
  title: string;
  content: string;
  imageUrl?: string;
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
        content: `Crée une présentation powerpoint sur le sujet suivant :  [${prompt}]. Ta réponse contiendra un texte brut au Format JSON array contenant title and content, sans aucun autre caractère de mise forme. Ta réponse doit contenir au moins 3 slides et maximum 10 slides. La présentation doit débuter obligatoirement par une slide avec le titre de la présentation et un sous titre.`
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

  async generateKeywordsForImage(slideContent: string): Promise<string> {
    const request = {
      model: 'mistral-large-latest',
      messages: [{
        role: 'user',
        content: `Analyse le contenu suivant d'une diapositive PowerPoint et génère 3 mots-clés maximum pertinents en français qui pourraient être utilisés pour rechercher une image illustrant parfaitement ce contenu. Réponds uniquement avec les mots-clés séparés par des virgules, sans phrases ni explications supplémentaires.\n\nContenu de la diapositive:\n${slideContent}`
      }],
      temperature: 0.3
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
      
      // Récupérer les mots-clés générés
      const keywords = response.data.choices[0].message.content.trim();
      return keywords;
    } catch (error) {
      log.error = error;
      this.logs.unshift(log);
      console.error('Mistral API Error:', error);
      throw new Error('Erreur lors de la génération des mots-clés');
    }
  }
}