import axios from 'axios';

// Interface pour les résultats de recherche d'images
export interface PexelsImage {
  id: number;
  width: number;
  height: number;
  url: string;
  photographer: string;
  photographer_url: string;
  photographer_id: number;
  avg_color: string;
  src: {
    original: string;
    large2x: string;
    large: string;
    medium: string;
    small: string;
    portrait: string;
    landscape: string;
    tiny: string;
  };
  liked: boolean;
  alt: string;
}

export interface PexelsSearchResult {
  total_results: number;
  page: number;
  per_page: number;
  photos: PexelsImage[];
  next_page: string;
}

export class PexelsService {
  private static readonly API_KEY = 'yoNhUCXRYoKKTtXURkdv4yETwNsVktigss19nmYcWBuNW7ePsO2hzygK'; // Clé API Pexels (remplacer par votre propre clé)
  private static readonly BASE_URL = 'https://api.pexels.com/v1';

  /**
   * Recherche des images sur Pexels en fonction des mots-clés
   * @param keywords Mots-clés pour la recherche
   * @param perPage Nombre d'images à retourner (max 80)
   * @returns Liste des images correspondant aux mots-clés
   */
  static async searchImages(keywords: string, perPage: number = 3): Promise<PexelsImage[]> {
    try {
      const response = await axios.get<PexelsSearchResult>(`${this.BASE_URL}/search`, {
        headers: {
          'Authorization': this.API_KEY
        },
        params: {
          query: keywords,
          per_page: perPage,
          locale: 'fr-FR'
        }
      });

      return response.data.photos;
    } catch (error) {
      console.error('Erreur lors de la recherche d\'images Pexels:', error);
      throw new Error('Impossible de récupérer les images');
    }
  }
}