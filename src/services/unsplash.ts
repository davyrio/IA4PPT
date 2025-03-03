import axios from 'axios';

// Interface pour les résultats de recherche d'images Unsplash
export interface UnsplashImage {
  id: string;
  width: number;
  height: number;
  urls: {
    raw: string;
    full: string;
    regular: string;
    small: string;
    thumb: string;
  };
  links: {
    self: string;
    html: string;
    download: string;
  };
  user: {
    id: string;
    username: string;
    name: string;
    portfolio_url: string | null;
    profile_image: {
      small: string;
      medium: string;
      large: string;
    };
    links: {
      self: string;
      html: string;
      photos: string;
      likes: string;
    };
  };
  alt_description: string | null;
  description: string | null;
}

export interface UnsplashSearchResult {
  total: number;
  total_pages: number;
  results: UnsplashImage[];
}

export class UnsplashService {
  private static readonly API_KEY = 'QC1nm8kz1Q-Fgk7XQ_qXss6AVCykWbrZvAZXILmDT6o'; // Remplacez par votre clé API Unsplash
  private static readonly BASE_URL = 'https://api.unsplash.com';
  
  // Stockage des dernières requêtes pour gérer la pagination
  private static searchCache: {
    [key: string]: {
      keywords: string;
      currentPage: number;
      totalPages: number;
    }
  } = {};

  /**
   * Recherche des images sur Unsplash en fonction des mots-clés
   * @param keywords Mots-clés pour la recherche
   * @param perPage Nombre d'images à retourner (max 30)
   * @param page Numéro de page pour la pagination
   * @returns Liste des images correspondant aux mots-clés
   */
  static async searchImages(keywords: string, perPage: number = 4, page: number = 1): Promise<UnsplashImage[]> {
    try {
      const response = await axios.get<UnsplashSearchResult>(`${this.BASE_URL}/search/photos`, {
        headers: {
          'Authorization': `Client-ID ${this.API_KEY}`,
          'Accept-Version': 'v1'
        },
        params: {
          query: keywords,
          per_page: perPage,
          page: page,
          lang: 'fr'
        }
      });
      
      // Mettre à jour le cache de recherche avec les informations de pagination
      this.searchCache[keywords] = {
        keywords,
        currentPage: page,
        totalPages: response.data.total_pages
      };

      return response.data.results;
    } catch (error) {
      console.error('Erreur lors de la recherche d\'images Unsplash:', error);
      throw new Error('Impossible de récupérer les images');
    }
  }
  
  /**
   * Charge la page suivante d'images pour les mots-clés donnés
   * @param keywords Mots-clés utilisés pour la recherche initiale
   * @param perPage Nombre d'images par page
   * @returns Liste des images de la page suivante
   */
  static async loadNextPage(keywords: string, perPage: number = 4): Promise<UnsplashImage[]> {
    // Vérifier si nous avons des informations de pagination pour ces mots-clés
    const searchInfo = this.searchCache[keywords];
    
    if (!searchInfo) {
      // Si aucune recherche précédente, commencer à la page 1
      return this.searchImages(keywords, perPage, 1);
    }
    
    // Calculer la prochaine page, avec retour à la page 1 si on atteint la dernière page
    const nextPage = searchInfo.currentPage < searchInfo.totalPages 
      ? searchInfo.currentPage + 1 
      : 1;
    
    return this.searchImages(keywords, perPage, nextPage);
  }
}