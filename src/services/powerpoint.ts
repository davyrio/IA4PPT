import { Slide } from './mistral';

export interface SlideOperationLog {
  operation: string;
  timestamp: string;
  slideIndex?: number;
  success: boolean;
  error?: string;
  details?: any;
}

export class PowerPointService {
  private static logs: SlideOperationLog[] = [];

  static getLogs(): SlideOperationLog[] {
    return this.logs;
  }

  static clearLogs(): void {
    this.logs = [];
  }

  private static logOperation(operation: string, success: boolean, slideIndex?: number, error?: string, details?: any): void {
    this.logs.unshift({
      operation,
      timestamp: new Date().toISOString(),
      slideIndex,
      success,
      error,
      details
    });
  }

  static async createSlides(slides: Slide[]): Promise<void> {
    return new Promise(async (resolve, reject) => {
      // Vérifier que l'API Office est disponible
      if (!window.Office || !Office.context) {
        this.logOperation('OfficeAPICheck', false, undefined, "L'API Office.js n'est pas disponible");
        reject(new Error("L'API Office.js n'est pas disponible"));
        return;
      }

      this.logOperation('OfficeAPICheck', true, undefined, undefined, { host: Office.context.host });

      // Déterminer la méthode à utiliser
      let methodSuccess = false;

      // Méthode 1: API PowerPoint spécifique (la plus fiable)
      if (!methodSuccess && typeof PowerPoint !== 'undefined' && PowerPoint.run) {
        try {
          this.logOperation('AttemptMethod', true, undefined, undefined, { method: 'PowerPointAPI' });
          await this.createSlidesViaPowerPointAPI(slides);
          methodSuccess = true;
          this.logOperation('MethodSuccess', true, undefined, undefined, { method: 'PowerPointAPI' });
        } catch (error) {
          this.logOperation('MethodFailed', false, undefined, error.message, { method: 'PowerPointAPI' });
        }
      }
/*
      // Méthode 2: API Document PowerPoint
      if (!methodSuccess) {
        try {
          this.logOperation('AttemptMethod', true, undefined, undefined, { method: 'DocumentAPI' });
          await this.createSlidesViaDocumentAPI(slides);
          methodSuccess = true;
          this.logOperation('MethodSuccess', true, undefined, undefined, { method: 'DocumentAPI' });
        } catch (error) {
          this.logOperation('MethodFailed', false, undefined, error.message, { method: 'DocumentAPI' });
        }
      }

      // Méthode 3: Insertion séquentielle
      if (!methodSuccess) {
        try {
          this.logOperation('AttemptMethod', true, undefined, undefined, { method: 'SequentialInsertion' });
          await this.createSlidesSequentially(slides);
          methodSuccess = true;
          this.logOperation('MethodSuccess', true, undefined, undefined, { method: 'SequentialInsertion' });
        } catch (error) {
          this.logOperation('MethodFailed', false, undefined, error.message, { method: 'SequentialInsertion' });
        }
      }

      // Méthode 4: Approche OOXML
      if (!methodSuccess) {
        try {
          this.logOperation('AttemptMethod', true, undefined, undefined, { method: 'OOXML' });
          await this.createSlidesViaOOXML(slides);
          methodSuccess = true;
          this.logOperation('MethodSuccess', true, undefined, undefined, { method: 'OOXML' });
        } catch (error) {
          this.logOperation('MethodFailed', false, undefined, error.message, { method: 'OOXML' });
        }
      }
*/
      // Méthode de dernier recours: texte formaté
      if (!methodSuccess) {
        try {
          this.logOperation('AttemptMethod', true, undefined, undefined, { method: 'TextFormat' });
          await this.createSlidesAsText(slides);
          methodSuccess = true;
          this.logOperation('MethodSuccess', true, undefined, undefined, { method: 'TextFormat' });
        } catch (error) {
          this.logOperation('MethodFailed', false, undefined, error.message, { method: 'TextFormat' });
          // Si même cette méthode échoue, rejeter avec toutes les informations des logs
          reject(new Error(`Échec de création des diapositives: ${this.logs.map(log => `${log.operation}:${log.success?'OK':'FAIL'}`).join(', ')}`));
          return;
        }
      }

      resolve();
    });
  }

  private static async createSlidesViaPowerPointAPI(slides: Slide[]): Promise<void> {
    return new Promise((resolve, reject) => {
      try {
        PowerPoint.run(async (context) => {
          for (let i = 0; i < slides.length; i++) {
            const slide = slides[i];
            try {
              await context.presentation.slides.add();
              const newSlide = context.presentation.slides.getItemAt(i);
              context.load(newSlide);
              await
              context.sync();
              this.logOperation('AddSlide', true, i, undefined, { newSlide: newSlide.id });
              const titleShape = newSlide.shapes.addTextBox(slide.title);
              titleShape.top = 50;
              titleShape.left = 50;
              titleShape.width = 600;
              titleShape.height = 50;
              
              const contentShape = newSlide.shapes.addTextBox(slide.content);
              contentShape.top = 120;
              contentShape.left = 50;
              contentShape.width = 600;
              contentShape.height = 300;
              
              this.logOperation('AddSlide', true, i, undefined, { title: slide.title.substring(0, 20) + '...' });
            } catch (slideError) {
              this.logOperation('AddSlide', false, i, slideError.message, { title: slide.title.substring(0, 20) + '...' });
              // Continue avec les autres slides malgré l'erreur
            }
          }
          
          await context.sync();
          resolve();
        }).catch((error) => {
          this.logOperation('PowerPointRunSync', false, undefined, error.message);
          reject(error);
        });
      } catch (error) {
        this.logOperation('PowerPointRunSetup', false, undefined, error.message);
        reject(error);
      }
    });
  }


  private static async createSlidesAsText(slides: Slide[]): Promise<void> {
    return new Promise((resolve, reject) => {
      try {
        let allSlidesText = "⚠️ INSTRUCTIONS POUR CRÉER MANUELLEMENT LA PRÉSENTATION ⚠️\n\n";
        allSlidesText += "Cette présentation n'a pas pu être automatiquement créée.\n";
        allSlidesText += "Veuillez suivre ces étapes manuelles:\n";
        allSlidesText += "1. Pour chaque section 'DIAPOSITIVE X' ci-dessous, créez une nouvelle diapositive\n";
        allSlidesText += "2. Copiez le titre et le contenu dans la diapositive appropriée\n\n";
        
        slides.forEach((slide, index) => {
          allSlidesText += `==== DIAPOSITIVE ${index + 1} ====\n\n`;
          allSlidesText += `TITRE: ${slide.title}\n\n`;
          allSlidesText += `CONTENU:\n${slide.content}\n\n`;
          allSlidesText += `-----------------\n\n`;
        });
        
        Office.context.document.setSelectedDataAsync(
          allSlidesText,
          { coercionType: Office.CoercionType.Text },
          (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
              this.logOperation('CreateSlidesAsText', true);
              resolve();
            } else {
              this.logOperation('CreateSlidesAsText', false, undefined, result.error.message);
              reject(new Error("Échec de l'insertion du texte"));
            }
          }
        );
      } catch (error) {
        this.logOperation('CreateSlidesAsTextSetup', false, undefined, error.message);
        reject(error);
      }
    });
  }
}