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

        PowerPoint.run(async function(context) {
          // Load information about all the slide masters and associated layouts.
          const slideMasters: PowerPoint.SlideMasterCollection = context.presentation.slideMasters.load("id, name, layouts/items/name, layouts/items/id");
          await context.sync();
        
          // Log the name and ID of each slide master.
          for (let i = 0; i < slideMasters.items.length; i++) {
            console.log("Master name: " + slideMasters.items[i].name);
            console.log("Master ID: " + slideMasters.items[i].id);
        
            // Log the name and ID of each slide layout in the slide master.
            const layoutsInMaster: PowerPoint.SlideLayoutCollection = slideMasters.items[i].layouts;
            for (let j = 0; j < layoutsInMaster.items.length; j++) {
              console.log("    Layout name: " + layoutsInMaster.items[j].name + " Layout ID: " + layoutsInMaster.items[j].id);
            }
          }
        });

        
        // Create a new slide using an existing master slide and layout.
          const newSlideOptions: PowerPoint.AddSlideOptions = {
            slideMasterId: '2147483648#93620447', /* An ID from `Presentation.slideMasters`. */
            layoutId: '2147483650#595629897' /* An ID from `SlideMaster.layouts`. */
          };
          const newSlideTitleOptions: PowerPoint.AddSlideOptions = {
            slideMasterId: '2147483648#93620447', /* An ID from `Presentation.slideMasters`. */
            layoutId: '2147483649#2160907878' /* An ID from `SlideMaster.layouts`. */
          };
      
        PowerPoint.run(async (context) => {
          for (let i = 0; i < slides.length; i++) {
            const slide = slides[i];
            try {
              await context.presentation.slides.load("items");
              await context.sync();
              let countSlides = context.presentation.slides.items.length;
              console.log(`Slide count: ${countSlides}`);
              if (countSlides == 0) {
                await context.presentation.slides.add(newSlideTitleOptions);
                await context.sync();
                console.log(`Added title slide`);
              } else {
                await context.presentation.slides.add(newSlideOptions);
                await context.sync();
                console.log(`Added content slide`);
              }
              const newSlide = await context.presentation.slides.getItemAt(i);
              await context.sync();
              this.logOperation('AddSlide', true, i, undefined, { newSlide: newSlide.id });
              let shapes = await newSlide.shapes.load("items,items/textFrame");

              await context.sync();
      
              // Modifier le texte du titre et du contenu
              shapes.items.forEach(async (shape) => {
                
                const textFrame: PowerPoint.TextFrame = shape.textFrame.load("textRange,hasText");
                await context.sync();
                console.log(`Shape has text: ${shape.textFrame.hasText}`);  
                const textRange: PowerPoint.TextRange = textFrame.textRange;
                textRange.load("text");
                
                await context.sync();
                let shapeId = shape.id;
                let shapeName = shape.name;
                if (shapeName.includes('Title')) {
                    textFrame.textRange.text = slide.title;
                } else if (shapeName.includes('Content') || shapeName.includes('Subtitle')) {
                    textFrame.textRange.text = slide.content;
                }
                await context.sync();
                console.log(`Updated text of shape ${shapeName} #${shapeId}: ${textFrame.textRange.text}`);
              });
      
              await context.sync();
              
              console.log(`Added slide ${j}: ${slide.title}`);
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