import React, { useState, useEffect } from 'react';
import { Send, Key, ChevronDown, ChevronUp, AlertCircle, RefreshCw } from 'lucide-react';
import { MistralService, ApiLog } from './services/mistral';
import { PowerPointService, SlideOperationLog } from './services/powerpoint';

function App() {
  const [apiKey, setApiKey] = useState('');
  const [prompt, setPrompt] = useState('');
  const [loading, setLoading] = useState(false);
  const [message, setMessage] = useState('');
  const [error, setError] = useState('');
  const [mistralService, setMistralService] = useState<MistralService | null>(null);
  const [showLogs, setShowLogs] = useState(false);
  const [showPPTLogs, setShowPPTLogs] = useState(false);
  const [slides, setSlides] = useState<any[]>([]);
  const [operationStatus, setOperationStatus] = useState<string>('');

  useEffect(() => {
    const savedApiKey = localStorage.getItem('mistral_api_key');
    if (savedApiKey) {
      setApiKey(savedApiKey);
      setMistralService(new MistralService(savedApiKey));
    }
    
    // Check if Office.js is available
    if (!window.Office) {
      setError("Cette application nécessite Microsoft Office pour fonctionner. Veuillez l'ouvrir dans un add-in Office.");
    } else {
      // Vérifier si c'est PowerPoint spécifiquement
      if (Office.context && Office.context.host !== "PowerPoint") {
        setError("Cette application est conçue pour fonctionner dans PowerPoint. Veuillez l'ouvrir dans PowerPoint.");
      }
    }
  }, []);

  const saveApiKey = () => {
    if (apiKey) {
      localStorage.setItem('mistral_api_key', apiKey);
      setMistralService(new MistralService(apiKey));
      setMessage('Clé API sauvegardée');
    }
  };

  const handleSubmit = async () => {
    if (!mistralService) {
      setMessage('Veuillez d\'abord configurer votre clé API');
      return;
    }

    if (!prompt) {
      setMessage('Veuillez entrer une description pour la présentation');
      return;
    }

    setLoading(true);
    setMessage('Génération de la présentation en cours...');
    setError('');
    setSlides([]);
    setOperationStatus('Initialisation...');

    try {
      setOperationStatus('Génération du contenu via API Mistral...');
      const generatedSlides = await mistralService.generatePresentation(prompt);
      setSlides(generatedSlides);
      
      try {
        setOperationStatus('Création des diapositives dans PowerPoint...');
        await PowerPointService.createSlides(generatedSlides);
        setOperationStatus('');
        setMessage('Présentation générée avec succès ! Les diapositives ont été insérées dans PowerPoint.');
      } catch (powerPointError) {
        setOperationStatus('');
        setError(`Erreur PowerPoint: ${powerPointError.message}`);
        setMessage('Présentation générée mais impossible de l\'insérer automatiquement dans PowerPoint. Vous pouvez copier manuellement le contenu ci-dessous.');
        // Afficher automatiquement les logs en cas d'erreur
        setShowPPTLogs(true);
      }
    } catch (error) {
      setOperationStatus('');
      setError(`Erreur: ${error.message}`);
      setMessage('');
    } finally {
      setLoading(false);
    }
  };

  const LogViewer = ({ log }: { log: ApiLog }) => (
    <div className="mt-2 p-4 rounded bg-gray-100 text-sm font-mono overflow-x-auto">
      <div className="text-gray-500">{new Date(log.timestamp).toLocaleString()}</div>
      <div className="mt-2">
        <strong>Request:</strong>
        <pre className="mt-1 text-xs">{JSON.stringify(log.request, null, 2)}</pre>
      </div>
      {log.response && (
        <div className="mt-2">
          <strong>Response:</strong>
          <pre className="mt-1 text-xs">{JSON.stringify(log.response, null, 2)}</pre>
        </div>
      )}
      {log.error && (
        <div className="mt-2 text-red-600">
          <strong>Error:</strong>
          <pre className="mt-1 text-xs">{JSON.stringify(log.error, null, 2)}</pre>
        </div>
      )}
    </div>
  );

  const PowerPointLogViewer = ({ log }: { log: SlideOperationLog }) => (
    <div className={`mt-2 p-3 rounded text-sm font-mono overflow-x-auto ${log.success ? 'bg-green-50' : 'bg-red-50'}`}>
      <div className="flex justify-between">
        <span className={`font-bold ${log.success ? 'text-green-700' : 'text-red-700'}`}>
          {log.operation} {log.slideIndex !== undefined ? `(Slide ${log.slideIndex + 1})` : ''}
        </span>
        <span className="text-gray-500 text-xs">{new Date(log.timestamp).toLocaleTimeString()}</span>
      </div>
      {log.error && (
        <div className="mt-1 text-red-600 text-xs">
          <strong>Error:</strong> {log.error}
        </div>
      )}
      {log.details && (
        <div className="mt-1 text-gray-600 text-xs">
          <strong>Details:</strong>
          <pre className="mt-1">{JSON.stringify(log.details, null, 2)}</pre>
        </div>
      )}
    </div>
  );

  const SlidePreview = ({ slides }: { slides: any[] }) => (
    <div className="mt-4 border rounded-md p-4 bg-gray-50">
      <h3 className="text-lg font-medium mb-3">Aperçu des diapositives générées</h3>
      <div className="space-y-4">
        {slides.map((slide, index) => (
          <div key={index} className="border rounded p-3 bg-white">
            <h4 className="font-bold">Diapositive {index + 1}: {slide.title}</h4>
            <p className="mt-2 text-sm whitespace-pre-line">{slide.content}</p>
          </div>
        ))}
      </div>
    </div>
  );

  return (
    <div className="min-h-screen bg-gray-50 p-4">
      <div className="max-w-4xl mx-auto bg-white rounded-xl shadow-md overflow-hidden">
        <div className="p-6">
          <h1 className="text-2xl font-bold text-gray-900 mb-4">Assistant Mistral PowerPoint</h1>
          
          {error && (
            <div className="mb-4 p-4 rounded-md bg-red-50 flex items-start">
              <AlertCircle className="h-5 w-5 text-red-500 mr-2 mt-0.5" />
              <div className="text-sm text-red-700">{error}</div>
            </div>
          )}
          
          <div className="space-y-4">
            <div>
              <label className="block text-sm font-medium text-gray-700">
                Clé API Mistral
              </label>
              <div className="mt-1 flex rounded-md shadow-sm">
                <input
                  type="password"
                  value={apiKey}
                  onChange={(e) => setApiKey(e.target.value)}
                  className="flex-1 min-w-0 block w-full px-3 py-2 rounded-l-md border border-gray-300 focus:ring-indigo-500 focus:border-indigo-500"
                  placeholder="sk-..."
                />
                <button
                  onClick={saveApiKey}
                  className="inline-flex items-center px-4 py-2 border border-l-0 border-gray-300 rounded-r-md bg-gray-50 hover:bg-gray-100"
                >
                  <Key className="h-5 w-5 text-gray-400" />
                </button>
              </div>
            </div>

            <div>
              <label className="block text-sm font-medium text-gray-700">
                Description de la présentation
              </label>
              <div className="mt-1">
                <textarea
                  value={prompt}
                  onChange={(e) => setPrompt(e.target.value)}
                  rows={4}
                  className="shadow-sm block w-full focus:ring-indigo-500 focus:border-indigo-500 sm:text-sm border border-gray-300 rounded-md"
                  placeholder="Décrivez la présentation que vous souhaitez créer..."
                />
              </div>
            </div>

            <button
              onClick={handleSubmit}
              disabled={loading}
              className="w-full flex justify-center items-center px-4 py-2 border border-transparent rounded-md shadow-sm text-sm font-medium text-white bg-indigo-600 hover:bg-indigo-700 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-indigo-500 disabled:opacity-50"
            >
              {loading ? (
                <div className="flex items-center">
                  <RefreshCw className="h-4 w-4 mr-2 animate-spin" />
                  {operationStatus || 'Génération en cours...'}
                </div>
              ) : (
                <>
                  <Send className="h-5 w-5 mr-2" />
                  Générer
                </>
              )}
            </button>

            {message && (
              <div className="mt-4 p-4 rounded-md bg-gray-50 text-sm text-gray-700">
                {message}
              </div>
            )}

            {slides.length > 0 && <SlidePreview slides={slides} />}

            <div className="mt-8 space-y-4">
              <div>
                <button
                  onClick={() => setShowPPTLogs(!showPPTLogs)}
                  className="flex items-center text-sm text-gray-600 hover:text-gray-900"
                >
                  {showPPTLogs ? (
                    <ChevronUp className="h-4 w-4 mr-1" />
                  ) : (
                    <ChevronDown className="h-4 w-4 mr-1" />
                  )}
                  Logs PowerPoint
                </button>
                
                {showPPTLogs && (
                  <div className="mt-2">
                    {PowerPointService.getLogs().length === 0 ? (
                      <div className="text-sm text-gray-500">Aucun log disponible</div>
                    ) : (
                      <div className="space-y-1 max-h-64 overflow-y-auto">
                        {PowerPointService.getLogs().map((log, index) => (
                          <PowerPointLogViewer key={index} log={log} />
                        ))}
                      </div>
                    )}
                  </div>
                )}
              </div>
              
              <div>
                <button
                  onClick={() => setShowLogs(!showLogs)}
                  className="flex items-center text-sm text-gray-600 hover:text-gray-900"
                >
                  {showLogs ? (
                    <ChevronUp className="h-4 w-4 mr-1" />
                  ) : (
                    <ChevronDown className="h-4 w-4 mr-1" />
                  )}
                  Logs API Mistral
                </button>
                
                {showLogs && mistralService && (
                  <div className="mt-2">
                    {mistralService.getLogs().length === 0 ? (
                      <div className="text-sm text-gray-500">Aucun log disponible</div>
                    ) : (
                      <div className="space-y-1 max-h-64 overflow-y-auto">
                        {mistralService.getLogs().map((log, index) => (
                          <LogViewer key={index} log={log} />
                        ))}
                      </div>
                    )}
                  </div>
                )}
              </div>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}

export default App;