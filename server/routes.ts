import type { Express } from "express";
import { createServer, type Server } from "http";
import { 
  googleSheetRequestSchema, 
  aiRewriteRequestSchema,
  proxyImageRequestSchema,
  type AIRewriteResponse
} from "@shared/schema";
import { log } from "./index";

export async function registerRoutes(
  httpServer: Server,
  app: Express
): Promise<Server> {
  
  // Google Sheets Proxy - Solve CORS issues
  app.post("/api/proxy/google-sheet", async (req, res) => {
    try {
      const { url } = googleSheetRequestSchema.parse(req.body);
      
      log(`Fetching Google Sheet: ${url}`);
      
      const response = await fetch(url);
      
      if (!response.ok) {
        return res.status(response.status).json({ 
          error: "Falha ao buscar planilha",
          details: response.statusText 
        });
      }
      
      // Get response as text with proper UTF-8 encoding
      const csvText = await response.text();
      
      res.json({ 
        success: true, 
        data: csvText,
        isText: true,
        contentType: response.headers.get('content-type') || 'text/csv; charset=utf-8'
      });
      
    } catch (error: any) {
      log(`Error proxying Google Sheet: ${error.message}`);
      res.status(500).json({ 
        error: "Erro ao processar planilha",
        details: error.message 
      });
    }
  });

  // Image Proxy - Fetch images from Google Drive and other sources
  const ALLOWED_IMAGE_HOSTS = [
    'drive.google.com',
    'lh3.googleusercontent.com',
    'docs.google.com',
    'drive.usercontent.google.com',
  ];
  const MAX_IMAGE_BYTES = 10 * 1024 * 1024;

  app.post("/api/proxy/image", async (req, res) => {
    try {
      const { url } = proxyImageRequestSchema.parse(req.body);

      let parsedUrl: URL;
      try {
        parsedUrl = new URL(url);
      } catch {
        return res.status(400).json({ error: "URL inválida" });
      }

      if (parsedUrl.protocol !== 'https:' || !ALLOWED_IMAGE_HOSTS.includes(parsedUrl.hostname)) {
        return res.status(403).json({ error: "Host não permitido" });
      }
      
      log(`Fetching image: ${url}`);
      
      const controller = new AbortController();
      const timeout = setTimeout(() => controller.abort(), 25000);
      
      try {
        const response = await fetch(url, { signal: controller.signal });
        
        if (!response.ok) {
          clearTimeout(timeout);
          return res.status(response.status).json({ 
            error: "Falha ao buscar imagem",
            details: response.statusText 
          });
        }
        
        const contentLength = response.headers.get('content-length');
        if (contentLength && parseInt(contentLength) > MAX_IMAGE_BYTES) {
          clearTimeout(timeout);
          controller.abort();
          return res.status(413).json({ 
            error: "Imagem muito grande",
            details: "Tamanho máximo: 5MB" 
          });
        }
        
        const contentType = response.headers.get('content-type') || '';

        // Reject non-image responses (e.g. HTML login redirect pages from Google Drive)
        const isImage = contentType.startsWith('image/') ||
          contentType.includes('jpeg') ||
          contentType.includes('png') ||
          contentType.includes('gif') ||
          contentType.includes('webp') ||
          contentType.includes('octet-stream');
        
        if (!isImage) {
          clearTimeout(timeout);
          log(`Image proxy rejected non-image content-type: ${contentType} for ${url}`);
          return res.status(415).json({ 
            error: "Conteúdo não é imagem",
            details: `Tipo recebido: ${contentType}. Verifique se o arquivo do Google Drive é público.`
          });
        }

        const imageData = await response.arrayBuffer();
        clearTimeout(timeout);
        
        if (imageData.byteLength > MAX_IMAGE_BYTES) {
          return res.status(413).json({ 
            error: "Imagem muito grande",
            details: "Tamanho máximo: 5MB" 
          });
        }
        
        const base64Data = Buffer.from(imageData).toString('base64');
        
        res.json({ 
          success: true, 
          data: base64Data,
          contentType: contentType || 'image/jpeg'
        });
      } catch (fetchError: any) {
        clearTimeout(timeout);
        if (fetchError.name === 'AbortError') {
          return res.status(408).json({ 
            error: "Timeout ao buscar imagem",
            details: "A requisição excedeu 10 segundos" 
          });
        }
        throw fetchError;
      }
      
    } catch (error: any) {
      log(`Error proxying image: ${error.message}`);
      res.status(500).json({ 
        error: "Erro ao buscar imagem",
        details: error.message 
      });
    }
  });

  // AI Text Rewriting - Professional technical language
  app.post("/api/ai/rewrite", async (req, res) => {
    try {
      const { text } = aiRewriteRequestSchema.parse(req.body);
      
      log(`AI rewriting text: ${text.substring(0, 50)}...`);
      
      // TODO: Implement AI rewriting using Hugging Face or OpenRouter
      // For now, return a placeholder that indicates AI is not configured
      
      const response: AIRewriteResponse = {
        originalText: text,
        rewrittenText: text, // Placeholder - will be replaced with AI
      };
      
      res.json(response);
      
    } catch (error: any) {
      log(`Error in AI rewrite: ${error.message}`);
      res.status(500).json({ 
        error: "Erro na reescrita de texto",
        details: error.message 
      });
    }
  });

  return httpServer;
}
