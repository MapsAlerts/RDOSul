import { useState, useMemo, useEffect } from "react";
import { Header } from "@/components/rdo/header";
import { FileUpload } from "@/components/rdo/file-upload";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Textarea } from "@/components/ui/textarea";
import { Calendar } from "@/components/ui/calendar";
import { Popover, PopoverContent, PopoverTrigger } from "@/components/ui/popover";
import { format, parse, isWithinInterval, startOfDay, endOfDay } from "date-fns";
import { ptBR } from "date-fns/locale";
import { 
  Calendar as CalendarIcon, 
  Download, 
  FileText, 
  RefreshCw, 
  ChevronRight, 
  ChevronLeft,
  FileCheck, 
  FileSpreadsheet, 
  CloudDownload, 
  ExternalLink,
  Filter,
  Search,
  CheckSquare,
  Square,
  FileOutput,
  FileType,
  Wand2,
  Eye,
  Image as ImageIcon,
  Sun,
  Cloud,
  Moon,
  Users,
  Wrench,
  Truck,
  MessageSquare,
  AlertTriangle,
  ClipboardCheck,
  Clock,
  Save,
  FileDown,
  Printer,
  PenLine
} from "lucide-react";
import { cn } from "@/lib/utils";
import { useToast } from "@/hooks/use-toast";
import * as XLSX from "xlsx";
import ExcelJS from "exceljs";
import { saveAs } from "file-saver";
import jsPDF from "jspdf";
import html2canvas from "html2canvas";
import Docxtemplater from "docxtemplater";
import PizZip from "pizzip";
import ImageModule from "docxtemplater-image-module-free";
import { Card, CardContent } from "@/components/ui/card";
import { Dialog, DialogContent, DialogHeader, DialogTitle, DialogDescription } from "@/components/ui/dialog";
import { ScrollArea } from "@/components/ui/scroll-area";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from "@/components/ui/table";
import { Checkbox } from "@/components/ui/checkbox";
import { Badge } from "@/components/ui/badge";
import { Progress } from "@/components/ui/progress";
import JSZip from "jszip";

// Setup interfaces
interface ReportData {
  id: string;
  rdoId: string; // Identificador único determinístico: IN-YYYY-NNN
  // OBRIGATÓRIAS
  data: Date | null;
  tecnico: string;    // Supervisor Responsável (campo principal de identificação)
  atividade: string;
  om: string;
  mina: string;
  localidade: string;
  // OPCIONAIS
  supervisor?: string;       // Supervisor Responsável
  encarregado?: string;      // Encarregado Responsável
  area?: string;             // Área
  tag?: string;              // TAG
  escalaTrabalho?: string;   // Escala de Trabalho
  obra?: string;             // Mantido para compatibilidade com planilhas antigas
  horarioInicio?: string;
  horarioTermino?: string;
  observacoes?: string;
  tarefaConcluida?: string;
  atividadePlanejada?: string;
  quemTrabalhou?: string;
  equipamentos?: string;
  veiculos?: string;
  climaMatutino?: string;
  climaVespertino?: string;
  climaNoturno?: string;
  fotosAntes: string[];
  fotosDepois: string[];
  status: "ready" | "error" | "generated";
  // CAMPOS EDITÁVEIS NO PREVIEW
  observacoesPreview?: string;
  correcaoFiscalizadora?: string;
  pendencias?: string;
}

// RDO_ID Generation System - Deterministic and Unique
interface RdoIdMapping {
  [year: string]: {
    [logicalKey: string]: number;
  };
}

const RDO_ID_STORAGE_KEY = 'rdo_id_mapping';

function normalizeString(str: string | number | null | undefined): string {
  return String(str || '')
    .toLowerCase()
    .trim()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, ' ');
}

function generateLogicalKey(report: { tecnico: string; obra?: string; localidade: string; data: Date | null; om: string | number }): string {
  const tecnico = normalizeString(report.tecnico);
  const obra = normalizeString(report.obra || '');
  const localidade = normalizeString(report.localidade);
  const dataStr = report.data ? format(report.data, 'yyyy-MM-dd') : '';
  const om = normalizeString(report.om);
  return `${tecnico}|${obra}|${localidade}|${dataStr}|${om}`;
}

function hashStringToNumber(str: string): number {
  let hash = 0;
  for (let i = 0; i < str.length; i++) {
    const char = str.charCodeAt(i);
    hash = ((hash << 5) - hash) + char;
    hash = hash & hash;
  }
  return Math.abs(hash);
}

function generateDeterministicSequence(logicalKey: string, year: string): number {
  const yearNum = parseInt(year, 10);
  const combinedKey = `${year}:${logicalKey}`;
  const hash = hashStringToNumber(combinedKey);
  return (hash % 900) + 1 + ((yearNum % 10) * 1000);
}

function getRdoIdMapping(): RdoIdMapping {
  try {
    const stored = localStorage.getItem(RDO_ID_STORAGE_KEY);
    return stored ? JSON.parse(stored) : {};
  } catch {
    return {};
  }
}

function saveRdoIdMapping(mapping: RdoIdMapping): void {
  localStorage.setItem(RDO_ID_STORAGE_KEY, JSON.stringify(mapping));
}

function generateRdoId(report: { tecnico: string; obra?: string; localidade: string; data: Date | null; om: string | number }): string {
  if (!report.data) return 'IN-0000-000';

  const year = format(report.data, 'yyyy');
  const logicalKey = generateLogicalKey(report);

  const mapping = getRdoIdMapping();

  if (!mapping[year]) {
    mapping[year] = {};
  }

  if (mapping[year][logicalKey]) {
    const sequence = mapping[year][logicalKey];
    return `IN-${year}-${sequence.toString().padStart(3, '0')}`;
  }

  let sequence = generateDeterministicSequence(logicalKey, year);
  const existingSequences = new Set(Object.values(mapping[year]));

  while (existingSequences.has(sequence)) {
    sequence++;
  }

  mapping[year][logicalKey] = sequence;
  saveRdoIdMapping(mapping);

  return `IN-${year}-${sequence.toString().padStart(3, '0')}`;
}

export default function Dashboard() {
  const DEFAULT_SHEET_URL = "https://docs.google.com/spreadsheets/d/1y79g7YYymbdWJ35PijwE7cL44-v1Hp8lKYEC3sPsmrA/edit?usp=sharing";

  const toCSVExportUrl = (url: string): string => {
    // Se já for URL de export CSV, retorna como está
    if (url.includes('/export?format=csv')) return url;
    // Extrai o ID da planilha e converte para URL de export CSV
    const match = url.match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
    if (match) {
      return `https://docs.google.com/spreadsheets/d/${match[1]}/export?format=csv`;
    }
    return url;
  };

  // Configuration State
  const [sheetUrl, setSheetUrl] = useState(DEFAULT_SHEET_URL);
  const [maskExcel, setMaskExcel] = useState<File | null>(null);
  const [maskWord, setMaskWord] = useState<File | Blob | null>(null);
  const [usingDefaultWord, setUsingDefaultWord] = useState(false);

  // Data State
  const [allReports, setAllReports] = useState<ReportData[]>([]);
  const [isFetching, setIsFetching] = useState(false);

  // Filter State
  const [dateStart, setDateStart] = useState<Date | undefined>();
  const [dateEnd, setDateEnd] = useState<Date | undefined>();
  const [selectedTechnicians, setSelectedTechnicians] = useState<string[]>([]);

  // Selection State
  const [selectedReportIds, setSelectedReportIds] = useState<Set<string>>(new Set());

  // Generation State
  const [isGeneratingExcel, setIsGeneratingExcel] = useState(false);
  const [excelProgress, setExcelProgress] = useState(0);

  const [isGeneratingWord, setIsGeneratingWord] = useState(false);
  const [wordProgress, setWordProgress] = useState(0);

  // Preview State
  const [previewOpen, setPreviewOpen] = useState(false);
  const [previewIndex, setPreviewIndex] = useState(0);
  const [isGeneratingPDF, setIsGeneratingPDF] = useState(false);

  // AI State - enabled by default
  const [aiEnabled, setAiEnabled] = useState(true);
  const [isCorrectingText, setIsCorrectingText] = useState(false);

  // Company Branding State
  const [companyName, setCompanyName] = useState("");
  const [companyLogo, setCompanyLogo] = useState<string | null>(null);
  const [prepostoSignature, setPrepostoSignature] = useState<string | null>(null);

  // Contract Configuration State
  const [companyCNPJ, setCompanyCNPJ] = useState("05.208.211/0001-38");
  const [contractLocal, setContractLocal] = useState("");
  const [contractNumber, setContractNumber] = useState("");
  const [contractStartDate, setContractStartDate] = useState<Date | undefined>();
  const [contractEndDate, setContractEndDate] = useState<Date | undefined>();
  const [contractExtensionDays, setContractExtensionDays] = useState(0);

  const { toast } = useToast();

  // Image cache to avoid re-downloading the same images
  const imageCache = useMemo(() => new Map<string, Promise<string | null>>(), []);

  // Flag to track if initial load is complete and auto-sync has run
  const [isInitialLoadComplete, setIsInitialLoadComplete] = useState(false);
  const [hasAutoSynced, setHasAutoSynced] = useState(false);

  // --- LocalStorage: Load saved configurations on mount ---
  useEffect(() => {
    try {
      const savedConfig = localStorage.getItem('rdo_config');
      if (savedConfig) {
        const config = JSON.parse(savedConfig);
        if (config.companyName) setCompanyName(config.companyName);
        if (config.companyLogo) setCompanyLogo(config.companyLogo);
        if (config.prepostoSignature) setPrepostoSignature(config.prepostoSignature);
        if (config.companyCNPJ) setCompanyCNPJ(config.companyCNPJ);
        if (config.contractLocal) setContractLocal(config.contractLocal);
        if (config.contractNumber) setContractNumber(config.contractNumber);
        if (config.contractStartDate) setContractStartDate(new Date(config.contractStartDate));
        if (config.contractEndDate) setContractEndDate(new Date(config.contractEndDate));
        if (typeof config.contractExtensionDays === 'number') setContractExtensionDays(config.contractExtensionDays);
      }

      const savedFilters = localStorage.getItem('rdo_filters');
      if (savedFilters) {
        const filters = JSON.parse(savedFilters);
        if (filters.dateStart) setDateStart(new Date(filters.dateStart));
        if (filters.dateEnd) setDateEnd(new Date(filters.dateEnd));
      }

      const OLD_DEFAULT_URL = "https://docs.google.com/spreadsheets/d/1PhA9jcvYoNKohMmQlgFHwZluKCbjQBweyXOHs3sMUCU/export?format=csv";
      const savedSheetUrl = localStorage.getItem('rdo_sheet_url');
      if (savedSheetUrl && savedSheetUrl !== OLD_DEFAULT_URL) {
        setSheetUrl(savedSheetUrl);
      } else if (savedSheetUrl === OLD_DEFAULT_URL) {
        // Migra automaticamente do URL antigo para o novo padrão
        localStorage.removeItem('rdo_sheet_url');
      }

      setIsInitialLoadComplete(true);
    } catch (error) {
      console.error('Erro ao carregar configurações salvas:', error);
      setIsInitialLoadComplete(true);
    }
  }, []);

  // --- LocalStorage: Save company & contract config when changed ---
  useEffect(() => {
    const config = {
      companyName,
      companyLogo,
      prepostoSignature,
      companyCNPJ,
      contractLocal,
      contractNumber,
      contractStartDate: contractStartDate?.toISOString(),
      contractEndDate: contractEndDate?.toISOString(),
      contractExtensionDays,
    };
    localStorage.setItem('rdo_config', JSON.stringify(config));
  }, [companyName, companyLogo, prepostoSignature, companyCNPJ, contractLocal, contractNumber, contractStartDate, contractEndDate, contractExtensionDays]);

  // --- LocalStorage: Save filter period when changed ---
  useEffect(() => {
    if (dateStart || dateEnd) {
      const filters = {
        dateStart: dateStart?.toISOString(),
        dateEnd: dateEnd?.toISOString(),
      };
      localStorage.setItem('rdo_filters', JSON.stringify(filters));
    }
  }, [dateStart, dateEnd]);

  // --- LocalStorage: Save sheet URL when changed ---
  useEffect(() => {
    if (sheetUrl) {
      localStorage.setItem('rdo_sheet_url', sheetUrl);
    }
  }, [sheetUrl]);

  // --- Helper: Text Correction with LanguageTool API (free, no API key) ---
  const correctTextWithLanguageTool = async (text: string): Promise<string> => {
    try {
      const response = await fetch('https://api.languagetoolplus.com/v2/check', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/x-www-form-urlencoded',
        },
        body: new URLSearchParams({
          text: text,
          language: 'pt-BR',
        }),
      });

      if (!response.ok) {
        if (response.status === 429) {
          throw new Error('RATE_LIMIT');
        }
        throw new Error(`LanguageTool API error: ${response.status}`);
      }

      const data = await response.json();

      if (!data.matches || data.matches.length === 0) {
        return text; // No corrections needed
      }

      // Apply corrections from end to start to preserve positions
      let correctedText = text;
      const sortedMatches = [...data.matches].sort((a, b) => b.offset - a.offset);

      for (const match of sortedMatches) {
        if (match.replacements && match.replacements.length > 0) {
          const replacement = match.replacements[0].value;
          const start = match.offset;
          const end = match.offset + match.length;
          correctedText = correctedText.slice(0, start) + replacement + correctedText.slice(end);
        }
      }

      return correctedText;
    } catch (error) {
      console.error("Erro na correção:", error);
      throw error;
    }
  };

  // --- Helper: Correct all activity texts for selected reports ---
  const correctAllTexts = async () => {
    if (!aiEnabled) return;

    const selectedReports = allReports.filter(r => selectedReportIds.has(r.id));
    if (selectedReports.length === 0) {
      toast({ title: "Aviso", description: "Selecione relatórios para corrigir.", variant: "destructive" });
      return;
    }

    setIsCorrectingText(true);
    let correctedCount = 0;
    let unchangedCount = 0;

    try {
      for (const report of selectedReports) {
        if (report.atividade && report.atividade.trim().length > 5) {
          try {
            const correctedText = await correctTextWithLanguageTool(report.atividade);
            if (correctedText !== report.atividade) {
              updateReportField(report.id, 'atividade', correctedText);
              correctedCount++;
            } else {
              unchangedCount++;
            }

            // Delay between requests to respect rate limits (20 req/min = 3s between requests)
            await new Promise(resolve => setTimeout(resolve, 3500));
          } catch (err: any) {
            if (err?.message === 'RATE_LIMIT') {
              toast({ 
                title: "Limite de requisições", 
                description: "Aguarde um momento e tente novamente.", 
                variant: "destructive" 
              });
              break; // Stop processing on rate limit
            }
            console.error(`Erro ao corrigir atividade do relatório ${report.id}:`, err);
          }
        }
      }

      toast({ 
        title: "Correção Concluída", 
        description: correctedCount > 0 
          ? `${correctedCount} texto(s) corrigido(s).${unchangedCount > 0 ? ` ${unchangedCount} já estavam corretos.` : ''}`
          : `Nenhuma correção necessária. ${unchangedCount} texto(s) já estavam corretos.`
      });
    } catch (error) {
      toast({ 
        title: "Erro", 
        description: "Falha na correção de textos.", 
        variant: "destructive" 
      });
    } finally {
      setIsCorrectingText(false);
    }
  };

  // --- Helper: Update report field in preview ---
  const updateReportField = (reportId: string, field: keyof ReportData, value: string) => {
    setAllReports(prev => prev.map(r => 
      r.id === reportId ? { ...r, [field]: value } : r
    ));
  };

  // --- Helper: Extract Google Drive file ID from various URL formats ---
  const extractFileId = (url: string): string => {
    if (url.includes('id=')) return url.split('id=')[1]?.split('&')[0] || '';
    if (url.includes('/d/')) return url.split('/d/')[1]?.split('/')[0] || '';
    return '';
  };

  // --- Helper: Single image fetch attempt via proxy ---
  const fetchImageFromUrl = async (proxyUrl: string): Promise<string | null> => {
    const response = await fetch('/api/proxy/image', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ url: proxyUrl })
    });

    if (!response.ok) return null;

    const result = await response.json();
    if (!result.success || !result.data) return null;

    let mimeType = 'image/jpeg';
    if (result.contentType?.includes('png')) {
      mimeType = 'image/png';
    } else if (result.contentType?.includes('gif')) {
      mimeType = 'image/gif';
    }

    return `data:${mimeType};base64,${result.data}`;
  };

  // --- Helper: Download image with multi-strategy fallback ---
  const downloadImageAsBase64 = async (url: string): Promise<string | null> => {
    if (imageCache.has(url)) {
      return imageCache.get(url)!;
    }

    const fetchPromise = (async (): Promise<string | null> => {
      try {
        const fileId = extractFileId(url);
        if (!fileId) return null;

        // Strategy 1: thumbnail (fast, smaller size)
        const thumbnailUrl = `https://drive.google.com/thumbnail?id=${fileId}&sz=w1200`;
        let result = await fetchImageFromUrl(thumbnailUrl);
        if (result) return result;

        // Strategy 2: drive.usercontent.google.com (more direct, avoids virus-scan redirect)
        await new Promise(r => setTimeout(r, 300));
        const ucontentUrl = `https://drive.usercontent.google.com/download?id=${fileId}&export=download&authuser=0`;
        result = await fetchImageFromUrl(ucontentUrl);
        if (result) return result;

        // Strategy 3: uc export download (classic fallback)
        await new Promise(r => setTimeout(r, 300));
        const directUrl = `https://drive.google.com/uc?export=download&id=${fileId}`;
        result = await fetchImageFromUrl(directUrl);
        return result;
      } catch (error) {
        console.error('Error downloading image:', error);
        return null;
      }
    })();

    imageCache.set(url, fetchPromise);
    return fetchPromise;
  };

  // --- Helper: Download multiple images in parallel with concurrency limit ---
  const downloadImagesParallel = async (urls: string[], concurrency: number = 3): Promise<(string | null)[]> => {
    const results: (string | null)[] = new Array(urls.length).fill(null);
    let currentIndex = 0;

    const worker = async () => {
      while (currentIndex < urls.length) {
        const index = currentIndex++;
        results[index] = await downloadImageAsBase64(urls[index]);
      }
    };

    const workers = Array(Math.min(concurrency, urls.length)).fill(null).map(() => worker());
    await Promise.all(workers);

    return results;
  };

  // --- Helper: Generate PDF from preview using html2canvas ---
  const generatePreviewPDF = async (report: ReportData, returnBlob: boolean = false): Promise<Blob | null> => {
    try {
      // Pre-download all images in PARALLEL (much faster than sequential)
      const antesUrls = report.fotosAntes.slice(0, 8);
      const depoisUrls = report.fotosDepois.slice(0, 8);
      const allUrls = [...antesUrls, ...depoisUrls];
      
      // Download all images in parallel with concurrency limit
      const allImages = await downloadImagesParallel(allUrls, 3);
      
      // Split results back into antes/depois arrays, replacing failures with placeholder
      const PLACEHOLDER_IMG = (() => {
        const canvas = document.createElement('canvas');
        canvas.width = 400;
        canvas.height = 267;
        const ctx = canvas.getContext('2d')!;
        ctx.fillStyle = '#f1f5f9';
        ctx.fillRect(0, 0, 400, 267);
        ctx.strokeStyle = '#cbd5e1';
        ctx.lineWidth = 2;
        ctx.strokeRect(1, 1, 398, 265);
        ctx.fillStyle = '#94a3b8';
        ctx.font = 'bold 16px system-ui, sans-serif';
        ctx.textAlign = 'center';
        ctx.fillText('Imagem indisponível', 200, 140);
        return canvas.toDataURL('image/png');
      })();
      const resolveImg = (img: string | null) => img ?? PLACEHOLDER_IMG;
      const fotosAntesBase64 = allImages.slice(0, antesUrls.length).map(resolveImg);
      const fotosDepoisBase64 = allImages.slice(antesUrls.length).map(resolveImg);

      // Create a hidden container for rendering the preview - A4 compatible
      const container = document.createElement('div');
      container.style.position = 'absolute';
      container.style.left = '-9999px';
      container.style.top = '0';
      container.style.width = '794px'; // A4 width at 96dpi (210mm)
      container.style.minHeight = '1123px'; // A4 height at 96dpi (297mm)
      container.style.boxSizing = 'border-box';
      container.style.backgroundColor = 'white';
      container.style.padding = '40px 45px'; // ~15mm margins for A4 printing
      container.style.fontFamily = 'system-ui, -apple-system, sans-serif';
      document.body.appendChild(container);

      // Helper: Format comma-separated list as bullet points
      const formatAsBulletList = (text: string | undefined): string => {
        if (!text || text.trim() === '' || text === '-' || text === 'N/A') return '-';
        const items = text.split(',').map(item => item.trim()).filter(item => item.length > 0);
        if (items.length === 0) return '-';
        if (items.length === 1) return `• ${items[0]}`;
        return items.map(item => `• ${item}`).join('<br/>');
      };

      // Generate photo HTML with actual images
      const generatePhotoGrid = (photos: string[], label: string, bgColor: string, borderColor: string, labelBg: string, labelText: string) => {
        if (photos.length === 0) {
          return `
            <div>
              <p style="text-align: center; background: ${labelBg}; color: ${labelText}; padding: 6px; font-size: 11px; font-weight: bold; margin: 0; border-radius: 6px 6px 0 0; border: 1px solid ${borderColor}; border-bottom: none;">${label}</p>
              <div style="display: flex; align-items: center; justify-content: center; padding: 20px; background: ${bgColor}; border-radius: 0 0 6px 6px; border: 1px solid ${borderColor}; border-top: none; min-height: 60px;">
                <p style="margin: 0; font-size: 11px; color: #64748b; font-weight: 500;">Sem fotos</p>
              </div>
            </div>
          `;
        }

        const photoHtml = photos.map((src, idx) => `
          <div style="aspect-ratio: 3/2; overflow: hidden; border-radius: 4px; border: 1px solid ${borderColor}; background: white;">
            <img src="${src}" alt="${label} ${idx + 1}" style="width: 100%; height: 100%; object-fit: cover;" crossorigin="anonymous" />
          </div>
        `).join('');

        return `
          <div>
            <p style="text-align: center; background: ${labelBg}; color: ${labelText}; padding: 6px; font-size: 11px; font-weight: bold; margin: 0; border-radius: 6px 6px 0 0; border: 1px solid ${borderColor}; border-bottom: none;">${label}</p>
            <div style="display: grid; grid-template-columns: repeat(2, 1fr); gap: 8px; padding: 12px; background: ${bgColor}; border-radius: 0 0 6px 6px; border: 1px solid ${borderColor}; border-top: none;">
              ${photoHtml}
            </div>
          </div>
        `;
      };

      // Render the preview HTML content
      container.innerHTML = `
        <div style="max-width: 750px; margin: 0 auto;">
          <!-- Header -->
          <div data-section="header" style="margin-bottom: 16px; border-bottom: 2px solid #1e3a5f; padding-bottom: 12px;">
            <div style="display: flex; align-items: center; justify-content: space-between; gap: 16px;">
              ${companyLogo ? `
                <img src="${companyLogo}" alt="Logo" style="height: 50px; width: auto; object-fit: contain;" />
              ` : `
                <div style="width: 50px;"></div>
              `}
              <div style="text-align: center; flex: 1;">
                ${companyName ? `
                  <p style="font-size: 14px; font-weight: bold; color: #1e3a5f; margin: 0 0 4px 0;">${companyName}</p>
                ` : ''}
                <h1 style="font-size: 16px; font-weight: bold; color: #1e3a5f; margin: 0;">RELATÓRIO DIÁRIO DE OBRA - RDO</h1>
                <p style="font-size: 11px; font-weight: 600; color: #64748b; margin: 4px 0 0 0; letter-spacing: 1px;">${report.rdoId}</p>
              </div>
              ${companyLogo ? `
                <div style="width: 50px;"></div>
              ` : `
                <div style="width: 50px;"></div>
              `}
            </div>
          </div>

          <!-- Contract Information -->
          ${(() => {
            const hasBothDates = contractStartDate && contractEndDate;
            const hasRdoAndStart = contractStartDate && report.data;
            const hasRdoAndEnd = contractEndDate && report.data;

            const prazoContratualRaw = hasBothDates 
              ? Math.ceil((contractEndDate.getTime() - contractStartDate.getTime()) / (1000 * 60 * 60 * 24))
              : null;
            const prazoContratual = prazoContratualRaw !== null && prazoContratualRaw >= 0 
              ? prazoContratualRaw + contractExtensionDays 
              : null;

            const diasDecorridosRaw = hasRdoAndStart 
              ? Math.ceil((report.data!.getTime() - contractStartDate.getTime()) / (1000 * 60 * 60 * 24)) + 1
              : null;
            const diasDecorridos = diasDecorridosRaw !== null && diasDecorridosRaw >= 1 
              ? diasDecorridosRaw 
              : null;

            const diasRestantesRaw = hasRdoAndEnd 
              ? Math.ceil((contractEndDate.getTime() - report.data!.getTime()) / (1000 * 60 * 60 * 24)) + contractExtensionDays
              : null;
            const diasRestantes = diasRestantesRaw !== null ? Math.max(0, diasRestantesRaw) : null;
            const diasAtraso = diasRestantesRaw !== null 
              ? (diasRestantesRaw < 0 ? Math.abs(diasRestantesRaw) : 0)
              : null;

            return `
          <div data-section="contract" style="background: #f8fafc; border-radius: 6px; padding: 12px 14px; border: 1px solid #e2e8f0; margin-bottom: 14px;">
            <div style="display: grid; grid-template-columns: repeat(2, 1fr); gap: 6px; font-size: 10px;">
              <div style="display: flex; gap: 4px;">
                <span style="color: #64748b;">Empresa:</span>
                <span style="color: #334155; font-weight: 500;">${companyName || '-'}</span>
              </div>
              <div style="display: flex; gap: 4px;">
                <span style="color: #64748b;">CNPJ:</span>
                <span style="color: #334155;">${companyCNPJ || '-'}</span>
              </div>
              <div style="display: flex; gap: 4px;">
                <span style="color: #64748b;">Local:</span>
                <span style="color: #334155;">${contractLocal || '-'}</span>
              </div>
              <div style="display: flex; gap: 4px;">
                <span style="color: #64748b;">Contrato:</span>
                <span style="color: #334155;">${contractNumber || '-'}</span>
              </div>
              <div style="display: flex; gap: 4px;">
                <span style="color: #64748b;">Data Inicial:</span>
                <span style="color: #334155;">${contractStartDate ? format(contractStartDate, 'dd/MM/yyyy') : '-'}</span>
              </div>
              <div style="display: flex; gap: 4px;">
                <span style="color: #64748b;">Data Final:</span>
                <span style="color: #334155;">${contractEndDate ? format(contractEndDate, 'dd/MM/yyyy') : '-'}</span>
              </div>
            </div>
            <div style="display: grid; grid-template-columns: repeat(5, 1fr); gap: 6px; margin-top: 8px; padding-top: 8px; border-top: 1px solid #e2e8f0;">
              <div style="text-align: center;">
                <p style="font-size: 8px; color: #64748b; margin: 0; text-transform: uppercase;">Prazo Contratual</p>
                <p style="font-size: 13px; font-weight: bold; color: #1e3a5f; margin: 2px 0 0 0;">${prazoContratual !== null ? prazoContratual : '-'}</p>
                <p style="font-size: 7px; color: #94a3b8; margin: 0;">${prazoContratual !== null ? 'dias' : ''}</p>
              </div>
              <div style="text-align: center;">
                <p style="font-size: 8px; color: #64748b; margin: 0; text-transform: uppercase;">Dias Decorridos</p>
                <p style="font-size: 13px; font-weight: bold; color: #1e3a5f; margin: 2px 0 0 0;">${diasDecorridos !== null ? diasDecorridos : '-'}</p>
                <p style="font-size: 7px; color: #94a3b8; margin: 0;">${diasDecorridos !== null ? 'dias' : ''}</p>
              </div>
              <div style="text-align: center;">
                <p style="font-size: 8px; color: #64748b; margin: 0; text-transform: uppercase;">Prorrogação</p>
                <p style="font-size: 13px; font-weight: bold; color: #1e3a5f; margin: 2px 0 0 0;">${contractExtensionDays}</p>
                <p style="font-size: 7px; color: #94a3b8; margin: 0;">dias</p>
              </div>
              <div style="text-align: center;">
                <p style="font-size: 8px; color: #64748b; margin: 0; text-transform: uppercase;">Dias Restantes</p>
                <p style="font-size: 13px; font-weight: bold; color: #1e3a5f; margin: 2px 0 0 0;">${diasRestantes !== null ? diasRestantes : '-'}</p>
                <p style="font-size: 7px; color: #94a3b8; margin: 0;">${diasRestantes !== null ? 'dias' : ''}</p>
              </div>
              <div style="text-align: center;">
                <p style="font-size: 8px; color: #64748b; margin: 0; text-transform: uppercase;">Dias de Atraso</p>
                <p style="font-size: 13px; font-weight: bold; color: #1e3a5f; margin: 2px 0 0 0;">${diasAtraso !== null ? diasAtraso : '-'}</p>
                <p style="font-size: 7px; color: #94a3b8; margin: 0;">${diasAtraso !== null ? 'dias' : ''}</p>
              </div>
            </div>
          </div>
            `;
          })()}

          <!-- Section 1: OM Identification - Clean Layout with Icons -->
          <div data-section="identification" style="background: #f8fafc; border-radius: 6px; padding: 16px; border: 1px solid #e2e8f0; margin-bottom: 14px;">
            <h3 style="font-size: 11px; font-weight: bold; color: #1e3a5f; margin: 0 0 12px 0; padding-bottom: 8px; border-bottom: 1px solid #e2e8f0; text-transform: uppercase; letter-spacing: 0.5px;">
              <span style="display: inline-flex; align-items: center; gap: 6px;">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#1e3a5f" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="4" width="18" height="18" rx="2" ry="2"></rect><line x1="16" y1="2" x2="16" y2="6"></line><line x1="8" y1="2" x2="8" y2="6"></line><line x1="3" y1="10" x2="21" y2="10"></line></svg>
                Identificação da Ordem de Serviço
              </span>
            </h3>

            <div style="display: grid; grid-template-columns: repeat(5, 1fr); gap: 10px; margin-bottom: 12px;">
              <div>
                <p style="font-size: 9px; color: #64748b; margin: 0 0 4px 0; text-transform: uppercase; display: flex; align-items: center; gap: 4px;">
                  <svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="#64748b" stroke-width="2"><rect x="3" y="4" width="18" height="18" rx="2" ry="2"></rect><line x1="16" y1="2" x2="16" y2="6"></line><line x1="8" y1="2" x2="8" y2="6"></line><line x1="3" y1="10" x2="21" y2="10"></line></svg>
                  Data
                </p>
                <p style="font-size: 12px; font-weight: 600; margin: 0; color: #334155;">${report.data ? format(report.data, 'dd/MM/yyyy') : '-'}</p>
              </div>
              <div>
                <p style="font-size: 9px; color: #64748b; margin: 0 0 4px 0; text-transform: uppercase; display: flex; align-items: center; gap: 4px;">
                  <svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="#64748b" stroke-width="2"><path d="M4 19.5A2.5 2.5 0 0 1 6.5 17H20"></path><path d="M6.5 2H20v20H6.5A2.5 2.5 0 0 1 4 19.5v-15A2.5 2.5 0 0 1 6.5 2z"></path></svg>
                  Nº OM
                </p>
                <p style="font-size: 14px; font-weight: bold; margin: 0; color: #1e3a5f;">${report.om || 'S/I'}</p>
              </div>
              <div>
                <p style="font-size: 9px; color: #64748b; margin: 0 0 4px 0; text-transform: uppercase; display: flex; align-items: center; gap: 4px;">
                  <svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="#64748b" stroke-width="2"><path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"></path><circle cx="12" cy="7" r="4"></circle></svg>
                  Supervisor
                </p>
                <p style="font-size: 11px; font-weight: 500; margin: 0; color: #334155;">${report.supervisor || report.tecnico}</p>
              </div>
              <div>
                <p style="font-size: 9px; color: #64748b; margin: 0 0 4px 0; text-transform: uppercase; display: flex; align-items: center; gap: 4px;">
                  <svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="#64748b" stroke-width="2"><path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"></path><circle cx="12" cy="7" r="4"></circle></svg>
                  Encarregado
                </p>
                <p style="font-size: 11px; font-weight: 500; margin: 0; color: #334155;">${report.encarregado || '-'}</p>
              </div>
              <div>
                <p style="font-size: 9px; color: #64748b; margin: 0 0 4px 0; text-transform: uppercase; display: flex; align-items: center; gap: 4px;">
                  <svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="#64748b" stroke-width="2"><polyline points="9 11 12 14 22 4"></polyline><path d="M21 12v7a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11"></path></svg>
                  Tarefa Concluída
                </p>
                <p style="font-size: 12px; font-weight: 600; margin: 0; color: #334155;">${report.tarefaConcluida || '-'}</p>
              </div>
            </div>

            <div style="display: grid; grid-template-columns: repeat(4, 1fr); gap: 12px; margin-bottom: 12px;">
              <div>
                <p style="font-size: 9px; color: #64748b; margin: 0 0 4px 0; text-transform: uppercase; display: flex; align-items: center; gap: 4px;">
                  <svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="#64748b" stroke-width="2"><polygon points="12 2 2 7 12 12 22 7 12 2"></polygon><polyline points="2 17 12 22 22 17"></polyline><polyline points="2 12 12 17 22 12"></polyline></svg>
                  Mina
                </p>
                <p style="font-size: 11px; font-weight: 500; margin: 0; color: #334155;">${report.mina || 'S/I'}</p>
              </div>
              <div>
                <p style="font-size: 9px; color: #64748b; margin: 0 0 4px 0; text-transform: uppercase; display: flex; align-items: center; gap: 4px;">
                  <svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="#64748b" stroke-width="2"><path d="M21 10c0 7-9 13-9 13s-9-6-9-13a9 9 0 0 1 18 0z"></path><circle cx="12" cy="10" r="3"></circle></svg>
                  Localidade
                </p>
                <p style="font-size: 11px; font-weight: 500; margin: 0; color: #334155;">${report.localidade || 'S/I'}</p>
              </div>
              <div>
                <p style="font-size: 9px; color: #64748b; margin: 0 0 4px 0; text-transform: uppercase; display: flex; align-items: center; gap: 4px;">
                  <svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="#64748b" stroke-width="2"><path d="M3 9l9-7 9 7v11a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z"></path><polyline points="9 22 9 12 15 12 15 22"></polyline></svg>
                  Área
                </p>
                <p style="font-size: 11px; font-weight: 500; margin: 0; color: #334155;">${report.area || 'S/I'}</p>
              </div>
              <div>
                <p style="font-size: 9px; color: #64748b; margin: 0 0 4px 0; text-transform: uppercase; display: flex; align-items: center; gap: 4px;">
                  <svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="#64748b" stroke-width="2"><circle cx="6" cy="6" r="3"></circle><circle cx="18" cy="18" r="3"></circle><path d="M6 9v12"></path><path d="M18 3v6"></path></svg>
                  TAG
                </p>
                <p style="font-size: 11px; font-weight: 500; margin: 0; color: #334155;">${report.tag || '-'}</p>
              </div>
            </div>

            <div style="display: grid; grid-template-columns: repeat(4, 1fr); gap: 12px;">
              <div>
                <p style="font-size: 9px; color: #64748b; margin: 0 0 4px 0; text-transform: uppercase; display: flex; align-items: center; gap: 4px;">
                  <svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="#64748b" stroke-width="2"><circle cx="12" cy="12" r="10"></circle><polyline points="12 6 12 12 16 14"></polyline></svg>
                  Horário Início
                </p>
                <p style="font-size: 12px; font-weight: 500; margin: 0; color: #334155;">${report.horarioInicio || '-'}</p>
              </div>
              <div>
                <p style="font-size: 9px; color: #64748b; margin: 0 0 4px 0; text-transform: uppercase; display: flex; align-items: center; gap: 4px;">
                  <svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="#64748b" stroke-width="2"><circle cx="12" cy="12" r="10"></circle><polyline points="12 6 12 12 16 14"></polyline></svg>
                  Horário Término
                </p>
                <p style="font-size: 12px; font-weight: 500; margin: 0; color: #334155;">${report.horarioTermino || '-'}</p>
              </div>
              <div>
                <p style="font-size: 9px; color: #64748b; margin: 0 0 4px 0; text-transform: uppercase; display: flex; align-items: center; gap: 4px;">
                  <svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="#64748b" stroke-width="2"><rect x="2" y="7" width="20" height="14" rx="2" ry="2"></rect><path d="M16 21V5a2 2 0 0 0-2-2h-4a2 2 0 0 0-2 2v16"></path></svg>
                  Escala
                </p>
                <p style="font-size: 11px; font-weight: 500; margin: 0; color: #334155;">${report.escalaTrabalho || '-'}</p>
              </div>
              <div>
                <p style="font-size: 9px; color: #64748b; margin: 0 0 4px 0; text-transform: uppercase; display: flex; align-items: center; gap: 4px;">
                  <svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="#64748b" stroke-width="2"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path><polyline points="14 2 14 8 20 8"></polyline></svg>
                  Ativ. Planejada
                </p>
                <p style="font-size: 12px; font-weight: 600; margin: 0; color: #334155;">${report.atividadePlanejada || '-'}</p>
              </div>
            </div>
          </div>

          <!-- Section 2: Activity - Clean with Icon -->
          <div data-section="activity" style="margin-bottom: 14px;">
            <h3 style="font-size: 11px; font-weight: bold; color: #1e3a5f; margin: 0 0 8px 0; text-transform: uppercase; letter-spacing: 0.5px; display: flex; align-items: center; gap: 6px;">
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#1e3a5f" stroke-width="2"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path><polyline points="14 2 14 8 20 8"></polyline><line x1="16" y1="13" x2="8" y2="13"></line><line x1="16" y1="17" x2="8" y2="17"></line><polyline points="10 9 9 9 8 9"></polyline></svg>
              Atividade Realizada
            </h3>
            <div style="background: #f8fafc; border: 1px solid #e2e8f0; padding: 12px; border-radius: 6px; font-size: 11px; white-space: pre-wrap; min-height: 50px; line-height: 1.5; color: #334155;">
              ${report.atividade || 'Sem descrição'}
            </div>
          </div>

          <!-- Section 3: Climate - Clean with Icons -->
          <div data-section="climate" style="margin-bottom: 14px;">
            <h3 style="font-size: 11px; font-weight: bold; color: #1e3a5f; margin: 0 0 8px 0; text-transform: uppercase; letter-spacing: 0.5px; display: flex; align-items: center; gap: 6px;">
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#1e3a5f" stroke-width="2"><path d="M17.5 19H9a7 7 0 1 1 6.71-9h1.79a4.5 4.5 0 1 1 0 9Z"></path></svg>
              Condições Climáticas
            </h3>
            <div style="display: grid; grid-template-columns: repeat(3, 1fr); gap: 10px;">
              <div style="background: #f8fafc; border: 1px solid #e2e8f0; padding: 10px; border-radius: 6px; text-align: center;">
                <p style="font-size: 9px; color: #64748b; margin: 0 0 4px 0; text-transform: uppercase; display: flex; align-items: center; justify-content: center; gap: 4px;">
                  <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="#f59e0b" stroke-width="2"><circle cx="12" cy="12" r="5"></circle><line x1="12" y1="1" x2="12" y2="3"></line><line x1="12" y1="21" x2="12" y2="23"></line><line x1="4.22" y1="4.22" x2="5.64" y2="5.64"></line><line x1="18.36" y1="18.36" x2="19.78" y2="19.78"></line><line x1="1" y1="12" x2="3" y2="12"></line><line x1="21" y1="12" x2="23" y2="12"></line><line x1="4.22" y1="19.78" x2="5.64" y2="18.36"></line><line x1="18.36" y1="5.64" x2="19.78" y2="4.22"></line></svg>
                  Matutino
                </p>
                <p style="font-size: 12px; font-weight: 500; margin: 0; color: #334155;">${report.climaMatutino || '-'}</p>
              </div>
              <div style="background: #f8fafc; border: 1px solid #e2e8f0; padding: 10px; border-radius: 6px; text-align: center;">
                <p style="font-size: 9px; color: #64748b; margin: 0 0 4px 0; text-transform: uppercase; display: flex; align-items: center; justify-content: center; gap: 4px;">
                  <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="#f97316" stroke-width="2"><path d="M12 2v2"></path><path d="m4.93 4.93 1.41 1.41"></path><path d="M20 12h2"></path><path d="m19.07 4.93-1.41 1.41"></path><path d="M15.947 12.65a4 4 0 0 0-5.925-4.128"></path><path d="M13 22H7a5 5 0 1 1 4.9-6H13a3 3 0 0 1 0 6Z"></path></svg>
                  Vespertino
                </p>
                <p style="font-size: 12px; font-weight: 500; margin: 0; color: #334155;">${report.climaVespertino || '-'}</p>
              </div>
              <div style="background: #f8fafc; border: 1px solid #e2e8f0; padding: 10px; border-radius: 6px; text-align: center;">
                <p style="font-size: 9px; color: #64748b; margin: 0 0 4px 0; text-transform: uppercase; display: flex; align-items: center; justify-content: center; gap: 4px;">
                  <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="#6366f1" stroke-width="2"><path d="M21 12.79A9 9 0 1 1 11.21 3 7 7 0 0 0 21 12.79z"></path></svg>
                  Noturno
                </p>
                <p style="font-size: 12px; font-weight: 500; margin: 0; color: #334155;">${report.climaNoturno || '-'}</p>
              </div>
            </div>
          </div>

          <!-- Section 4: Resources - Clean with Icons -->
          <div data-section="resources" style="margin-bottom: 14px;">
            <div style="display: grid; grid-template-columns: repeat(2, 1fr); gap: 10px; align-items: stretch;">
              <div style="background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 6px; overflow: hidden;">
                <div style="background: #1e3a5f; padding: 8px 10px; display: flex; align-items: center; gap: 6px;">
                  <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="white" stroke-width="2"><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"></path><circle cx="9" cy="7" r="4"></circle><path d="M23 21v-2a4 4 0 0 0-3-3.87"></path><path d="M16 3.13a4 4 0 0 1 0 7.75"></path></svg>
                  <span style="color: white; font-size: 10px; font-weight: 600; text-transform: uppercase;">Equipe</span>
                </div>
                <div style="padding: 10px; font-size: 10px; min-height: 40px; color: #334155; line-height: 1.6;">
                  ${formatAsBulletList(report.quemTrabalhou)}
                </div>
              </div>
              <div style="background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 6px; overflow: hidden;">
                <div style="background: #1e3a5f; padding: 8px 10px; display: flex; align-items: center; gap: 6px;">
                  <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="white" stroke-width="2"><path d="M14.7 6.3a1 1 0 0 0 0 1.4l1.6 1.6a1 1 0 0 0 1.4 0l3.77-3.77a6 6 0 0 1-7.94 7.94l-6.91 6.91a2.12 2.12 0 0 1-3-3l6.91-6.91a6 6 0 0 1 7.94-7.94l-3.76 3.76z"></path></svg>
                  <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="white" stroke-width="2" style="margin-left:4px"><path d="M19 17h2c.6 0 1-.4 1-1v-3c0-.9-.7-1.7-1.5-1.9C18.7 10.6 16 10 16 10s-1.3-1.4-2.2-2.3c-.5-.6-1.1-.7-1.8-.7H5c-.6 0-1.1.4-1.4.9l-1.4 2.9A3.7 3.7 0 0 0 2 12v4c0 .6.4 1 1 1h2"></path><circle cx="7" cy="17" r="2"></circle><path d="M9 17h6"></path><circle cx="17" cy="17" r="2"></circle></svg>
                  <span style="color: white; font-size: 10px; font-weight: 600; text-transform: uppercase;">Equipamentos e Veículos</span>
                </div>
                <div style="padding: 10px; font-size: 10px; min-height: 40px; color: #334155; line-height: 1.6;">
                  ${formatAsBulletList([report.equipamentos, report.veiculos].filter(v => v && v.trim() && v.trim() !== '-').join(', '))}
                </div>
              </div>
            </div>
          </div>

          <!-- Section 5: Photos - Clean with Icon -->
          <div data-section="photos" style="margin-bottom: 14px; page-break-inside: avoid;">
            <h3 style="font-size: 11px; font-weight: bold; color: #1e3a5f; margin: 0 0 8px 0; text-transform: uppercase; letter-spacing: 0.5px; display: flex; align-items: center; gap: 6px;">
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#1e3a5f" stroke-width="2"><rect x="3" y="3" width="18" height="18" rx="2" ry="2"></rect><circle cx="8.5" cy="8.5" r="1.5"></circle><polyline points="21 15 16 10 5 21"></polyline></svg>
              Registro Fotográfico
            </h3>
            <div style="display: grid; grid-template-columns: repeat(2, 1fr); gap: 12px;">
              ${generatePhotoGrid(fotosAntesBase64, 'ANTES', '#f8fafc', '#e2e8f0', '#1e3a5f', '#ffffff')}
              ${generatePhotoGrid(fotosDepoisBase64, 'DEPOIS', '#f8fafc', '#e2e8f0', '#1e3a5f', '#ffffff')}
            </div>
          </div>

          <!-- Section 6: Annotations - Clean with Icon -->
          <div data-section="annotations" style="background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 6px; padding: 14px; margin-bottom: 14px;">
            <h3 style="font-size: 11px; font-weight: bold; color: #1e3a5f; margin: 0 0 12px 0; padding-bottom: 8px; border-bottom: 1px solid #e2e8f0; text-transform: uppercase; letter-spacing: 0.5px; display: flex; align-items: center; gap: 6px;">
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#1e3a5f" stroke-width="2"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"></path><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"></path></svg>
              Anotações e Correções
            </h3>
            <div style="display: grid; gap: 10px;">
              <div>
                <p style="font-size: 9px; color: #64748b; margin: 0 0 4px 0; text-transform: uppercase; display: flex; align-items: center; gap: 4px;">
                  <svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="#64748b" stroke-width="2"><line x1="21" y1="10" x2="3" y2="10"></line><line x1="21" y1="6" x2="3" y2="6"></line><line x1="21" y1="14" x2="3" y2="14"></line><line x1="21" y1="18" x2="3" y2="18"></line></svg>
                  Observações Adicionais
                </p>
                <div style="background: white; border: 1px solid #e2e8f0; padding: 8px; border-radius: 4px; font-size: 10px; white-space: pre-wrap; min-height: 24px; color: #334155;">
                  ${report.observacoesPreview || '-'}
                </div>
              </div>
              <div>
                <p style="font-size: 9px; color: #64748b; margin: 0 0 4px 0; text-transform: uppercase; display: flex; align-items: center; gap: 4px;">
                  <svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="#64748b" stroke-width="2"><path d="M12 20h9"></path><path d="M16.5 3.5a2.121 2.121 0 0 1 3 3L7 19l-4 1 1-4L16.5 3.5z"></path></svg>
                  Correções da Fiscalização
                </p>
                <div style="background: white; border: 1px solid #e2e8f0; padding: 8px; border-radius: 4px; font-size: 10px; white-space: pre-wrap; min-height: 24px; color: #334155;">
                  ${report.correcaoFiscalizadora || '-'}
                </div>
              </div>
              <div>
                <p style="font-size: 9px; color: #64748b; margin: 0 0 4px 0; text-transform: uppercase; display: flex; align-items: center; gap: 4px;">
                  <svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="#64748b" stroke-width="2"><path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"></path><line x1="12" y1="9" x2="12" y2="13"></line><line x1="12" y1="17" x2="12.01" y2="17"></line></svg>
                  Pendências
                </p>
                <div style="background: white; border: 1px solid #e2e8f0; padding: 8px; border-radius: 4px; font-size: 10px; white-space: pre-wrap; min-height: 24px; color: #334155;">
                  ${report.pendencias || '-'}
                </div>
              </div>
            </div>
          </div>

          <!-- Section 7: Signatures - Clean with Icon -->
          <div data-section="signatures" style="margin-top: 24px; margin-bottom: 16px;">
            <h3 style="font-size: 11px; font-weight: bold; color: #1e3a5f; margin: 0 0 12px 0; text-transform: uppercase; letter-spacing: 0.5px; display: flex; align-items: center; gap: 6px;">
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#1e3a5f" stroke-width="2"><path d="M17 3a2.828 2.828 0 1 1 4 4L7.5 20.5 2 22l1.5-5.5L17 3z"></path></svg>
              Assinaturas
            </h3>
            <div style="display: grid; grid-template-columns: repeat(3, 1fr); gap: 16px;">
              <div style="text-align: center;">
                <div style="border-bottom: 1px solid #334155; margin-bottom: 8px; height: 40px;"></div>
                <p style="font-size: 9px; font-weight: 600; color: #334155; margin: 0; text-transform: uppercase;">Supervisor Responsável</p>
              </div>
              <div style="text-align: center;">
                <div style="border-bottom: 1px solid #334155; margin-bottom: 8px; height: 40px;"></div>
                <p style="font-size: 9px; font-weight: 600; color: #334155; margin: 0; text-transform: uppercase;">Fiscalizadora</p>
              </div>
              <div style="text-align: center;">
                ${prepostoSignature ? `
                  <div style="border-bottom: 1px solid #334155; margin-bottom: 8px; height: 40px; display: flex; align-items: center; justify-content: center;">
                    <img src="${prepostoSignature}" alt="Assinatura Preposto" style="max-height: 38px; max-width: 100%; object-fit: contain;" />
                  </div>
                ` : `
                  <div style="border-bottom: 1px solid #334155; margin-bottom: 8px; height: 40px;"></div>
                `}
                <p style="font-size: 9px; font-weight: 600; color: #334155; margin: 0; text-transform: uppercase;">Preposto</p>
              </div>
            </div>
          </div>

          <!-- Footer -->
          <div style="display: flex; justify-content: space-between; align-items: center; padding-top: 10px; border-top: 1px solid #e2e8f0; margin-top: 12px;">
            <p style="font-size: 8px; color: #94a3b8; margin: 0;">${report.rdoId}</p>
            <p style="font-size: 8px; color: #94a3b8; margin: 0;">Gerado em ${format(new Date(), 'dd/MM/yyyy')} às ${format(new Date(), 'HH:mm')}</p>
          </div>
        </div>
      `;

      // Wait for images to load
      const images = container.querySelectorAll('img');
      await Promise.all(Array.from(images).map(img => 
        new Promise<void>((resolve) => {
          if (img.complete) {
            resolve();
          } else {
            img.onload = () => resolve();
            img.onerror = () => resolve();
          }
        })
      ));

      // Give a small delay for fonts to load
      await new Promise(resolve => setTimeout(resolve, 100));

      // Get section boundaries before capturing
      const sections = container.querySelectorAll('[data-section]');
      const sectionBounds: { name: string; top: number; bottom: number }[] = [];
      sections.forEach((section) => {
        const el = section as HTMLElement;
        sectionBounds.push({
          name: el.dataset.section || '',
          top: el.offsetTop,
          bottom: el.offsetTop + el.offsetHeight
        });
      });

      // Capture with html2canvas
      const canvas = await html2canvas(container, {
        scale: 2,
        useCORS: true,
        allowTaint: true,
        backgroundColor: '#ffffff',
        logging: false,
      });

      // Remove the container
      document.body.removeChild(container);

      // Create PDF from canvas - A4 format with proper margins
      const pdf = new jsPDF('p', 'mm', 'a4');
      const pageWidth = pdf.internal.pageSize.getWidth(); // 210mm
      const pageHeight = pdf.internal.pageSize.getHeight(); // 297mm
      const margin = 10; // 10mm margins on each side
      const contentWidth = pageWidth - (margin * 2); // 190mm usable width

      // Calculate how the canvas maps to PDF dimensions
      const imgWidth = contentWidth;
      const imgHeight = (canvas.height * imgWidth) / canvas.width;

      // Calculate pixels per mm (at scale 2)
      const scale = 2;
      const pxPerMm = (canvas.height / scale) / imgHeight * scale;
      const pageHeightPx = pageHeight * pxPerMm;

      // If content fits in one page, simple case
      if (imgHeight <= pageHeight - (margin * 2)) {
        const imgData = canvas.toDataURL('image/png');
        pdf.addImage(imgData, 'PNG', margin, margin, imgWidth, imgHeight);
      } else {
        // Multi-page with intelligent section-aware page breaks
        let sourceY = 0;
        let pageNum = 0;

        // Scale section bounds to match canvas scale
        const scaledBounds = sectionBounds.map(s => ({
          name: s.name,
          top: s.top * scale,
          bottom: s.bottom * scale
        }));

        while (sourceY < canvas.height) {
          if (pageNum > 0) {
            pdf.addPage();
          }

          // Calculate theoretical page end
          let pageEndY = sourceY + pageHeightPx;
          const originalPageEndY = pageEndY;

          // Check if we would cut through a section
          if (pageEndY < canvas.height) {
            // Find if page break falls inside any section
            for (const section of scaledBounds) {
              // If page break would cut this section
              if (pageEndY > section.top && pageEndY < section.bottom) {
                // Only move if section starts after current position (avoid zero progress)
                if (section.top > sourceY) {
                  pageEndY = section.top;
                }
                // If section is taller than a page, we must cut through it
                // In this case, keep the original pageEndY
                break;
              }
            }
          }

          // Safety check: ensure we always make progress
          if (pageEndY <= sourceY) {
            pageEndY = originalPageEndY;
          }

          // Calculate how much of the source canvas to grab
          const remainingPx = canvas.height - sourceY;
          const thisChunkPx = Math.min(pageEndY - sourceY, remainingPx);
          const thisChunkMm = thisChunkPx / pxPerMm;

          // Safety: minimum chunk size to prevent infinite loop
          if (thisChunkPx <= 0) {
            break;
          }

          // Create a temporary canvas for this page section
          const pageCanvas = document.createElement('canvas');
          pageCanvas.width = canvas.width;
          pageCanvas.height = thisChunkPx;
          const ctx = pageCanvas.getContext('2d');
          if (ctx) {
            ctx.fillStyle = '#ffffff';
            ctx.fillRect(0, 0, pageCanvas.width, pageCanvas.height);
            ctx.drawImage(canvas, 0, sourceY, canvas.width, thisChunkPx, 0, 0, canvas.width, thisChunkPx);
          }

          const imgData = pageCanvas.toDataURL('image/png');
          pdf.addImage(imgData, 'PNG', margin, margin, imgWidth, thisChunkMm);

          sourceY = pageEndY;
          pageNum++;

          // Safety: prevent infinite loops (max 50 pages)
          if (pageNum > 50) {
            console.warn('PDF generation: max pages reached');
            break;
          }
        }
      }

      // Generate filename
      const dateStr = report.data ? format(report.data, 'yyyy-MM-dd') : 'sem-data';
      const tecSan = report.tecnico.replace(/[^a-zA-Z0-9]/g, '_').substring(0, 20);
      const filename = `RDO_${dateStr}_${report.om || 'sem-om'}_${tecSan}.pdf`;

      if (returnBlob) {
        return pdf.output('blob');
      }

      pdf.save(filename);
      return null;
    } catch (error) {
      console.error('Error generating PDF:', error);
      throw error;
    }
  };

  // --- Helper: Generate PDF and handle UI state ---
  const handleGeneratePDF = async (report: ReportData) => {
    setIsGeneratingPDF(true);
    try {
      await generatePreviewPDF(report);
      toast({
        title: "PDF Gerado!",
        description: "Arquivo baixado com sucesso.",
      });
    } catch (error) {
      toast({
        title: "Erro ao gerar PDF",
        description: "Ocorreu um erro ao gerar o arquivo PDF.",
        variant: "destructive"
      });
    } finally {
      setIsGeneratingPDF(false);
    }
  };

  // --- Helper: Generate all PDFs as ZIP (optimized with parallel image loading) ---
  const [isGeneratingZip, setIsGeneratingZip] = useState(false);
  const [zipProgress, setZipProgress] = useState(0);
  const [zipStatus, setZipStatus] = useState<string>("");

  const generateAllPDFsZip = async () => {
    const reportsToProcess = filteredReports.filter(r => selectedReportIds.has(r.id));
    if (reportsToProcess.length === 0) {
      toast({
        title: "Nenhum relatório selecionado",
        description: "Selecione ao menos um relatório para gerar o ZIP.",
        variant: "destructive"
      });
      return;
    }

    setIsGeneratingZip(true);
    setZipProgress(0);
    setZipStatus("Preparando imagens...");

    try {
      // PHASE 1: Pre-load ALL images in parallel (biggest time saver)
      const allImageUrls = new Set<string>();
      for (const report of reportsToProcess) {
        report.fotosAntes.slice(0, 8).forEach(url => allImageUrls.add(url));
        report.fotosDepois.slice(0, 8).forEach(url => allImageUrls.add(url));
      }
      
      // Pre-load all unique images with high concurrency
      const urlArray = Array.from(allImageUrls);
      if (urlArray.length > 0) {
        let loadedCount = 0;
        const batchSize = 3;
        for (let i = 0; i < urlArray.length; i += batchSize) {
          const batch = urlArray.slice(i, i + batchSize);
          await Promise.all(batch.map(url => downloadImageAsBase64(url)));
          loadedCount += batch.length;
          setZipProgress(Math.round((loadedCount / urlArray.length) * 30));
          setZipStatus(`Baixando imagens: ${Math.min(loadedCount, urlArray.length)}/${urlArray.length}`);
        }
      }

      // PHASE 2: Generate PDFs (images are now cached)
      setZipStatus("Gerando PDFs...");
      const zip = new JSZip();
      let completed = 0;

      // Build filename occurrence map to add sequential index
      const filenameCount: Record<string, number> = {};
      const filenameIndex: Record<string, number> = {};

      for (const report of reportsToProcess) {
        const dateStr = report.data ? format(report.data, 'yyyy-MM-dd') : 'sem-data';
        const tecSan = report.tecnico.replace(/[^a-zA-Z0-9]/g, '_').substring(0, 20);
        const key = `${dateStr}_${report.om || 'sem-om'}_${tecSan}`;
        filenameCount[key] = (filenameCount[key] || 0) + 1;
      }

      // Generate PDFs in batches of 2 to avoid memory issues
      const pdfBatchSize = 2;
      for (let i = 0; i < reportsToProcess.length; i += pdfBatchSize) {
        const batch = reportsToProcess.slice(i, i + pdfBatchSize);
        
        await Promise.all(batch.map(async (report) => {
          try {
            const blob = await generatePreviewPDF(report, true);
            if (blob) {
              const dateStr = report.data ? format(report.data, 'yyyy-MM-dd') : 'sem-data';
              const tecSan = report.tecnico.replace(/[^a-zA-Z0-9]/g, '_').substring(0, 20);
              const key = `${dateStr}_${report.om || 'sem-om'}_${tecSan}`;

              let filename = `RDO_${key}.pdf`;
              if (filenameCount[key] > 1) {
                filenameIndex[key] = (filenameIndex[key] || 0) + 1;
                filename = `RDO_${key}_${filenameIndex[key]}.pdf`;
              }

              zip.file(filename, blob);
            }
          } catch (error) {
            console.error(`Error generating PDF for report ${report.id}:`, error);
          }
        }));

        completed += batch.length;
        const pdfProgress = 30 + Math.round((completed / reportsToProcess.length) * 60); // 30-90%
        setZipProgress(pdfProgress);
        setZipStatus(`Gerando PDF: ${completed}/${reportsToProcess.length}`);
        
        // Small delay to let UI update
        await new Promise(r => setTimeout(r, 10));
      }

      // PHASE 3: Compress and download
      setZipStatus("Compactando...");
      setZipProgress(95);
      
      const zipBlob = await zip.generateAsync({ 
        type: 'blob',
        compression: 'DEFLATE',
        compressionOptions: { level: 6 }
      });
      
      const dateStr = format(new Date(), 'yyyy-MM-dd_HHmm');
      saveAs(zipBlob, `RDOs_${dateStr}.zip`);

      setZipProgress(100);
      toast({
        title: "ZIP Gerado!",
        description: `${reportsToProcess.length} PDFs compactados com sucesso.`,
      });
    } catch (error) {
      console.error('Error generating ZIP:', error);
      toast({
        title: "Erro ao gerar ZIP",
        description: "Ocorreu um erro ao gerar o arquivo ZIP.",
        variant: "destructive"
      });
    } finally {
      setIsGeneratingZip(false);
      setZipProgress(0);
      setZipStatus("");
    }
  };

  // --- Helper: Download image from Google Drive via proxy (for Excel) ---
  const downloadGoogleDriveImage = async (url: string): Promise<{ base64: string; extension: string } | null> => {
    try {
      const fileId = extractFileId(url);
      if (!fileId) return null;

      const tryFetch = async (fetchUrl: string) => {
        const response = await fetch('/api/proxy/image', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ url: fetchUrl })
        });
        if (!response.ok) return null;
        const result = await response.json();
        if (!result.success || !result.data) return null;
        let extension = 'jpeg';
        if (result.contentType?.includes('png')) extension = 'png';
        else if (result.contentType?.includes('gif')) extension = 'gif';
        return { base64: result.data as string, extension };
      };

      const thumbnailUrl = `https://drive.google.com/thumbnail?id=${fileId}&sz=w800`;
      let result = await tryFetch(thumbnailUrl);
      if (result) return result;

      await new Promise(r => setTimeout(r, 500));
      const directUrl = `https://drive.google.com/uc?export=download&id=${fileId}`;
      result = await tryFetch(directUrl);
      return result;
    } catch (error) {
      console.error('Error downloading image:', error);
      return null;
    }
  };

  // --- Derived Data (Filters) ---
  const technicians = useMemo(() => {
    const techs = new Set(allReports.map(r => r.tecnico).filter(Boolean));
    return Array.from(techs).sort();
  }, [allReports]);

  const filteredReports = useMemo(() => {
    return allReports.filter(report => {
      // Date Filter
      if (dateStart && report.data && report.data < startOfDay(dateStart)) return false;
      if (dateEnd && report.data && report.data > endOfDay(dateEnd)) return false;

      // Technician Filter
      if (selectedTechnicians.length > 0 && !selectedTechnicians.includes(report.tecnico)) return false;

      return true;
    });
  }, [allReports, dateStart, dateEnd, selectedTechnicians]);

  // --- Handlers ---

  const handleSelectAll = (checked: boolean) => {
    if (checked) {
      const allIds = new Set(filteredReports.map(r => r.id));
      setSelectedReportIds(allIds);
    } else {
      setSelectedReportIds(new Set());
    }
  };

  const handleSelectReport = (id: string, checked: boolean) => {
    const newSet = new Set(selectedReportIds);
    if (checked) newSet.add(id);
    else newSet.delete(id);
    setSelectedReportIds(newSet);
  };

  const fetchGoogleSheet = async () => {
    if (!sheetUrl) return;
    setIsFetching(true);
    try {
      const response = await fetch('/api/proxy/google-sheet', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ url: toCSVExportUrl(sheetUrl) })
      });

      if (!response.ok) {
        throw new Error('Falha ao buscar planilha');
      }

      const result = await response.json();

      if (!result.success) {
        throw new Error(result.error || 'Erro desconhecido');
      }

      // Process CSV text directly (preserves UTF-8 encoding)
      // raw: true preserves dates as strings to avoid MM/DD vs DD/MM confusion
      const workbook = XLSX.read(result.data, { type: "string", raw: true, cellDates: false });
      processWorkbook(workbook);

    } catch (error: any) {
      console.error(error);
      toast({
        title: "Erro de Conexão",
        description: error.message || "Não foi possível buscar a planilha.",
        variant: "destructive"
      });
    } finally {
      setIsFetching(false);
    }
  };

  // --- Auto-sync: Fetch Google Sheet automatically on page load (once per session) ---
  useEffect(() => {
    if (isInitialLoadComplete && sheetUrl && !hasAutoSynced) {
      setHasAutoSynced(true);
      fetchGoogleSheet();
    }
  }, [isInitialLoadComplete, sheetUrl, hasAutoSynced]);

  const handleFileUpload = (file: File) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        // raw: true preserves dates as strings to avoid MM/DD vs DD/MM confusion
        const workbook = XLSX.read(data, { type: "array", raw: true, cellDates: false });
        processWorkbook(workbook);
      } catch (error) {
        toast({ title: "Erro", description: "Arquivo inválido", variant: "destructive" });
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const processWorkbook = (workbook: XLSX.WorkBook) => {
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];

    // Read all rows as raw arrays to detect where the real headers are
    const allRows: any[][] = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: true });

    // Find the row that contains "Carimbo de data/hora" — that is the real header row
    // This handles Excel files where row 1 has column letters (A, B, C...) and row 2 has real headers
    const headerRowIndex = allRows.findIndex((row: any[]) =>
      row.some((cell: any) => String(cell ?? '').trim() === 'Carimbo de data/hora')
    );

    // If not found, fall back to default first-row-as-header parsing
    let jsonData: Record<string, any>[];
    if (headerRowIndex === -1) {
      const rawJsonData = XLSX.utils.sheet_to_json(worksheet, { raw: true });
      jsonData = rawJsonData.map((row: any) => {
        const normalized: Record<string, any> = {};
        for (const key of Object.keys(row)) {
          normalized[key.trim()] = row[key];
        }
        return normalized;
      });
    } else {
      // Use the found row as headers; data starts on the next row
      const headers: string[] = (allRows[headerRowIndex] as any[]).map((h: any) =>
        String(h ?? '').trim()
      );
      jsonData = allRows.slice(headerRowIndex + 1)
        .map((row: any[]) => {
          const obj: Record<string, any> = {};
          headers.forEach((h, i) => {
            if (h) obj[h] = row[i] ?? '';
          });
          return obj;
        })
        // Keep only rows that have a real timestamp or date value (skip blank/formatting rows)
        .filter((obj: Record<string, any>) => {
          const ts = obj['Carimbo de data/hora'];
          const dt = obj['Data'];
          return (ts !== null && ts !== undefined && String(ts).trim() !== '') ||
                 (dt !== null && dt !== undefined && String(dt).trim() !== '' && String(dt).trim() !== '0');
        });
    }

    // Final safety: remove rows without any meaningful data (applies to both paths)
    const jsonData2 = jsonData.filter((row: Record<string, any>) => {
      const ts = row['Carimbo de data/hora'];
      const dt = row['Data'];
      return (ts !== null && ts !== undefined && String(ts).trim() !== '') ||
             (dt !== null && dt !== undefined && String(dt).trim() !== '' && String(dt).trim() !== '0');
    });

    console.log("Raw Data:", jsonData2.slice(0, 2));
    console.log("Column Names:", Object.keys(jsonData2[0] || {}));

    const parsed: ReportData[] = jsonData2.map((row: any, index: number) => {
       let parsedDate: Date | null = null;
       if (row['Data']) {
         if (typeof row['Data'] === 'number') {
           parsedDate = new Date((row['Data'] - (25567 + 1)) * 86400 * 1000);
         } else {
           const dateStr = String(row['Data']).trim();
           // Tentar formatos brasileiros (dd/mm/yyyy) - NUNCA usar new Date() diretamente
           const formats = ['dd/MM/yyyy', 'd/M/yyyy', 'dd-MM-yyyy', 'd-M-yyyy'];
           for (const fmt of formats) {
             try {
               const testDate = parse(dateStr, fmt, new Date());
               if (!isNaN(testDate.getTime())) {
                 parsedDate = testDate;
                 break;
               }
             } catch (e) { /* tentar próximo formato */ }
           }
           // Se ainda não conseguiu, tentar parse manual dd/mm/yyyy
           if (!parsedDate) {
             const parts = dateStr.split(/[\/\-]/);
             if (parts.length === 3) {
               const day = parseInt(parts[0], 10);
               const month = parseInt(parts[1], 10) - 1;
               const year = parseInt(parts[2], 10);
               if (!isNaN(day) && !isNaN(month) && !isNaN(year)) {
                 parsedDate = new Date(year, month, day);
               }
             }
           }
         }
       }

       const observacoesGerais = row['Observações Gerais'] || row['ObservaÃ§Ãµes Gerais'] || '';

       // Planilha nova: "Supervisor Responsável"; planilha antiga: "Técnico Responsável"
       const supervisor = row['Supervisor Responsável'] || row['TÃ©cnico ResponsÃ¡vel'] || row['Técnico Responsável'] || "S/I";
       const tecnico = supervisor; // campo principal de identificação
       const encarregado = row['Encarregado Responsável'] || '';
       const obra = row['Obra'] || ''; // mantido para planilhas antigas
       const localidade = row['Localidade'] || "S/I";
       const area = row['Área'] || row['Area'] || '';
       const tag = row['TAG'] || '';
       const escalaTrabalho = row['Escala de Trabalho'] || '';
       const om = row['Nº OM'] || row['Nº OM/Forms'] || row['NÂº OM/Forms'] || "S/I";

       const rdoId = generateRdoId({ tecnico, obra, localidade, data: parsedDate, om });

       return {
         id: `rep-${index}`,
         rdoId,
         // OBRIGATÓRIAS
         data: parsedDate,
         tecnico,
         atividade: row['Atividade Realizada'] || "S/I",
         om,
         mina: row['Mina'] || "S/I",
         localidade,
         // NOVOS CAMPOS
         supervisor,
         encarregado,
         area,
         tag,
         escalaTrabalho,
         // OPCIONAIS (com fallbacks para encoding)
         obra,
         horarioInicio: row['Horário de Início'] || row['HorÃ¡rio de InÃ­cio'],
         horarioTermino: row['Horário de término '] || row['Horário de término'] || row['HorÃ¡rio de tÃ©rmino '],
         observacoes: observacoesGerais,
         observacoesPreview: observacoesGerais,
         tarefaConcluida: row['Tarefa concluída ou haverá retorno?'] || row['Tarefa concluÃ­da ou haverÃ¡ retorno?'],
         atividadePlanejada: row['Atividade estava planejada?'] || row['Atividade estava planejada para a semana?'] || row['Atividade estava planejada para a semana'] || '',
         quemTrabalhou: row['Quem trabalhou com você?'] || row['Quem trabalhou com vocÃª?'],
         equipamentos: row['Equipamentos'] || '',
         veiculos: row['Veículos e Equipamentos Móveis'] || row['VeÃ­culos e Equipamentos MÃ³veis'] || '',
         climaMatutino: row['Clima Matutino'] || '',
         climaVespertino: row['Clima Vespertino'] || '',
         climaNoturno: row['Clima Noturno'] || '',
         fotosAntes: [
           ...String(row['Fotos das Atividades (Antes)'] || "").split(',').map((s: string) => s.trim()).filter(Boolean),
           ...String(row['Fotos das Atividades (Antes) Extra'] || "").split(',').map((s: string) => s.trim()).filter(Boolean),
         ].filter((url, idx, arr) => arr.indexOf(url) === idx),
         fotosDepois: [
           ...String(row['Fotos das Atividades (Depois)'] || "").split(',').map((s: string) => s.trim()).filter(Boolean),
           ...String(row['Fotos das Atividades (Depois) Extra'] || "").split(',').map((s: string) => s.trim()).filter(Boolean),
         ].filter((url, idx, arr) => arr.indexOf(url) === idx),
         status: "ready"
       };
    });

    console.log("Parsed Reports:", parsed.slice(0, 2));
    setAllReports(parsed);

    toast({ 
      title: "✓ Dados Carregados", 
      description: `${parsed.length} relatórios importados.` 
    });
  };

  // --- Generation Logic ---

  const generateExcelRDOs = async () => {
    if (!maskExcel) {
      toast({ title: "Falta Máscara Excel", description: "Carregue o arquivo .xlsx de máscara.", variant: "destructive" });
      return;
    }

    const reportsToProcess = filteredReports.filter(r => selectedReportIds.has(r.id));
    if (reportsToProcess.length === 0) return;

    setIsGeneratingExcel(true);
    setExcelProgress(0);

    try {
      const zip = new JSZip();
      const maskBuffer = await maskExcel.arrayBuffer();

      let completed = 0;

      // Build filename occurrence map to add sequential index and prevent overwrites
      const filenameCount: Record<string, number> = {};
      const filenameIndex: Record<string, number> = {};

      for (const report of reportsToProcess) {
        const obraSan = (report.obra || 'SEM-OBRA').replace(/[^a-zA-Z0-9\-_]/g, '_').substring(0, 30);
        const tecSan = report.tecnico.replace(/[^a-zA-Z0-9\-_]/g, '_').substring(0, 30);
        const dateStr = report.data ? format(report.data, 'dd-MM-yyyy') : 'DATA';
        const key = `${dateStr}_${obraSan}_${tecSan}`;
        filenameCount[key] = (filenameCount[key] || 0) + 1;
      }

      for (const report of reportsToProcess) {
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(maskBuffer);
        const sheet = workbook.getWorksheet(1);

        if (sheet) {
          // ========== MAPEAMENTO RDO EXCEL (conforme script Python original) ==========
          // Cabeçalho RDO:
          sheet.getCell('S1').value = report.data ? format(report.data, 'dd/MM/yyyy') : "S/I"; // Data → row=1, col=19
          sheet.getCell('Q57').value = report.supervisor || report.tecnico; // Supervisor Responsável → row=57, col=17
          sheet.getCell('F57').value = report.area || report.obra || "N/A"; // Área → row=57, col=6
          sheet.getCell('F58').value = report.om; // Nº OM → row=58, col=6
          sheet.getCell('N58').value = report.horarioInicio || ""; // Horário de Início → row=58, col=14
          sheet.getCell('S58').value = report.horarioTermino || ""; // Horário de término → row=58, col=19
          sheet.getCell('E59').value = report.atividade; // Atividade Realizada → row=59, col=5
          sheet.getCell('E67').value = report.observacoes || ""; // Observações Gerais → row=67, col=5
          sheet.getCell('R67').value = report.tarefaConcluida || ""; // Tarefa concluída → row=67, col=18

          // Listas (linha 17 pra baixo):
          if (report.quemTrabalhou) {
            const pessoas = report.quemTrabalhou.split(',');
            pessoas.forEach((pessoa, i) => {
              sheet.getCell(17 + i, 4).value = pessoa.trim(); // coluna D
            });
          }
          if (report.equipamentos) {
            const equips = report.equipamentos.split(',');
            equips.forEach((eq, i) => {
              sheet.getCell(17 + i, 13).value = eq.trim(); // coluna M
            });
          }
          if (report.veiculos) {
            const veics = report.veiculos.split(',');
            veics.forEach((v, i) => {
              sheet.getCell(17 + i, 19).value = v.trim(); // coluna S
            });
          }

          // Clima (linhas 12-14, colunas 12-14 com "✓" ou "N/A"):
          const climaMapping = {
            'Bom': { row: 12, symbol: '✓' },
            'Chuvoso': { row: 13, symbol: '✓' },
            'Nublado': { row: 14, symbol: '✓' }
          };
          // Clima Matutino (col 12 = L)
          if (report.climaMatutino && climaMapping[report.climaMatutino as keyof typeof climaMapping]) {
            const { row, symbol } = climaMapping[report.climaMatutino as keyof typeof climaMapping];
            sheet.getCell(row, 12).value = symbol;
          }
          // Clima Vespertino (col 13 = M)
          if (report.climaVespertino && climaMapping[report.climaVespertino as keyof typeof climaMapping]) {
            const { row, symbol } = climaMapping[report.climaVespertino as keyof typeof climaMapping];
            sheet.getCell(row, 13).value = symbol;
          }
          // Clima Noturno (col 14 = N)
          if (report.climaNoturno && climaMapping[report.climaNoturno as keyof typeof climaMapping]) {
            const { row, symbol } = climaMapping[report.climaNoturno as keyof typeof climaMapping];
            sheet.getCell(row, 14).value = symbol;
          }

          // Fotos:
          // Antes → E65, I65, E66, I66 (slots 1-4) + E67, I67, E68, I68 (slots 5-8)
          // Depois → N65, S65, N66, S66 (slots 1-4) + N67, S67, N68, S68 (slots 5-8)
          const fotosAntesCells = [
            { cell: 'E65', col: 5, row: 65 },
            { cell: 'I65', col: 9, row: 65 },
            { cell: 'E66', col: 5, row: 66 },
            { cell: 'I66', col: 9, row: 66 },
            { cell: 'E67', col: 5, row: 67 },
            { cell: 'I67', col: 9, row: 67 },
            { cell: 'E68', col: 5, row: 68 },
            { cell: 'I68', col: 9, row: 68 }
          ];
          const fotosDepoisCells = [
            { cell: 'N65', col: 14, row: 65 },
            { cell: 'S65', col: 19, row: 65 },
            { cell: 'N66', col: 14, row: 66 },
            { cell: 'S66', col: 19, row: 66 },
            { cell: 'N67', col: 14, row: 67 },
            { cell: 'S67', col: 19, row: 67 },
            { cell: 'N68', col: 14, row: 68 },
            { cell: 'S68', col: 19, row: 68 }
          ];

          // Download and insert images for "Antes"
          let skippedImages = 0;
          for (let i = 0; i < Math.min(report.fotosAntes.length, 8); i++) {
            const imageData = await downloadGoogleDriveImage(report.fotosAntes[i]);
            if (imageData) {
              const imageId = workbook.addImage({
                base64: imageData.base64,
                extension: imageData.extension as 'png' | 'jpeg' | 'gif',
              });
              sheet.addImage(imageId, {
                tl: { col: fotosAntesCells[i].col - 1, row: fotosAntesCells[i].row - 1 },
                ext: { width: 244, height: 162 }
              });
            } else {
              // Fallback: write URL as text when image download fails
              sheet.getCell(fotosAntesCells[i].cell).value = report.fotosAntes[i];
              skippedImages++;
            }
          }

          // Download and insert images for "Depois"
          for (let i = 0; i < Math.min(report.fotosDepois.length, 8); i++) {
            const imageData = await downloadGoogleDriveImage(report.fotosDepois[i]);
            if (imageData) {
              const imageId = workbook.addImage({
                base64: imageData.base64,
                extension: imageData.extension as 'png' | 'jpeg' | 'gif',
              });
              sheet.addImage(imageId, {
                tl: { col: fotosDepoisCells[i].col - 1, row: fotosDepoisCells[i].row - 1 },
                ext: { width: 244, height: 162 }
              });
            } else {
              // Fallback: write URL as text when image download fails
              sheet.getCell(fotosDepoisCells[i].cell).value = report.fotosDepois[i];
              skippedImages++;
            }
          }

          if (skippedImages > 0) {
            console.log(`Report ${report.id}: ${skippedImages} images could not be downloaded, URLs preserved as text`);
          }
        }

        const buffer = await workbook.xlsx.writeBuffer();
        const obraSanitized = (report.obra || 'SEM-OBRA').replace(/[^a-zA-Z0-9\-_]/g, '_').substring(0, 30);
        const tecnicoSanitized = report.tecnico.replace(/[^a-zA-Z0-9\-_]/g, '_').substring(0, 30);
        const dateStr = report.data ? format(report.data, 'dd-MM-yyyy') : 'DATA';
        const key = `${dateStr}_${obraSanitized}_${tecnicoSanitized}`;

        // Increment index for this key
        filenameIndex[key] = (filenameIndex[key] || 0) + 1;
        const indexStr = filenameCount[key] > 1 ? `_${String(filenameIndex[key]).padStart(2, '0')}` : '';

        const fileName = `RDO_${dateStr}_${obraSanitized}_${tecnicoSanitized}${indexStr}.xlsx`;
        zip.file(fileName, buffer);

        completed++;
        setExcelProgress((completed / reportsToProcess.length) * 100);
      }

      const content = await zip.generateAsync({ type: "blob" });
      saveAs(content, "RDOs_Gerados_Excel.zip");
      toast({ title: "Concluído", description: "RDOs Excel gerados com sucesso." });

    } catch (error) {
      console.error(error);
      toast({ title: "Erro", description: "Falha na geração dos Excel.", variant: "destructive" });
    } finally {
      setIsGeneratingExcel(false);
    }
  };

  // Helper to convert base64 to Uint8Array for docxtemplater image module
  const base64ToUint8Array = (base64: string): Uint8Array => {
    const binaryString = atob(base64);
    const len = binaryString.length;
    const bytes = new Uint8Array(len);
    for (let i = 0; i < len; i++) {
      bytes[i] = binaryString.charCodeAt(i);
    }
    return bytes;
  };

  const generateWordReports = async () => {
    const reportsToProcess = filteredReports.filter(r => selectedReportIds.has(r.id));
    if (reportsToProcess.length === 0) return;

    if (!maskWord) {
      toast({ title: "Modelo Word", description: "Por favor, carregue um modelo Word (.docx) com placeholders.", variant: "destructive" });
      return;
    }

    setIsGeneratingWord(true);
    setWordProgress(0);

    try {
      // Read the template
      const templateArrayBuffer = await (maskWord as Blob).arrayBuffer();
      const templateZip = new PizZip(templateArrayBuffer);

      const zip = new JSZip();
      let completed = 0;
      let totalImagesFailed = 0;

      // Build filename occurrence map to add sequential index and prevent overwrites
      const filenameCountWord: Record<string, number> = {};
      const filenameIndexWord: Record<string, number> = {};

      for (const report of reportsToProcess) {
        const obraSan = (report.obra || 'SEM-OBRA').replace(/[^a-zA-Z0-9\-_]/g, '_').substring(0, 30);
        const tecSan = report.tecnico.replace(/[^a-zA-Z0-9\-_]/g, '_').substring(0, 30);
        const dateStr = report.data ? format(report.data, 'dd-MM-yyyy') : 'DATA';
        const key = `${dateStr}_${obraSan}_${tecSan}`;
        filenameCountWord[key] = (filenameCountWord[key] || 0) + 1;
      }

      for (const report of reportsToProcess) {
        // Download all images for this report (with extension info)
        // Keys include extension for image module type detection: foto_antes_1.jpg
        const imageBuffers: Record<string, Uint8Array> = {};
        const imageKeys: Record<string, string> = {}; // Maps placeholder name to actual key with extension

        console.log(`[WORD IMG] Report ${report.rdoId} — fotosAntes:`, report.fotosAntes);
        console.log(`[WORD IMG] Report ${report.rdoId} — fotosDepois:`, report.fotosDepois);

        // Download ANTES images (up to 8)
        for (let i = 0; i < Math.min(report.fotosAntes.length, 8); i++) {
          const url = report.fotosAntes[i];
          console.log(`[WORD IMG] Downloading foto_antes_${i + 1}:`, url);
          const imgResult = await downloadGoogleDriveImage(url);
          if (imgResult) {
            const ext = imgResult.extension === 'jpeg' ? 'jpg' : imgResult.extension;
            const key = `foto_antes_${i + 1}.${ext}`;
            imageBuffers[key] = base64ToUint8Array(imgResult.base64);
            imageKeys[`foto_antes_${i + 1}`] = key;
            console.log(`[WORD IMG] ✓ foto_antes_${i + 1} OK → key=${key}, bytes=${imgResult.base64.length}`);
          } else {
            console.warn(`[WORD IMG] ✗ foto_antes_${i + 1} FAILED for URL:`, url);
          }
        }

        // Download DEPOIS images (up to 8)
        for (let i = 0; i < Math.min(report.fotosDepois.length, 8); i++) {
          const url = report.fotosDepois[i];
          console.log(`[WORD IMG] Downloading foto_depois_${i + 1}:`, url);
          const imgResult = await downloadGoogleDriveImage(url);
          if (imgResult) {
            const ext = imgResult.extension === 'jpeg' ? 'jpg' : imgResult.extension;
            const key = `foto_depois_${i + 1}.${ext}`;
            imageBuffers[key] = base64ToUint8Array(imgResult.base64);
            imageKeys[`foto_depois_${i + 1}`] = key;
            console.log(`[WORD IMG] ✓ foto_depois_${i + 1} OK → key=${key}, bytes=${imgResult.base64.length}`);
          } else {
            console.warn(`[WORD IMG] ✗ foto_depois_${i + 1} FAILED for URL:`, url);
          }
        }

        const totalFotos = report.fotosAntes.length + report.fotosDepois.length;
        const totalBaixadas = Object.keys(imageKeys).length;
        const failedHere = totalFotos - totalBaixadas;
        if (failedHere > 0) {
          totalImagesFailed += failedHere;
          console.warn(`[WORD IMG] ${failedHere} de ${totalFotos} fotos não puderam ser baixadas para ${report.rdoId} (verifique se são públicas no Drive).`);
        }
        console.log(`[WORD IMG] imageKeys populated:`, imageKeys);

        // Image size: 6.44 cm × 4.29 cm (converted to EMUs for docx)
        // 1 cm = 360000 EMUs, so 6.44 cm = 2318400 EMUs, 4.29 cm = 1544400 EMUs
        // For pixels: 6.44 cm = 244 px, 4.29 cm = 162 px (at 96 DPI)

        // 1x1 transparent PNG as placeholder for missing images
        const transparentPngBase64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==";
        const transparentPngBuffer = base64ToUint8Array(transparentPngBase64);

        const imageModule = new ImageModule({
          centered: false,
          getImage: (tagValue: string) => {
            // tagValue will be like "foto_antes_1.jpg", "foto_antes_1.png", or "placeholder.png"
            // Return image data if exists, otherwise return transparent 1x1 PNG
            if (tagValue === "placeholder.png") {
              return transparentPngBuffer;
            }
            const img = imageBuffers[tagValue];
            if (img) {
              return img;
            }
            // Return transparent placeholder for missing images
            return transparentPngBuffer;
          },
          getSize: (img: Uint8Array, tagValue: string) => {
            // Return actual size for real images, tiny size for placeholders
            if (tagValue !== "placeholder.png" && imageBuffers[tagValue]) {
              return [244, 162]; // Width x Height in pixels (6.44 cm × 4.29 cm)
            }
            return [1, 1]; // Tiny size for missing images
          }
        });

        // Clone the template for each report
        const docZip = new PizZip(templateZip.generate({ type: "arraybuffer" }));

        // Pre-process: replace <<TABELA>> paragraph with a real ANTES/DEPOIS photo table
        // (mirrors what the original Python code did with python-docx)
        const docXmlFile = docZip.file("word/document.xml");
        if (docXmlFile) {
          let docXml = docXmlFile.asText();
          const TABELA_MARKER = "&lt;&lt;TABELA&gt;&gt;";
          const tabelaIdx = docXml.indexOf(TABELA_MARKER);
          if (tabelaIdx !== -1) {
            // Find the enclosing <w:p ...>...</w:p> element
            const pStart = docXml.lastIndexOf("<w:p ", tabelaIdx);
            const pEndPos = docXml.indexOf("</w:p>", tabelaIdx);
            if (pStart !== -1 && pEndPos !== -1) {
              const pEnd = pEndPos + "</w:p>".length;

              // Determine how many photo rows to build
              const antesCount = Object.keys(imageKeys).filter(k => k.startsWith("foto_antes_")).length;
              const depoisCount = Object.keys(imageKeys).filter(k => k.startsWith("foto_depois_")).length;
              const rowCount = Math.max(antesCount, depoisCount, 1);

              // Cell border XML (reusable)
              const cellBorders = `<w:tcBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/></w:tcBorders>`;

              // Build photo rows
              const photoRows: string[] = [];
              for (let ri = 1; ri <= rowCount; ri++) {
                const hasAntes = !!imageKeys[`foto_antes_${ri}`];
                const hasDepois = !!imageKeys[`foto_depois_${ri}`];

                const antesContent = hasAntes
                  ? `<w:r><w:t>&lt;&lt;%foto_antes_${ri}&gt;&gt;</w:t></w:r>`
                  : `<w:r><w:t xml:space="preserve"> </w:t></w:r>`;
                const depoisContent = hasDepois
                  ? `<w:r><w:t>&lt;&lt;%foto_depois_${ri}&gt;&gt;</w:t></w:r>`
                  : `<w:r><w:t xml:space="preserve"> </w:t></w:r>`;

                photoRows.push(
                  `<w:tr><w:trPr><w:trHeight w:val="2160" w:hRule="atLeast"/></w:trPr>` +
                  `<w:tc><w:tcPr><w:tcW w:w="4675" w:type="dxa"/>${cellBorders}</w:tcPr>` +
                  `<w:p><w:pPr><w:jc w:val="center"/></w:pPr>${antesContent}</w:p></w:tc>` +
                  `<w:tc><w:tcPr><w:tcW w:w="4675" w:type="dxa"/>${cellBorders}</w:tcPr>` +
                  `<w:p><w:pPr><w:jc w:val="center"/></w:pPr>${depoisContent}</w:p></w:tc>` +
                  `</w:tr>`
                );
              }

              const tableXml =
                `<w:tbl>` +
                `<w:tblPr>` +
                `<w:tblW w:w="9350" w:type="dxa"/>` +
                `<w:tblBorders>` +
                `<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>` +
                `<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>` +
                `<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>` +
                `<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>` +
                `<w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/>` +
                `<w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/>` +
                `</w:tblBorders>` +
                `<w:tblCellMar><w:top w:w="80" w:type="dxa"/><w:left w:w="80" w:type="dxa"/><w:bottom w:w="80" w:type="dxa"/><w:right w:w="80" w:type="dxa"/></w:tblCellMar>` +
                `</w:tblPr>` +
                // Header row
                `<w:tr>` +
                `<w:tc><w:tcPr><w:tcW w:w="4675" w:type="dxa"/>${cellBorders}</w:tcPr>` +
                `<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>ANTES</w:t></w:r></w:p></w:tc>` +
                `<w:tc><w:tcPr><w:tcW w:w="4675" w:type="dxa"/>${cellBorders}</w:tcPr>` +
                `<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>DEPOIS</w:t></w:r></w:p></w:tc>` +
                `</w:tr>` +
                photoRows.join("") +
                `</w:tbl>`;

              docXml = docXml.slice(0, pStart) + tableXml + docXml.slice(pEnd);
              docZip.file("word/document.xml", docXml);
            }
          }
        }

        const doc = new Docxtemplater(docZip, {
          modules: [imageModule],
          paragraphLoop: true,
          linebreaks: true,
          delimiters: { start: "<<", end: ">>" },
          nullGetter: () => "" // Return empty string for missing/null values
        });

        // Prepare data for placeholders
        // Both lowercase and UPPERCASE keys are provided so the template can use either style
        // e.g. <<mina>> or <<MINA>>, <<tecnico>> or <<NOME_TECNICO>>, <<atividade>> or <<ATIVIDADE_REALIZADA>>
        const obraVal = (report.obra || "S/I").toUpperCase();
        const dataVal = report.data ? format(report.data, "dd/MM/yyyy") : "S/I";
        const tecnicoVal = report.tecnico.toUpperCase();
        const minaVal = (report.mina || "S/I").toUpperCase();
        const localidadeVal = (report.localidade || "S/I").toUpperCase();
        const omVal = String(report.om || "S/I");
        const atividadeVal = report.atividade || "";
        const observacoesVal = report.observacoes || "";
        const data: Record<string, string> = {
          // Lowercase aliases
          obra: obraVal,
          data: dataVal,
          tecnico: tecnicoVal,
          mina: minaVal,
          localidade: localidadeVal,
          om: omVal,
          atividade: atividadeVal,
          observacoes: observacoesVal,
          clima_matutino: report.climaMatutino || "-",
          clima_vespertino: report.climaVespertino || "-",
          clima_noturno: report.climaNoturno || "-",
          quem_trabalhou: report.quemTrabalhou || "-",
          equipamentos: report.equipamentos || "-",
          veiculos: report.veiculos || "-",
          horario_inicio: report.horarioInicio || "-",
          horario_termino: report.horarioTermino || "-",
          tarefa_concluida: report.tarefaConcluida || "-",
          observacoes_preview: report.observacoesPreview || "",
          correcao_fiscalizadora: report.correcaoFiscalizadora || "",
          pendencias: report.pendencias || "",
          // UPPERCASE aliases — matches <<MINA>>, <<NOME_TECNICO>>, <<ATIVIDADE_REALIZADA>>, etc.
          OBRA: obraVal,
          DATA: dataVal,
          NOME_TECNICO: tecnicoVal,
          MINA: minaVal,
          LOCALIDADE: localidadeVal,
          OM: omVal,
          ATIVIDADE_REALIZADA: atividadeVal,
          OBSERVACOES: observacoesVal,
          CLIMA_MATUTINO: report.climaMatutino || "-",
          CLIMA_VESPERTINO: report.climaVespertino || "-",
          CLIMA_NOTURNO: report.climaNoturno || "-",
          QUEM_TRABALHOU: report.quemTrabalhou || "-",
          EQUIPAMENTOS: report.equipamentos || "-",
          VEICULOS: report.veiculos || "-",
          HORARIO_INICIO: report.horarioInicio || "-",
          HORARIO_TERMINO: report.horarioTermino || "-",
          TAREFA_CONCLUIDA: report.tarefaConcluida || "-",
          // Novos campos da planilha
          supervisor: (report.supervisor || report.tecnico).toUpperCase(),
          encarregado: (report.encarregado || "S/I").toUpperCase(),
          area: (report.area || "S/I").toUpperCase(),
          tag: (report.tag || "S/I").toUpperCase(),
          escala_trabalho: (report.escalaTrabalho || "S/I"),
          SUPERVISOR: (report.supervisor || report.tecnico).toUpperCase(),
          ENCARREGADO: (report.encarregado || "S/I").toUpperCase(),
          AREA: (report.area || "S/I").toUpperCase(),
          TAG: (report.tag || "S/I").toUpperCase(),
          ESCALA_TRABALHO: (report.escalaTrabalho || "S/I"),
          TABELA: "",
          // Image placeholders (1-8) — use key with extension if image exists, otherwise placeholder.png
          foto_antes_1: imageKeys["foto_antes_1"] || "placeholder.png",
          foto_antes_2: imageKeys["foto_antes_2"] || "placeholder.png",
          foto_antes_3: imageKeys["foto_antes_3"] || "placeholder.png",
          foto_antes_4: imageKeys["foto_antes_4"] || "placeholder.png",
          foto_antes_5: imageKeys["foto_antes_5"] || "placeholder.png",
          foto_antes_6: imageKeys["foto_antes_6"] || "placeholder.png",
          foto_antes_7: imageKeys["foto_antes_7"] || "placeholder.png",
          foto_antes_8: imageKeys["foto_antes_8"] || "placeholder.png",
          foto_depois_1: imageKeys["foto_depois_1"] || "placeholder.png",
          foto_depois_2: imageKeys["foto_depois_2"] || "placeholder.png",
          foto_depois_3: imageKeys["foto_depois_3"] || "placeholder.png",
          foto_depois_4: imageKeys["foto_depois_4"] || "placeholder.png",
          foto_depois_5: imageKeys["foto_depois_5"] || "placeholder.png",
          foto_depois_6: imageKeys["foto_depois_6"] || "placeholder.png",
          foto_depois_7: imageKeys["foto_depois_7"] || "placeholder.png",
          foto_depois_8: imageKeys["foto_depois_8"] || "placeholder.png",
        };

        doc.render(data);

        const blob = doc.getZip().generate({ 
          type: "blob",
          mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        });

        const obraSanitizedWord = (report.obra || 'SEM-OBRA').replace(/[^a-zA-Z0-9\-_]/g, '_').substring(0, 30);
        const tecnicoSanitizedWord = report.tecnico.replace(/[^a-zA-Z0-9\-_]/g, '_').substring(0, 30);
        const dateStrWord = report.data ? format(report.data, 'dd-MM-yyyy') : 'DATA';
        const keyWord = `${dateStrWord}_${obraSanitizedWord}_${tecnicoSanitizedWord}`;

        // Increment index for this key
        filenameIndexWord[keyWord] = (filenameIndexWord[keyWord] || 0) + 1;
        const indexStrWord = filenameCountWord[keyWord] > 1 ? `_${String(filenameIndexWord[keyWord]).padStart(2, '0')}` : '';

        const fileName = `RELATORIO_${dateStrWord}_${obraSanitizedWord}_${tecnicoSanitizedWord}${indexStrWord}.docx`;
        zip.file(fileName, blob);

        completed++;
        setWordProgress((completed / reportsToProcess.length) * 100);
      }

      const content = await zip.generateAsync({ type: "blob" });
      saveAs(content, "Relatorios_Word_Gerados.zip");
      if (totalImagesFailed > 0) {
        toast({
          title: "Word gerado com avisos",
          description: `${totalImagesFailed} foto(s) não puderam ser incorporadas. Verifique se os arquivos no Google Drive estão compartilhados como "Qualquer pessoa com o link".`,
          variant: "destructive"
        });
      } else {
        toast({ title: "Concluído", description: "Relatórios Word gerados com sucesso." });
      }

    } catch (error) {
      console.error(error);
      toast({ title: "Erro", description: "Falha na geração dos Word.", variant: "destructive" });
    } finally {
      setIsGeneratingWord(false);
    }
  };

  const loadDefaultWord = async () => {
    try {
      const response = await fetch("/template.docx");
      const blob = await response.blob();
      setMaskWord(blob);
      setUsingDefaultWord(true);
      toast({ title: "Modelo Padrão", description: "Modelo Word carregado." });
    } catch(e) {
      toast({ title: "Erro", description: "Não foi possível carregar modelo padrão.", variant: "destructive" });
    }
  };

  return (
    <div className="min-h-screen bg-background font-sans flex flex-col">
      <Header />

      <main className="flex-1 container mx-auto px-4 py-6 flex flex-col lg:flex-row gap-4 lg:gap-6 min-h-0 lg:h-[calc(100vh-80px)]">
        {/* SIDEBAR CONFIG */}
        <Card className="w-full lg:w-[320px] xl:w-[350px] flex-shrink-0 lg:h-full flex flex-col shadow-lg border-t-4 border-t-primary">
           <div className="p-4 bg-muted/20 border-b">
             <h2 className="font-display font-bold text-lg flex items-center gap-2">
               <CloudDownload className="w-5 h-5 text-primary" />
               Entrada de Dados
             </h2>
           </div>
           <CardContent className="p-4 space-y-6 overflow-y-auto flex-1">

             {/* 1. DATA SOURCE */}
             <div className="space-y-3">
               <Label className="text-xs font-bold text-muted-foreground uppercase tracking-wider">1. Planilha Google</Label>
               <div className="flex gap-2">
                 <Input 
                   value={sheetUrl} 
                   onChange={(e) => setSheetUrl(e.target.value)} 
                   className="h-9 text-xs" 
                   placeholder="URL CSV..."
                 />
                 <Button size="icon" className="h-9 w-9" onClick={fetchGoogleSheet} disabled={isFetching}>
                   {isFetching ? <RefreshCw className="w-4 h-4 animate-spin" /> : <Download className="w-4 h-4" />}
                 </Button>
               </div>
               <div className="flex items-center justify-center gap-2 text-xs text-muted-foreground">
                 <span className="h-px w-8 bg-border"></span>
                 OU
                 <span className="h-px w-8 bg-border"></span>
               </div>
               <FileUpload 
                  label="Upload CSV/XLSX Local" 
                  accept=".csv, .xlsx" 
                  onFileSelect={handleFileUpload}
                  icon={<FileSpreadsheet className="w-4 h-4" />}
               />
             </div>

             <div className="h-px bg-border w-full" />

             {/* 2. EXCEL MASK */}
             <div className="space-y-3">
               <Label className="text-xs font-bold text-muted-foreground uppercase tracking-wider">2. Máscara RDO (Excel)</Label>
               <FileUpload 
                  label={maskExcel ? maskExcel.name : "Carregar MascaraRDO.xlsx"}
                  accept=".xlsx" 
                  onFileSelect={setMaskExcel}
                  icon={<FileSpreadsheet className="w-4 h-4 text-green-600" />}
               />
               <p className="text-[10px] text-muted-foreground">
                 Preenche células e insere fotos nas coordenadas E65, I65...
               </p>
             </div>

             <div className="h-px bg-border w-full" />

             {/* 3. WORD MASK */}
             <div className="space-y-3">
               <Label className="text-xs font-bold text-muted-foreground uppercase tracking-wider">3. Modelo Relatório (Word)</Label>
               {usingDefaultWord ? (
                 <div className="flex items-center justify-between p-2 bg-blue-50 border border-blue-100 rounded text-blue-800 text-xs">
                   <span className="flex items-center gap-2 font-medium"><FileCheck className="w-3 h-3"/> Padrão Carregado</span>
                   <Button variant="ghost" size="sm" className="h-5 w-5 p-0" onClick={() => { setMaskWord(null); setUsingDefaultWord(false); }}>×</Button>
                 </div>
               ) : (
                 <div className="space-y-2">
                   <FileUpload 
                      label={maskWord instanceof File ? maskWord.name : "Carregar Modelo.docx"}
                      accept=".docx" 
                      onFileSelect={setMaskWord}
                      icon={<FileText className="w-4 h-4 text-blue-600" />}
                   />
                   <Button variant="outline" size="sm" className="w-full h-7 text-xs" onClick={loadDefaultWord}>
                     Usar Modelo Padrão
                   </Button>
                 </div>
               )}
               <p className="text-[10px] text-muted-foreground">
                 Modelo com placeholders: {"<<obra>>"}, {"<<data>>"}, {"<<foto_antes_1>>"}, etc.
               </p>
             </div>

             <div className="h-px bg-border w-full" />

             {/* 4. COMPANY & CONTRACT CONFIG */}
             <div className="space-y-3">
               <Label className="text-xs font-bold text-muted-foreground uppercase tracking-wider">4. Empresa e Contrato</Label>
               <div className="space-y-2">
                 <div className="grid grid-cols-2 gap-2">
                   <div className="col-span-2">
                     <label className="text-[10px] text-muted-foreground">Empresa</label>
                     <Input 
                       value={companyName} 
                       onChange={(e) => setCompanyName(e.target.value)} 
                       className="h-8 text-xs" 
                       placeholder="Nome da Empresa..."
                       data-testid="input-company-name"
                     />
                   </div>
                   <div className="col-span-2">
                     <label className="text-[10px] text-muted-foreground">CNPJ</label>
                     <Input 
                       value={companyCNPJ} 
                       onChange={(e) => setCompanyCNPJ(e.target.value)} 
                       className="h-8 text-xs" 
                       placeholder="00.000.000/0000-00"
                       data-testid="input-company-cnpj"
                     />
                   </div>
                   <div className="col-span-2">
                     <label className="text-[10px] text-muted-foreground">Local</label>
                     <Input 
                       value={contractLocal} 
                       onChange={(e) => setContractLocal(e.target.value)} 
                       className="h-8 text-xs" 
                       placeholder="Local da obra..."
                       data-testid="input-contract-local"
                     />
                   </div>
                   <div className="col-span-2">
                     <label className="text-[10px] text-muted-foreground">Contrato</label>
                     <Input 
                       value={contractNumber} 
                       onChange={(e) => setContractNumber(e.target.value)} 
                       className="h-8 text-xs" 
                       placeholder="Nº do Contrato..."
                       data-testid="input-contract-number"
                     />
                   </div>
                   <div>
                     <label className="text-[10px] text-muted-foreground">Data Inicial</label>
                     <Popover>
                       <PopoverTrigger asChild>
                         <Button variant="outline" className="w-full justify-start text-left font-normal h-8 text-xs">
                           <CalendarIcon className="mr-1 h-3 w-3" />
                           {contractStartDate ? format(contractStartDate, "dd/MM/yy") : "Início"}
                         </Button>
                       </PopoverTrigger>
                       <PopoverContent className="w-auto p-0">
                         <Calendar mode="single" selected={contractStartDate} onSelect={setContractStartDate} locale={ptBR} />
                       </PopoverContent>
                     </Popover>
                   </div>
                   <div>
                     <label className="text-[10px] text-muted-foreground">Data Final</label>
                     <Popover>
                       <PopoverTrigger asChild>
                         <Button variant="outline" className="w-full justify-start text-left font-normal h-8 text-xs">
                           <CalendarIcon className="mr-1 h-3 w-3" />
                           {contractEndDate ? format(contractEndDate, "dd/MM/yy") : "Fim"}
                         </Button>
                       </PopoverTrigger>
                       <PopoverContent className="w-auto p-0">
                         <Calendar mode="single" selected={contractEndDate} onSelect={setContractEndDate} locale={ptBR} />
                       </PopoverContent>
                     </Popover>
                   </div>
                   <div className="col-span-2">
                     <label className="text-[10px] text-muted-foreground">Prorrogação (Dias)</label>
                     <Input 
                       type="number"
                       value={contractExtensionDays} 
                       onChange={(e) => setContractExtensionDays(parseInt(e.target.value) || 0)} 
                       className="h-8 text-xs" 
                       placeholder="0"
                       data-testid="input-contract-extension"
                     />
                   </div>
                 </div>

                 <div className="space-y-1 pt-2">
                   <label className="text-[10px] text-muted-foreground">Logo da Empresa (PNG, JPG)</label>
                   {companyLogo ? (
                     <div className="flex items-center gap-2 p-2 bg-green-50 border border-green-200 rounded">
                       <img src={companyLogo} alt="Logo" className="w-8 h-8 object-contain" />
                       <span className="text-xs text-green-700 flex-1 truncate">Logo carregado</span>
                       <Button 
                         variant="ghost" 
                         size="sm" 
                         className="h-5 w-5 p-0 text-red-500 hover:text-red-700" 
                         onClick={() => setCompanyLogo(null)}
                       >
                         ×
                       </Button>
                     </div>
                   ) : (
                     <label className="flex items-center gap-2 p-2 border border-dashed rounded cursor-pointer hover:bg-muted/50 transition-colors">
                       <ImageIcon className="w-4 h-4 text-muted-foreground" />
                       <span className="text-xs text-muted-foreground">Clique para carregar logo</span>
                       <input 
                         type="file" 
                         accept="image/png,image/jpeg,image/jpg" 
                         className="hidden"
                         onChange={(e) => {
                           const file = e.target.files?.[0];
                           if (file) {
                             const reader = new FileReader();
                             reader.onload = (event) => {
                               setCompanyLogo(event.target?.result as string);
                             };
                             reader.readAsDataURL(file);
                           }
                         }}
                         data-testid="input-company-logo"
                       />
                     </label>
                   )}
                 </div>
               </div>

                 <div className="space-y-1 pt-2">
                   <label className="text-[10px] text-muted-foreground">Assinatura do Preposto (PNG, JPG)</label>
                   {prepostoSignature ? (
                     <div className="flex items-center gap-2 p-2 bg-green-50 border border-green-200 rounded">
                       <img src={prepostoSignature} alt="Assinatura" className="w-12 h-8 object-contain" />
                       <span className="text-xs text-green-700 flex-1 truncate">Assinatura carregada</span>
                       <Button
                         variant="ghost"
                         size="sm"
                         className="h-5 w-5 p-0 text-red-500 hover:text-red-700"
                         onClick={() => setPrepostoSignature(null)}
                       >
                         ×
                       </Button>
                     </div>
                   ) : (
                     <label className="flex items-center gap-2 p-2 border border-dashed rounded cursor-pointer hover:bg-muted/50 transition-colors">
                       <PenLine className="w-4 h-4 text-muted-foreground" />
                       <span className="text-xs text-muted-foreground">Clique para carregar assinatura</span>
                       <input
                         type="file"
                         accept="image/png,image/jpeg,image/jpg"
                         className="hidden"
                         onChange={(e) => {
                           const file = e.target.files?.[0];
                           if (file) {
                             const reader = new FileReader();
                             reader.onload = (event) => {
                               setPrepostoSignature(event.target?.result as string);
                             };
                             reader.readAsDataURL(file);
                           }
                         }}
                         data-testid="input-preposto-signature"
                       />
                     </label>
                   )}
                 </div>
               <p className="text-[10px] text-muted-foreground">
                 Informações exibidas no cabeçalho do PDF.
               </p>
             </div>

             <div className="h-px bg-border w-full" />

             {/* 5. TEXT CORRECTION */}
             <div className="space-y-3">
               <Label className="text-xs font-bold text-muted-foreground uppercase tracking-wider flex items-center gap-2">
                 <Wand2 className="w-3 h-3" /> Correção de Texto
               </Label>
               <div className="flex items-center space-x-2 border p-2 rounded-md bg-muted/10">
                 <Checkbox 
                   id="ai-mode" 
                   checked={aiEnabled}
                   onCheckedChange={(checked) => setAiEnabled(checked === true)}
                 />
                 <label
                   htmlFor="ai-mode"
                   className="text-xs font-medium leading-none peer-disabled:cursor-not-allowed peer-disabled:opacity-70 cursor-pointer"
                 >
                   Ativar Correção Automática
                 </label>
               </div>
               {aiEnabled && (
                 <Button 
                   variant="outline" 
                   size="sm" 
                   className="w-full h-8 text-xs bg-purple-50 border-purple-200 text-purple-700 hover:bg-purple-100"
                   onClick={correctAllTexts}
                   disabled={isCorrectingText || selectedReportIds.size === 0}
                 >
                   {isCorrectingText ? (
                     <>
                       <RefreshCw className="w-3 h-3 mr-2 animate-spin" />
                       Corrigindo...
                     </>
                   ) : (
                     <>
                       <Wand2 className="w-3 h-3 mr-2" />
                       Corrigir {selectedReportIds.size} Selecionados
                     </>
                   )}
                 </Button>
               )}
               <p className="text-[10px] text-muted-foreground">
                 {aiEnabled 
                   ? "Selecione relatórios e clique no botão para corrigir os textos."
                   : "Corrige erros ortográficos e gramaticais automaticamente."
                 }
               </p>
             </div>

           </CardContent>
        </Card>

        {/* MAIN CONTENT */}
        <div className="flex-1 flex flex-col gap-4 min-h-0 lg:h-full overflow-hidden">

          {/* FILTERS BAR */}
          <Card className="flex-shrink-0">
            <CardContent className="p-3 sm:p-4">
              <div className="flex flex-col sm:flex-row gap-3 sm:gap-4 sm:items-end">
                <div className="space-y-2 flex-1 min-w-0">
                  <Label className="text-sm">Período</Label>
                  <div className="flex items-center gap-2">
                    <Popover>
                      <PopoverTrigger asChild>
                        <Button variant="outline" className="flex-1 justify-start text-left font-normal h-9 text-xs sm:text-sm">
                          <CalendarIcon className="mr-2 h-4 w-4 flex-shrink-0" />
                          <span className="truncate">{dateStart ? format(dateStart, "dd/MM/yy") : "Início"}</span>
                        </Button>
                      </PopoverTrigger>
                      <PopoverContent className="w-auto p-0"><Calendar mode="single" selected={dateStart} onSelect={setDateStart} locale={ptBR} /></PopoverContent>
                    </Popover>
                    <span className="text-muted-foreground flex-shrink-0">-</span>
                    <Popover>
                      <PopoverTrigger asChild>
                        <Button variant="outline" className="flex-1 justify-start text-left font-normal h-9 text-xs sm:text-sm">
                          <CalendarIcon className="mr-2 h-4 w-4 flex-shrink-0" />
                          <span className="truncate">{dateEnd ? format(dateEnd, "dd/MM/yy") : "Fim"}</span>
                        </Button>
                      </PopoverTrigger>
                      <PopoverContent className="w-auto p-0"><Calendar mode="single" selected={dateEnd} onSelect={setDateEnd} locale={ptBR} /></PopoverContent>
                    </Popover>
                  </div>
                </div>

                <div className="space-y-2 flex-1 min-w-0">
                  <Label className="text-sm">Supervisores</Label>
                  <Select onValueChange={(val) => {
                     if(!selectedTechnicians.includes(val)) setSelectedTechnicians([...selectedTechnicians, val]);
                  }}>
                    <SelectTrigger className="h-9 text-xs sm:text-sm">
                      <SelectValue placeholder="Filtrar Supervisor..." />
                    </SelectTrigger>
                    <SelectContent>
                      {technicians.map(t => <SelectItem key={t} value={t}>{t}</SelectItem>)}
                    </SelectContent>
                  </Select>
                </div>
              </div>

              {selectedTechnicians.length > 0 && (
                <div className="flex flex-wrap gap-2 mt-3 pt-3 border-t">
                   {selectedTechnicians.map(t => (
                     <Badge key={t} variant="secondary" className="gap-1 text-xs">
                       <span className="truncate max-w-[100px]">{t}</span>
                       <XIcon className="w-3 h-3 cursor-pointer flex-shrink-0" onClick={() => setSelectedTechnicians(selectedTechnicians.filter(x => x !== t))} />
                     </Badge>
                   ))}
                   <Button variant="ghost" size="sm" className="h-6 px-2 text-xs text-muted-foreground" onClick={() => setSelectedTechnicians([])}>Limpar</Button>
                </div>
              )}
            </CardContent>
          </Card>

          {/* DATA GRID */}
          <Card className="flex-1 overflow-hidden flex flex-col border-2">
            <div className="p-2 border-b bg-muted/30 flex justify-between items-center">
               <div className="flex items-center gap-2">
                 <Checkbox 
                    checked={filteredReports.length > 0 && selectedReportIds.size === filteredReports.length}
                    onCheckedChange={handleSelectAll}
                 />
                 <span className="text-sm font-medium text-muted-foreground">
                   {selectedReportIds.size} selecionados de {filteredReports.length} filtrados
                 </span>
               </div>
               <Badge variant="outline" className="font-mono text-xs">{allReports.length} TOTAL</Badge>
            </div>
            <div className="overflow-auto flex-1">
              <Table>
                <TableHeader>
                  <TableRow>
                    <TableHead className="w-[50px]"></TableHead>
                    <TableHead className="font-mono text-xs">RDO_ID</TableHead>
                    <TableHead>Data</TableHead>
                    <TableHead>Supervisor</TableHead>
                    <TableHead>Área</TableHead>
                    <TableHead>Mina</TableHead>
                    <TableHead>Local</TableHead>
                    <TableHead>OM</TableHead>
                    <TableHead className="w-[300px]">Atividade</TableHead>
                  </TableRow>
                </TableHeader>
                <TableBody>
                  {filteredReports.map((report) => (
                    <TableRow key={report.id} className={selectedReportIds.has(report.id) ? "bg-primary/5" : ""}>
                      <TableCell>
                        <Checkbox 
                          checked={selectedReportIds.has(report.id)}
                          onCheckedChange={(c) => handleSelectReport(report.id, c as boolean)}
                        />
                      </TableCell>
                      <TableCell className="font-mono text-xs font-medium text-primary">{report.rdoId}</TableCell>
                      <TableCell className="font-medium">{report.data ? format(report.data, 'dd/MM/yyyy') : '-'}</TableCell>
                      <TableCell>{report.supervisor || report.tecnico}</TableCell>
                      <TableCell>{report.area || report.obra || '-'}</TableCell>
                      <TableCell>{report.mina}</TableCell>
                      <TableCell>{report.localidade}</TableCell>
                      <TableCell>{report.om}</TableCell>
                      <TableCell className="truncate max-w-[300px] text-xs text-muted-foreground" title={report.atividade}>
                        {report.atividade}
                      </TableCell>
                    </TableRow>
                  ))}
                  {filteredReports.length === 0 && (
                    <TableRow>
                      <TableCell colSpan={9} className="h-24 text-center text-muted-foreground">
                        Nenhum relatório encontrado para os filtros atuais.
                      </TableCell>
                    </TableRow>
                  )}
                </TableBody>
              </Table>
            </div>
          </Card>

          {/* ACTION BAR */}
          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-3">
             {/* PREVIEW ACTION */}
             <Card className="bg-purple-50/50 border-purple-100 hover:border-purple-300 transition-all">
               <CardContent className="p-3 flex flex-col justify-between h-full min-h-[90px]">
                 <div className="flex items-center gap-2 text-purple-800 mb-2">
                   <Eye className="w-4 h-4 flex-shrink-0" />
                   <span className="font-bold text-sm truncate">Pré-visualizar</span>
                 </div>
                 <Button 
                   className="w-full bg-purple-600 hover:bg-purple-700 text-white text-sm h-9" 
                   onClick={() => { setPreviewIndex(0); setPreviewOpen(true); }}
                   disabled={filteredReports.length === 0}
                 >
                   Ver RDOs
                 </Button>
               </CardContent>
             </Card>

             {/* EXCEL ACTION */}
             <Card className="bg-green-50/50 border-green-100 hover:border-green-300 transition-all">
               <CardContent className="p-3 flex flex-col justify-between h-full min-h-[90px]">
                 <div className="flex justify-between items-center gap-2 mb-2">
                   <div className="flex items-center gap-2 text-green-800 min-w-0">
                     <FileSpreadsheet className="w-4 h-4 flex-shrink-0" />
                     <span className="font-bold text-sm truncate">Excel</span>
                   </div>
                   <Badge variant="secondary" className="bg-green-100 text-green-800 hover:bg-green-200 text-xs flex-shrink-0">
                     {selectedReportIds.size}
                   </Badge>
                 </div>
                 {isGeneratingExcel ? (
                   <div className="space-y-1">
                     <Progress value={excelProgress} className="h-2 bg-green-200" />
                     <p className="text-xs text-green-700 text-center animate-pulse">Gerando...</p>
                   </div>
                 ) : (
                   <Button 
                     className="w-full bg-green-600 hover:bg-green-700 text-white text-sm h-9" 
                     onClick={generateExcelRDOs}
                     disabled={selectedReportIds.size === 0}
                   >
                     Baixar ZIP
                   </Button>
                 )}
               </CardContent>
             </Card>

             {/* WORD ACTION */}
             <Card className="bg-blue-50/50 border-blue-100 hover:border-blue-300 transition-all">
               <CardContent className="p-3 flex flex-col justify-between h-full min-h-[90px]">
                 <div className="flex justify-between items-center gap-2 mb-2">
                   <div className="flex items-center gap-2 text-blue-800 min-w-0">
                     <FileText className="w-4 h-4 flex-shrink-0" />
                     <span className="font-bold text-sm truncate">Word</span>
                   </div>
                   <Badge variant="secondary" className="bg-blue-100 text-blue-800 hover:bg-blue-200 text-xs flex-shrink-0">
                     {selectedReportIds.size}
                   </Badge>
                 </div>
                 {isGeneratingWord ? (
                   <div className="space-y-1">
                     <Progress value={wordProgress} className="h-2 bg-blue-200" />
                     <p className="text-xs text-blue-700 text-center animate-pulse">Processando...</p>
                   </div>
                 ) : (
                   <Button 
                     className="w-full bg-blue-600 hover:bg-blue-700 text-white text-sm h-9"
                     onClick={generateWordReports}
                     disabled={selectedReportIds.size === 0}
                   >
                     Baixar ZIP
                   </Button>
                 )}
               </CardContent>
             </Card>

             {/* PDF ACTION */}
             <Card className="bg-red-50/50 border-red-100 hover:border-red-300 transition-all">
               <CardContent className="p-3 flex flex-col justify-between h-full min-h-[90px]">
                 <div className="flex justify-between items-center gap-2 mb-2">
                   <div className="flex items-center gap-2 text-red-800 min-w-0">
                     <FileDown className="w-4 h-4 flex-shrink-0" />
                     <span className="font-bold text-sm truncate">PDF</span>
                   </div>
                   <Badge variant="secondary" className="bg-red-100 text-red-800 hover:bg-red-200 text-xs flex-shrink-0">
                     {selectedReportIds.size}
                   </Badge>
                 </div>
                 {isGeneratingZip ? (
                   <div className="space-y-1">
                     <Progress value={zipProgress} className="h-2 bg-red-200" />
                     <p className="text-xs text-red-700 text-center animate-pulse">{zipStatus || 'Processando...'}</p>
                   </div>
                 ) : (
                   <Button 
                     className="w-full bg-red-600 hover:bg-red-700 text-white text-sm h-9"
                     onClick={generateAllPDFsZip}
                     disabled={selectedReportIds.size === 0}
                     data-testid="button-generate-all-pdfs"
                   >
                     Baixar ZIP
                   </Button>
                 )}
               </CardContent>
             </Card>
          </div>

        </div>
      </main>

      {/* PREVIEW DIALOG */}
      <Dialog open={previewOpen} onOpenChange={setPreviewOpen}>
        <DialogContent className="max-w-4xl h-[90vh] flex flex-col">
          <DialogHeader className="flex-shrink-0">
            <DialogTitle className="flex items-center gap-2">
              <Eye className="w-5 h-5 text-purple-600" />
              Pré-visualização dos RDOs
            </DialogTitle>
            <DialogDescription>
              Visualize os dados antes de gerar os documentos
            </DialogDescription>
          </DialogHeader>

          {filteredReports.length > 0 && (
            <>
              {/* Navigation */}
              <div className="flex items-center justify-between py-2 border-b flex-shrink-0">
                <Button
                  variant="outline"
                  size="sm"
                  onClick={() => setPreviewIndex(Math.max(0, previewIndex - 1))}
                  disabled={previewIndex === 0}
                >
                  <ChevronLeft className="w-4 h-4 mr-1" /> Anterior
                </Button>
                <div className="flex items-center gap-2">
                  <span className="text-sm font-medium">
                    Registro {previewIndex + 1} de {filteredReports.length}
                  </span>
                  <Select 
                    value={String(previewIndex)} 
                    onValueChange={(val) => setPreviewIndex(Number(val))}
                  >
                    <SelectTrigger className="w-[200px] h-8">
                      <SelectValue />
                    </SelectTrigger>
                    <SelectContent>
                      {filteredReports.map((r, idx) => (
                        <SelectItem key={r.id} value={String(idx)}>
                          {r.data ? format(r.data, 'dd/MM/yyyy') : '-'} - {r.tecnico}
                        </SelectItem>
                      ))}
                    </SelectContent>
                  </Select>
                </div>
                <div className="flex items-center gap-2">
                  <Button
                    variant="outline"
                    size="sm"
                    onClick={() => setPreviewIndex(Math.min(filteredReports.length - 1, previewIndex + 1))}
                    disabled={previewIndex === filteredReports.length - 1}
                  >
                    Próximo <ChevronRight className="w-4 h-4 ml-1" />
                  </Button>
                  <Button
                    size="sm"
                    className="bg-red-600 hover:bg-red-700 text-white"
                    onClick={() => {
                      const report = filteredReports[previewIndex];
                      if (report) handleGeneratePDF(report);
                    }}
                    disabled={isGeneratingPDF}
                    data-testid="button-generate-pdf"
                  >
                    {isGeneratingPDF ? (
                      <RefreshCw className="w-4 h-4 mr-1 animate-spin" />
                    ) : (
                      <FileDown className="w-4 h-4 mr-1" />
                    )}
                    Gerar PDF
                  </Button>
                </div>
              </div>

              {/* Content */}
              <ScrollArea className="flex-1 pr-4">
                {(() => {
                  const report = filteredReports[previewIndex];
                  if (!report) return null;

                  const prazoContratualRaw = contractStartDate && contractEndDate 
                    ? Math.ceil((contractEndDate.getTime() - contractStartDate.getTime()) / (1000 * 60 * 60 * 24))
                    : null;
                  const prazoContratual = prazoContratualRaw !== null && prazoContratualRaw >= 0 
                    ? prazoContratualRaw + contractExtensionDays 
                    : null;
                  const diasDecorridosRaw = contractStartDate && report.data 
                    ? Math.ceil((report.data.getTime() - contractStartDate.getTime()) / (1000 * 60 * 60 * 24)) + 1
                    : null;
                  const diasDecorridos = diasDecorridosRaw !== null && diasDecorridosRaw >= 1 
                    ? diasDecorridosRaw 
                    : null;
                  const diasRestantesRaw = contractEndDate && report.data 
                    ? Math.ceil((contractEndDate.getTime() - report.data.getTime()) / (1000 * 60 * 60 * 24)) + contractExtensionDays
                    : null;
                  const diasRestantes = diasRestantesRaw !== null ? Math.max(0, diasRestantesRaw) : null;
                  const diasAtraso = diasRestantesRaw !== null 
                    ? (diasRestantesRaw < 0 ? Math.abs(diasRestantesRaw) : 0)
                    : null;

                  return (
                    <div className="space-y-6 py-4">
                      {/* CABEÇALHO: Empresa e Contrato */}
                      <div className="border-b-2 border-primary pb-4 mb-4">
                        <div className="flex items-center justify-between gap-4">
                          {companyLogo && (
                            <img src={companyLogo} alt="Logo" className="h-12 w-auto object-contain" />
                          )}
                          <div className="text-center flex-1">
                            {companyName && (
                              <p className="font-bold text-primary text-base">{companyName}</p>
                            )}
                            <h2 className="font-bold text-primary text-lg">RELATÓRIO DIÁRIO DE OBRA - RDO</h2>
                          </div>
                          {companyLogo && <div className="w-12" />}
                        </div>

                        {/* Contract Info Grid */}
                        <div className="mt-4 bg-sky-50 rounded-lg p-3 border border-sky-200">
                          <div className="grid grid-cols-2 gap-x-6 gap-y-2 text-sm">
                            <div className="flex gap-2">
                              <span className="text-muted-foreground font-medium">Empresa:</span>
                              <span className="text-primary font-semibold">{companyName || '-'}</span>
                            </div>
                            <div className="flex gap-2">
                              <span className="text-muted-foreground font-medium">CNPJ:</span>
                              <span className="text-primary">{companyCNPJ || '-'}</span>
                            </div>
                            <div className="flex gap-2">
                              <span className="text-muted-foreground font-medium">Local:</span>
                              <span className="text-primary">{contractLocal || '-'}</span>
                            </div>
                            <div className="flex gap-2">
                              <span className="text-muted-foreground font-medium">Contrato:</span>
                              <span className="text-primary">{contractNumber || '-'}</span>
                            </div>
                            <div className="flex gap-2">
                              <span className="text-muted-foreground font-medium">Data Inicial:</span>
                              <span className="text-primary">{contractStartDate ? format(contractStartDate, 'dd/MM/yyyy') : '-'}</span>
                            </div>
                            <div className="flex gap-2">
                              <span className="text-muted-foreground font-medium">Data Final:</span>
                              <span className="text-primary">{contractEndDate ? format(contractEndDate, 'dd/MM/yyyy') : '-'}</span>
                            </div>
                          </div>

                          {/* Contract Metrics */}
                          <div className="grid grid-cols-5 gap-2 mt-3 pt-3 border-t border-sky-200 border-dashed">
                            <div className="text-center">
                              <p className="text-xs text-muted-foreground">Prazo Contratual</p>
                              <p className="text-lg font-bold text-primary">{prazoContratual !== null ? prazoContratual : '-'}</p>
                              {prazoContratual !== null && <p className="text-[10px] text-muted-foreground">dias</p>}
                            </div>
                            <div className="text-center">
                              <p className="text-xs text-muted-foreground">Dias Decorridos</p>
                              <p className="text-lg font-bold text-sky-600">{diasDecorridos !== null ? diasDecorridos : '-'}</p>
                              {diasDecorridos !== null && <p className="text-[10px] text-muted-foreground">dias</p>}
                            </div>
                            <div className="text-center">
                              <p className="text-xs text-muted-foreground">Prorrogação</p>
                              <p className="text-lg font-bold text-violet-600">{contractExtensionDays}</p>
                              <p className="text-[10px] text-muted-foreground">dias</p>
                            </div>
                            <div className="text-center">
                              <p className="text-xs text-muted-foreground">Dias Restantes</p>
                              <p className={`text-lg font-bold ${diasRestantes !== null && diasRestantes > 0 ? 'text-green-600' : 'text-muted-foreground'}`}>
                                {diasRestantes !== null ? diasRestantes : '-'}
                              </p>
                              {diasRestantes !== null && <p className="text-[10px] text-muted-foreground">dias</p>}
                            </div>
                            <div className="text-center">
                              <p className="text-xs text-muted-foreground">Dias de Atraso</p>
                              <p className={`text-lg font-bold ${diasAtraso !== null && diasAtraso > 0 ? 'text-red-600' : 'text-green-600'}`}>
                                {diasAtraso !== null ? diasAtraso : '-'}
                              </p>
                              {diasAtraso !== null && <p className="text-[10px] text-muted-foreground">dias</p>}
                            </div>
                          </div>
                        </div>
                      </div>

                      {/* SEÇÃO 1: Identificação da OM */}
                      <div className="bg-slate-50 rounded-lg p-4 border border-slate-200">
                        <h3 className="text-sm font-bold text-slate-700 mb-3 flex items-center gap-2">
                          <ClipboardCheck className="w-4 h-4" />
                          Identificação da Ordem de Serviço
                        </h3>
                        <div className="grid grid-cols-4 gap-4">
                          <div className="space-y-1">
                            <Label className="text-xs text-muted-foreground">Data</Label>
                            <p className="font-medium flex items-center gap-2 text-sm">
                              <CalendarIcon className="w-4 h-4 text-primary" />
                              {report.data ? format(report.data, 'dd/MM/yyyy') : '-'}
                            </p>
                          </div>
                          <div className="space-y-1">
                            <Label className="text-xs text-muted-foreground">Nº OM</Label>
                            <p className="font-bold text-primary text-lg">{report.om || 'S/I'}</p>
                          </div>
                          <div className="space-y-1">
                            <Label className="text-xs text-muted-foreground">Supervisor Responsável</Label>
                            <p className="font-medium text-sm">{report.supervisor || report.tecnico}</p>
                          </div>
                          <div className="space-y-1">
                            <Label className="text-xs text-muted-foreground">Tarefa Concluída</Label>
                            <Badge variant={report.tarefaConcluida === 'Sim' ? 'default' : 'secondary'} className={report.tarefaConcluida === 'Sim' ? 'bg-green-600' : ''}>
                              {report.tarefaConcluida || '-'}
                            </Badge>
                          </div>
                        </div>
                        <div className="grid grid-cols-4 gap-4 mt-3">
                          <div className="space-y-1">
                            <Label className="text-xs text-muted-foreground">Encarregado</Label>
                            <p className="font-medium text-sm">{report.encarregado || '-'}</p>
                          </div>
                          <div className="space-y-1">
                            <Label className="text-xs text-muted-foreground">Mina</Label>
                            <p className="font-medium text-sm">{report.mina || 'S/I'}</p>
                          </div>
                          <div className="space-y-1">
                            <Label className="text-xs text-muted-foreground">Localidade</Label>
                            <p className="font-medium text-sm">{report.localidade || 'S/I'}</p>
                          </div>
                          <div className="space-y-1">
                            <Label className="text-xs text-muted-foreground">Área</Label>
                            <p className="font-medium text-sm">{report.area || report.obra || 'S/I'}</p>
                          </div>
                        </div>
                        <div className="grid grid-cols-4 gap-4 mt-3">
                          <div className="space-y-1">
                            <Label className="text-xs text-muted-foreground">TAG</Label>
                            <p className="font-medium text-sm">{report.tag || '-'}</p>
                          </div>
                          <div className="space-y-1">
                            <Label className="text-xs text-muted-foreground">Escala</Label>
                            <p className="font-medium text-sm">{report.escalaTrabalho || '-'}</p>
                          </div>
                          <div className="space-y-1">
                            <Label className="text-xs text-muted-foreground flex items-center gap-1">
                              <Clock className="w-3 h-3" /> Horário Início
                            </Label>
                            <p className="font-medium text-sm">{report.horarioInicio || '-'}</p>
                          </div>
                          <div className="space-y-1">
                            <Label className="text-xs text-muted-foreground flex items-center gap-1">
                              <Clock className="w-3 h-3" /> Horário Término
                            </Label>
                            <p className="font-medium text-sm">{report.horarioTermino || '-'}</p>
                          </div>
                        </div>
                      </div>

                      {/* SEÇÃO 2: Atividade Realizada */}
                      <div className="space-y-2">
                        <Label className="text-sm font-bold flex items-center gap-2">
                          <FileText className="w-4 h-4 text-blue-600" />
                          Atividade Realizada
                        </Label>
                        <div className="bg-blue-50 border border-blue-200 p-3 rounded-md text-sm whitespace-pre-wrap min-h-[60px]">
                          {report.atividade || 'Sem descrição'}
                        </div>
                      </div>

                      {/* SEÇÃO 3: Clima */}
                      <div className="space-y-2">
                        <Label className="text-sm font-bold">Condições Climáticas</Label>
                        <div className="grid grid-cols-3 gap-3">
                          <div className="flex items-center gap-2 bg-yellow-50 p-3 rounded-md border border-yellow-200">
                            <Sun className="w-5 h-5 text-yellow-600" />
                            <div>
                              <p className="text-xs text-muted-foreground">Matutino</p>
                              <p className="text-sm font-medium">{report.climaMatutino || '-'}</p>
                            </div>
                          </div>
                          <div className="flex items-center gap-2 bg-orange-50 p-3 rounded-md border border-orange-200">
                            <Cloud className="w-5 h-5 text-orange-600" />
                            <div>
                              <p className="text-xs text-muted-foreground">Vespertino</p>
                              <p className="text-sm font-medium">{report.climaVespertino || '-'}</p>
                            </div>
                          </div>
                          <div className="flex items-center gap-2 bg-indigo-50 p-3 rounded-md border border-indigo-200">
                            <Moon className="w-5 h-5 text-indigo-600" />
                            <div>
                              <p className="text-xs text-muted-foreground">Noturno</p>
                              <p className="text-sm font-medium">{report.climaNoturno || '-'}</p>
                            </div>
                          </div>
                        </div>
                      </div>

                      {/* SEÇÃO 4: Recursos */}
                      <div className="grid grid-cols-2 gap-4">
                        <div className="space-y-2">
                          <Label className="text-sm font-bold flex items-center gap-2">
                            <Users className="w-4 h-4 text-green-600" />
                            Equipe
                          </Label>
                          <div className="bg-green-50 border border-green-200 p-2 rounded-md text-sm min-h-[50px] leading-relaxed">
                            {report.quemTrabalhou ? (
                              report.quemTrabalhou.split(',').map((item, idx) => (
                                <div key={idx}>• {item.trim()}</div>
                              ))
                            ) : '-'}
                          </div>
                        </div>
                        <div className="space-y-2">
                          <Label className="text-sm font-bold flex items-center gap-2">
                            <Wrench className="w-4 h-4 text-orange-600" />
                            <Truck className="w-4 h-4 text-blue-600" />
                            Equipamentos e Veículos
                          </Label>
                          <div className="bg-slate-50 border border-slate-200 p-2 rounded-md text-sm min-h-[50px] leading-relaxed">
                            {[report.equipamentos, report.veiculos].filter(v => v && v.trim() && v.trim() !== '-').length > 0 ? (
                              [report.equipamentos, report.veiculos]
                                .filter(v => v && v.trim() && v.trim() !== '-')
                                .join(',')
                                .split(',')
                                .map((item, idx) => (
                                  <div key={idx}>• {item.trim()}</div>
                                ))
                            ) : '-'}
                          </div>
                        </div>
                      </div>

                      {/* SEÇÃO 5: Registro Fotográfico */}
                      <div className="space-y-3">
                        <Label className="text-sm font-bold flex items-center gap-2">
                          <ImageIcon className="w-4 h-4 text-purple-600" />
                          Registro Fotográfico
                        </Label>
                        <div className="grid grid-cols-2 gap-4">
                          {/* ANTES */}
                          <div className="space-y-2">
                            <p className="text-xs font-bold text-center bg-red-100 text-red-800 py-1.5 rounded-t border border-red-200 border-b-0">ANTES</p>
                            <div className="grid grid-cols-2 gap-2 p-2 bg-red-50/50 rounded-b border border-red-200 border-t-0">
                              {report.fotosAntes.length > 0 ? (
                                report.fotosAntes.slice(0, 8).map((url, idx) => {
                                  const fileId = url.includes('id=') 
                                    ? url.split('id=')[1]?.split('&')[0] 
                                    : url.includes('/d/') 
                                      ? url.split('/d/')[1]?.split('/')[0] 
                                      : '';
                                  const thumbnailUrl = fileId 
                                    ? `https://drive.google.com/thumbnail?id=${fileId}&sz=w200` 
                                    : '';
                                  return (
                                    <div key={idx} className="aspect-[3/2] bg-white rounded overflow-hidden border shadow-sm">
                                      {thumbnailUrl ? (
                                        <img 
                                          src={thumbnailUrl} 
                                          alt={`Antes ${idx + 1}`}
                                          className="w-full h-full object-cover"
                                          onError={(e) => {
                                            (e.target as HTMLImageElement).style.display = 'none';
                                            (e.target as HTMLImageElement).parentElement!.innerHTML = '<div class="flex items-center justify-center h-full text-xs text-muted-foreground">Erro</div>';
                                          }}
                                        />
                                      ) : (
                                        <div className="flex items-center justify-center h-full text-xs text-muted-foreground">
                                          Sem preview
                                        </div>
                                      )}
                                    </div>
                                  );
                                })
                              ) : (
                                <div className="col-span-2 text-center text-sm text-muted-foreground py-6">
                                  Sem fotos
                                </div>
                              )}
                            </div>
                          </div>

                          {/* DEPOIS */}
                          <div className="space-y-2">
                            <p className="text-xs font-bold text-center bg-green-100 text-green-800 py-1.5 rounded-t border border-green-200 border-b-0">DEPOIS</p>
                            <div className="grid grid-cols-2 gap-2 p-2 bg-green-50/50 rounded-b border border-green-200 border-t-0">
                              {report.fotosDepois.length > 0 ? (
                                report.fotosDepois.slice(0, 8).map((url, idx) => {
                                  const fileId = url.includes('id=') 
                                    ? url.split('id=')[1]?.split('&')[0] 
                                    : url.includes('/d/') 
                                      ? url.split('/d/')[1]?.split('/')[0] 
                                      : '';
                                  const thumbnailUrl = fileId 
                                    ? `https://drive.google.com/thumbnail?id=${fileId}&sz=w200` 
                                    : '';
                                  return (
                                    <div key={idx} className="aspect-[3/2] bg-white rounded overflow-hidden border shadow-sm">
                                      {thumbnailUrl ? (
                                        <img 
                                          src={thumbnailUrl} 
                                          alt={`Depois ${idx + 1}`}
                                          className="w-full h-full object-cover"
                                          onError={(e) => {
                                            (e.target as HTMLImageElement).style.display = 'none';
                                            (e.target as HTMLImageElement).parentElement!.innerHTML = '<div class="flex items-center justify-center h-full text-xs text-muted-foreground">Erro</div>';
                                          }}
                                        />
                                      ) : (
                                        <div className="flex items-center justify-center h-full text-xs text-muted-foreground">
                                          Sem preview
                                        </div>
                                      )}
                                    </div>
                                  );
                                })
                              ) : (
                                <div className="col-span-2 text-center text-sm text-muted-foreground py-6">
                                  Sem fotos
                                </div>
                              )}
                            </div>
                          </div>
                        </div>
                      </div>

                      <div className="h-px bg-border" />

                      {/* SEÇÃO 6: Campos Editáveis */}
                      <div className="bg-amber-50 rounded-lg p-4 border border-amber-200 space-y-4">
                        <h3 className="text-sm font-bold text-amber-800 flex items-center gap-2">
                          <AlertTriangle className="w-4 h-4" />
                          Anotações e Correções (Editável)
                        </h3>

                        {/* Observações do Preview */}
                        <div className="space-y-2">
                          <Label className="text-sm font-medium flex items-center gap-2">
                            <MessageSquare className="w-4 h-4 text-blue-600" />
                            Observações Adicionais
                          </Label>
                          <Textarea 
                            placeholder="Digite observações adicionais sobre este RDO..."
                            value={report.observacoesPreview || ''}
                            onChange={(e) => updateReportField(report.id, 'observacoesPreview', e.target.value)}
                            className="min-h-[80px] bg-white"
                            data-testid={`textarea-observacoes-${report.id}`}
                          />
                        </div>

                        {/* Correção Fiscalizadora */}
                        <div className="space-y-2">
                          <Label className="text-sm font-medium flex items-center gap-2">
                            <ClipboardCheck className="w-4 h-4 text-red-600" />
                            Correções da Fiscalização
                          </Label>
                          <Textarea 
                            placeholder="Registre aqui correções ou apontamentos da fiscalização..."
                            value={report.correcaoFiscalizadora || ''}
                            onChange={(e) => updateReportField(report.id, 'correcaoFiscalizadora', e.target.value)}
                            className="min-h-[80px] bg-white border-red-200 focus:border-red-400"
                            data-testid={`textarea-correcao-${report.id}`}
                          />
                        </div>

                        {/* Pendências */}
                        <div className="space-y-2">
                          <Label className="text-sm font-medium flex items-center gap-2">
                            <AlertTriangle className="w-4 h-4 text-amber-600" />
                            Pendências
                          </Label>
                          <Textarea 
                            placeholder="Liste pendências ou itens a serem resolvidos..."
                            value={report.pendencias || ''}
                            onChange={(e) => updateReportField(report.id, 'pendencias', e.target.value)}
                            className="min-h-[80px] bg-white border-amber-200 focus:border-amber-400"
                            data-testid={`textarea-pendencias-${report.id}`}
                          />
                        </div>

                        <p className="text-xs text-amber-700 flex items-center gap-1">
                          <Save className="w-3 h-3" />
                          As alterações são salvas automaticamente e incluídas nos documentos gerados.
                        </p>
                      </div>
                    </div>
                  );
                })()}
              </ScrollArea>
            </>
          )}
        </DialogContent>
      </Dialog>
    </div>
  );
}

function XIcon({ className, onClick }: { className?: string, onClick?: () => void }) {
  return (
    <svg 
      xmlns="http://www.w3.org/2000/svg" 
      width="24" height="24" 
      viewBox="0 0 24 24" 
      fill="none" 
      stroke="currentColor" 
      strokeWidth="2" 
      strokeLinecap="round" 
      strokeLinejoin="round" 
      className={className}
      onClick={onClick}
    >
      <path d="M18 6 6 18"/>
      <path d="m6 6 12 12"/>
    </svg>
  )
}
