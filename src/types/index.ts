// Compound types
export interface Compound {
    id: string;
    code: string;
    name: string;
    description?: string;
}

// Template types
export interface Template {
    id: string;
    compoundId: string;
    region: 'CN' | 'EU' | 'US';
    templateContent: string;
    fieldMapping?: Record<string, string>;
}

// COA Document types
export interface COADocument {
    id: string;
    compoundId: string;
    filename: string;
    filePath: string;
    processingStatus: 'pending' | 'processing' | 'completed' | 'failed';
    uploadedAt: Date;
    processedAt?: Date;
}

// Extracted data types
export interface ExtractedField {
    id: string;
    documentId: string;
    fieldName: string;
    fieldValue: string;
    confidenceScore: number;
    originalText?: string;
}

// API Response types
export interface ApiResponse<T> {
    success: boolean;
    data?: T;
    error?: string;
    message?: string;
}

// File upload response
export interface FileUploadResponse {
    documentId: string;
    filename: string;
    status: string;
}

// Processing result
export interface ProcessingResult {
    documentId: string;
    extractedData: ExtractedField[];
    status: 'success' | 'partial' | 'failed';
    message?: string;
}

// App state types
export interface AppState {
    selectedCompound?: Compound;
    selectedTemplate?: Template;
    uploadedDocument?: COADocument;
    extractedData: ExtractedField[];
    isLoading: boolean;
    error?: string;
}

// Form data types
export interface COAFormData {
    lotNumber: string;
    manufacturer: string;
    storageCondition: string;
}

// API Endpoints
declare const process: {
    env: {
        [key: string]: string | undefined;
    };
};


export const API_BASE_URL = (() => {
  const hostname = window.location.hostname;

  // 本地开发环境（localhost 或 127.0.0.1）
  if (hostname === "localhost" || hostname === "127.0.0.1") {
    return "https://localhost:8000";
  }
  return "https://10.8.63.207:8000"; 
})();
