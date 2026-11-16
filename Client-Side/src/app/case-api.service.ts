import { Injectable } from '@angular/core';
import { HttpClient } from '@angular/common/http';
import { firstValueFrom } from 'rxjs';

export type FieldType = 'string' | 'date' | 'number' | 'boolean';

/** One entry in the merge field schema (drives the missing-field dialog). */
export interface FieldSchemaEntry {
    type: FieldType;
    label?: string;
    description?: string;
    /** number fields: whole numbers only. */
    integer?: boolean;
    /** number fields: optional bounds. */
    min?: number | null;
    max?: number | null;
}

/** Merge field schema from GET /api/schema: region (group) name -> field name -> typed entry.
 * No region or field names are hardcoded on the client. */
export interface FieldSchema {
    [groupName: string]: { [fieldName: string]: FieldSchemaEntry };
}

export interface CaseInfo {
    CaseId: number;
    CaseNumber: string;
    CaseCaption: string;
}

export interface PartyData {
    FirstName: string;
    LastName: string;
    Address: string;
}

export interface CaseData {
    CaseId: number;
    CaseNumber: string;
    CaseCaption: string;
    Parties: PartyData[];
}

export interface MissingField {
    caseId: number;
    caseNumber: string;
    groupName: string;
    rowIndex: number;
    fieldName: string;
}

export interface AdditionalMergeValue {
    caseId: number;
    caseNumber: string;
    groupName: string;
    rowIndex: number;
    fieldName: string;
    value: string;
}

/**
 * A supporting document available in the configured templates folder.
 * Returned by GET /api/templates.
 */
export interface SupportingDocument {
    id: string;
    displayName: string;
    order: number;
    /** True for the default main template ("Notice of Hearing"). */
    isMainTemplate?: boolean;
    /** False when the file has not been saved to the templates folder yet. */
    exists?: boolean;
}

/** Result of a successful template save. */
export interface SaveTemplateResult {
    id: string;
    displayName: string;
    path: string;
}

export interface GenerateRequest {
    caseIds: number[];
    additionalMergeData?: AdditionalMergeValue[];
    templateSfdt?: string;
    appendedDocumentIds?: string[];
}

/** Response of GET /api/template - the default main template. */
export interface DefaultTemplate {
    /** The document serialized as SFDT, ready for the editor. */
    sfdt: string;
    /** File name inside the templates folder that Save Document writes to. */
    fileName: string;
    /** False when the file does not exist on disk yet (the built-in sample is shown). */
    exists: boolean;
}

export type GenerateResult =
    | { kind: 'requires-input'; missingFields: MissingField[] }
    | { kind: 'merged'; mergedSfdt: string; documentName: string };

/**
 * API client for the Notice of Hearing POC. Mirrors the ASP.NET Core endpoints:
 * GET  /api/schema           - merge field schema (loaded on start, independently of search)
 * GET  /api/template         - bundled sample template as SFDT (opened by the editor on start)
 * GET  /api/cases/search     - targeted case search
 * GET  /api/cases/{caseId}   - full case details
 * POST /api/generate         - mail merge; returns SFDT for the editor or missing fields
 */
@Injectable({ providedIn: 'root' })
export class CaseApiService {
    constructor(private http: HttpClient) { }

    getSchema(): Promise<FieldSchema> {
        return firstValueFrom(this.http.get<FieldSchema>('/api/schema'));
    }

    /** Default main template as SFDT (the Notice of Hearing from the templates folder),
     * opened in the editor on application start. Falls back to the bundled sample
     * template server-side if the configured file is not present. */
    getDefaultTemplate(): Promise<DefaultTemplate> {
        return firstValueFrom(this.http.get<DefaultTemplate>('/api/template'));
    }

    searchCases(keyword: string): Promise<CaseInfo[]> {
        const kw = keyword && keyword.trim() ? keyword.trim() : '';
        const url = `/api/cases/search${kw ? `?keyword=${encodeURIComponent(kw)}` : ''}`;
        return firstValueFrom(this.http.get<CaseInfo[]>(url));
    }

    getCase(caseId: number): Promise<CaseData> {
        return firstValueFrom(this.http.get<CaseData>(`/api/cases/${caseId}`));
    }

    /** Lists the supporting documents available for appending after the merge. */
    listTemplates(): Promise<SupportingDocument[]> {
        return firstValueFrom(this.http.get<SupportingDocument[]>('/api/templates'));
    }

    /** Loads a supporting document as SFDT so the editor can open it. */
    getTemplate(id: string): Promise<string> {
        return firstValueFrom(this.http.get<{ sfdt: string }>(`/api/templates/${encodeURIComponent(id)}`))
            .then(r => r.sfdt);
    }

    /**
     * Saves the template currently open in the editor back to the templates folder.
     * The server converts the SFDT payload back to DOCX and overwrites the file.
     */
    saveTemplate(id: string, sfdt: string): Promise<SaveTemplateResult> {
        return firstValueFrom(this.http.post<SaveTemplateResult>(
            `/api/templates/${encodeURIComponent(id)}/save`, { sfdt }));
    }

    /**
     * Converts an uploaded PDF to a Word document on the server and returns the
     * result as SFDT so the Document Editor can open it directly.
     * Server side uses Syncfusion's DataExtractor.ExtractDataAsWordDocument.
     */
    convertPdfToWord(file: File): Promise<string> {
        const form = new FormData();
        form.append('file', file, file.name);
        return firstValueFrom(this.http.post<{ sfdt: string }>('/api/pdf/convert', form))
            .then(r => r.sfdt);
    }

    /**
     * Converts the XSL/XSLT stylesheet from the templates folder (InputXSL.txt) into
     * a Word mail-merge template and returns it as SFDT for the editor.
     *
     * Server side turns xsl:for-each into merge regions and xsl:value-of into
     * merge fields, so the result works with the same mail-merge pipeline as the
     * other templates.
     */
    convertXslToWord(fileName?: string): Promise<string> {
        return firstValueFrom(this.http.post<{ sfdt: string }>(
            '/api/templates/xsl-to-word', { fileName: fileName ?? 'InputXSL.txt' }))
            .then(r => r.sfdt);
    }

    /**
     * Runs the same mail merge as /api/generate and returns the result as a PDF
     * (server uses Syncfusion DocIORenderer for the Word -> PDF step). The browser
     * triggers a download via the returned Blob URL.
     *
     * Returns null when the server reports missing fields, in which case `missingFields`
     * is populated so the client can prompt and re-call /api/generate as usual.
     */
    async generatePdf(request: GenerateRequest): Promise<{ kind: 'pdf'; blob: Blob; fileName: string } | { kind: 'requires-input'; missingFields: MissingField[] }> {
        const response = await firstValueFrom(this.http.post(
            '/api/generate-pdf',
            request,
            { observe: 'response', responseType: 'blob' }));
        const contentType = response.headers.get('content-type') ?? '';

        if (contentType.includes('application/json')) {
            const text = await response.body!.text();
            const parsed = JSON.parse(text) as Partial<{ requiresInput: boolean; missingFields: MissingField[] }>;
            if (parsed.requiresInput) {
                return { kind: 'requires-input', missingFields: parsed.missingFields ?? [] };
            }
            throw new Error('Unexpected generate-pdf response.');
        }

        const disposition = response.headers.get('content-disposition') ?? '';
        const match = /filename="?([^";]+)"?/i.exec(disposition);
        const fileName = match ? match[1] : 'Merged-Notice-of-Hearing.pdf';
        return { kind: 'pdf', blob: response.body!, fileName };
    }

    /**
     * Mail merge for the selected cases.
     *
     * The server answers with either
     *   { requiresInput: true, missingFields: [...] }  -> prompt the user, or
     *   { mergedSfdt: '...', documentName: '...' }      -> open the merged document
     *                                                    in the editor.
     * Both are JSON: the response is inspected and mapped to a discriminated union so
     * the component can branch on result.kind.
     */
    async generate(request: GenerateRequest): Promise<GenerateResult> {
        const parsed = await firstValueFrom(
            this.http.post<Partial<GenerateRequiresInputResponse & GenerateMergedResponse>>('/api/generate', request));
        if (parsed && parsed.requiresInput) {
            return { kind: 'requires-input', missingFields: parsed.missingFields || [] };
        }
        if (parsed && parsed.mergedSfdt) {
            return { kind: 'merged', mergedSfdt: parsed.mergedSfdt, documentName: parsed.documentName || 'Merged document' };
        }
        throw new Error('Unexpected generate response.');
    }
}

/** Server contract for the missing-field response. */
interface GenerateRequiresInputResponse {
    requiresInput: boolean;
    missingFields: MissingField[];
}

/** Server contract for the merged response. */
interface GenerateMergedResponse {
    mergedSfdt: string;
    documentName: string;
}