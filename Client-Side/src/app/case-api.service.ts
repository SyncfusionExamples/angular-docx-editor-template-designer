import { Injectable } from '@angular/core';
import { HttpClient } from '@angular/common/http';
import { firstValueFrom } from 'rxjs';

export interface SchemaGroup { [fieldName: string]: string; }
/** Merge field schema: group name -> field name -> field type. */
export interface FieldSchema { [groupName: string]: SchemaGroup; }

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

export interface GenerateRequest {
    caseIds: number[];
    additionalMergeData?: AdditionalMergeValue[];
    templateSfdt?: string;
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

    /** Bundled sample template as SFDT, opened in the editor on application start. */
    getSampleTemplate(): Promise<string> {
        return firstValueFrom(this.http.get<{ sfdt: string }>('/api/template'))
            .then(r => r.sfdt);
    }

    searchCases(keyword: string): Promise<CaseInfo[]> {
        const kw = keyword && keyword.trim() ? keyword.trim() : '';
        const url = `/api/cases/search${kw ? `?keyword=${encodeURIComponent(kw)}` : ''}`;
        return firstValueFrom(this.http.get<CaseInfo[]>(url));
    }

    getCase(caseId: number): Promise<CaseData> {
        return firstValueFrom(this.http.get<CaseData>(`/api/cases/${caseId}`));
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