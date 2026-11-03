import { Component, ViewChild, signal, AfterViewInit, OnDestroy } from '@angular/core';
import { FormsModule } from '@angular/forms';
import { CommonModule } from '@angular/common';
import { DocumentEditorContainerComponent } from '@syncfusion/ej2-angular-documenteditor';
import {
    ToolbarService,
    RibbonService
} from '@syncfusion/ej2-angular-documenteditor';
import {
    CaseApiService,
    FieldSchema,
    CaseInfo,
    MissingField,
    AdditionalMergeValue
} from './case-api.service';

interface MissingFieldInput {
    caseId: number;
    caseNumber: string;
    groupName: string;
    rowIndex: number;
    fieldName: string;
    label: string;
    value: string;
}

interface SchemaGroupView {
    name: string;
    fields: { name: string; type: string }[];
}

/**
 *
 * Startup: schema loads and the bundled sample template opens in the editor, so the
 * user can immediately search, select cases and generate (no blank editor).
 *
 * Layout (unchanged):
 * - Left:  Merge Fields (click-to-insert, driven by GET /api/schema)
 * - Center: Document Editor in Ribbon mode with a custom "Mail Merge" tab
 * - Right: Search Cases + Generate Documents (business data selection)
 *
 * Generate flow: the current editor document is serialized as SFDT, sent to
 * POST /api/generate, DocIO executes the relational nested mail merge for every
 * selected case into ONE document, and the merged SFDT is opened directly in the
 * editor (no download).
 *
 * State uses signals because the app runs zoneless (Angular 21 default): promise
 * callbacks do not trigger change detection by themselves.
 */
@Component({
    selector: 'app-root',
    standalone: true,
    imports: [DocumentEditorContainerComponent, FormsModule, CommonModule],
    providers: [ToolbarService, RibbonService],
    template: `
    <div class="poc-shell">
        <header class="poc-header">
            <h1>Syncfusion DOCX Editor - Document Generation</h1>
            <p>Template authoring in the Document Editor, relational mail merge on the server.</p>
        </header>
        <div class="poc-body">
            <!-- LEFT PANEL: Merge Field Explorer (click-to-insert) -->
            <aside class="poc-panel poc-fields">
                <h2>Merge Fields</h2>
                <div class="poc-tree">
                    @for (group of schemaGroups(); track group.name) {
                        <div class="poc-group-header">{{ group.name }}</div>
                        @for (field of group.fields; track field.name) {
                            <button type="button" class="poc-field"
                                (click)="insertMergeField(field.name)"
                                [title]="'Insert merge field ' + field.name">
                                {{ field.name }}
                            </button>
                        }
                    }
                </div>
                <p class="poc-muted">
                    Click a field to insert it at the cursor. Party fields repeat inside
                    a Parties region; Case fields resolve inside it too.
                </p>

                <!-- Manual merge field creation (Change 5) -->
                <div class="poc-manual-section">
                    <div class="poc-divider"></div>
                    <h3>Add Custom Merge Field</h3>
                    <label class="poc-label">
                        Field Name:
                        <input type="text" class="poc-input" name="customFieldName"
                            [(ngModel)]="customFieldName" placeholder="e.g. HearingLocation"
                            (keyup.enter)="insertCustomField()">
                    </label>
                    <button type="button" class="poc-btn poc-btn-primary poc-btn-block"
                        (click)="insertCustomField()"
                        [disabled]="!customFieldName || !customFieldName.trim()">Insert Field</button>
                </div>

                <!-- Manual region creation (Change 6) -->
                <div class="poc-manual-section">
                    <div class="poc-divider"></div>
                    <h3>Add Merge Region</h3>
                    <label class="poc-label">
                        Region Name:
                        <input type="text" class="poc-input" name="customRegionName"
                            [(ngModel)]="customRegionName" placeholder="e.g. Parties"
                            (keyup.enter)="insertCustomRegion()">
                    </label>
                    <button type="button" class="poc-btn poc-btn-primary poc-btn-block"
                        (click)="insertCustomRegion()"
                        [disabled]="!customRegionName || !customRegionName.trim()">Insert Region</button>
                </div>
            </aside>

            <!-- CENTER: Document Editor (Ribbon mode + Mail Merge tab) -->
            <section class="poc-editor">
                <ejs-documenteditorcontainer
                    #container
                    height="700px"
                    style="display:block"
                    [enableToolbar]="true"
                    [toolbarMode]="'Ribbon'"
                    serviceUrl="/api/documenteditor/">
                </ejs-documenteditorcontainer>
            </section>

            <!-- RIGHT PANEL: Search Cases + Generate -->
            <aside class="poc-panel poc-cases">
                <h2>Search Cases</h2>
                <div class="poc-search-row">
                    <input type="text" [(ngModel)]="keyword" placeholder="e.g. Smith"
                        (keyup.enter)="search()" class="poc-input" name="keyword">
                    <button type="button" class="poc-btn" (click)="search()">Search</button>
                </div>

                <table class="poc-table" *ngIf="results().length > 0">
                    <thead>
                        <tr>
                            <th class="poc-col-check">Select</th>
                            <th>Case Number</th>
                            <th>Caption</th>
                        </tr>
                    </thead>
                    <tbody>
                        @for (c of results(); track c.CaseId) {
                            <tr>
                                <td class="poc-col-check">
                                    <input type="checkbox"
                                        [checked]="isSelected(c.CaseId)"
                                        (change)="toggleSelection(c.CaseId)">
                                </td>
                                <td>{{ c.CaseNumber }}</td>
                                <td>{{ c.CaseCaption }}</td>
                            </tr>
                        }
                    </tbody>
                </table>
                <p class="poc-muted" *ngIf="searched() && results().length === 0">No cases match the keyword.</p>

                <p class="poc-selected">Selected Cases : {{ selectedIds().size }}</p>
                <button type="button" class="poc-btn poc-btn-primary" (click)="generate()"
                    [disabled]="selectedIds().size === 0 || generating()">
                    {{ generating() ? 'Generating...' : 'Generate Documents' }}
                </button>
                <p class="poc-muted" *ngIf="statusMessage()">{{ statusMessage() }}</p>
            </aside>
        </div>
    </div>

    <!-- Missing merge field prompt (Requirement 1) -->
    <div class="poc-dialog-backdrop" *ngIf="showMissingDialog()">
        <div class="poc-dialog" role="dialog" aria-label="Enter missing merge field values">
            <h2>Missing merge field values</h2>
            <p class="poc-muted">The template contains merge fields the case data does not fill. Enter the missing values to continue.</p>
            <div class="poc-missing-list">
                @for (item of missingInputs(); track $index) {
                    <label class="poc-missing-item">
                        <span>Enter {{ item.label }}:</span>
                        <input type="text" class="poc-input" [(ngModel)]="item.value" name="missingValue{{ $index }}">
                    </label>
                }
            </div>
            <div class="poc-dialog-actions">
                <button type="button" class="poc-btn" (click)="cancelMissingDialog()">Cancel</button>
                <button type="button" class="poc-btn poc-btn-primary" (click)="submitMissingValues()">Continue</button>
            </div>
        </div>
    </div>
    `
})
export class App implements AfterViewInit, OnDestroy {
    @ViewChild('container', { static: true })
    public container?: DocumentEditorContainerComponent;

    /** Merge field explorer tree, built from GET /api/schema on startup. */
    public schemaGroups = signal<SchemaGroupView[]>([]);

    public keyword = '';
    public results = signal<CaseInfo[]>([]);
    public searched = signal(false);
    public selectedIds = signal<Set<number>>(new Set<number>());
    public generating = signal(false);
    public statusMessage = signal('');

    public showMissingDialog = signal(false);
    public missingInputs = signal<MissingFieldInput[]>([]);

    /** Left-panel manual merge field / region creation (Changes 5 + 6). */
    public customFieldName = '';
    public customRegionName = '';

    private disposed = false;

    constructor(private api: CaseApiService) { }

    ngAfterViewInit(): void {
        // Startup: schema first (independent of any search)...
        this.api.getSchema().then(schema => {
            if (this.disposed) { return; }
            this.schemaGroups.set(this.buildSchemaGroups(schema));
        }).catch(() => {
            this.statusMessage.set('Could not load the merge field schema from /api/schema.');
        });

        // ...then the bundled sample template into the editor (no blank start).
        this.api.getSampleTemplate().then(sfdt => {
            if (this.disposed) { return; }
            this.openInEditor(sfdt);
        }).catch(() => {
            this.statusMessage.set('Could not load the sample template from /api/template.');
        });
    }

    ngOnDestroy(): void {
        this.disposed = true;
    }

    private buildSchemaGroups(schema: FieldSchema): SchemaGroupView[] {
        return Object.keys(schema).map(groupName => ({
            name: groupName,
            fields: Object.keys(schema[groupName]).map(fieldName => ({
                name: fieldName,
                type: schema[groupName][fieldName]
            }))
        }));
    }

    /** The editor instance behind the container. */
    private get editor(): any | null {
        return this.container?.documentEditor ?? null;
    }

    private openInEditor(sfdt: string): void {
        const editor = this.editor;
        if (editor) { editor.open(sfdt); }
    }

    /**
     * Inserts a real MERGEFIELD (not plain text) at the cursor.
     * Used by the left-panel field explorer (click-to-insert).
     */
    public insertMergeField(fieldName: string): void {
        const editor = this.editor;
        if (!editor) { return; }
        editor.focusIn();
        // insertField builds a Word field: code MERGEFIELD <name>, result «<name>».
        editor.editor.insertField(`MERGEFIELD ${fieldName} \\* MERGEFORMAT`, `«${fieldName}»`);
    }

    /**
     * Manual merge field creation (Change 5): inserts a real MERGEFIELD for a
     * user-specified name that is not in the schema.
     */
    public insertCustomField(): void {
        const name = (this.customFieldName || '').trim();
        if (!name) { return; }
        this.insertMergeField(name);
        this.customFieldName = '';
    }

    /**
     * Manual region creation (Change 6): inserts a repeating region with the
     * user-specified name (not hardcoded):
     *
     *     «BeginName»
     *     «EndName»
     *
     * Both delimiters are real MERGEFIELDs, the same field types DocIO uses to
     * discover and execute regions server-side.
     */
    public insertCustomRegion(): void {
        const name = (this.customRegionName || '').trim();
        if (!name) { return; }
        const editor = this.editor;
        if (!editor) { return; }
        editor.focusIn();
        const e = editor.editor;
        e.insertField(
            `MERGEFIELD BeginGroup:${name} \\* MERGEFORMAT`,
            `«BeginGroup:${name}»`
        );

        e.handleEnterKey();
        e.insertField(
            `MERGEFIELD EndGroup:${name} \\* MERGEFORMAT`,
            `«EndGroup:${name}»`
        );
        this.customRegionName = '';
    }

    // ---------------- Search / selection ----------------

    public search(): void {
        this.searched.set(true);
        this.api.searchCases(this.keyword).then(results => {
            this.results.set(results);
        }).catch(() => {
            this.results.set([]);
            this.statusMessage.set('Case search failed. Is the ASP.NET Core service running?');
        });
    }

    public isSelected(caseId: number): boolean {
        return this.selectedIds().has(caseId);
    }

    public toggleSelection(caseId: number): void {
        const next = new Set(this.selectedIds());
        if (next.has(caseId)) { next.delete(caseId); } else { next.add(caseId); }
        this.selectedIds.set(next); // new Set instance so the signal fires
    }

    // ---------------- Generate ----------------

    /**
     * Serialize the current editor document as SFDT, merge the selected cases on the
     * server, and open the merged SFDT back in the editor. The result is previewed,
     * never downloaded.
     */
    public generate(): void {
        if (this.selectedIds().size === 0) { return; }
        this.generating.set(true);
        this.statusMessage.set('');
        this.runGenerate([]);
    }

    private runGenerate(additionalMergeData: AdditionalMergeValue[]): void {
        const templateSfdt = this.editor?.serialize();
        const request = {
            caseIds: Array.from(this.selectedIds()),
            additionalMergeData,
            templateSfdt
        };

        this.api.generate(request).then(result => {
            this.generating.set(false);
            if (result.kind === 'requires-input') {
                this.openMissingDialog(result.missingFields);
                return;
            }
            this.openInEditor(result.mergedSfdt);
            this.statusMessage.set(`Merged ${this.selectedIds().size} case(s) into one document.`);
        }).catch(() => {
            this.generating.set(false);
            this.statusMessage.set('Generation failed. Check the service log.');
        });
    }

    private openMissingDialog(missing: MissingField[]): void {
        this.missingInputs.set(missing.map(m => ({
            caseId: m.caseId,
            caseNumber: m.caseNumber,
            groupName: m.groupName,
            rowIndex: m.rowIndex,
            fieldName: m.fieldName,
            label: m.groupName
                ? `${m.fieldName} for item ${m.rowIndex} in group ${m.groupName} (case ${m.caseNumber})`
                : `${m.fieldName} (case ${m.caseNumber})`,
            value: ''
        })));
        this.showMissingDialog.set(true);
    }

    public cancelMissingDialog(): void {
        this.showMissingDialog.set(false);
        this.missingInputs.set([]);
        this.statusMessage.set('Generation cancelled.');
    }

    /**
     * Second request (Requirement 1): re-run the merge with the API data plus the
     * values entered in the dialog.
     */
    public submitMissingValues(): void {
        const additionalMergeData: AdditionalMergeValue[] = this.missingInputs()
            .filter(item => item.value && item.value.trim())
            .map(item => ({
                caseId: item.caseId,
                caseNumber: item.caseNumber,
                groupName: item.groupName,
                rowIndex: item.rowIndex,
                fieldName: item.fieldName,
                value: item.value.trim()
            }));

        this.showMissingDialog.set(false);
        this.generating.set(true);
        this.runGenerate(additionalMergeData);
    }
}