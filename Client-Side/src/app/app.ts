import { Component, ViewChild, signal, computed, effect, AfterViewInit, OnDestroy } from '@angular/core';
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
    FieldSchemaEntry,
    FieldType,
    CaseInfo,
    MissingField,
    AdditionalMergeValue,
    SupportingDocument
} from './case-api.service';

interface MissingFieldInput {
    caseId: number;
    caseNumber: string;
    groupName: string;
    rowIndex: number;
    fieldName: string;
    label: string;
    type: FieldType;
    value: string;
    /** number rules from the schema */
    integer?: boolean;
    min?: number | null;
    max?: number | null;
}

interface SchemaGroupView {
    name: string;
    fields: { name: string; type: FieldType }[];
}

/**
 * Static field metadata map: field name -> { type, label, groupName }.
 *
 * Built once from GET /api/schema on startup and used to (a) drive the missing-field
 * dialog input type and label and (b) validate typed values before they are
 * submitted back to the server as strings.
 *
 * Fields not present in the schema (custom fields typed in via the "Add Custom
 * Merge Field" button) default to string input.
 */
interface FieldMeta {
    type: FieldType;
    label?: string;
    groupName: string;
    integer?: boolean;
    min?: number | null;
    max?: number | null;
}

/**
 * Builds a human-readable label for a missing merge field. Case-level fields
 * (empty group or rowIndex === 0, which is how the server tags them) are labeled as
 * "FieldName (case CaseNumber)". Row-level fields inside a repeating group
 * add "for item N in group GroupName".
 */
function buildMissingFieldLabel(
    m: { fieldName: string; groupName: string; rowIndex: number; caseNumber: string },
    fieldMeta: Record<string, FieldMeta>
): string {
    const isCaseLevel = !m.groupName
        || m.rowIndex === 0;
    const friendly = fieldMeta[m.fieldName]?.label || m.fieldName;
    return isCaseLevel
        ? `${friendly} (case ${m.caseNumber})`
        : `${friendly} for item ${m.rowIndex} in group ${m.groupName} (case ${m.caseNumber})`;
}

/**
 * Validates one typed value. Returns null when valid, otherwise an error message.
 * - date:    non-empty + parseable by Date
 * - number:  (schema may require whole numbers and min/max) non-empty + strict decimal form (optional sign, digits, optional
 *            fractional digits). Rejects '1e5', 'Infinity', '0x11', blank input.
 * - boolean: always valid (a toggle always has a value)
 * - string:  any value is valid
 */
const NUMBER_PATTERN = /^-?\d+(\.\d+)?$/;

const INTEGER_PATTERN = /^-?\d+$/;

/** Number rules read from the schema (integer / min / max). */
interface NumberRules {
    integer?: boolean;
    min?: number | null;
    max?: number | null;
}

function validateTypedValue(type: FieldType, value: string, rules?: NumberRules): string | null {
    if (type === 'boolean') {
        return null;
    }
    if (!value || !value.trim()) {
        // String fields may be left empty; every other typed field is required.
        return type === 'string' ? null : 'Value is required.';
    }
    if (type === 'number') {
        const text = value.trim();
        if (rules?.integer) {
            if (!INTEGER_PATTERN.test(text)) {
                return 'Enter a whole number (digits only, e.g. 10).';
            }
        } else if (!NUMBER_PATTERN.test(text)) {
            return 'Enter a valid number (digits only, e.g. 250 or 19.99).';
        }
        const n = Number(text);
        if (rules?.min != null && n < rules.min) {
            return `Value must be at least ${rules.min}.`;
        }
        if (rules?.max != null && n > rules.max) {
            return `Value must be at most ${rules.max}.`;
        }
    } else if (type === 'date') {
        const d = new Date(value);
        if (Number.isNaN(d.getTime())) {
            return 'Enter a valid date.';
        }
    }
    return null;
}

/**
 * Coerces a typed value to the string that the server stores in
 * AdditionalMergeValue.value. Booleans become 'true' / 'false'; everything
 * else is returned as-is.
 */
function coerceTypedValue(type: FieldType, value: string): string {
    if (type === 'boolean') {
        return value === 'true' || value === 'on' ? 'true' : 'false';
    }
    return value;
}

/**
 *
 * Startup: schema loads and the default main template ("Notice of Hearing" from
 * wwwroot/Templates) opens in the editor, so the user can immediately search, select
 * cases and generate (no blank editor).
 *
 * Layout:
 * - Left:  Merge Fields (click-to-insert, driven by GET /api/schema)
 * - Center: Document Editor in Ribbon mode with a custom "Mail Merge" tab
 * - Right: Template loader + Search Cases + Documents to Merge + Generate
 *
 * Generate flow: the current editor document is serialized as SFDT, sent to
 * POST /api/generate with the selected case ids and selected supporting-document ids.
 * DocIO executes the relational nested mail merge for every selected case against the
 * main template AND every selected supporting document, then appends the merged
 * supporting documents to the main document to produce a single consolidated output.
 * The merged SFDT is opened directly in the editor (no download).
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
            <h1>Syncfusion DOCX Editor - Template Designer</h1>
            <p>Template authoring in the Document Editor, mail merge on the server.</p>
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
                <div class="poc-preview-banner" *ngIf="previewMode()">
                    <span><strong>Preview mode.</strong> The editor is read-only. Review the assembled document, then click <strong>Generate Document</strong> to finalize or <strong>Back to Edit</strong> to return to the template.</span>
                    <button type="button" class="poc-btn poc-btn-banner" (click)="exitPreview()">Back to Edit</button>
                </div>
                <div class="poc-action-toolbar">
                    <button type="button" class="poc-btn poc-btn-reset" (click)="reset()"
                        [disabled]="resetting()">
                        {{ resetting() ? 'Resetting...' : 'Reset' }}
                    </button>
                    <button type="button" class="poc-btn poc-btn-save" (click)="save()"
                        [disabled]="saving() || !editorHoldsTemplate() || !currentDocumentId || (!documentDirty() && savedFileExists())">
                        {{ saving() ? 'Saving...' : (documentDirty() ? 'Save Document *' : 'Save Document') }}
                    </button>
                    <button type="button" class="poc-btn poc-btn-preview" (click)="preview()"
                        [disabled]="selectedIds().size === 0 || previewing() || generating()">
                        {{ previewing() ? 'Building Preview...' : 'Preview Document' }}
                    </button>
                    <button type="button" class="poc-btn poc-btn-export" (click)="exportPdf()"
                        [disabled]="selectedIds().size === 0 || exportingPdf() || generating()">
                        {{ exportingPdf() ? 'Exporting...' : 'Export as PDF' }}
                    </button>
                    <button type="button" class="poc-btn poc-btn-open-pdf" (click)="openPdfDocument()"
                        [disabled]="convertingPdf()">
                        {{ convertingPdf() ? 'Converting...' : 'Open PDF Document' }}
                    </button>
                    <button type="button" class="poc-btn poc-btn-xsl" (click)="openXslDocument()"
                        [disabled]="convertingXsl()">
                        {{ convertingXsl() ? 'Converting...' : 'XSLT to Word Document' }}
                    </button>
                    <!-- Hidden file input: clicking Open PDF Document opens the browser's
                         native file menu dialog, which lets the user pick a PDF. The
                         input is given an id so the click handler can address it directly
                         regardless of ViewChild resolution timing. -->
                    <input id="pdfFileInput" type="file" accept="application/pdf,.pdf"
                        style="display:none"
                        (change)="onPdfPicked($event)">
                </div>
                <ejs-documenteditorcontainer
                    #container
                    height="100%"
                    style="display:block"
                    [enableToolbar]="!previewMode()"
                    [toolbarMode]="'Ribbon'"
                    (created)="onEditorCreated()"
                    (contentChange)="onContentChange()"
                    serviceUrl="/api/documenteditor/">
                </ejs-documenteditorcontainer>
            </section>

            <!-- RIGHT PANEL: Template + Search Cases + Documents to Merge + Generate -->
            <aside class="poc-panel poc-cases">
                <h2>Template</h2>
                <div class="poc-search-row">
                    <select class="poc-input" [(ngModel)]="selectedTemplateId" name="templateSelect"
                        [disabled]="loadingTemplate() || loadableTemplates().length === 0">
                        @for (t of loadableTemplates(); track t.id) {
                            <option [value]="t.id">{{ t.displayName }}</option>
                        }
                    </select>
                    <button type="button" class="poc-btn" (click)="loadSelectedTemplate()"
                        [disabled]="!selectedTemplateId || loadingTemplate()">
                        {{ loadingTemplate() ? 'Loading...' : 'Load' }}
                    </button>
                </div>
                <p class="poc-muted">The editor always starts from a template created from scratch. Save it, then open the saved document from this list.</p>

                <div class="poc-divider"></div>

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

                <div class="poc-divider"></div>

                <h2>Documents to Merge</h2>
                <p class="poc-muted">Select supporting documents. Their merge fields are filled with the case data and combined with the main template into a single document.</p>
                <div class="poc-template-list">
                    @for (t of supportingDocuments(); track t.id) {
                        <label class="poc-template-item">
                            <input type="checkbox"
                                [checked]="isAppendedSelected(t.id)"
                                (change)="toggleAppendedSelection(t.id)">
                            <span>{{ t.displayName }}</span>
                        </label>
                    }
                    <p class="poc-muted" *ngIf="supportingDocuments().length === 0">No supporting documents available.</p>
                </div>
                <p class="poc-selected">Selected Documents : {{ selectedAppendedIds().size }}</p>

                <div class="poc-divider"></div>

                <button type="button" class="poc-btn poc-btn-primary poc-btn-block" (click)="generate()"
                    [disabled]="selectedIds().size === 0 || generating()">
                    {{ generating() ? 'Generating...' : 'Generate Document' }}
                </button>
                <p class="poc-muted" *ngIf="statusMessage()">{{ statusMessage() }}</p>
            </aside>
        </div>
    </div>

    <!-- Missing merge field prompt (typed inputs based on /api/schema) -->
    <div class="poc-dialog-backdrop" *ngIf="showMissingDialog()">
        <div class="poc-dialog" role="dialog" aria-label="Enter missing merge field values">
            <h2>Missing merge field values</h2>
            <p class="poc-muted">The template contains merge fields the case data does not fill. Enter the missing values to continue.</p>
            <div class="poc-missing-list">
                @for (item of missingInputs(); track $index) {
                    <div class="poc-missing-item">
                        <span>Enter {{ item.label }}:</span>

                        @switch (item.type) {
                            @case ('date') {
                                <input type="date" class="poc-input"
                                    [(ngModel)]="item.value"
                                    name="missingValue{{ $index }}"
                                    [class.poc-input-invalid]="invalidFor(item)">
                            }
                            @case ('number') {
                                <!-- type=text + inputmode=decimal: the browser reports
                                     '' for invalid native number input, which would
                                     lose the raw text and show the wrong message. -->
                                <input type="text" class="poc-input"
                                    [attr.inputmode]="item.integer ? 'numeric' : 'decimal'"
                                    [(ngModel)]="item.value"
                                    name="missingValue{{ $index }}"
                                    [placeholder]="item.integer ? 'Whole number, e.g. 10' : 'e.g. 250 or 19.99'"
                                    [class.poc-input-invalid]="invalidFor(item)">
                            }
                            @case ('boolean') {
                                <label class="poc-checkbox-row">
                                    <input type="checkbox"
                                        [checked]="item.value === 'true'"
                                        (change)="onBooleanChange(item, $event)"
                                        name="missingValue{{ $index }}">
                                    <span>{{ item.value === 'true' ? 'Yes' : 'No' }}</span>
                                </label>
                            }
                            @default {
                                <input type="text" class="poc-input"
                                    [(ngModel)]="item.value"
                                    name="missingValue{{ $index }}">
                            }
                        }

                        @if (validationError(item); as err) {
                            <span class="poc-field-error">{{ err }}</span>
                        }
                    </div>
                }
            </div>
            <div class="poc-dialog-actions">
                <button type="button" class="poc-btn" (click)="cancelMissingDialog()">Cancel</button>
                <button type="button" class="poc-btn poc-btn-primary"
                    (click)="submitMissingValues()"
                    [disabled]="hasValidationErrors()">Continue</button>
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

    /** Static field metadata map: field name -> { type, label, groupName }.
     * Used to drive the missing-field dialog (input type, label) and to
     * validate typed values before they are submitted as strings. */
    public fieldMeta: Record<string, FieldMeta> = {};

    public keyword = '';
    public results = signal<CaseInfo[]>([]);
    public searched = signal(false);
    public selectedIds = signal<Set<number>>(new Set<number>());
    public generating = signal(false);
    public previewing = signal(false);
    /** True when the editor is showing the assembled output in viewer/read-only mode. */
    public previewMode = signal(false);
    /** Cached SFDT of the last successful preview so Generate can reuse it without re-merging. */
    private previewedSfdt: string | null = null;
    /** Mode to resume after the missing-field dialog is submitted. */
    private pendingMode: 'preview' | 'generate' | 'export-pdf' | null = null;
    /** File id of the default main template, resolved from the template list. */
    private mainTemplateId = '';
    /** True when the editor currently holds the default main template. */
    private editorHoldsDefaultTemplate = false;
    public statusMessage = signal('');

    /** Supporting documents (templates) loaded from GET /api/templates on startup. */
    public templates = signal<SupportingDocument[]>([]);
    public selectedTemplateId = '';
    public loadingTemplate = signal(false);

    /**
     * The subset of templates that can be APPENDED to the merged output - i.e.
     * everything except the default main template (which is the document being
     * merged, so appending it to itself makes no sense).
     */
    public supportingDocuments = computed(() => this.templates().filter(t => !t.isMainTemplate));

    /**
     * Templates that exist as files in the templates folder and can be opened from the
     * dropdown. The main template appears here only after it has been saved once.
     */
    public loadableTemplates = computed(() => this.templates().filter(t => t.exists !== false));

    /**
     * Id (file name) of the template currently open in the editor. Tracked so the
     * "Save Document" action knows which file in the templates folder to overwrite.
     * Empty when the editor holds a document that is not a template (e.g. a merged
     * output or a brand-new template authored from scratch).
     */
    public currentDocumentId = '';
    /** True when the editor holds a template that can be saved back to disk. */
    public editorHoldsTemplate = signal(false);
    /** True when the editor content differs from the last loaded/saved state. */
    public documentDirty = signal(false);
    /**
     * True when the file Save writes to already exists in the templates folder. While it
     * does not, Save is enabled even without edits, so the template created from scratch
     * can be saved as-is to create the file.
     */
    public savedFileExists = signal(true);
    public saving = signal(false);
    public exportingPdf = signal(false);
    public convertingPdf = signal(false);
    public convertingXsl = signal(false);
    public resetting = signal(false);

    /** Selected supporting documents to append after the merge. */
    public selectedAppendedIds = signal<Set<string>>(new Set<string>());

    public showMissingDialog = signal(false);
    public missingInputs = signal<MissingFieldInput[]>([]);

    /** Left-panel manual merge field / region creation (Changes 5 + 6). */
    public customFieldName = '';
    public customRegionName = '';

    private disposed = false;
    /** True while a document is being opened - suppresses dirty tracking. */
    private loadingDocument = false;
    /**
     * The template (SFDT) that was in the editor when a preview opened the merged
     * output. While a preview is showing, the editor holds the MERGED document, so
     * Generate / Export must merge THIS snapshot, and Back to Edit restores it.
     */
    private templateSnapshot: string | null = null;
    /** Whether the snapshot had unsaved changes (restored on Back to Edit). */
    private snapshotDirty = false;
    /** Handle of the settle timer that ends the load-suppression window. */
    private loadingSettleHandle: number | null = null;

    constructor(private api: CaseApiService) {
        // Mirror the previewMode signal onto the editor's actual read-only state.
        // The container's `restrictEditing` input only toggles the Restrict-Editing
        // pane; the editor's editing permission is controlled by
        // documentEditor.isReadOnly. The effect ensures the two stay in sync on
        // every preview-mode transition.
        effect(() => {
            this.applyEditorReadOnly(this.previewMode());
        });
    }

    ngAfterViewInit(): void {
        // NOTE: the inner documentEditor is not created yet when this lifecycle hook
        // runs - the EJ2 container builds it asynchronously and fires the `created`
        // event when ready (see onEditorCreated, where the contentChange
        // subscription is attached).

        // Startup: schema first (independent of any search)...
        this.api.getSchema().then(schema => {
            if (this.disposed) { return; }
            this.fieldMeta = this.buildFieldMeta(schema);
            this.schemaGroups.set(this.buildSchemaGroups(schema));
        }).catch(() => {
            this.statusMessage.set('Could not load the merge field schema from /api/schema.');
        });

        // ...then the default main template ("Notice of Hearing") from the templates folder
        // into the editor (no blank start). The server falls back to the bundled sample
        // template if no DOCX is present in the configured folder yet - in that case it
        // still reports the target file name so Save can CREATE the file on first save.
        this.api.getDefaultTemplate().then(result => {
            if (this.disposed) { return; }
            this.openInEditor(result.sfdt);
            // The editor now holds the main template; Save Document writes to the
            // reported file name (creating it if it does not exist yet). Older
            // backends do not send fileName - fall back to the constant default.
            this.editorHoldsDefaultTemplate = true;
            this.editorHoldsTemplate.set(true);
            this.currentDocumentId = result.fileName || this.mainTemplateId || 'Notice of Hearing.docx';
            this.savedFileExists.set(result.exists !== false);
            this.markClean();
        }).catch(() => {
            this.statusMessage.set('Could not load the default template from /api/template.');
        });

        // ...and the template list (drives the "Documents to Merge" panel and the
        // template picker dropdown).
        this.api.listTemplates().then(items => {
            if (this.disposed) { return; }
            this.templates.set(items);
            this.mainTemplateId = items.find(t => t.isMainTemplate)?.id ?? '';
            // Backfill the save target if the default template loaded before the list
            // and the server did not report a file name.
            if (this.editorHoldsDefaultTemplate && !this.currentDocumentId) {
                this.currentDocumentId = this.mainTemplateId || 'Notice of Hearing.docx';
            }
            // Default the dropdown to the first SAVED template (the main template is
            // listed only once it exists as a file).
            const loadable = items.filter(t => t.exists !== false);
            if (loadable.length > 0 && !this.selectedTemplateId) {
                this.selectedTemplateId = loadable[0].id;
            }
        }).catch(() => {
            this.statusMessage.set('Could not list supporting documents from /api/templates.');
        });
    }

    /**
     * Re-reads the template list (the folder is read on every call). Called after a
     * save so a newly created file appears in the dropdown and can be opened again.
     */
    private refreshTemplateList(selectId?: string): void {
        this.api.listTemplates().then(items => {
            if (this.disposed) { return; }
            this.templates.set(items);
            this.mainTemplateId = items.find(t => t.isMainTemplate)?.id ?? this.mainTemplateId;
            const loadable = items.filter(t => t.exists !== false);
            if (selectId && loadable.some(t => t.id === selectId)) {
                this.selectedTemplateId = selectId;
            } else if (!loadable.some(t => t.id === this.selectedTemplateId)) {
                this.selectedTemplateId = loadable.length > 0 ? loadable[0].id : '';
            }
        }).catch(() => { /* the list is refreshed again on the next page load */ });
    }

    ngOnDestroy(): void {
        this.disposed = true;
    }

    /**
     * The EJ2 container fires `created` once the inner DocumentEditor exists. This is
     * the earliest reliable point to correct the editor's default read-only state
     * (it defaults to read-only).
     */
    public onEditorCreated(): void {
        // The editor just became available; make sure its read-only state matches
        // the current preview mode (it defaults to read-only).
        this.applyEditorReadOnly(this.previewMode());
    }

    /**
     * Bound to the container's (contentChange) output. Fires for every edit in the
     * document - enabling the Save Document action.
     */
    public onContentChange(): void {
        this.markDirty();
    }

    private buildSchemaGroups(schema: FieldSchema): SchemaGroupView[] {
        return Object.keys(schema).map(groupName => {
            const group = (schema as Record<string, { [fieldName: string]: FieldSchemaEntry } | undefined>)[groupName];
            if (!group) {
                return { name: groupName, fields: [] };
            }
            return {
                name: groupName,
                fields: Object.keys(group).map(fieldName => ({
                    name: fieldName,
                    type: group[fieldName].type
                }))
            };
        });
    }

    /**
     * Flattens the grouped schema into a fieldName -> { type, label, groupName } map
     * the rest of the app can index directly. Custom fields typed in via the
     * "Add Custom Merge Field" button are not in the schema; they default to a
     * string input at the dialog site.
     */
    private buildFieldMeta(schema: FieldSchema): Record<string, FieldMeta> {
        const meta: Record<string, FieldMeta> = {};
        const groups = schema as Record<string, { [fieldName: string]: FieldSchemaEntry } | undefined>;
        for (const groupName of Object.keys(groups)) {
            const group = groups[groupName];
            if (!group) continue;
            for (const fieldName of Object.keys(group)) {
                meta[fieldName] = {
                    type: group[fieldName].type,
                    label: group[fieldName].label,
                    groupName,
                    integer: group[fieldName].integer,
                    min: group[fieldName].min,
                    max: group[fieldName].max
                };
            }
        }
        return meta;
    }

    /** The editor instance behind the container. */
    private get editor(): any | null {
        return this.container?.documentEditor ?? null;
    }

    private openInEditor(sfdt: string): void {
        const editor = this.editor;
        if (!editor) { return; }
        // Suppress dirty tracking while the document loads: editor.open() can fire
        // contentChange for the initial render, sometimes on a later task (after
        // layout), which would otherwise re-dirty a freshly loaded document. A short
        // settle window is robust against both sync and async firing.
        this.loadingDocument = true;
        if (this.loadingSettleHandle !== null) {
            clearTimeout(this.loadingSettleHandle);
        }
        this.loadingSettleHandle = window.setTimeout(() => {
            this.loadingDocument = false;
            this.loadingSettleHandle = null;
        }, 300);

        editor.open(sfdt);
        // The editor's read-only state may have been reset by open(). Re-apply the
        // current preview-mode setting so a freshly opened document stays in the
        // expected read-only / editable state.
        this.applyEditorReadOnly(this.previewMode());
        // A freshly opened document has no unsaved changes.
        this.markClean();
    }

    /** Marks the editor content as modified - enables the Save Document action. */
    private markDirty(): void {
        if (this.loadingDocument) { return; }
        this.documentDirty.set(true);
    }

    /** Marks the editor content as saved/unchanged - disables the Save action. */
    private markClean(): void {
        this.documentDirty.set(false);
    }

    /**
     * Lock or unlock the editor's read-only state. The container's `restrictEditing`
     * input only toggles the Restrict-Editing pane - the editor's actual editing
     * permission is controlled by `documentEditor.isReadOnly`. Every preview-mode
     * transition calls this so the banner and the editing state stay in sync.
     */
    private applyEditorReadOnly(readOnly: boolean): void {
        const editor = this.editor;
        if (editor && typeof editor.isReadOnly === 'boolean') {
            try { editor.isReadOnly = readOnly; } catch { /* the editor may not be ready yet */ }
        }
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

    // ---------------- Template loader + supporting document selection ----------------

    /**
     * Replaces the editor's current document with the chosen template.
     * Used by the "Template" dropdown at the top of the right panel.
     */
    public loadSelectedTemplate(): void {
        const id = this.selectedTemplateId;
        if (!id) { return; }
        this.loadingTemplate.set(true);
        this.api.getTemplate(id).then(sfdt => {
            if (this.disposed) { return; }
            this.openInEditor(sfdt);
            // Remember which file the editor now holds so "Save Document" can
            // overwrite it. Leaving preview mode makes sense here too.
            this.currentDocumentId = id;
            this.savedFileExists.set(true);
            this.editorHoldsDefaultTemplate = false;
            this.editorHoldsTemplate.set(true);
            this.previewMode.set(false);
            this.previewedSfdt = null;
            this.templateSnapshot = null;
            this.statusMessage.set(`Loaded template '${id}'.`);
        }).catch(() => {
            this.statusMessage.set(`Could not load template '${id}'.`);
        }).finally(() => {
            this.loadingTemplate.set(false);
        });
    }

    public isAppendedSelected(id: string): boolean {
        return this.selectedAppendedIds().has(id);
    }

    // ---------------- Save ----------------

    /**
     * Saves the template currently open in the editor back to the templates folder
     * (wwwroot/Templates). The editor document is serialized as SFDT and POSTed to
     * /api/templates/{id}/save; the server converts it back to DOCX and overwrites
    /**
     * Resets the editor to the default main template ("Notice of Hearing" from the
     * templates folder). Clears any in-progress merge / preview state and reloads
     * the bundled default template.
     */
    public reset(): void {
        if (this.resetting()) { return; }
        this.resetting.set(true);
        this.statusMessage.set('Resetting: creating the template from scratch...');

        // Drop every transient state so the editor and tracking signals agree with
        // the freshly loaded default template.
        this.previewMode.set(false);
        this.applyEditorReadOnly(false);
        this.previewedSfdt = null;
        this.templateSnapshot = null;
        this.previewing.set(false);
        this.generating.set(false);
        this.exportingPdf.set(false);
        this.statusMessage.set('');

        this.api.getDefaultTemplate().then(result => {
            if (this.disposed) { return; }
            this.openInEditor(result.sfdt);
            // The editor now holds the default template again - it is savable, not a
            // merged output.
            this.currentDocumentId = result.fileName || this.mainTemplateId || 'Notice of Hearing.docx';
            this.savedFileExists.set(result.exists !== false);
            this.editorHoldsDefaultTemplate = true;
            this.editorHoldsTemplate.set(true);
            this.statusMessage.set('Reset: the template was created from scratch.');
        }).catch(() => {
            this.statusMessage.set('Could not reload the default template.');
        }).finally(() => {
            this.resetting.set(false);
        });
    }

    /**
     * Saves the template currently open in the editor back to the templates folder:
     * the server converts the SFDT payload back to DOCX and overwrites the file the
     * user opened.
     */
    public save(): void {
        // The editor holds the assembled/merged output in preview mode - that is not
        // a template and must not overwrite a template file.
        if (this.previewMode()) {
            this.statusMessage.set('Exit preview mode before saving a template.');
            return;
        }

        // The editor may hold a generated (merged) document. Saving that over a
        // template would replace the template with merged content.
        if (!this.editorHoldsTemplate()) {
            this.statusMessage.set('The editor holds a generated document, not a template. Load a template to save.');
            return;
        }

        const id = this.currentDocumentId;
        if (!id) {
            this.statusMessage.set('Load a template before saving. There is no file to write to.');
            return;
        }

        const sfdt = this.editor?.serialize();
        if (!sfdt) {
            this.statusMessage.set('There is nothing to save.');
            return;
        }

        this.saving.set(true);
        this.statusMessage.set('');
        this.api.saveTemplate(id, sfdt).then(result => {
            // The on-disk file now matches the editor content.
            this.markClean();
            this.savedFileExists.set(true);
            this.statusMessage.set(`Saved '${result.displayName}' to the templates folder.`);
            // A first save creates the file: list it so it can be opened from the dropdown.
            this.refreshTemplateList(result.id);
        }).catch(() => {
            this.statusMessage.set(`Could not save '${id}'. Check the service log.`);
        }).finally(() => {
            this.saving.set(false);
        });
    }

    // ---------------- Open PDF Document ----------------

    /**
     * Opens the browser's native file menu dialog so the user can pick a PDF.
     * The picked file is handled by onPdfPicked().
     */
    public openPdfDocument(): void {
        if (this.convertingPdf()) { return; }
        const input = document.getElementById('pdfFileInput') as HTMLInputElement | null;
        input?.click();
    }

    /**
     * Sends the picked PDF to the server, which converts it to a Word document and
     * returns the SFDT. The converted document opens in the editor.
     *
     * The converted document is NOT written to the templates folder, so the editor
     * no longer holds a template that Save Document could overwrite - the user can
     * still save it explicitly once a template file name is chosen.
     */
    public onPdfPicked(ev: Event): void {
        const input = ev.target as HTMLInputElement;
        const file = input.files && input.files.length > 0 ? input.files[0] : null;
        // Reset the input so picking the same file twice in a row still fires.
        input.value = '';

        if (!file) { return; }
        if (!file.name.toLowerCase().endsWith('.pdf')) {
            this.statusMessage.set('Only PDF files are supported.');
            return;
        }

        this.convertingPdf.set(true);
        this.statusMessage.set(`Converting '${file.name}'...`);

        this.api.convertPdfToWord(file).then(sfdt => {
            if (this.disposed) { return; }
            this.openInEditor(sfdt);
            // The editor holds a converted document, not a template from the
            // templates folder - Save Document must not overwrite one.
            this.currentDocumentId = '';
            this.editorHoldsDefaultTemplate = false;
            this.editorHoldsTemplate.set(false);
            this.previewMode.set(false);
            this.previewedSfdt = null;
            this.templateSnapshot = null;
            this.statusMessage.set(`Converted '${file.name}' to Word and opened it in the editor.`);
        }).catch((err: unknown) => {
            let message: string;
            if (err && typeof err === 'object' && 'error' in err) {
                const inner = (err as { error?: { message?: string } }).error;
                message = inner?.message ?? `Could not convert '${file.name}'.`;
            } else {
                message = `Could not convert '${file.name}'.`;
            }
            this.statusMessage.set(message);
        }).finally(() => {
            this.convertingPdf.set(false);
        });
    }

    // ---------------- XSL to Word Document ----------------

    /**
     * Converts the XSL/XSLT stylesheet from the templates folder (InputXSL.txt) into
     * a Word mail-merge template and opens it in the editor. The server turns the
     * stylesheet's xsl:for-each into merge regions and xsl:value-of into merge
     * fields, so the result works with the same mail-merge pipeline as the other
     * templates.
     */
    public openXslDocument(): void {
        if (this.convertingXsl()) { return; }
        this.convertingXsl.set(true);
        this.statusMessage.set('Converting the XSL template to a Word document...');

        this.api.convertXslToWord().then(sfdt => {
            if (this.disposed) { return; }
            this.openInEditor(sfdt);
            // The editor holds a freshly generated template - it is not backed by a
            // file in the templates folder yet, so Save Document stays disabled until
            // the user saves it to a real template.
            this.currentDocumentId = '';
            this.editorHoldsDefaultTemplate = false;
            this.editorHoldsTemplate.set(false);
            this.previewMode.set(false);
            this.previewedSfdt = null;
            this.templateSnapshot = null;
            this.statusMessage.set('Converted the XSL template to a Word document and opened it in the editor.');
        }).catch((err: unknown) => {
            let message: string;
            if (err && typeof err === 'object' && 'error' in err) {
                const inner = (err as { error?: { message?: string } }).error;
                message = inner?.message ?? 'Could not convert the XSL template.';
            } else {
                message = 'Could not convert the XSL template.';
            }
            this.statusMessage.set(message);
        }).finally(() => {
            this.convertingXsl.set(false);
        });
    }

    // ---------------- Supporting document selection ----------------

    public toggleAppendedSelection(id: string): void {
        const next = new Set(this.selectedAppendedIds());
        if (next.has(id)) { next.delete(id); } else { next.add(id); }
        this.selectedAppendedIds.set(next);
    }

    // ---------------- Generate ----------------

    /**
     * Preview click: run the mail merge using the selected case data + supporting
     * documents, then display the assembled SFDT in the editor in viewer/read-only
     * mode. The user can review and either click Generate Document to finalize or
     * Back to Edit to return to the template.
     *
     * The merge call is identical to Generate - same endpoint, same data - only the
     * editor's mode differs.
     */
    public preview(): void {
        if (this.selectedIds().size === 0) { return; }
        this.previewing.set(true);
        this.generating.set(true);
        this.statusMessage.set('');
        this.runMerge('preview', []);
    }

    /**
     * Final click: serialise the current editor document, POST to /api/generate on the
     * server, and open the merged SFDT back in the editor. The result is previewed,
     * never downloaded.
     *
     * Always runs a fresh merge with the currently selected case data and supporting
     * documents, even after a Preview - the user may have changed the selection and
     * the latest data must drive the output. The cached previewedSfdt is intentionally
     * discarded so the editor ends up with the freshly generated document.
     */
    public generate(): void {
        if (this.selectedIds().size === 0) { return; }

        // If a preview is showing, stay in preview mode until the merge succeeds (the
        // success handler leaves it). Cancelling the missing-value dialog then leaves
        // the editor exactly as it was. The merge uses the template snapshot, not the
        // merged preview that the editor currently holds.
        this.previewedSfdt = null;

        this.generating.set(true);
        this.statusMessage.set('');
        this.runMerge('generate', []);
    }

    /**
     * Mail-merge the selected cases + supporting documents and download the result
     * as a PDF. Runs the same merge pipeline as Preview / Generate; the server
     * converts the merged Word document to PDF using Syncfusion DocIORenderer.
     */
    public exportPdf(): void {
        if (this.selectedIds().size === 0) { return; }
        this.exportingPdf.set(true);
        this.generating.set(true);
        this.statusMessage.set('');
        this.runExportPdf([]);
    }

    /**
     * Re-runs the PDF export after the user supplied missing-field values.
     * Stays in export mode across the prompt so the second pass actually delivers
     * a PDF download.
     */
    private runExportPdf(additionalMergeData: AdditionalMergeValue[]): void {
        const templateSfdt = this.templateSnapshot ?? this.editor?.serialize();
        const request = {
            caseIds: Array.from(this.selectedIds()),
            additionalMergeData,
            templateSfdt,
            appendedDocumentIds: Array.from(this.selectedAppendedIds())
        };

        this.api.generatePdf(request).then(result => {
            this.exportingPdf.set(false);
            this.generating.set(false);

            if (result.kind === 'requires-input') {
                this.openMissingDialog(result.missingFields, 'export-pdf');
                return;
            }

            // Trigger the browser download.
            const url = URL.createObjectURL(result.blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = result.fileName;
            document.body.appendChild(a);
            a.click();
            a.remove();
            URL.revokeObjectURL(url);

            const appendedCount = this.selectedAppendedIds().size;
            const caseCount = this.selectedIds().size;
            const suffix = appendedCount > 0
                ? ` and ${appendedCount} supporting document${appendedCount === 1 ? '' : 's'}`
                : '';
            this.statusMessage.set(`Downloaded ${result.fileName} (${caseCount} case(s)${suffix}).`);
        }).catch(() => {
            this.exportingPdf.set(false);
            this.generating.set(false);
            this.statusMessage.set('PDF export failed. Check the service log.');
        });
    }

    /**
     * Drops preview mode and reloads the template that was open before the preview
     * so the user can continue editing it.
     */
    public exitPreview(): void {
        if (!this.previewMode()) { return; }
        this.previewMode.set(false);
        this.previewedSfdt = null;

        // Restore exactly what was in the editor before the preview (including unsaved
        // edits, and a template created from scratch that has no file yet).
        const snapshot = this.templateSnapshot;
        const wasDirty = this.snapshotDirty;
        this.templateSnapshot = null;
        if (snapshot) {
            this.openInEditor(snapshot);
            this.documentDirty.set(wasDirty);
            this.editorHoldsTemplate.set(true);
            this.statusMessage.set('Returned to template editing.');
            return;
        }

        // No snapshot (should not happen): reload the saved file, else start from scratch.
        const id = this.currentDocumentId;
        const fromScratch = () => this.api.getDefaultTemplate().then(r => r.sfdt);
        const load: Promise<string> = id
            ? this.api.getTemplate(id).catch(() => fromScratch())
            : fromScratch();

        load.then(sfdt => {
            if (this.disposed) { return; }
            this.openInEditor(sfdt);
            this.editorHoldsTemplate.set(true);
            this.statusMessage.set('Returned to template editing.');
        }).catch(() => {
            this.statusMessage.set('Could not reload the template.');
        });
    }

    /**
     * Shared merge call used by both preview and generate. Mode controls whether the
     * result is loaded into the editor in viewer mode (preview) or as the final
     * editable document (generate).
     */
    private runMerge(mode: 'preview' | 'generate', additionalMergeData: AdditionalMergeValue[]): void {
        // While a preview is showing the editor holds the merged output, so merge the
        // template that was in the editor before the preview instead.
        const templateSfdt = this.templateSnapshot ?? this.editor?.serialize();
        const request = {
            caseIds: Array.from(this.selectedIds()),
            additionalMergeData,
            templateSfdt,
            appendedDocumentIds: Array.from(this.selectedAppendedIds())
        };

        this.api.generate(request).then(result => {
            this.previewing.set(false);
            this.generating.set(false);
            if (result.kind === 'requires-input') {
                this.openMissingDialog(result.missingFields, mode);
                return;
            }

            if (mode === 'preview') {
                // Remember the template (and whether it had unsaved edits) so Back to Edit
                // and Generate can use it while the editor shows the merged output.
                if (!this.templateSnapshot) {
                    this.templateSnapshot = templateSfdt ?? null;
                    this.snapshotDirty = this.documentDirty();
                }
                this.previewedSfdt = result.mergedSfdt;
                this.previewMode.set(true);
                this.openInEditor(result.mergedSfdt);
                const appendedCount = this.selectedAppendedIds().size;
                const caseCount = this.selectedIds().size;
                const suffix = appendedCount > 0
                    ? ` and merged ${appendedCount} supporting document${appendedCount === 1 ? '' : 's'}`
                    : '';
                this.statusMessage.set(`Preview ready: merged ${caseCount} case(s)${suffix}. Review and Generate to finalize.`);
            } else {
                this.previewMode.set(false);
                this.previewedSfdt = null;
                this.templateSnapshot = null;
                this.openInEditor(result.mergedSfdt);
                // The editor now holds the GENERATED output, not a template - saving it
                // back would overwrite a template file with merged content.
                this.editorHoldsTemplate.set(false);
                const appendedCount = this.selectedAppendedIds().size;
                const caseCount = this.selectedIds().size;
                const suffix = appendedCount > 0
                    ? ` and merged ${appendedCount} supporting document${appendedCount === 1 ? '' : 's'}`
                    : '';
                this.statusMessage.set(`Generated ${caseCount} case(s) into one document${suffix}.`);
            }
        }).catch(() => {
            this.previewing.set(false);
            this.generating.set(false);
            this.statusMessage.set('Merge failed. Check the service log.');
        });
    }

    private openMissingDialog(missing: MissingField[], mode: 'preview' | 'generate' | 'export-pdf'): void {
        this.pendingMode = mode;
        this.missingInputs.set(missing.map(m => {
            // Resolve the type from the static field-metadata map; unknown fields
            // (custom fields typed in via "Add Custom Merge Field") default to string.
            const meta = this.fieldMeta[m.fieldName];
            const type: FieldType = meta?.type ?? 'string';
            return {
                caseId: m.caseId,
                caseNumber: m.caseNumber,
                groupName: m.groupName,
                rowIndex: m.rowIndex,
                fieldName: m.fieldName,
                type,
                label: buildMissingFieldLabel(m, this.fieldMeta),
                value: this.defaultValueFor(type),
                integer: meta?.integer,
                min: meta?.min,
                max: meta?.max
            };
        }));
        this.showMissingDialog.set(true);
    }

    public cancelMissingDialog(): void {
        this.showMissingDialog.set(false);
        this.missingInputs.set([]);
        this.pendingMode = null;
        this.statusMessage.set('Generation cancelled.');
    }

    /**
     * Called from the boolean checkbox. Checkbox state ('true'/'false') is the
     * only thing the server cares about; the visible label is a side effect.
     */
    public onBooleanChange(item: MissingFieldInput, ev: Event): void {
        const checked = (ev.target as HTMLInputElement).checked;
        const next = this.missingInputs().map(x => x === item ? { ...x, value: checked ? 'true' : 'false' } : x);
        this.missingInputs.set(next);
    }

    /** Validation helper used by the template. */
    public validationError(item: MissingFieldInput): string | null {
        return validateTypedValue(item.type, item.value, item);
    }

    public invalidFor(item: MissingFieldInput): boolean {
        return this.validationError(item) != null;
    }

    public hasValidationErrors(): boolean {
        return this.missingInputs().some(item => this.validationError(item) != null);
    }

    private defaultValueFor(type: FieldType): string {
        if (type === 'boolean') return 'false';
        return '';
    }

    /**
     * Second request (Requirement 1): re-run the merge with the API data plus the
     * values entered in the dialog. Values are coerced to strings here so the
     * server contract is unchanged.
     */
    public submitMissingValues(): void {
        const additionalMergeData: AdditionalMergeValue[] = this.missingInputs()
            .map(item => ({
                caseId: item.caseId,
                caseNumber: item.caseNumber,
                groupName: item.groupName,
                rowIndex: item.rowIndex,
                fieldName: item.fieldName,
                value: coerceTypedValue(item.type, item.value)
            }));

        // Preserve the originating action: Preview -> back to preview, Generate ->
        // finalize, Export as PDF -> trigger the PDF download on the second pass.
        const mode: 'preview' | 'generate' | 'export-pdf' = this.pendingMode
            ?? (this.exportingPdf() ? 'export-pdf'
                : this.previewing() ? 'preview' : 'generate');
        this.pendingMode = null;

        this.showMissingDialog.set(false);

        if (mode === 'export-pdf') {
            this.runExportPdf(additionalMergeData);
            return;
        }

        this.generating.set(true);
        this.runMerge(mode, additionalMergeData);
    }
}