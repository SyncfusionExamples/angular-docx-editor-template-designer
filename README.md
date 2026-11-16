# Syncfusion® Angular DOCX Editor – Template Designer

## Introduction

This sample extends the [Syncfusion® Angular DOCX Editor](https://www.syncfusion.com/docx-editor-sdk/angular-docx-editor?utm_source=github&utm_medium=listing&utm_campaign=github-github-documenteditor-examples) and an ASP.NET Core backend into a legal document generation workflow. A user authors a "Notice of Hearing" template in the Angular Document Editor, searches for case records, selects one or more cases, and generates a merged document with Syncfusion DocIO on the server. The merged result opens directly inside the editor. It is intended for developers building document generation for business records such as court cases and their parties.

## Project Features

- **Merge field panel** — click a field to insert a real merge field into the template, or add custom fields and regions by name.
- **Case search and selection** — search by case number or caption; full details load only for the selected cases.
- **Repeating regions** — one block per party, one block per case, all in a single document.
- **Missing value prompt** — fields with no data are asked for in a dialog (date picker, number box, Yes/No checkbox), then the merge re-runs.
- **Supporting documents** — any `.docx` dropped in the templates folder can be appended to the merged output.
- **Preview, Generate, Export as PDF** — the merged document opens in the editor, or downloads as PDF.
- **Open PDF Document** and **XSLT to Word Document** — convert a PDF or an XSLT stylesheet into an editable Word template.

## How to Run the Sample

Both projects must run: the ASP.NET Core backend and the Angular client.

```bash
# 1. Backend  (http://localhost:62870)
cd Server-Side/src
dotnet restore
dotnet run

# 2. Client   (http://localhost:4200)
cd Client-Side
npm install
npm start
```

The Angular dev server forwards `/api` calls to the backend. The target is in `Client-Side/proxy.conf.json` and must match the backend URL in `Server-Side/src/Properties/launchSettings.json` (both use `http://localhost:62870`). Add your Syncfusion license key in `Server-Side/src/Program.cs`.

The editor always starts from a template created from scratch (also after **Reset**). Templates are saved to and read from `Server-Side/src/wwwroot/Templates`: **Save Document** creates `Notice of Hearing.docx` there, and you open the saved document again from the Template list. Every other `.docx` in the folder is a supporting document.

## How Each Requirement Is Achieved

| # | Requirement | How it is achieved | Where |
|---|---|---|---|
| 1 | Prompt for missing values, then re-run the merge | The server finds fields with no data and returns them without creating a document. The client shows a dialog, then resends the request with the entered values. | `DocumentMergeService.cs` (`BeforeClearField`), `CasesController.Generate`, `app.ts` (`openMissingDialog`, `submitMissingValues`) |
| 2 | Add new merge fields and regions | The left panel inserts a Word merge field, or a `BeginGroup` / `EndGroup` pair for a region, with any name. | `app.ts` (`insertMergeField`, `insertCustomField`, `insertCustomRegion`) |
| 3 | Targeted data retrieval | Search returns only case summaries. Full details are loaded only for the selected case ids. | `CaseRepository.cs`, `CasesController.cs` |
| 4 | One block per party | The `Parties` region repeats once per party. | `DocumentMergeService.cs` (`BuildPartyGroupTable`) |
| 5 | Case fields inside the party block | `«Case:CaseNumber»` is resolved from the parent case using nested (relational) mail merge. | `DocumentMergeService.cs` (`BuildNestedDataSet`, `ExecuteNestedGroup`) |
| 6 | Several cases in one document | Each selected case is one row of the `Case` table, so the case block repeats per case and the party block per party. | `DocumentMergeService.cs` (`BuildNestedDataSet`) |
| 7 | Type-based prompts (date, number, Yes/No) | Each field has a type that picks the dialog control and validation. Yes/No merges as "Yes" / "No". | `Models/FieldSchema.cs` (`MergeFieldConfig`), `app.ts` (`validateTypedValue`) |
| 8 | Merge fields defined in one place | All field names, types, labels and regions live in one class; the client and the merge code read from it. | `Models/FieldSchema.cs` (`MergeFieldConfig`) |
| 9 | Supporting documents | Selected `.docx` files are appended to the main template, then merged once. The list is read from the folder, so new files appear after a reload. | `DocumentMergeService.MergeAsDocument`, `SupportingDocumentRepository.cs` |

## How It Works, Step by Step

Each flow below shows who does what: **User**, **Client** (the Angular app) and **Server** (ASP.NET Core).

### 1. Start up

1. **User** opens the page.
2. **Client** asks the server for the merge field list (`GET /api/schema`).
3. **Server** returns each field's name, type and label.
4. **Client** asks for the starting template (`GET /api/template`) and the template list (`GET /api/templates`).
5. **Server** builds the template from scratch and returns it, with the name of the file Save writes to. It also returns the saved documents found in the templates folder.
6. **Client** opens the template in the editor and fills the Merge Fields panel and the document list.

### 2. Search and select cases

1. **User** types a keyword and clicks **Search**.
2. **Client** calls `GET /api/cases/search?keyword=...`.
3. **Server** returns matching case summaries only (no parties).
4. **Client** shows the results table.
5. **User** ticks the cases to merge. The client keeps only their ids.

### 3. Generate or preview, with missing values

1. **User** selects cases (and optionally supporting documents), then clicks **Preview** or **Generate Document**.
2. **Client** sends the template as it is in the editor, the case ids and the document ids to `POST /api/generate`.
3. **Server** loads only the selected cases, adds the supporting documents to the template, and runs the mail merge.
4. If every field has data, go to step 9. Otherwise **Server** replies with the list of missing fields and creates no document.
5. **Client** shows a dialog with one input per missing field: a date picker, a number box, a Yes/No checkbox or a text box, chosen from the field's type.
6. **User** enters the values and clicks **Continue**. Invalid values block the button.
7. **Client** sends the same request again, now with the entered values.
8. **Server** merges again using those values.
9. **Server** returns the merged document.
10. **Client** opens it in the editor (read-only for Preview, editable for Generate).

### 4. Export as PDF

1. **User** clicks **Export as PDF**.
2. **Client** sends the same request as Generate to `POST /api/generate-pdf`.
3. If values are missing, the same dialog and re-send as in flow 3 happen.
4. **Server** merges, converts the Word document to PDF and returns the file.
5. **Client** downloads it in the browser.

### 5. Edit and save a template

1. **User** edits the template created from scratch and inserts merge fields or regions from the left panel. This part runs in the browser only.
2. **User** clicks **Save Document**. (Before the first save the button is enabled even without edits.)
3. **Client** sends the document to `POST /api/templates/{id}/save`.
4. **Server** converts it to DOCX and writes `Notice of Hearing.docx` in the templates folder, creating it on the first save and overwriting it afterwards.
5. **Client** shows a "saved" message and refreshes the Template list, so the saved document now appears in it.
6. Later, **User** picks the saved template in the list and clicks **Load**; **Client** calls `GET /api/templates/{id}` and **Server** returns the saved document, which opens in the editor.

Save is disabled when the editor holds merged, previewed or converted content, so a template file cannot be overwritten by mistake.

### 6. Convert a PDF or an XSLT file to a template

1. **User** clicks **Open PDF Document** and chooses a file, or clicks **XSLT to Word Document**.
2. **Client** sends the PDF to `POST /api/pdf/convert`, or asks `POST /api/templates/xsl-to-word` to convert `InputXSL.txt` from the templates folder.
3. **Server** converts the PDF to Word, or turns the stylesheet into a Word template (`for-each` becomes a merge region, `value-of` becomes a merge field).
4. **Server** returns the Word document.
5. **Client** opens it in the editor. It is not tied to a file, so Save stays disabled until it is saved as a template.

### 7. Reset

1. **User** clicks **Reset**.
2. **Client** clears the preview state and calls `GET /api/template`.
3. **Server** builds the template from scratch again (the saved file is not used).
4. **Client** opens it in the editor, discarding unsaved changes.

## APIs Used

| Endpoint | Method | Purpose |
|---|---|---|
| `/api/schema` | GET | Merge field names, types and labels |
| `/api/template` | GET | The template created from scratch, plus the file name Save writes to |
| `/api/templates` | GET | List the saved templates and supporting documents |
| `/api/templates/{id}` | GET | Load a template |
| `/api/templates/{id}/save` | POST | Save the edited template |
| `/api/templates/xsl-to-word` | POST | Convert an XSLT stylesheet to a Word template |
| `/api/cases/search?keyword=` | GET | Search cases |
| `/api/cases/{caseId}` | GET | Full case with parties |
| `/api/generate` | POST | Mail merge; returns missing fields or the merged document |
| `/api/generate-pdf` | POST | Mail merge returned as a PDF |
| `/api/pdf/convert` | POST | Convert a PDF to a Word document |
| `/api/documenteditor/*` | POST | The editor's own service (import, spell check, paste) |

## Resources

- **Product page:** [Syncfusion Angular DOCX Editor](https://www.syncfusion.com/docx-editor-sdk/angular-docx-editor?utm_source=github&utm_medium=listing&utm_campaign=github-github-documenteditor-examples)
- **Documentation:** [Syncfusion Angular DOCX Editor - Documentation](https://help.syncfusion.com/document-processing/word/word-processor/angular/overview?utm_source=github&utm_medium=listing&utm_campaign=github-github-documenteditor-examples)
- **Online demo:** [Syncfusion Angular DOCX Editor - Online demo](https://document.syncfusion.com/demos/docx-editor/angular/#/tailwind3/document-editor/default?utm_source=github&utm_medium=listing&utm_campaign=github-github-documenteditor-examples)

## Support and feedback

For any other queries, reach our [Syncfusion support team](https://support.syncfusion.com/?utm_source=github&utm_medium=listing&utm_campaign=github-github-documenteditor-examples) or post the queries through the [community forums](https://www.syncfusion.com/forums?utm_source=github&utm_medium=listing&utm_campaign=github-github-documenteditor-examples). 

Request new feature through [Syncfusion feedback portal](https://www.syncfusion.com/feedback?utm_source=github&utm_medium=listing&utm_campaign=github-github-documenteditor-examples). 

## License

This is a commercial product and requires a paid license for possession or use. Syncfusion's licensed software, including this component, is subject to the terms and conditions of [Syncfusion's EULA](https://www.syncfusion.com/license/studio/35.1.37/syncfusion_essential_studio_eula.pdf?utm_source=github&utm_medium=listing&utm_campaign=github-github-documenteditor-examples). You can purchase a license [here](https://www.syncfusion.com/sales/products?utm_source=github&utm_medium=listing&utm_campaign=github-github-documenteditor-examples) or start a free 30-day trial [here](https://www.syncfusion.com/account/manage-trials/start-trials?utm_source=github&utm_medium=listing&utm_campaign=github-github-documenteditor-examples).