# Syncfusion® Angular DOCX Editor – Document Generation

## Introduction

This sample extends the [Syncfusion® Angular DOCX Editor](https://www.syncfusion.com/docx-editor-sdk/angular-docx-editor?utm_source=github&utm_medium=listing&utm_campaign=github-github-documenteditor-examples) and an ASP.NET Core backend into a legal document generation workflow. A user authors a "Notice of Hearing" template in the Angular Document Editor, searches for case records, selects one or more cases, and generates a merged document with Syncfusion DocIO on the server. The merged result opens directly inside the editor. It is intended for developers building document generation for business records such as court cases and their parties.

## Project Features

- **Merge field explorer** — click a case or party field in the left panel to insert a real merge field into the template, or create custom fields and regions with a user-specified name.
- **Case search and selection** — search cases by case number or caption, select the records to merge, and load only the selected business data.
- **Standard and nested repeating regions** — a simple `Parties` region repeats once per party; a `Case` region wrapping it repeats once per selected case and resolves case fields inside each party block.
- **Missing value prompting** — when a template field has no data (for example `JudgeName` or `NoticeDate`), generation pauses, the user is prompted for the value, and the merge runs again with the entered data.
- **In-editor result preview** — the merged document opens directly in the Document Editor instead of downloading a file.

## How to Run the Sample

This sample has two projects that must both run: an Angular client and an ASP.NET Core backend. The merge, case search, and template APIs are served by the backend.

### Prerequisites

- Node.js 20.19+ or 22.12+ with npm (for the Angular client)
- .NET SDK 10.0 (for the ASP.NET Core backend)

### 1. Start the backend

The backend serves the case APIs (`/api/schema`, `/api/cases/...`, `/api/generate`) and the Document Editor processing endpoints (`/api/documenteditor/...`).

```bash
cd Server-Side/src
dotnet restore
dotnet run
```

The `EJ2APIServices` profile in `Properties/launchSettings.json` starts the app at `http://localhost:62870/`. If you host it elsewhere, note the base URL for the next step.

### 2. Configure the client

The Angular dev-server proxy forwards `/api` calls to the backend. The setting lives in `Client-Side/proxy.conf.json`:

```json
{
  "/api": {
    "target": "http://localhost:62870",
    "secure": false,
    "changeOrigin": false
  }
}
```

Change `target` only if your backend uses a different base URL. With the default `dotnet run` command above, no change is required.

### 3. Start the client

```bash
cd Client-Side
npm install
npm start
```

`npm start` runs `ng serve` in watch mode. Open the URL printed in the terminal (by default `http://localhost:4200/`).

### Using the sample

The template opens automatically. Search by case number or caption (for example `Smith`), select one or more cases, and click **Generate Documents**. If the template contains fields without data (the bundled template uses `JudgeName` and `NoticeDate` for this), enter the values in the prompt and continue; the merged document opens in the editor.

Editing the template, inserting fields from the left panel, and custom field/region creation are client-side editor features and work as soon as the page loads; search and generation require the backend.

## APIs Used

Custom sample APIs (served by the backend's `CasesController`):

| API / Endpoint | HTTP Method | Purpose |
|---|---|---|
| `/api/schema` | GET | Merge field schema for the explorer panel (field names and types) |
| `/api/template` | GET | Bundled sample template, opened by the editor on application load |
| `/api/cases/search?keyword=` | GET | Case search by case number or caption; returns matching case summaries |
| `/api/cases/{caseId}` | GET | Full case record (case details plus parties) for a selected case |
| `/api/generate` | POST | Mail merge for the selected cases; returns either missing-field prompts or the merged document as SFDT |

Document Editor processing service (`DocumentEditorController`, used through the editor's `serviceUrl` of `/api/documenteditor/` in `Client-Side/src/app/app.ts`):

| API / Endpoint | HTTP Method | Purpose |
|---|---|---|
| `/api/documenteditor/import` | POST | Converts an uploaded document to the editor's internal format |
| `/api/documenteditor/spellcheck` | POST | Spell check support for the editor |
| `/api/documenteditor/paste` | POST | Paste formatting (HTML to document conversion) |

## Resources

- **Product page:** [Syncfusion Angular DOCX Editor](https://www.syncfusion.com/docx-editor-sdk/angular-docx-editor?utm_source=github&utm_medium=listing&utm_campaign=github-github-documenteditor-examples)
- **Documentation:** [Syncfusion Angular DOCX Editor - Documentation](https://help.syncfusion.com/document-processing/word/word-processor/angular/overview?utm_source=github&utm_medium=listing&utm_campaign=github-github-documenteditor-examples)
- **Online demo:** [Syncfusion Angular DOCX Editor - Online demo](https://document.syncfusion.com/demos/docx-editor/angular/#/tailwind3/document-editor/default?utm_source=github&utm_medium=listing&utm_campaign=github-github-documenteditor-examples)

## Support and feedback

For any other queries, reach our [Syncfusion support team](https://support.syncfusion.com/?utm_source=github&utm_medium=listing&utm_campaign=github-github-documenteditor-examples) or post the queries through the [community forums](https://www.syncfusion.com/forums?utm_source=github&utm_medium=listing&utm_campaign=github-github-documenteditor-examples). 

Request new feature through [Syncfusion feedback portal](https://www.syncfusion.com/feedback?utm_source=github&utm_medium=listing&utm_campaign=github-github-documenteditor-examples). 

## License

This is a commercial product and requires a paid license for possession or use. Syncfusion's licensed software, including this component, is subject to the terms and conditions of [Syncfusion's EULA](https://www.syncfusion.com/license/studio/35.1.37/syncfusion_essential_studio_eula.pdf?utm_source=github&utm_medium=listing&utm_campaign=github-github-documenteditor-examples). You can purchase a license [here](https://www.syncfusion.com/sales/products?utm_source=github&utm_medium=listing&utm_campaign=github-github-documenteditor-examples) or start a free 30-day trial [here](https://www.syncfusion.com/account/manage-trials/start-trials?utm_source=github&utm_medium=listing&utm_campaign=github-github-documenteditor-examples).