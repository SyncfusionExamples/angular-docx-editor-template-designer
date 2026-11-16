using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using EJ2APIServices.Data;
using EJ2APIServices.Models;
using EJ2APIServices.Services;
using Microsoft.AspNetCore.Cors;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using Syncfusion.Pdf;

namespace EJ2APIServices.Controllers
{
    /// <summary>
    /// Document generation API for the Notice of Hearing POC.
    ///
    /// GET  /api/schema          - merge field schema for the Merge Field Explorer
    /// GET  /api/template        - bundled sample template as SFDT (opened by the editor on start)
    /// GET  /api/cases/search    - targeted case search, no full load
    /// GET  /api/cases/{caseId}  - full case details for a selected record
    /// POST /api/generate        - mail merge with missing-field prompting; returns SFDT
    ///                              for the editor, never a download
    /// </summary>
    [Route("api/[controller]")]
    public class CasesController : Controller
    {
        private readonly IWebHostEnvironment _env;
        private readonly IConfiguration _config;

        public CasesController(IWebHostEnvironment env, IConfiguration config)
        {
            _env = env;
            _config = config;
        }

        /// <summary>
        /// Merge field schema for the explorer, the Insert Merge Field dialog, and the
        /// missing-field prompt. Each entry carries a type tag that drives the UI input
        /// control: "string" (default text box), "date" (date picker), "number" (numeric
        /// text box), or "boolean" (checkbox). All values are submitted back to the
        /// server as strings regardless of type.
        ///
        /// Field names are source-aware: fields reachable inside the Parties region
        /// carry the case source with a "Case:" prefix («Case:CaseNumber»), party
        /// fields stay plain («FirstName») because they resolve per party row.
        /// </summary>
        [HttpGet]
        [Route("/api/schema")]
        public IActionResult Schema()
        {
            // All merge field definitions live in MergeFieldConfig (Models/FieldSchema.cs).
            return Ok(MergeFieldConfig.Schema);
        }

        /// <summary>
        /// The template the editor opens on application start and after Reset. It is
        /// ALWAYS created from scratch (DocumentMergeService.BuildSampleTemplate) - the
        /// "Notice of Hearing" DOCX file in the templates folder is never loaded here.
        ///
        /// That file exists only as the target of "Save Document": the user saves the
        /// template to it, then opens it again from the Template dropdown
        /// (GET /api/templates/{id}).
        ///
        /// The response also tells the client which file name Save writes to and
        /// whether that file exists yet.
        /// </summary>
        [HttpGet]
        [Route("/api/template")]
        public IActionResult Template()
        {
            var folder = SupportingDocumentRepository.GetFolder(
                _env.ContentRootPath,
                _config["TemplatesFolder"]);
            var mainPath = SupportingDocumentRepository.GetMainTemplatePath(
                folder,
                _config["MainTemplateFileName"]);

            var fileName = SupportingDocumentRepository.GetMainTemplateFileName(
                _config["MainTemplateFileName"]);

            return Ok(new
            {
                sfdt = DocumentMergeService.BuildSampleTemplateSfdt(),
                fileName,
                exists = mainPath != null
            });
        }

        /// <summary>
        /// Keyword search over case number and caption. Returns matching case
        /// summaries only; parties are not loaded until a case is selected.
        /// </summary>
        [HttpGet]
        [Route("/api/cases/search")]
        public IActionResult Search([FromQuery] string keyword)
        {
            var results = CaseRepository.Search(keyword);
            return Ok(results);
        }

        /// <summary>
        /// Full case details for one selected case.
        /// </summary>
        [HttpGet]
        [Route("/api/cases/{caseId:int}")]
        public IActionResult GetCase(int caseId)
        {
            var caseData = CaseRepository.GetCase(caseId);
            if (caseData == null)
            {
                return NotFound(new { message = $"Case {caseId} was not found." });
            }
            return Ok(caseData);
        }

        /// <summary>
        /// Mail merge over the selected cases.
        ///
        /// First request (fields missing): JSON { requiresInput: true, missingFields: [...] },
        /// no document. The UI prompts for the values.
        ///
        /// Second request (with additionalMergeData): the merge runs again with the API
        /// data plus the entered values. The response is the ONE merged document
        /// serialized as SFDT (mergedSfdt), which the editor opens directly.
        ///
        /// All selected cases are merged into one document: one notice block per case,
        /// one party block per party (ExecuteNestedGroup).
        ///
        /// If appendedDocumentIds are supplied, each is loaded from the configured
        /// templates folder and appended to the merged document in the order received.
        /// </summary>
        [HttpPost]
        [Route("/api/generate")]
        public IActionResult Generate([FromBody] GenerateRequest request)
        {
            if (request?.CaseIds == null || request.CaseIds.Count == 0)
            {
                return BadRequest(new { message = "Select at least one case." });
            }

            // Load only the selected cases (targeted data loading).
            var cases = new List<CaseData>();
            foreach (var id in request.CaseIds)
            {
                var caseData = CaseRepository.GetCase(id);
                if (caseData == null)
                {
                    return NotFound(new { message = $"Case {id} was not found." });
                }
                cases.Add(caseData);
            }

            WordDocument template;
            try
            {
                template = string.IsNullOrEmpty(request.TemplateSfdt)
                    ? DocumentMergeService.BuildSampleTemplate()
                    : DocumentMergeService.LoadTemplateFromSfdt(request.TemplateSfdt);
            }
            catch (Exception)
            {
                // Fall back to the bundled sample template when the editor payload is unusable.
                template = DocumentMergeService.BuildSampleTemplate();
            }

            // Resolve supporting document ids -> file paths. Skip unknown ids so a stale
            // selection from a previous session does not break the request.
            var supportingPaths = new List<string>();
            if (request.AppendedDocumentIds != null && request.AppendedDocumentIds.Count > 0)
            {
                var folder = SupportingDocumentRepository.GetFolder(
                    _env.ContentRootPath,
                    _config["TemplatesFolder"]);

                foreach (var id in request.AppendedDocumentIds)
                {
                    // The main template is the document being merged - appending it to
                    // itself makes no sense, so it is never accepted as a supporting doc.
                    if (string.Equals(id, SupportingDocumentRepository.DefaultMainTemplateFileName, StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    var path = SupportingDocumentRepository.GetSupportingFilePath(folder, id);
                    if (path != null)
                    {
                        supportingPaths.Add(path);
                    }
                }
            }

            try
            {
                var additional = request.AdditionalMergeData ?? new List<AdditionalMergeValue>();
                var outcome = DocumentMergeService.Merge(template, cases, additional, supportingPaths);

                if (outcome.MissingFields.Count > 0)
                {
                    // Do not generate anything yet: prompt the user first.
                    return Ok(new GenerateRequiresInputResponse
                    {
                        RequiresInput = true,
                        MissingFields = outcome.MissingFields
                    });
                }

                // One merged document for every selected case + appended supporting documents,
                // opened in the editor.
                return Ok(new GenerateMergedResponse
                {
                    MergedSfdt = outcome.MergedSfdt,
                    DocumentName = outcome.DocumentName
                });
            }
            finally
            {
                template.Close();
            }
        }

        /// <summary>
        /// Mail merge + Word -> PDF in one round trip. Runs the same merge pipeline as
        /// /api/generate (same templateSfdt, caseIds, supporting docs, missing-field
        /// handling) and converts the merged WordDocument to PDF using the canonical
        /// DocIORenderer pattern from
        /// https://help.syncfusion.com/document-processing/word/conversions/word-to-pdf/net/word-to-pdf
        /// so the user can download the finalised document as a PDF.
        ///
        /// Returns the PDF as application/pdf so the browser can save it directly.
        /// If the merge produced missing fields, the response is the same JSON shape
        /// as /api/generate (requiresInput + missingFields) with HTTP 200, so the
        /// client can prompt the user and retry.
        /// </summary>
        [HttpPost]
        [Route("/api/generate-pdf")]
        public IActionResult GeneratePdf([FromBody] GenerateRequest request)
        {
            if (request?.CaseIds == null || request.CaseIds.Count == 0)
            {
                return BadRequest(new { message = "Select at least one case." });
            }

            var cases = new List<CaseData>();
            foreach (var id in request.CaseIds)
            {
                var caseData = CaseRepository.GetCase(id);
                if (caseData == null)
                {
                    return NotFound(new { message = $"Case {id} was not found." });
                }
                cases.Add(caseData);
            }

            WordDocument template;
            try
            {
                template = string.IsNullOrEmpty(request.TemplateSfdt)
                    ? DocumentMergeService.BuildSampleTemplate()
                    : DocumentMergeService.LoadTemplateFromSfdt(request.TemplateSfdt);
            }
            catch (Exception)
            {
                template = DocumentMergeService.BuildSampleTemplate();
            }

            var supportingPaths = new List<string>();
            if (request.AppendedDocumentIds != null && request.AppendedDocumentIds.Count > 0)
            {
                var folder = SupportingDocumentRepository.GetFolder(
                    _env.ContentRootPath,
                    _config["TemplatesFolder"]);

                foreach (var id in request.AppendedDocumentIds)
                {
                    if (string.Equals(id, SupportingDocumentRepository.DefaultMainTemplateFileName, StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }
                    var path = SupportingDocumentRepository.GetSupportingFilePath(folder, id);
                    if (path != null)
                    {
                        supportingPaths.Add(path);
                    }
                }
            }

            try
            {
                var additional = request.AdditionalMergeData ?? new List<AdditionalMergeValue>();
                var outcome = new MergeOutcome();
                var merged = DocumentMergeService.MergeAsDocument(template, cases, additional, supportingPaths, outcome);

                if (merged == null)
                {
                    // Missing fields - return the same JSON shape as /api/generate so the
                    // client can reuse its missing-field dialog flow.
                    return Ok(new GenerateRequiresInputResponse
                    {
                        RequiresInput = true,
                        MissingFields = outcome.MissingFields
                    });
                }

                try
                {
                    // Canonical Word -> PDF pattern from the Syncfusion docs:
                    //   using (DocIORenderer renderer = new DocIORenderer())
                    //   using (PdfDocument pdfDocument = renderer.ConvertToPDF(wordDocument))
                    using (var renderer = new DocIORenderer())
                    using (var pdfDocument = renderer.ConvertToPDF(merged))
                    {
                        using (var outputStream = new MemoryStream())
                        {
                            pdfDocument.Save(outputStream);
                            return File(outputStream.ToArray(), "application/pdf", "Merged-Notice-of-Hearing.pdf");
                        }
                    }
                }
                finally
                {
                    merged.Close();
                }
            }
            finally
            {
                template.Close();
            }
        }
    }
}