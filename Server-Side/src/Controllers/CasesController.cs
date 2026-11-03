using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using EJ2APIServices.Data;
using EJ2APIServices.Models;
using EJ2APIServices.Services;
using Microsoft.AspNetCore.Cors;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Syncfusion.DocIO.DLS;

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
        /// <summary>
        /// Merge field schema for the explorer and the Insert Merge Field dialog.
        ///
        /// Field names are source-aware: fields reachable inside the Parties region
        /// carry the case source with a "Case:" prefix («Case:CaseNumber»), party
        /// fields stay plain («FirstName») because they resolve per party row.
        /// </summary>
        [HttpGet]
        [Route("/api/schema")]
        public IActionResult Schema()
        {
            return Ok(new Dictionary<string, object>
            {
                ["Case"] = new Dictionary<string, string>
                {
                    ["Case:CaseNumber"] = "string",
                    ["Case:CaseCaption"] = "string"
                },
                ["Parties"] = new Dictionary<string, string>
                {
                    ["FirstName"] = "string",
                    ["LastName"] = "string",
                    ["Address"] = "string"
                }
            });
        }

        /// <summary>
        /// The bundled sample Notice of Hearing template serialized as SFDT.
        /// The Angular editor opens this on application start so the app never starts
        /// blank and the sample already demonstrates case fields, the repeating party
        /// region and parent fields inside the region.
        /// </summary>
        [HttpGet]
        [Route("/api/template")]
        public IActionResult Template()
        {
            string sfdt = DocumentMergeService.BuildSampleTemplateSfdt();
            return Ok(new { sfdt });
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

            try
            {
                var additional = request.AdditionalMergeData ?? new List<AdditionalMergeValue>();
                var outcome = DocumentMergeService.Merge(template, cases, additional);

                if (outcome.MissingFields.Count > 0)
                {
                    // Do not generate anything yet: prompt the user first.
                    return Ok(new GenerateRequiresInputResponse
                    {
                        RequiresInput = true,
                        MissingFields = outcome.MissingFields
                    });
                }

                // One merged document for every selected case, opened in the editor.
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
    }
}