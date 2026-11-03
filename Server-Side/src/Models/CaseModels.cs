using System.Collections.Generic;
using System.Text.Json.Serialization;

namespace EJ2APIServices.Models
{
    /// <summary>
    /// Party record inside a case.
    /// </summary>
    public class PartyData
    {
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string Address { get; set; }
    }

    /// <summary>
    /// Case summary returned by GET /api/cases/search. No parties are included,
    /// so search loads targeted data only (Requirement 3).
    /// </summary>
    public class CaseInfo
    {
        public int CaseId { get; set; }
        public string CaseNumber { get; set; }
        public string CaseCaption { get; set; }
    }

    /// <summary>
    /// Full case record returned by GET /api/cases/{caseId}.
    /// </summary>
    public class CaseData
    {
        public int CaseId { get; set; }
        public string CaseNumber { get; set; }
        public string CaseCaption { get; set; }
        public List<PartyData> Parties { get; set; }
    }

    /// <summary>
    /// One user-supplied value for a merge field the case data did not fill.
    /// Keys: CaseId (0 = apply to every case), GroupName, RowIndex (1-based inside
    /// the group; 0 for case-level fields), FieldName.
    /// </summary>
    public class AdditionalMergeValue
    {
        public int CaseId { get; set; }
        public string CaseNumber { get; set; }
        public string GroupName { get; set; }
        public int RowIndex { get; set; }
        public string FieldName { get; set; }
        public string Value { get; set; }
    }

    /// <summary>
    /// One missing merge field reported back to the UI. Carries the case identity so
    /// the dialog can prompt per case (Requirement 1, multi-case handling).
    /// </summary>
    public class MissingFieldInfo
    {
        [JsonPropertyName("caseId")]
        public int CaseId { get; set; }

        [JsonPropertyName("caseNumber")]
        public string CaseNumber { get; set; }

        [JsonPropertyName("groupName")]
        public string GroupName { get; set; }

        [JsonPropertyName("rowIndex")]
        public int RowIndex { get; set; }

        [JsonPropertyName("fieldName")]
        public string FieldName { get; set; }
    }

    /// <summary>
    /// Request body of POST /api/generate.
    /// </summary>
    public class GenerateRequest
    {
        public List<int> CaseIds { get; set; }
        public List<AdditionalMergeValue> AdditionalMergeData { get; set; }
        /// <summary>Serialized template (SFDT) from the Angular Document Editor. Falls back to the bundled sample template.</summary>
        public string TemplateSfdt { get; set; }
    }

    /// <summary>
    /// Response of POST /api/generate when input is still missing. No document is produced;
    /// the UI prompts and re-sends the request with additionalMergeData.
    /// </summary>
    public class GenerateRequiresInputResponse
    {
        [JsonPropertyName("requiresInput")]
        public bool RequiresInput { get; set; }

        [JsonPropertyName("missingFields")]
        public List<MissingFieldInfo> MissingFields { get; set; }
    }

    /// <summary>
    /// Response of POST /api/generate when the merged document was produced:
    /// the merged DOCX converted back to SFDT so the editor can open it directly.
    /// </summary>
    public class GenerateMergedResponse
    {
        [JsonPropertyName("mergedSfdt")]
        public string MergedSfdt { get; set; }

        [JsonPropertyName("documentName")]
        public string DocumentName { get; set; }
    }
}