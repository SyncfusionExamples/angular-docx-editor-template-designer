using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text.Json;
using EJ2APIServices.Models;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using EJ2WordDocument = Syncfusion.EJ2.DocumentEditor.WordDocument;

namespace EJ2APIServices.Services
{
    /// <summary>
    /// Result of a merge attempt: either the missing fields to prompt for,
    /// or the merged document as SFDT (for the editor) plus DOCX bytes.
    /// </summary>
    public class MergeOutcome
    {
        public List<MissingFieldInfo> MissingFields { get; set; } = new List<MissingFieldInfo>();
        /// <summary>Merged document serialized as SFDT, ready for documentEditor.open().</summary>
        public string MergedSfdt { get; set; }
        public string DocumentName { get; set; }
    }

    /// <summary>
    /// DocIO mail merge engine for the Notice of Hearing POC.
    ///
    /// === Requirement 4: repeating Party regions ===
    /// The template contains a mail merge region delimited by the merge fields
    /// «BeginGroup:Parties» ... «EndGroup:Parties». DocIO discovers the region from
    /// the template's MERGEFIELD delimiters (MailMerge.GetMergeGroupNames() lists it),
    /// and repeats the region body once per matching data row during the merge.
    ///
    /// === Requirement 5: parent fields inside child regions ===
    /// Implemented with DocIO relational (nested) mail merge — NOT flattening.
    /// The data set has a parent table "Case" and a child table "Parties" related by
    /// CaseId. The template wraps the party region in an outer region «BeginGroup:Case».
    /// DocIO resolves «Case:CaseNumber» inside «BeginParties» from the current parent
    /// row while iterating the child rows (verified: parent values appear in every
    /// repeated party block). See BuildNestedDataSet() and Merge() for the exact code path.
    ///
    /// === Multiple selected cases ===
    /// All selected cases are merged into ONE document: each selected case becomes one
    /// row of the parent "Case" table, so the outer region repeats once per case and the
    /// inner region repeats per party of that case. No per-case files and no ZIP.
    /// </summary>
    public static class DocumentMergeService
    {
        /// <summary>Parent region/table name. «Case:CaseNumber» resolves against this table.</summary>
        public const string CaseRegionName = "Case";

        /// <summary>Child region/table name (repeating party region).</summary>
        public const string PartiesRegionName = "Parties";

        /// <summary>Relation key between Case and Parties.</summary>
        private const string RelationKey = "CaseId";

        /// <summary>
        /// Converts the SFDT template serialized by the Angular Document Editor into a
        /// DocIO WordDocument with materialized merge fields.
        ///
        /// EJ2WordDocument.Save(sfdt) returns a DocIO document whose field structures are
        /// not yet exposed to MailMerge (GetMergeFieldNames() is empty). Saving that
        /// instance to DOCX and reloading it materializes the fields, so the merge works
        /// on editor-authored templates. (Verified in this POC's integration runs.)
        /// </summary>
        public static WordDocument LoadTemplateFromSfdt(string sfdt)
        {
            WordDocument raw = EJ2WordDocument.Save(sfdt);
            using (var ms = new MemoryStream())
            {
                raw.Save(ms, FormatType.Docx);
                raw.Close();
                ms.Position = 0;
                return new WordDocument(ms, FormatType.Docx);
            }
        }

        /// <summary>
        /// Converts a merged DocIO document into SFDT so the Angular Document Editor can
        /// open the result directly (open(sfdt)) instead of downloading a file.
        /// </summary>
        public static string ToSfdt(WordDocument merged)
        {
            var sfdtDoc = EJ2WordDocument.Load(merged);
            string json = Newtonsoft.Json.JsonConvert.SerializeObject(sfdtDoc);
            sfdtDoc.Dispose();
            return json;
        }

        /// <summary>
        /// Builds the bundled sample Notice of Hearing template. Used on application start
        /// so the editor never opens blank, and as the fallback when no SFDT is sent.
        ///
        /// The document is laid out like a real court notice so the generation scenario is
        /// obvious at a glance:
        ///
        ///   NOTICE OF HEARING (title)
        ///   Case Number: «CaseNumber» / Case Caption: «CaseCaption»   (case fields)
        ///   Notice Date: «NoticeDate»                                  (NOT in API data -> prompts)
        ///   TO ALL PARTIES OF RECORD ...
        ///   «BeginGroup:Parties»
        ///     Dear «FirstName» «LastName»,
        ///     Case Number: «Case:CaseNumber»  Case Caption: «Case:CaseCaption»
        ///     Mailing Address: «Address»
        ///     Presiding Judge: «JudgeName»    (NOT in API data -> prompts)
        ///     ... court boilerplate ...
        ///   «EndGroup:Parties»
        ///   Respectfully Submitted, Court Administration
        ///
        /// JudgeName and NoticeDate are deliberately absent from the sample API data so a
        /// normal demo run goes: Generate -> missing fields detected -> prompt -> values
        /// entered -> merge re-executed -> merged document opens in the editor.
        /// </summary>
        public static WordDocument BuildSampleTemplate()
        {
            var doc = new WordDocument();
            var section = doc.AddSection();

            // ---- Letterhead-style title ----
            var title = section.AddParagraph();
            title.AppendText("NOTICE OF HEARING");
            title.ApplyStyle(BuiltinStyle.Heading1);
            title.ParagraphFormat.HorizontalAlignment = Syncfusion.DocIO.DLS.HorizontalAlignment.Center;

            // ---- Body (outside the repeating regions: shown once) ----
            section.AddParagraph().AppendText("TO ALL PARTIES OF RECORD:");
            section.AddParagraph().AppendText("PLEASE TAKE NOTICE that a hearing has been scheduled in the above-captioned matter.");

            // Outer region: one block per selected case (nested merge parent).
            // This region is what makes multiple selected cases merge into ONE document
            // and what lets «Case:CaseNumber» resolve inside the party region.
            // The case identification block lives inside it (per case), as before.
            section.AddParagraph().AppendField($"BeginGroup:{CaseRegionName}", FieldType.FieldMergeField);

            // ---- Case identification block (per selected case) ----
            var caption = section.AddParagraph();
            caption.AppendText("Case Number: ");
            caption.AppendField("CaseNumber", FieldType.FieldMergeField);
            var captionPara = section.AddParagraph();
            captionPara.AppendText("Case Caption: ");
            captionPara.AppendField("CaseCaption", FieldType.FieldMergeField);
            var noticeDate = section.AddParagraph();
            noticeDate.AppendText("Notice Date: ");
            noticeDate.AppendField("NoticeDate", FieldType.FieldMergeField);
            section.AddParagraph().AppendText("---------------------------------------------------------------------");

            // ---- Repeating party region (Requirement 4 + 5) ----
            section.AddParagraph().AppendField($"BeginGroup:{PartiesRegionName}", FieldType.FieldMergeField);

            var dear = section.AddParagraph();
            dear.AppendText("Dear ");
            dear.AppendField("FirstName", FieldType.FieldMergeField);
            dear.AppendText(" ");
            dear.AppendField("LastName", FieldType.FieldMergeField);
            dear.AppendText(",");

            section.AddParagraph().AppendText("You are hereby notified that a hearing has been scheduled in the following case:");
            section.AddParagraph().AppendText("Case Number: ");
            section.AddParagraph().AppendField("Case:CaseNumber", FieldType.FieldMergeField);
            section.AddParagraph().AppendText("Case Caption: ");
            section.AddParagraph().AppendField("Case:CaseCaption", FieldType.FieldMergeField);
            section.AddParagraph().AppendText("Mailing Address: ");
            section.AddParagraph().AppendField("Address", FieldType.FieldMergeField);
            section.AddParagraph().AppendText("Presiding Judge: ");
            section.AddParagraph().AppendField("JudgeName", FieldType.FieldMergeField);
            section.AddParagraph().AppendText("Please appear before the Court at the designated hearing date and time.");
            section.AddParagraph().AppendText("---------------------------------------------------------------------");

            section.AddParagraph().AppendField($"EndGroup:{PartiesRegionName}", FieldType.FieldMergeField);
            section.AddParagraph().AppendField($"EndGroup:{CaseRegionName}", FieldType.FieldMergeField);

            // ---- Signature block ----
            section.AddParagraph().AppendText("Respectfully Submitted,");
            section.AddParagraph().AppendText("Court Administration");
            return doc;
        }

        /// <summary>
        /// Serializes the sample template as SFDT for GET /api/template (the editor opens
        /// this on application load).
        /// </summary>
        public static string BuildSampleTemplateSfdt()
        {
            using (var doc = BuildSampleTemplate())
            {
                return ToSfdt(doc);
            }
        }

        /// <summary>
        /// The relational data set consumed by ExecuteNestedGroup. This is the object
        /// model the review asked about, expressed as ADO.NET tables DocIO can relate:
        ///
        ///   Case                          Parties
        ///   ------                         ---------
        ///   CaseId   (relation key)  ->   CaseId
        ///   CaseNumber                     FirstName
        ///   CaseCaption                    LastName
        ///                                  Address
        ///
        /// One parent row per selected case; child rows relate by CaseId.
        /// </summary>
        public static DataSet BuildNestedDataSet(List<CaseData> cases)
        {
            var ds = new DataSet();

            var caseTable = new DataTable(CaseRegionName);
            caseTable.Columns.Add("CaseId");
            caseTable.Columns.Add("CaseNumber");
            caseTable.Columns.Add("CaseCaption");
            ds.Tables.Add(caseTable);

            var partyTable = new DataTable(PartiesRegionName);
            partyTable.Columns.Add(RelationKey);
            partyTable.Columns.Add("FirstName");
            partyTable.Columns.Add("LastName");
            partyTable.Columns.Add("Address");
            ds.Tables.Add(partyTable);

            foreach (var c in cases)
            {
                caseTable.Rows.Add(c.CaseId.ToString(), c.CaseNumber ?? "", c.CaseCaption ?? "");
                foreach (var p in c.Parties ?? Enumerable.Empty<PartyData>())
                {
                    partyTable.Rows.Add(c.CaseId.ToString(), p.FirstName ?? "", p.LastName ?? "", p.Address);
                }
            }
            return ds;
        }

        /// <summary>
        /// The merge commands for ExecuteNestedGroup: merge the parent "Case" table,
        /// then the child "Parties" table related by CaseId.
        /// </summary>
        private static ArrayList BuildNestedCommands()
        {
            return new ArrayList
            {
                new DictionaryEntry(CaseRegionName, ""),
                new DictionaryEntry(PartiesRegionName, $"{RelationKey} = %{CaseRegionName}.{RelationKey}%")
            };
        }

        /// <summary>
        /// The flat party table consumed by ExecuteGroup in the standard path
        /// (Scenario A: only a «BeginGroup:Parties» region, no parent Case region).
        /// One row per party across all selected cases; the region repeats per row.
        /// </summary>
        public static DataTable BuildPartyGroupTable(List<CaseData> cases)
        {
            var dt = new DataTable(PartiesRegionName); // table name == region name
            dt.Columns.Add("FirstName");
            dt.Columns.Add("LastName");
            dt.Columns.Add("Address");
            foreach (var c in cases)
            {
                foreach (var p in c.Parties ?? Enumerable.Empty<PartyData>())
                {
                    dt.Rows.Add(p.FirstName ?? "", p.LastName ?? "", p.Address ?? "");
                }
            }
            return dt;
        }

        /// <summary>
        /// Runs the mail merge over the selected cases (one merged document) and returns
        /// either the missing fields to prompt for, or the merged document as SFDT.
        ///
        /// Region detection decides the merge strategy before anything is executed:
        ///
        ///   Scenario A (standard regions): the template has «BeginGroup:Parties» but no
        ///   parent «BeginGroup:Case» region. Detected groups: [Parties]. Strategy:
        ///   MailMerge.ExecuteGroup(partyTable) repeats the region per party row
        ///   (Pattern 1), and MailMerge.Execute(...) fills the case fields outside
        ///   the region. No parent Case region is required.
        ///
        ///   Scenario B (nested regions): the template wraps «BeginGroup:Parties» inside
        ///   «BeginGroup:Case». Detected groups: [Case, Parties]. Strategy: the existing
        ///   relational MailMerge.ExecuteNestedGroup(dataSet, commands) with parent-child
        ///   resolution (Pattern 2 / Requirement 5) - unchanged.
        ///
        /// Exact code path:
        ///   1. Region discovery:  doc.MailMerge.GetMergeGroupNames() / GetMergeFieldNames()
        ///      (logged so manually inserted regions can be verified).
        ///   2. Data mapping:     BuildPartyGroupTable() for the standard path;
        ///                        BuildNestedDataSet()/BuildNestedCommands() for the nested path.
        ///   3. Mail merge:       ExecuteGroup(...) or ExecuteNestedGroup(...) per scenario.
        ///   4. Missing detection: BeforeClearField fires for every field DocIO could not
        ///      fill. MergeField events track which case row is current (nested path) so
        ///      each missing field is tagged with its CaseId/CaseNumber. AdditionalMergeData
        ///      values are injected through the same event in both paths.
        /// </summary>
        public static MergeOutcome Merge(
            WordDocument template,
            List<CaseData> cases,
            List<AdditionalMergeValue> additionalData)
        {
            var doc = template.Clone();
            var outcome = new MergeOutcome();

            // ---- Region detection (decides the merge strategy) ----
            var detectedGroups = doc.MailMerge.GetMergeGroupNames() ?? new string[0];
            var detectedFields = doc.MailMerge.GetMergeFieldNames() ?? new string[0];
            bool hasCaseRegion = detectedGroups.Any(g =>
                string.Equals(g, CaseRegionName, StringComparison.OrdinalIgnoreCase));
            bool hasPartiesRegion = detectedGroups.Any(g =>
                string.Equals(g, PartiesRegionName, StringComparison.OrdinalIgnoreCase));
            Console.WriteLine($"[merge] detected groups: [{string.Join(", ", detectedGroups)}]");
            Console.WriteLine($"[merge] detected fields: [{string.Join(", ", detectedFields)}]");
            Console.WriteLine($"[merge] strategy: {(hasCaseRegion && hasPartiesRegion ? "nested (ExecuteNestedGroup)" : "standard (ExecuteGroup)")}");

            // Parent index (1-based) of the case row currently being merged. Tracked from
            // MergeField events of the parent group so missing child fields can be tagged
            // with their case.
            int currentCaseId = 0;
            string currentCaseNumber = "";
            int currentCaseRowIndex = 0;
            var caseIds = cases.Select(c => c.CaseId).ToList();
            var caseNumbers = cases.Select(c => c.CaseNumber).ToList();

            doc.MailMerge.ClearFields = false; // keep unmerged fields until we know what is missing

            doc.MailMerge.MergeField += (sender, args) =>
            {
                if (string.Equals(args.GroupName, CaseRegionName, StringComparison.OrdinalIgnoreCase))
                {
                    int parentRow = args.RowIndex < 0 ? 0 : args.RowIndex;
                    currentCaseRowIndex = parentRow;
                    if (parentRow >= 0 && parentRow < caseIds.Count)
                    {
                        currentCaseId = caseIds[parentRow];
                        currentCaseNumber = caseNumbers[parentRow];
                    }
                }
            };

            doc.MailMerge.BeforeClearField += (sender, args) =>
            {
                // DocIO reports a 0-based row inside the group; the UI contract is 1-based.
                // Case-level fields (GroupName "Case") get RowIndex 0.
                string groupName = args.GroupName ?? "";
                int rowIndex = args.RowIndex < 0 ? 0 : args.RowIndex + 1;
                if (string.Equals(groupName, CaseRegionName, StringComparison.OrdinalIgnoreCase))
                {
                    rowIndex = 0; // case-level field: no row differentiation
                }

                // Which case is this field in? Child fields inherit the tracked case;
                // case-level fields are the current case row.
                int caseIdForField = currentCaseId;
                string caseNumberForField = currentCaseNumber;

                var supplied = FindSuppliedValue(additionalData, caseIdForField, groupName, rowIndex, args.FieldName);

                if (supplied != null && !string.IsNullOrEmpty(supplied.Value))
                {
                    // Requirement 1, second request: re-execute the field with the entered value.
                    args.FieldValue = supplied.Value;
                    args.ClearField = false;
                }
                else
                {
                    AddMissing(outcome, caseIdForField, caseNumberForField, groupName, rowIndex, args.FieldName);
                    args.ClearField = true;
                }
            };

            doc.MailMerge.BeforeClearGroupField += (sender, args) =>
            {
                if (!args.HasMappedGroupInDataSource)
                {
                    // e.g. a case without parties: remove the empty region
                    args.ClearGroup = true;
                }
            };

            // ---- The actual merge: strategy chosen from the detected regions ----
            if (hasCaseRegion && hasPartiesRegion)
            {
                // Scenario B: nested relational merge (Requirement 5), unchanged.
                doc.MailMerge.ExecuteNestedGroup(BuildNestedDataSet(cases), BuildNestedCommands());
            }
            else if (hasPartiesRegion)
            {
                // Scenario A: standard repeating region (Pattern 1). No parent Case region.
                doc.MailMerge.ExecuteGroup(BuildPartyGroupTable(cases));
                // Case-level fields outside the region are filled per selection: when
                // several cases are selected in the standard scenario there is no
                // per-case repetition, so the first selected case supplies them.
                var first = cases.FirstOrDefault();
                doc.MailMerge.Execute(
                    new[] { "CaseNumber", "CaseCaption" },
                    new[] { first?.CaseNumber ?? "", first?.CaseCaption ?? "" });
            }
            else
            {
                // No regions at all: fill the plain case fields only.
                var first = cases.FirstOrDefault();
                doc.MailMerge.Execute(
                    new[] { "CaseNumber", "CaseCaption" },
                    new[] { first?.CaseNumber ?? "", first?.CaseCaption ?? "" });
            }

            if (outcome.MissingFields.Count > 0)
            {
                doc.Close();
                return outcome; // no document: the UI must prompt first
            }

            outcome.MergedSfdt = ToSfdt(doc);
            outcome.DocumentName = "Merged-Notice-of-Hearing.docx";
            doc.Close();
            return outcome;
        }

        /// <summary>
        /// Matches a user-supplied value for a missing field. A value tagged with
        /// CaseId 0 applies to every case; otherwise the case id must match, so one
        /// case's answer never leaks into another case.
        /// </summary>
        private static AdditionalMergeValue FindSuppliedValue(
            List<AdditionalMergeValue> additionalData,
            int caseId,
            string groupName,
            int rowIndex,
            string fieldName)
        {
            if (additionalData == null) return null;
            return additionalData.FirstOrDefault(x =>
                (x.CaseId == 0 || x.CaseId == caseId)
                && string.Equals(x.GroupName ?? "", groupName, StringComparison.OrdinalIgnoreCase)
                && x.RowIndex == rowIndex
                && string.Equals(x.FieldName, fieldName, StringComparison.OrdinalIgnoreCase));
        }

        private static void AddMissing(
            MergeOutcome outcome,
            int caseId,
            string caseNumber,
            string groupName,
            int rowIndex,
            string fieldName)
        {
            bool exists = outcome.MissingFields.Any(m =>
                m.CaseId == caseId
                && string.Equals(m.GroupName, groupName, StringComparison.OrdinalIgnoreCase)
                && m.RowIndex == rowIndex
                && string.Equals(m.FieldName, fieldName, StringComparison.OrdinalIgnoreCase));
            if (exists) return;

            outcome.MissingFields.Add(new MissingFieldInfo
            {
                CaseId = caseId,
                CaseNumber = caseNumber,
                GroupName = groupName,
                RowIndex = rowIndex,
                FieldName = fieldName
            });
        }

        /// <summary>
        /// Serializes the sample template as SFDT (GET /api/template response body).
        /// Kept for API symmetry; the controller exposes BuildSampleTemplateSfdt().
        /// </summary>
        public static string SerializeTemplate(WordDocument doc)
        {
            return ToSfdt(doc);
        }
    }
}