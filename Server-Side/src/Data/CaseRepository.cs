using System;
using System.Collections.Generic;
using System.Linq;
using EJ2APIServices.Models;

namespace EJ2APIServices.Data
{
    /// <summary>
    /// In-memory sample data set for the POC. No database is used.
    /// Address is intentionally null/empty on some parties so the
    /// missing-field prompt flow (Requirement 1) can be demonstrated.
    /// </summary>
    public static class CaseRepository
    {
        private static readonly List<CaseData> _cases = new List<CaseData>
        {
            new CaseData
            {
                CaseNumber = "2026-CV-001234",
                CaseCaption = "Smith v. Johnson",
                Parties = new List<PartyData>
                {
                    new PartyData { FirstName = "John", LastName = "Smith", Address = "New York" },
                    new PartyData { FirstName = "Emily", LastName = "Johnson", Address = null }
                }
            },
            new CaseData
            {
                CaseNumber = "2026-CV-001237",
                CaseCaption = "Smith v. Anderson",
                Parties = new List<PartyData>
                {
                    new PartyData { FirstName = "Michael", LastName = "Smith", Address = "Chicago" },
                    new PartyData { FirstName = "Sarah", LastName = "Anderson", Address = null },
                    new PartyData { FirstName = "David", LastName = "Lee", Address = "Seattle" }
                }
            },
            new CaseData
            {
                CaseNumber = "2026-CV-001238",
                CaseCaption = "Martin v. Smith",
                Parties = new List<PartyData>
                {
                    new PartyData { FirstName = "Anna", LastName = "Martin", Address = "Denver" }
                }
            }
        };

        // CaseId is stable for the sample run: index + 1 style ids (101, 102, 103).
        private static int IdFor(CaseData c) => 100 + _cases.IndexOf(c) + 1;

        /// <summary>Search by keyword over case number and caption. Returns only matching summaries, never parties.</summary>
        public static List<CaseInfo> Search(string keyword)
        {
            if (string.IsNullOrWhiteSpace(keyword))
            {
                return _cases.Select(ToInfo).ToList();
            }

            return _cases
                .Where(c => (c.CaseNumber ?? "").IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0
                         || (c.CaseCaption ?? "").IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0)
                .Select(ToInfo)
                .ToList();
        }

        /// <summary>Load one full case (with parties) by id. Called only for cases the user selected.</summary>
        public static CaseData GetCase(int caseId)
        {
            var c = _cases.FirstOrDefault(x => IdFor(x) == caseId);
            if (c == null) return null;
            return new CaseData
            {
                CaseId = caseId,
                CaseNumber = c.CaseNumber,
                CaseCaption = c.CaseCaption,
                Parties = c.Parties.Select(p => new PartyData
                {
                    FirstName = p.FirstName,
                    LastName = p.LastName,
                    Address = p.Address
                }).ToList()
            };
        }

        private static CaseInfo ToInfo(CaseData c) => new CaseInfo
        {
            CaseId = IdFor(c),
            CaseNumber = c.CaseNumber,
            CaseCaption = c.CaseCaption
        };
    }
}